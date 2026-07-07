# Phase 6 test: real end-to-end mutation of a live Word document through the
# same chat pipeline the taskpane UI uses - `opencode run --attach` against a
# dedicated `opencode serve` instance, using the "build" agent which has
# access to the word-mcp-live MCP tools declared in opencode.json. This is
# intentionally not mocked: it drives real word COM automation (word-mcp-live
# attaches to the active Word instance via GetActiveObject, so Word must
# already be open with a document before the prompt runs) and inspects the
# resulting document via COM to prove the mutation actually happened, not
# just that the model claimed to do it.
#
# We attach to a pre-started, health-checked `opencode serve` rather than
# letting `opencode run` spin up its own ephemeral per-call instance: a first
# version of this test that used a bare `opencode run` (own embedded server)
# intermittently got a reply claiming "I don't have access to Word
# automation tools" when run back-to-back with other phases - most likely a
# race where the freshly-spawned local MCP subprocess (uvx word_mcp_server)
# hadn't finished connecting before the first agent turn. Attaching to an
# already-running, already-health-checked server (same pattern proven
# reliable by phase3/phase4) also matches how the real task pane talks to
# OpenCode - it always uses a long-lived `opencode serve`, never the
# ephemeral `opencode run` mode.

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$failures = @()
$opencodePort = 4099

function Release-Com($obj) {
    if ($obj) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null
    }
}

Push-Location $root
try {
    $opencodeCmd = Get-Command opencode -ErrorAction SilentlyContinue
    if (-not $opencodeCmd) {
        $failures += "'opencode' CLI not found on PATH"
    } else {
        $serveJob = Start-Job -ScriptBlock {
            param($root, $port)
            Set-Location $root
            opencode serve --port $port --hostname 127.0.0.1
        } -ArgumentList $root, $opencodePort

        $word = $null
        $doc = $null
        try {
            $serverReady = $false
            for ($i = 0; $i -lt 20; $i++) {
                Start-Sleep -Seconds 1
                try {
                    $resp = Invoke-WebRequest -Uri "http://127.0.0.1:$opencodePort/global/health" -TimeoutSec 3 -UseBasicParsing
                    if ($resp.StatusCode -eq 200) {
                        $body = $resp.Content | ConvertFrom-Json
                        if ($body.healthy -eq $true) { $serverReady = $true; break }
                    }
                } catch { }
            }
            # Give the word-mcp-live MCP subprocess (spawned by opencode.json) a moment to
            # finish connecting after the server reports healthy, before the first prompt.
            if ($serverReady) { Start-Sleep -Seconds 3 }

            if (-not $serverReady) {
                $failures += "http://127.0.0.1:$opencodePort/global/health did not report healthy within 20s"
            } else {
                $word = New-Object -ComObject Word.Application
                $word.Visible = $true
                $doc = $word.Documents.Add()

                $prompt = "In the currently open Word document, add a Heading 1 saying 'Phase 6 Test' " +
                    "at the very top, then insert a 2x2 table with A and B in the first row and " +
                    "C and D in the second row. Do this now using the available Word tools."

                # Run via a job with a hard timeout: a first attempt at this test hung
                # indefinitely (opencode.exe sat at ~0% CPU with no network activity for 18+
                # minutes), most likely because `opencode run` blocked on an interactive
                # tool-call approval prompt that a non-interactive background shell can never
                # answer. The Wait-Job timeout below guarantees the script fails fast instead
                # of hanging forever if that (or anything else) happens again.
                $runJob = Start-Job -ScriptBlock {
                    param($root, $prompt, $port)
                    Set-Location $root
                    opencode run $prompt --agent build -m sunshine-coder/ssc-chat-latest --attach "http://127.0.0.1:$port"
                } -ArgumentList $root, $prompt, $opencodePort

                $jobDone = Wait-Job $runJob -Timeout 150
                if (-not $jobDone) {
                    Stop-Job $runJob -ErrorAction SilentlyContinue | Out-Null
                    $failures += "opencode run did not finish within 150s (likely stuck waiting for input or a stalled tool call)"
                } else {
                    $runOutput = Receive-Job $runJob -ErrorAction SilentlyContinue | Out-String
                    Write-Host $runOutput
                    if ($runJob.State -eq "Failed") {
                        $failures += "opencode run job failed:`n$runOutput"
                    }
                }
                Remove-Job $runJob -Force -ErrorAction SilentlyContinue | Out-Null

                # Re-fetch the heading style's localized name so the check works regardless
                # of Word's UI language (e.g. "Heading 1" vs a localized equivalent).
                $heading1Name = $doc.Styles.Item(-2).NameLocal

                $headingFound = $false
                foreach ($para in $doc.Paragraphs) {
                    $text = $para.Range.Text.Trim()
                    $styleName = $para.Range.Style.NameLocal
                    if ($text -eq "Phase 6 Test" -and $styleName -eq $heading1Name) {
                        $headingFound = $true
                        break
                    }
                }
                if (-not $headingFound) {
                    $failures += "Heading 1 paragraph with text 'Phase 6 Test' not found in document"
                }

                if ($doc.Tables.Count -lt 1) {
                    $failures += "No table found in document"
                } else {
                    $table = $doc.Tables.Item(1)
                    if ($table.Rows.Count -ne 2 -or $table.Columns.Count -ne 2) {
                        $failures += "Table dimensions wrong: expected 2x2, got $($table.Rows.Count)x$($table.Columns.Count)"
                    } else {
                        $expected = @("A", "B", "C", "D")
                        $actual = @()
                        for ($r = 1; $r -le 2; $r++) {
                            for ($c = 1; $c -le 2; $c++) {
                                # Cell text ends with a cell-mark (chr 13 + chr 7); strip those.
                                $actual += $table.Cell($r, $c).Range.Text.TrimEnd([char]13, [char]7).Trim()
                            }
                        }
                        for ($i = 0; $i -lt 4; $i++) {
                            if ($actual[$i] -ne $expected[$i]) {
                                $failures += "Table cell $i mismatch: expected '$($expected[$i])', got '$($actual[$i])'"
                            }
                        }
                    }
                }
            }
        } finally {
            if ($doc) {
                $doc.Close([ref]$false) | Out-Null
            }
            if ($word) {
                $word.Quit() | Out-Null
            }
            Release-Com $doc
            Release-Com $word
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()

            Stop-Job $serveJob -ErrorAction SilentlyContinue | Out-Null
            Remove-Job $serveJob -Force -ErrorAction SilentlyContinue | Out-Null
            Get-NetTCPConnection -LocalPort $opencodePort -ErrorAction SilentlyContinue |
                ForEach-Object { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue }
        }
    }
} finally {
    Pop-Location
}

if ($failures.Count -eq 0) {
    Write-Host "PASS: Phase 6 Word mutation via chat (heading + table)" -ForegroundColor Green
    exit 0
} else {
    Write-Host "FAIL: Phase 6 Word mutation via chat" -ForegroundColor Red
    $failures | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
    exit 1
}
