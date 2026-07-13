# Phase 4 test: verify opencode.json correctly registers word-mcp-live as a local
# MCP server and OpenCode can actually connect to it (spawn it over stdio and
# complete the MCP handshake). Relies on `uvx` having word-mcp-live already
# cached locally (first-time `uvx --from word-mcp-live word_mcp_server` resolves
# and installs ~80 packages, which can take a while - run it once manually
# before this test if it's never been run on this machine).

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$failures = @()

Push-Location $root
try {
    $configPath = Join-Path $root "opencode.json"
    if (-not (Test-Path $configPath)) {
        $failures += "opencode.json not found"
    } else {
        $config = Get-Content $configPath -Raw | ConvertFrom-Json
        $wordMcp = $config.mcp.word
        if (-not $wordMcp) {
            $failures += "opencode.json does not declare an mcp.word entry"
        } else {
            if ($wordMcp.type -ne "local") {
                $failures += "mcp.word.type is '$($wordMcp.type)', expected 'local'"
            }
            $expectedCommand1 = @("uvx", "--from", "word-mcp-live", "word_mcp_server")
            $expectedCommand2 = @("uv", "run", "--with", "word-mcp-live", "scripts/word_mcp_launcher.py")
            $actualCommand = @($wordMcp.command)
            $actualStr = $actualCommand -join " "
            if ($actualStr -ne ($expectedCommand1 -join " ") -and $actualStr -ne ($expectedCommand2 -join " ")) {
                $failures += "mcp.word.command is '$actualStr', expected '$($expectedCommand2 -join ' ')' (to support multiple Word desktop instances)"
            }
        }
    }

    $opencodeCmd = Get-Command opencode -ErrorAction SilentlyContinue
    if (-not $opencodeCmd) {
        $failures += "'opencode' CLI not found on PATH"
    } elseif ($failures.Count -eq 0) {
        # `opencode mcp list` spawns every configured MCP server and reports connection
        # status. Give it a generous timeout since the first cold start can be slow.
        # Note: do not redirect stderr (2>&1) on this native exe under Windows PowerShell 5.1 -
        # it wraps stderr lines as terminating NativeCommandError records even on exit code 0.
        $mcpOutput = & opencode mcp list | Out-String
        if ($mcpOutput -notmatch "word" -or $mcpOutput -notmatch "connected") {
            $failures += "opencode mcp list did not report 'word' as connected:`n$mcpOutput"
        }
    }
} finally {
    Pop-Location
}

if ($failures.Count -eq 0) {
    Write-Host "PASS: Phase 4 word-mcp-live checks (opencode.json config, opencode mcp list connected)" -ForegroundColor Green
    exit 0
} else {
    Write-Host "FAIL: Phase 4 word-mcp-live checks" -ForegroundColor Red
    $failures | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
    exit 1
}
