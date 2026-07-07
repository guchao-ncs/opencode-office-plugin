# Phase 1 test: verify the Office Add-in project was scaffolded correctly.
# Pure structural checks only (no network/npm calls) so this stays fast and safe to re-run.

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$failures = @()

function Check-FileExists($relPath) {
    $full = Join-Path $root $relPath
    if (-not (Test-Path $full)) {
        $script:failures += "Missing file: $relPath"
    }
}

# Required generated files
@(
    "manifest.xml",
    "package.json",
    "webpack.config.js",
    "babel.config.json",
    "src\taskpane\taskpane.html",
    "src\taskpane\taskpane.js",
    "src\taskpane\taskpane.css",
    "assets\icon-16.png",
    "assets\icon-32.png",
    "assets\icon-80.png"
) | ForEach-Object { Check-FileExists $_ }

# manifest.xml must be well-formed XML, point at the right dev server URL,
# and no longer carry the generator's placeholder GUID/branding.
$manifestPath = Join-Path $root "manifest.xml"
if (Test-Path $manifestPath) {
    try {
        [xml]$manifest = Get-Content $manifestPath -Raw
    } catch {
        $failures += "manifest.xml is not well-formed XML: $($_.Exception.Message)"
    }
    if ($manifest) {
        $sourceLocation = $manifest.OfficeApp.DefaultSettings.SourceLocation.DefaultValue
        if ($sourceLocation -ne "https://localhost:3000/taskpane.html") {
            $failures += "manifest.xml SourceLocation is '$sourceLocation', expected https://localhost:3000/taskpane.html"
        }
        if ($manifest.OfficeApp.Id -eq "05c2e1c9-3e1d-406e-9a91-e9ac64854143") {
            $failures += "manifest.xml still has the generator's placeholder GUID"
        }
        if ($manifest.OfficeApp.ProviderName -eq "Contoso") {
            $failures += "manifest.xml still has placeholder ProviderName 'Contoso'"
        }
        $hostNames = $manifest.OfficeApp.Hosts.Host | ForEach-Object { $_.Name }
        if ($hostNames -notcontains "Document") {
            $failures += "manifest.xml does not declare Host Name='Document' (Word)"
        }
    }
}

# package.json must declare the scripts the rest of the plan depends on.
$pkgPath = Join-Path $root "package.json"
if (Test-Path $pkgPath) {
    $pkg = Get-Content $pkgPath -Raw | ConvertFrom-Json
    foreach ($script in @("start", "stop", "validate", "dev-server", "build")) {
        if (-not ($pkg.scripts.PSObject.Properties.Name -contains $script)) {
            $failures += "package.json is missing npm script '$script'"
        }
    }
    foreach ($dep in @("office-addin-debugging", "office-addin-dev-certs", "office-addin-manifest", "webpack-dev-server")) {
        if (-not ($pkg.devDependencies.PSObject.Properties.Name -contains $dep)) {
            $failures += "package.json devDependencies is missing '$dep'"
        }
    }
}

if ($failures.Count -eq 0) {
    Write-Host "PASS: Phase 1 scaffold checks (manifest.xml, package.json, src/taskpane, assets)" -ForegroundColor Green
    exit 0
} else {
    Write-Host "FAIL: Phase 1 scaffold checks" -ForegroundColor Red
    $failures | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
    exit 1
}
