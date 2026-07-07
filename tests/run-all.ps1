# Regression runner: executes every phaseN-*.ps1 test script in order and
# fails fast on the first regression, so later phases never ship on top of a
# broken earlier one.

$ErrorActionPreference = "Stop"
$testsDir = $PSScriptRoot
$scripts = Get-ChildItem -Path $testsDir -Filter "phase*-*.ps1" | Sort-Object Name

$overallFailed = $false
foreach ($script in $scripts) {
    Write-Host "==> Running $($script.Name)" -ForegroundColor Cyan
    & $script.FullName
    if ($LASTEXITCODE -ne 0) {
        Write-Host "==> $($script.Name) FAILED" -ForegroundColor Red
        $overallFailed = $true
        break
    }
}

if ($overallFailed) {
    Write-Host "`nRegression check FAILED." -ForegroundColor Red
    exit 1
} else {
    Write-Host "`nAll phase tests passed." -ForegroundColor Green
    exit 0
}
