¾# AUDIT : crÃ©ation 20251106_1426 -> temporary monitor script created by assistant
param(
    [string]$WorkbookPath = 'index_import_test_20251028.xlsm',
    [int]$Attempts = 12,
    [int]$IntervalSeconds = 300
)

$ErrorActionPreference = 'Stop'
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = 'C:\VBA\SGQ 1.65'
$testScript = Join-Path $repoRoot 'vba-files\scripts\test-vbide-and-compile.ps1'
$log = Join-Path $repoRoot ('logs\compile-monitor-' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.log')
if (-not (Test-Path -LiteralPath (Split-Path -Parent $log))) { New-Item -ItemType Directory -Path (Split-Path -Parent $log) -Force | Out-Null }
Write-Host "Monitor starting; logging to: $log"

for ($i = 1; $i -le $Attempts; $i++) {
    $ts = Get-Date -Format 's'
    "=== Attempt $i - $ts ===" | Tee-Object -FilePath $log -Append
    try {
        pwsh -NoProfile -ExecutionPolicy Bypass -File $testScript -AttemptCompile -WorkbookPath $WorkbookPath 2>&1 | Tee-Object -FilePath $log -Append
    }
    catch {
        "Test script invocation failed: $($_.Exception.Message)" | Tee-Object -FilePath $log -Append
    }
    $rc = $LASTEXITCODE
    "ExitCode: $rc" | Tee-Object -FilePath $log -Append
    if ($rc -eq 0) {
        'Compile clean - stopping monitor.' | Tee-Object -FilePath $log -Append
        break
    }
    if ($i -lt $Attempts) {
        "Sleeping $IntervalSeconds seconds before next attempt..." | Tee-Object -FilePath $log -Append
        Start-Sleep -Seconds $IntervalSeconds
    }
}
Write-Host "Monitor finished. Log: $log"
¾*cascade08"(9e377b08b88e30057baf56c074b3d12f1d2237232?file:///c:/VBA/SGQ%201.65/scripts/monitor-compile-automated.ps1:file:///c:/VBA/SGQ%201.65