�	<#
.SYNOPSIS
    Runs the VBA Unit Test Suite via COM.
.DESCRIPTION
    Opens the target workbook (default: index.xlsm) and executes modUnitTestEngine.RunAllTests.
#>
param(
    [string]$WorkbookPath = ".\index.xlsm"
)

$ErrorActionPreference = "Stop"
$path = Convert-Path $WorkbookPath

Write-Host "🚀 Launching Unit Tests in '$path'..." -ForegroundColor Cyan

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Open($path)
    
    # Run the test engine
    Write-Host "⏳ Running tests..." -ForegroundColor Yellow
    $excel.Run("modUnitTestEngine.RunAllTests")
    
    # Note: modUnitTestEngine.RunAllTests displays a MsgBox at the end.
    # In a CI environment, we would want to write to a log file or cell instead of MsgBox.
    # For local dev, MsgBox is fine.
    
    Write-Host "✅ Tests execution triggered." -ForegroundColor Green
}
catch {
    Write-Host "❌ Error running tests: $_" -ForegroundColor Red
}
finally {
    # Keep Excel open for review
    # $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
�	*cascade08"(b3eea256776ced57c23cbf87c776d2461e607ebe2Ifile:///C:/Users/AbelBoudreau/Workspace_CPA_AI/scripts/run-unit-tests.ps1:.file:///C:/Users/AbelBoudreau/Workspace_CPA_AI