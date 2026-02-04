è<#
.SYNOPSIS
    Clean up duplicate components created by import bugs (e.g. ThisWorkbook1, SheetXXX1)
.DESCRIPTION
    Scans the open Excel workbook for VBComponents that:
    1. Are of Type 2 (Class Module) - implying they are not the real Document object.
    2. Have names starting with 'Sheet' or 'ThisWorkbook'.
    3. Are NOT listed in the manifest.json.
    
    Prompts for deletion or auto-deletes.
#>
param(
    [string]$WorkbookPath = "$PSScriptRoot\..\index.xlsm",
    [switch]$Force = $false
)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$vbaRoot = Join-Path $scriptDir '..\vba-files'
$manifestPath = Join-Path $vbaRoot 'manifest.json'

Write-Host "Reading manifest..." -ForegroundColor Cyan
if (-not (Test-Path $manifestPath)) { Write-Error "Manifest not found."; exit 1 }
$manifestJson = Get-Content $manifestPath -Raw | ConvertFrom-Json
$validNames = @{}

# Harvest valid names from manifest (assuming filename base matches component name)
$manifestJson.imports | ForEach-Object {
    $fname = Split-Path $_ -Leaf
    $baseDetail = [System.IO.Path]::GetFileNameWithoutExtension($fname)
    $validNames[$baseDetail] = $true
}

Write-Host "Found $($validNames.Count) valid component names in manifest." -ForegroundColor Gray

# Connect to Excel (Launch new instance)
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3 # Disable Macros
    Write-Host "Opening workbook '$WorkbookPath'..." -ForegroundColor Cyan
    $wb = $excel.Workbooks.Open($WorkbookPath)
}
catch {
    Write-Error "Failed to launch Excel or open workbook: $_"
    exit 1
}

$vbProj = $wb.VBProject
Write-Host "Scanning $($vbProj.VBComponents.Count) components in '$($wb.Name)'..." -ForegroundColor Cyan

$toDelete = @()

foreach ($comp in $vbProj.VBComponents) {
    # Filter for potential duplicates
    # Type 2 = Class Module. Type 100 = Document.
    # The duplicates we want to kill are Type 2 masquerading as Document objects.
    if ($comp.Type -eq 2) { 
        if ($comp.Name -match "^(Sheet\d+|ThisWorkbook)\d+$") {
            # It looks like a duplicate (starts with Sheet/ThisWorkbook, ends with extra digits)
            # AND it is NOT in the valid list.
            if (-not $validNames.ContainsKey($comp.Name)) {
                $toDelete += $comp
            }
        }
        # Also catch literal "ThisWorkbook1" if regex above didn't catch it
        if ($comp.Name -eq "ThisWorkbook1" -and -not $validNames.ContainsKey("ThisWorkbook1")) {
            if ($toDelete.Name -notcontains "ThisWorkbook1") { $toDelete += $comp }
        }
    }
}

Write-Host "Found $($toDelete.Count) candidate(s) for deletion:" -ForegroundColor Yellow
$toDelete | ForEach-Object { Write-Host " - $($_.Name) (Type: $($_.Type))" }

if ($toDelete.Count -eq 0) {
    Write-Host "No duplicates found." -ForegroundColor Green
    exit 0
}

if (-not $Force) {
    $confirm = Read-Host "Delete these components? (y/n)"
    if ($confirm -ne 'y') { Write-Host "Aborted."; exit 0 }
}

foreach ($comp in $toDelete) {
    Write-Host "Deleting $($comp.Name)..." -NoNewline
    try {
        $vbProj.VBComponents.Remove($comp)
        Write-Host "OK" -ForegroundColor Green
    }
    catch {
        Write-Host "FAILED: $_" -ForegroundColor Red
    }
}

Write-Host "Saving workbook..." -ForegroundColor Cyan
$wb.Save()
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host "Cleanup complete." -ForegroundColor Cyan
è"(9e377b08b88e30057baf56c074b3d12f1d22372328file:///c:/VBA/SGQ%201.65/scripts/cleanup-duplicates.ps1:file:///c:/VBA/SGQ%201.65