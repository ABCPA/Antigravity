¿<#
.SYNOPSIS
    Static Analysis for VBA: Checks for missing method references.
.DESCRIPTION
    1. Indexes all Public Sub/Function/Property in vba-files.
    2. Scans files for "Module.Method" calls.
    3. Reports references to Methods that do not exist (Member not found).
#>
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$vbaRoot = Join-Path $scriptDir '..\vba-files'

$definitions = @{} # Key: "Module.Method", Value: True
$modules = @{}     # Key: "Module", Value: True

Write-Host "Indexing definitions..." -ForegroundColor Cyan

# 1. Index Definitions
$files = Get-ChildItem -Path $vbaRoot -Include *.bas, *.cls -Recurse
foreach ($file in $files) {
    if ($file.Name -match '\.bak$') { continue }
    
    $modName = $file.BaseName
    $modules[$modName] = $true
    
    $content = Get-Content $file.FullName
    foreach ($line in $content) {
        # Match Public Sub/Function/Property
        # Ignore comments
        if ($line -match "^\s*('|REM)") { continue }
        
        if ($line -match '^\s*(Public|Private|Friend)?\s*(Sub|Function|Property\s+Get|Property\s+Let|Property\s+Set)\s+([a-zA-Z0-9_]+)') {
            $methodName = $matches[3]
            $key = "$modName.$methodName"
            $definitions[$key] = $true
            # Write-Host "Defined: $key" -ForegroundColor Gray
        }
    }
}

Write-Host "Indexed $($definitions.Count) methods in $($modules.Count) modules." -ForegroundColor Green

# 2. Scan for Usages
$errors = @()

foreach ($file in $files) {
    $content = Get-Content $file.FullName
    $lineNum = 0
    foreach ($line in $content) {
        $lineNum++
        if ($line -match "^\s*('|REM)") { continue }
        
        # Check for Module.Method calls
        # Regex to capture "ModName.MethodName"
        # We ignore "Me.", "Debug.", "Err.", "Excel.", "VBA.", "Scripting."
        # And common objects like "Worksheets.", "Cells."
        
        # Matches: Word.Word
        $matchesList = [regex]::Matches($line, '\b([a-zA-Z][a-zA-Z0-9_]*)\.([a-zA-Z][a-zA-Z0-9_]*)')
        foreach ($m in $matchesList) {
            $modRef = $m.Groups[1].Value
            $methRef = $m.Groups[2].Value
            
            # Skip built-ins
            if ($modRef -match '^(Me|Debug|Err|VBA|Excel|Office|Scripting|System|Application|ActiveWorkbook|ActiveSheet|ThisWorkbook|Worksheets|Sheets|Range|Cells|Rows|Columns)$') { continue }
            if ($modRef -in @("msg", "fso", "ts", "wb", "ws", "sh", "target", "dict", "coll", "json")) { continue } # Variables
            
            # Heuristic: If ModRef is one of OUR modules
            if ($modules.ContainsKey($modRef)) {
                $checkKey = "$modRef.$methRef"
                if (-not $definitions.ContainsKey($checkKey)) {
                    # Special Case: Enums or Constants? (Not indexed by regex above)
                    # For now, just report as potential error
                    $errors += "[MISSING] $checkKey referenced in $($file.Name):$lineNum"
                }
            }
        }
    }
}

if ($errors.Count -gt 0) {
    Write-Host "Found $($errors.Count) potential errors:" -ForegroundColor Red
    $errors | ForEach-Object { Write-Host $_ }
    exit 1
}
else {
    Write-Host "No missing references found." -ForegroundColor Green
    exit 0
}
¿"(9e377b08b88e30057baf56c074b3d12f1d22372325file:///c:/VBA/SGQ%201.65/scripts/static-analysis.ps1:file:///c:/VBA/SGQ%201.65