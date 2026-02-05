ö¨<#
.SYNOPSIS
  Provides a single, robust entry point for all VBA-related CI/CD actions.

.DESCRIPTION
  This guardian script manages the entire lifecycle of an Excel interaction, including
  process management, pre-flight checks, action execution (Import, Export, Compile),
  and robust saving. It combines the best features of several existing utility scripts
  to provide a stable and reliable automation tool.

.PARAMETER Action
  The action to perform. Valid values: Import, Export, Compile, RunMacro.

.PARAMETER MacroName
  The name of the macro to run when -Action is 'RunMacro'.

.PARAMETER WorkbookPath
  The absolute path to the target Excel workbook. Defaults to 'index.xlsm' in the repo root.

.PARAMETER TimeoutSeconds
  The timeout in seconds for the watchdog to terminate a hung process. Defaults to 300.

.EXAMPLE
  # Robustly import all modules
  pwsh -File .\Start-VbaAction.ps1 -Action Import

.EXAMPLE
  # Run a compile check
  pwsh -File .\Start-VbaAction.ps1 -Action Compile

.EXAMPLE
  # Export all modules with a 10-minute timeout
  pwsh -File .\Start-VbaAction.ps1 -Action Export -TimeoutSeconds 600
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Import', 'Export', 'Compile', 'RunMacro')]
    [string]$Action,

    [string]$MacroName,

    [string]$WorkbookPath, # Default value assigned below

    [int]$TimeoutSeconds = 300
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Setup and Path Resolution
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = (Resolve-Path (Join-Path $scriptRoot '..')).ProviderPath

if (-not $WorkbookPath) {
    $WorkbookPath = (Resolve-Path (Join-Path $repoRoot 'index.xlsm')).ProviderPath
}

function Write-Info($msg) { Write-Host "[INFO] $msg" -ForegroundColor Green }
function Write-Warn($msg) { Write-Warning $msg }
function Write-Err ($msg) { Write-Error $msg }
# AUDIT : sauvegarde 20251107_095047 -> backups/20251107_095047/Start-VbaAction.ps1.bak
# Sanitize source when falling back to AddFromString: remove 'Attribute VB_*' lines
function Sanitize-VbaSource([string]$text) { (($text -split "`r?`n") | Where-Object { $_ -notmatch '^(?i)\s*(Attribute\s+VB_|VERSION\s+\d|BEGIN|MultiUse\s+|End$)' }) -join "`r`n" }
#endregion

#region Watchdog and Process Management (from export-modules-robust.ps1)
$startTime = Get-Date
$sentinel = Join-Path $env:TEMP ("vba_action_" + $startTime.ToString('yyyyMMdd_HHmmss') + ".sentinel")
New-Item -Path $sentinel -ItemType File -Force | Out-Null

Write-Info "Starting watchdog with a ${TimeoutSeconds}s timeout."
$watchJob = Start-Job -ScriptBlock {
    param($sentinelPath, $maxSeconds, $originTime)
    Start-Sleep -Seconds $maxSeconds
    if (Test-Path $sentinelPath) {
        try {
            $logDir = Join-Path (Split-Path $sentinelPath -Parent) "VBA_ACTION_LOGS"
            New-Item -ItemType Directory -Path $logDir -Force -ErrorAction SilentlyContinue | Out-Null
            $logFile = Join-Path $logDir 'watchdog.log'
            
            $candidates = Get-CimInstance Win32_Process -Filter "Name = 'EXCEL.EXE'" -ErrorAction SilentlyContinue | Select-Object ProcessId, CreationDate
            foreach ($p in $candidates) {
                try {
                    $createTime = [Management.ManagementDateTimeConverter]::ToDateTime($p.CreationDate)
                    if ($createTime -ge $originTime) {
                        Stop-Process -Id $p.ProcessId -Force -ErrorAction SilentlyContinue
                        $logLine = "$(Get-Date -Format s) - Watchdog stopped hung EXCEL process Id=$($p.ProcessId) Created=$createTime"
                        Add-Content -Path $logFile -Value $logLine
                    }
                }
                catch {}
            }
        }
        catch {}
    }
} -ArgumentList $sentinel, $TimeoutSeconds, $startTime

#endregion

$excel = $null
$wb = $null
$vbProj = $null

try {
    #region Pre-flight Checks
    
    # 1. Clean up existing Excel processes (from export-modules-robust.ps1)
    Write-Info "Checking for and closing existing Excel instances... (DISABLED TO AVOID HANG)"
    # try {
    #     $candidates = Get-CimInstance Win32_Process -Filter "Name = 'EXCEL.EXE'" -ErrorAction SilentlyContinue | Select-Object ProcessId, CreationDate
    #     foreach ($p in $candidates) {
    #         # SKIPPED
    #     }
    #     }
    # } catch {
    #     Write-Warning "Auto-close failed: $_"
    # }
    # $proc = Get-Process -Id $p.ProcessId -ErrorAction SilentlyContinue
    # if ($proc.MainWindowTitle -like "*${WorkbookPath}*") {
    #    Stop-Process -Id $p.ProcessId -Force
    #    Write-Info "Closed pre-existing Excel process $($p.ProcessId) holding the workbook."
    # }
    #}
    #} catch {}
    #}
    #} catch {
    #    Write-Warn "Could not fully check for pre-existing excel processes. Manual check may be required if script fails."
    #}


    # 2. Check for lock file (from import-modules-robust.ps1)
    $wbDir = Split-Path -Parent $WorkbookPath
    $wbName = Split-Path -Leaf $WorkbookPath
    $lockFile = Join-Path $wbDir ("~$" + $wbName)
    if (Test-Path -LiteralPath $lockFile) {
        Write-Info "Lock file found. Attempting to remove."
        try { Remove-Item -LiteralPath $lockFile -Force -ErrorAction Stop } catch {
            Write-Err "Workbook appears to be open and locked: $lockFile. Aborting."
            exit 10
        }
    }

    # 3. Remove ReadOnly attribute (from import-modules-robust.ps1)
    try {
        $fi = Get-Item -LiteralPath $WorkbookPath
        if ($fi.Attributes -band [IO.FileAttributes]::ReadOnly) {
            attrib -R -- $WorkbookPath | Out-Null
            Write-Info "Removed ReadOnly attribute from workbook."
        }
    }
    catch { Write-Warn "Could not inspect or change workbook attributes: $_" }

    #endregion

    #region Excel COM Initialization
    Write-Info "Creating new Excel COM instance."
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try { $excel.AutomationSecurity = 3 } catch {} # msoAutomationSecurityForceDisable

    # 4. VBE Access Check (from test-vbide-and-compile.ps1)
    try {
        $null = $excel.VBE
        Write-Info "VBE access: OK (Trust access appears enabled)."
    }
    catch {
        Write-Err "VBE access is not enabled. Please enable 'Trust access to the VBA project object model' in Excel's Trust Center."
        exit 11
    }
    
    Write-Info "Opening workbook: $WorkbookPath"
    $wb = $excel.Workbooks.Open($WorkbookPath)
    try { $excel.EnableEvents = $false } catch {}
    try { $excel.ScreenUpdating = $false } catch {}
    $vbProj = $wb.VBProject
    #endregion

    #region Main Action Execution
    switch ($Action) {
        'Import' {
            Write-Info "Executing Action: Import"
            $vbaRoot = (Resolve-Path (Join-Path $repoRoot 'vba-files')).ProviderPath
            
            # Import Standard Modules (.bas)
            $moduleDir = Join-Path $vbaRoot 'Module'
            Get-ChildItem -Path $moduleDir -Filter '*.bas' -File | ForEach-Object {
                $targetName = $_.BaseName
                try { $existing = $vbProj.VBComponents.Item($targetName); $vbProj.VBComponents.Remove($existing) } catch {}
                try {
                    $vbProj.VBComponents.Import($_.FullName) | Out-Null
                    Write-Info "Imported module: $targetName"
                }
                catch {
                    Write-Warn "Import() failed for module '$targetName'. Trying fallback AddFromString..."
                    try {
                        $text = Sanitize-VbaSource ([System.IO.File]::ReadAllText($_.FullName, [System.Text.Encoding]::GetEncoding(1252)))
                        $comp = $vbProj.VBComponents.Add(1) # 1 = vbext_ct_StdModule
                        $comp.Name = $targetName
                        $comp.CodeModule.AddFromString($text)
                        Write-Info "Successfully imported module '$targetName' via AddFromString fallback."
                    }
                    catch {
                        Write-Err "Fallback AddFromString also failed for module '$targetName': $($_.Exception.Message)"
                        throw
                    }
                }
            }

            # Import Class Modules (.cls)
            $classDir = Join-Path $vbaRoot 'Class'
            Get-ChildItem -Path $classDir -Filter '*.cls' -File | ForEach-Object {
                $targetName = $_.BaseName
                $existing = $null
                try { $existing = $vbProj.VBComponents.Item($targetName) } catch {}

                if ($existing -ne $null -and $existing.Type -eq 100) {
                    # Document Object (ThisWorkbook, Sheets) - Update Code Only
                    Write-Info "Updating Document Object: $targetName"
                    try {
                        $text = Sanitize-VbaSource ([System.IO.File]::ReadAllText($_.FullName, [System.Text.Encoding]::GetEncoding(1252)))
                        if ($existing.CodeModule.CountOfLines -gt 0) {
                            $existing.CodeModule.DeleteLines(1, $existing.CodeModule.CountOfLines)
                        }
                        $existing.CodeModule.AddFromString($text)
                        Write-Info "Successfully updated Document Object '$targetName'."
                    }
                    catch {
                        Write-Err "Failed to update Document Object '$targetName': $($_.Exception.Message)"
                    }
                }
                else {
                    # Standard Class Import (Remove, Import, Fallback)
                    if ($existing) { try { $vbProj.VBComponents.Remove($existing) } catch {} }
                    try {
                        $vbProj.VBComponents.Import($_.FullName) | Out-Null
                        Write-Info "Imported class: $targetName"
                    }
                    catch {
                        Write-Warn "Import() failed for class '$targetName'. Trying fallback AddFromString..."
                        try {
                            $text = Sanitize-VbaSource ([System.IO.File]::ReadAllText($_.FullName, [System.Text.Encoding]::GetEncoding(1252)))
                            $comp = $vbProj.VBComponents.Add(2) # 2 = vbext_ct_ClassModule
                            $comp.Name = $targetName
                            $comp.CodeModule.AddFromString($text)
                            Write-Info "Successfully imported class '$targetName' via AddFromString fallback."
                        }
                        catch {
                            Write-Err "Fallback AddFromString also failed for class '$targetName': $($_.Exception.Message)"
                            throw
                        }
                    }
                }
            }
            
            # Audit log for successful import
            Write-Info "Writing audit log for Import action..."
            try {
                $moduleCount = (Get-ChildItem -Path $moduleDir -Filter '*.bas' -File).Count
                $classCount = (Get-ChildItem -Path $classDir -Filter '*.cls' -File).Count
                $totalCount = $moduleCount + $classCount
                $auditScript = Join-Path $scriptRoot 'append-audit-entry.ps1'
                & $auditScript -Action "import" -Module "bulk_import" -Details "Imported $totalCount modules ($moduleCount .bas, $classCount .cls) via Start-VbaAction.ps1"
            }
            catch {
                Write-Warn "Failed to write audit log: $_"
            }
        }
        'Export' {
            Write-Info "Executing Action: Export"
            # Core logic from export-modules-robust.ps1
            $moduleDir = Join-Path $repoRoot 'vba-files\Module'
            $classDir = Join-Path $repoRoot 'vba-files\Class'
            $exportCount = 0
            foreach ($vbComp in $vbProj.VBComponents) {
                $destPath = $null
                switch ($vbComp.Type) {
                    1 { $destPath = Join-Path $moduleDir ($vbComp.Name + '.bas') }
                    2 { $destPath = Join-Path $classDir ($vbComp.Name + '.cls') }
                }
                if ($destPath) {
                    $vbComp.Export($destPath)
                    Write-Info "Exported $($vbComp.Name) to $destPath"
                    $exportCount++
                }
            }
            
            # Audit log for successful export
            Write-Info "Writing audit log for Export action..."
            try {
                $auditScript = Join-Path $scriptRoot 'append-audit-entry.ps1'
                & $auditScript -Action "export" -Module "bulk_export" -Details "Exported $exportCount modules via Start-VbaAction.ps1"
            }
            catch {
                Write-Warn "Failed to write audit log: $_"
            }
        }
        'Compile' {
            Write-Info "Executing Action: Compile"
            # Full logic from test-vbide-and-compile.ps1
            $compileInvoked = $false
            $global:__CompileClean = $false
            try {
                $menuBar = $excel.VBE.CommandBars.Item('Menu Bar')
                $topMenus = $menuBar.Controls | ForEach-Object {
                    $cap = $_.Caption
                    $norm = ($cap -replace '&', '') -replace '\.\.\.', '' -replace '\s+', ' ' | ForEach-Object { $_.Trim() }
                    [pscustomobject]@{ Control = $_; Caption = $cap; Norm = $norm }
                }
                $debugMenu = ($topMenus | Where-Object { $_.Norm -match '(?i)debug|debog|d.ebog' } | Select-Object -First 1).Control
                if ($null -eq $debugMenu) { throw "Could not find Debug menu on VBE.CommandBars" }

                $controlsWithNorm = $debugMenu.Controls | ForEach-Object {
                    $cap = $_.Caption
                    $norm = ($cap -replace '&', '') -replace '\.\.\.', '' -replace '\s+', ' ' | ForEach-Object { $_.Trim() }
                    [pscustomobject]@{ Control = $_; Caption = $cap; Norm = $norm }
                }
                $compileCtrlObj = $controlsWithNorm | Where-Object { $_.Norm -match '^(?i)compile|compiler|compilar|kompil|compilare|kompilieren' } | Select-Object -First 1
                if ($null -eq $compileCtrlObj) { $compileCtrlObj = $controlsWithNorm | Where-Object { $_.Norm -match '^(?i)comp' } | Select-Object -First 1 }

                if ($null -eq $compileCtrlObj) {
                    Write-Warn "No control matching Compile found under Debug menu."
                }
                elseif (-not $compileCtrlObj.Control.Enabled) {
                    Write-Info ("Compile control disabled: '" + $compileCtrlObj.Caption + "' (project likely compiled/clean)")
                    $compileInvoked = $false
                    $global:__CompileClean = $true
                }
                else {
                    Write-Info ("Invoking: '" + $compileCtrlObj.Caption + "'")
                    $compileCtrlObj.Control.Execute()
                    $compileInvoked = $true
                    Start-Sleep -Milliseconds 300
                    
                    # Re-evaluate to see if compile succeeded
                    $compileAfter = ($debugMenu.Controls | ForEach-Object { [pscustomobject]@{ Control = $_; Norm = ($_.Caption -replace '&' -replace '\.\.\.' -replace '\s+', ' ').Trim() } } | Where-Object { $_.Norm -match '^(?i)compile|compiler' } | Select -First 1)
                    if ($compileAfter -and -not $compileAfter.Control.Enabled) {
                        Write-Info "Compile state after invoke: Compile control is disabled (clean)."
                        $global:__CompileClean = $true
                    }
                    else {
                        Write-Warn "Compile state after invoke: Compile control is still enabled (errors likely remain)."
                        $global:__CompileClean = $false
                    }
                }
            }
            catch {
                Write-Warn ("Could not invoke compile: " + $_.Exception.Message)
            }

            if ($compileInvoked -and ($global:__CompileClean -ne $true)) {
                try {
                    $activePane = $excel.VBE.ActiveCodePane
                    if ($null -ne $activePane) {
                        $cm = $activePane.CodeModule
                        $sl = 0; $sc = 0; $el = 0; $ec = 0
                        $null = $activePane.GetSelection([ref]$sl, [ref]$sc, [ref]$el, [ref]$ec)
                        $focusLine = if ($sl -gt 0) { $sl } else { 1 }
                        $startLine = [Math]::Max(1, $focusLine - 2)
                        $endLine = [Math]::Min($cm.CountOfLines, $focusLine + 2)
                        $lines = $cm.Lines($startLine, ($endLine - $startLine + 1))
                        Write-Warn ("Compile error detected in module: " + $cm.Name + ", around line: " + $focusLine)
                        $ln = $startLine
                        foreach ($l in $lines -split "`r?`n") {
                            Write-Host ("  [" + $ln + "] " + $l)
                            $ln++
                        }
                    }
                }
                catch {
                    Write-Warn ("Could not capture active pane context: " + $_.Exception.Message)
                }
            }
            
            if ($global:__CompileClean -ne $true) { exit 1 }
        }
        'RunMacro' {
            Write-Info "Executing Action: RunMacro"
            if (-not $MacroName) {
                Write-Err "MacroName parameter is required for RunMacro action."
                exit 12
            }
            Write-Info "Running macro: $MacroName"
            $excel.Run($MacroName)
        }
    }
    #endregion

    #region Robust Save (from import-modules-robust.ps1)
    if ($wb.Saved -eq $false) {
        Write-Info "Workbook has unsaved changes. Initiating robust save..."
        $saved = $false
        try {
            if (-not $wb.ReadOnly) { $wb.Save(); $saved = $true; Write-Info "Primary Save successful." }
        }
        catch {
            Write-Warn "Primary Save failed: $_"
        }

        if (-not $saved) {
            $tempDir = [System.IO.Path]::GetTempPath()
            $tempCopy = Join-Path $tempDir ("sgq_action_tmp_" + [guid]::NewGuid().ToString() + ".xlsm")
            try {
                $wb.SaveAs($tempCopy, 52) # 52 = xlOpenXMLWorkbookMacroEnabled
                Write-Info "SaveAs to temp file successful: $tempCopy"
                $saved = $true
            }
            catch {
                Write-Warn "SaveAs fallback failed: $_"
            }

            # This block executes outside the main try/catch to ensure cleanup happens
            try {
                $wb.Close($false)
            }
            catch {}
            if ($excel) { try { $excel.Quit() } catch {} }
            # Release COM objects before file move
            if ($vbProj) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($vbProj) | Out-Null; $vbProj = $null }
            if ($wb) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null; $wb = $null }
            if ($excel) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null; $excel = $null }
            [GC]::Collect(); [GC]::WaitForPendingFinalizers()

            if ($saved -and (Test-Path -LiteralPath $tempCopy)) {
                try {
                    Move-Item -LiteralPath $tempCopy -Destination $WorkbookPath -Force
                    Write-Info "Replaced original workbook with saved temp copy."
                }
                catch {
                    Write-Err "Failed to replace original workbook with temp copy: $_"
                    exit 13
                }
            }
        }
    }
    #endregion

}
catch {
    Write-Err "An unhandled error occurred: $($_.Exception.Message)"
    # You can add more specific error handling here
    exit 99
}
finally {
    #region Final Cleanup (from export-modules-robust.ps1)
    Write-Info "Performing final cleanup."
    try {
        if ($wb -and -not $wb.Saved) { $wb.Close($false) } elseif ($wb) { $wb.Close() }
    }
    catch {}
    try {
        if ($excel) { $excel.Quit() }
    }
    catch {}

    # Release COM objects
    try { if ($vbProj) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($vbProj) | Out-Null } } catch {}
    try { if ($wb) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null } } catch {}
    try { if ($excel) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } } catch {}
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    # Stop watchdog
    try {
        if (Test-Path $sentinel) { Remove-Item -Path $sentinel -Force -ErrorAction SilentlyContinue }
    }
    catch {}
    try { if ($watchJob) { Stop-Job -Job $watchJob -Force -ErrorAction SilentlyContinue; Receive-Job -Job $watchJob -ErrorAction SilentlyContinue } } catch {}
    Write-Info "Cleanup complete."
    #endregion
}
ö¨"(9e377b08b88e30057baf56c074b3d12f1d22372325file:///c:/VBA/SGQ%201.65/scripts/Start-VbaAction.ps1:file:///c:/VBA/SGQ%201.65