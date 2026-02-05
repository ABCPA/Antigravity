®^param(
    [string]$WorkbookPath = (Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) 'index.xlsm'),
    [switch]$AttemptCompile,
    [switch]$ImportBeforeCompile,
    [switch]$ForceCompile
)

# Purpose: Safely test VBIDE access (Trust access to the VBA project object model)
#          and optionally attempt a compile via the VBE Debug menu.
# Notes:   This avoids complicated -Command quoting issues by using a script file.
#          Compatible with Windows PowerShell and PowerShell 7 on Windows.

$ErrorActionPreference = 'Stop'

function Write-Info($msg) { Write-Host "[INFO] $msg" }
function Write-Warn($msg) { Write-Warning $msg }
function Write-Err ($msg) { Write-Error $msg }

# Create Excel COM
$excel = $null
$wb = $null
$vbideAccessible = $false
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent (Split-Path -Parent $scriptRoot)
$importRan = $false

try {
    Write-Info 'Creating Excel COM instance...'
    $excel = New-Object -ComObject Excel.Application
    $excel.DisplayAlerts = $false

    # Test VBIDE access
    try {
        $null = $excel.VBE
        $vbideAccessible = $true
        Write-Info 'VBIDE access: OK (Trust access appears enabled)'
    }
    catch {
        $vbideAccessible = $false
        Write-Warn ('VBIDE access error: ' + $_.Exception.Message)
    }

    if (-not $vbideAccessible) {
        Write-Output 'VBIDE access allowed: False'
        return
    }

    if ($AttemptCompile) {
        if ($ImportBeforeCompile) {
            try {
                $importWrapper = Join-Path (Split-Path -Parent $scriptRoot) 'import-modules-noninteractive.ps1'
                if (-not (Test-Path -LiteralPath $importWrapper)) {
                    Write-Warn "Import wrapper not found: $importWrapper"
                }
                else {
                    Write-Info "Running import wrapper before compile: $importWrapper"
                    & pwsh -NoProfile -ExecutionPolicy Bypass -File $importWrapper | Write-Host
                    Write-Info 'Import wrapper completed'
                    $importRan = $true
                }
            }
            catch {
                Write-Warn ('Import-before-compile failed: ' + $_.Exception.Message)
            }
        }
        if (-not $WorkbookPath) {
            throw 'WorkbookPath is empty; provide a path to an .xlsm workbook.'
        }
        # Normalize to absolute file system path string
        try {
            $resolvedPath = (Resolve-Path -LiteralPath $WorkbookPath -ErrorAction Stop).ProviderPath
        }
        catch {
            throw "Workbook does not exist: $WorkbookPath"
        }

        if ($importRan) {
            try {
                $workbookDir = Split-Path -Parent $resolvedPath
                if (-not $workbookDir) { $workbookDir = $repoRoot }
                $archiveDir = Join-Path $repoRoot 'logs\workbooks'
                if (-not (Test-Path -LiteralPath $archiveDir)) {
                    New-Item -ItemType Directory -Path $archiveDir -Force | Out-Null
                }
                $backups = Get-ChildItem -Path $workbookDir -Filter 'import_backup_*.xlsm' -ErrorAction SilentlyContinue
                foreach ($bk in $backups) {
                    $destination = Join-Path $archiveDir $bk.Name
                    Move-Item -LiteralPath $bk.FullName -Destination $destination -Force
                    Write-Info "Relocated workbook backup to: $destination"
                }
            }
            catch {
                Write-Warn ('Failed to relocate workbook backups: ' + $_.Exception.Message)
            }
        }

        # Guard: if a workbook with same Name is already open, reuse it instead of opening
        $targetName = [System.IO.Path]::GetFileName($resolvedPath)
        $wb = $null
        foreach ($w in $excel.Workbooks) {
            if ($w.Name -ieq $targetName) { $wb = $w; break }
        }

        if ($null -eq $wb) {
            Write-Info "Opening workbook: $resolvedPath"
            $wb = $excel.Workbooks.Open($resolvedPath)
        }
        else {
            Write-Info "Reusing already-open workbook: $targetName"
        }

        # Retry logic for RPC_E_CALL_REJECTED
        $maxRetries = 3
        $retryDelaySeconds = 5
        $retryCount = 0

        while ($retryCount -lt $maxRetries) {
            try {
                $wb.Activate() | Out-Null
                break
            }
            catch {
                if ($_.Exception.HResult -eq 0x80010001) {
                    Write-Warn "RPC_E_CALL_REJECTED encountered. Retrying in $retryDelaySeconds seconds..."
                    Start-Sleep -Seconds $retryDelaySeconds
                    $retryCount++
                }
                else {
                    throw $_
                }
            }
        }

        if ($retryCount -ge $maxRetries) {
            throw "Failed to activate workbook after $maxRetries retries due to RPC_E_CALL_REJECTED."
        }

        # Try to locate the Debug->Compile command regardless of localization by scanning captions
        $compileInvoked = $false
        try {
            $menuBar = $excel.VBE.CommandBars.Item('Menu Bar')
            # Normalize top-level menu captions to remove '&' and '...'
            $topMenus = $menuBar.Controls | ForEach-Object {
                $cap = $_.Caption
                $norm = ($cap -replace '&', '')
                $norm = ($norm -replace '\.\.\.', '')
                $norm = ($norm -replace '\s+', ' ').Trim()
                [pscustomobject]@{ Control = $_; Caption = $cap; Norm = $norm }
            }
            # Match Debug menu in several locales (English/French), without start anchor to allow prefixes
            $debugMenu = ($topMenus | Where-Object { $_.Norm -match '(?i)debug|debog|d.ebog' } | Select-Object -First 1).Control
            if ($null -eq $debugMenu) { throw 'Could not find Debug menu on VBE.CommandBars' }

            # List available controls for diagnostics
            Write-Info 'Debug menu controls:'
            $i = 0
            foreach ($c in $debugMenu.Controls) {
                $i++
                Write-Host ("  [{0}] Caption='{1}' Enabled={2}" -f $i, ($c.Caption), ($c.Enabled))
            }

            # Normalize captions: remove accelerator '&', ellipses, compress spaces
            $controlsWithNorm = $debugMenu.Controls | ForEach-Object {
                $cap = $_.Caption
                $norm = ($cap -replace '&', '')
                $norm = ($norm -replace '\.\.\.', '')
                $norm = ($norm -replace '\s+', ' ').Trim()
                [pscustomobject]@{ Control = $_; Caption = $cap; Norm = $norm }
            }

            # Look for common variants of Compile in multiple languages
            $compileCtrlObj = $controlsWithNorm | Where-Object {
                $_.Norm -match '^(?i)compile|compiler|compilar|kompil|compilare|kompilieren'
            } | Select-Object -First 1

            if ($null -eq $compileCtrlObj) {
                # Fallback heuristic: first control whose normalized caption starts with 'Comp'
                $compileCtrlObj = $controlsWithNorm | Where-Object { $_.Norm -match '^(?i)comp' } | Select-Object -First 1
            }

            if ($null -eq $compileCtrlObj) {
                Write-Warn 'No control matching Compile found under Debug menu.'
            }
            elseif (-not $compileCtrlObj.Control.Enabled) {
                if ($ForceCompile.IsPresent) {
                    Write-Warn ("Compile control disabled: '" + $compileCtrlObj.Caption + "' but -ForceCompile specified; attempting to Execute() anyway.")
                    try {
                        $compileCtrlObj.Control.Execute()
                        $compileInvoked = $true
                    }
                    catch {
                        Write-Warn ('Forced compile invocation failed: ' + $_.Exception.Message)
                        $compileInvoked = $false
                    }
                }
                else {
                    Write-Info ("Compile control disabled: '" + $compileCtrlObj.Caption + "' (project likely compiled/clean)")
                    $compileInvoked = $false
                }
            }
            else {
                Write-Info ("Invoking: '" + $compileCtrlObj.Caption + "'")
                $compileCtrlObj.Control.Execute()
                $compileInvoked = $true

                # Allow UI to update
                Start-Sleep -Milliseconds 300

                # Re-evaluate Debug menu after invocation to determine if compile cleaned project
                $controlsAfter = $debugMenu.Controls | ForEach-Object {
                    $cap = $_.Caption
                    $norm = ($cap -replace '&', '')
                    $norm = ($norm -replace '\.\.\.', '')
                    $norm = ($norm -replace '\s+', ' ').Trim()
                    [pscustomobject]@{ Control = $_; Caption = $cap; Norm = $norm }
                }
                $compileAfter = $controlsAfter | Where-Object { $_.Norm -match '^(?i)compile|compiler|compilar|kompil|compilare|kompilieren' } | Select-Object -First 1
                if ($compileAfter -and -not $compileAfter.Control.Enabled) {
                    Write-Info 'Compile state after invoke: Compile control is disabled (clean).'
                    $global:__CompileClean = $true
                }
                else {
                    Write-Info 'Compile state after invoke: Compile control is still enabled (errors likely remain).'
                    $global:__CompileClean = $false
                }
            }
        }
        catch {
            Write-Warn ('Could not invoke compile: ' + $_.Exception.Message)
        }

        Write-Output ('Compile attempted: ' + $compileInvoked)
        if ($compileInvoked) {
            Write-Output ('Compile clean: ' + ($global:__CompileClean -eq $true))
        }

        # If compile was attempted and not clean, try to log current active pane context (best-effort)
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
                    Write-Info ('Active module: ' + $cm.Name + ', around line: ' + $focusLine)
                    $ln = $startLine
                    foreach ($l in $lines -split "`r?`n") {
                        Write-Host ('  [' + $ln + '] ' + $l)
                        $ln++
                    }
                }
            }
            catch {
                Write-Warn ('Could not capture active pane context: ' + $_.Exception.Message)
            }
        }
    }

    Write-Output ('VBIDE access allowed: ' + $vbideAccessible)
}
finally {
    if ($null -ne $wb) {
        try { $wb.Close($false) } catch {}
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
    }
    [gc]::Collect(); [gc]::WaitForPendingFinalizers()
}

# Set exit code based on compile cleanliness when attempted
if ($AttemptCompile.IsPresent) {
    if ($global:__CompileClean -eq $true) { exit 0 }
    elseif ($global:__CompileClean -eq $false) { exit 1 }
}
®^*cascade08"(9e377b08b88e30057baf56c074b3d12f1d2237232Ffile:///c:/VBA/SGQ%201.65/vba-files/scripts/test-vbide-and-compile.ps1:file:///c:/VBA/SGQ%201.65