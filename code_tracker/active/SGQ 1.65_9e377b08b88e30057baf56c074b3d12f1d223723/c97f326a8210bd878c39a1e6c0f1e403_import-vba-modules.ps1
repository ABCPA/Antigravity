‰## AUDIT : sauvegarde 20251128_090702 -> backups/20251128_090702/import-vba-modules.ps1
param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,

    [string[]]$ModulesToImport = @(
        "vba-files/Module/modAppStateGuard.bas",
        "vba-files/Module/modExcelUtils.bas",
        "vba-files/Module/modSGQUtilitaires.bas"
    ),

    [switch]$VerboseLog
)

# Remarque : Active la reference VBE (Trust access to the VBA project object model) avant d'executer.
# Etape 11.4 : utilise VBComponents.Import pour conserver les attributs de module.

function Write-Log([string]$message) {
    if ($VerboseLog) { Write-Host "[import] $message" }
}

if (-not (Test-Path $WorkbookPath)) {
    throw "Classeur introuvable : $WorkbookPath"
}

$absWb = (Resolve-Path $WorkbookPath).Path

# √âtape de s√©curit√© : Normalisation de l'encodage (D√©sactiv√©e pour ce run car d√©j√† fait)
# Write-Log "Normalisation des encodages des fichiers sources..."
# & "$PSScriptRoot\ensure-encoding.ps1" -SearchPath (Join-Path $PSScriptRoot "..\vba-files")

$absModules = @()
foreach ($m in $ModulesToImport) {
    if (-not (Test-Path $m)) { throw "Module introuvable : $m" }
    $absModules += (Resolve-Path $m).Path
}

# Ferme Excel proprement apres import.
$excel = $null
$workbook = $null
try {
    Write-Log "Ouverture Excel et du classeur $absWb"
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open($absWb)

    # Si projet protege, lever une erreur claire.
    if ($workbook.VBProject.Protection -ne 0) {
        throw "VBProject protege : impossible d'importer les modules."
    }

    $vbProj = $workbook.VBProject

    foreach ($modPath in $absModules) {
        $targetName = [IO.Path]::GetFileNameWithoutExtension($modPath)
        try {
            # Recherche du composant existant
            $existing = $null
            foreach ($comp in $vbProj.VBComponents) {
                if ($comp.Name -ieq $targetName) { $existing = $comp; break }
            }
            
            if ($null -ne $existing) {
                if ($existing.Type -eq 100) {
                    # Document module (Sheet/ThisWorkbook)
                    Write-Log "Mise √† jour code Document Module $targetName"
                    $codeMod = $existing.CodeModule
                    
                    # Charger et filtrer les lignes de m√©tadonn√©es
                    $lines = Get-Content $modPath -Encoding Default
                    $filteredLines = $lines | Where-Object { 
                        $_ -notmatch '^VERSION' -and 
                        $_ -notmatch '^Attribute' -and 
                        $_ -notmatch '^BEGIN' -and 
                        $_ -notmatch '^END' -and
                        $_ -notmatch '^\s+(MultiUse|Persistable|DataBindingBehavior|DataSourceBehavior|MTSTransactionMode)'
                    }
                    $cleanCode = ($filteredLines -join "`r`n").Trim()
                    
                    if ($codeMod.CountOfLines -gt 0) { $codeMod.DeleteLines(1, $codeMod.CountOfLines) }
                    if ($cleanCode.Length -gt 0) { $codeMod.AddFromString($cleanCode) }
                }
                else {
                    Write-Log "Suppression et r√©-import du module $targetName"
                    $vbProj.VBComponents.Remove($existing)
                    $vbProj.VBComponents.Import($modPath) | Out-Null
                }
            }
            else {
                Write-Log "Achat/Import du nouveau module $targetName"
                $vbProj.VBComponents.Import($modPath) | Out-Null
            }
            
            # Petit d√©lai pour laisser souffler Excel et √©viter les collisions COM
            Start-Sleep -Milliseconds 100
        }
        catch {
            Write-Host "?? ERREUR sur $targetName : $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    Write-Log "Sauvegarde finale du classeur"
    $workbook.Save()
    Write-Host "Synchronisation termin√©e avec succ√®s pour $absWb"
}
finally {
    if ($null -ne $workbook) { 
        $workbook.Close($false) 
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null 2>$null
    }
    if ($null -ne $excel) { 
        $excel.Quit() 
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null 2>$null
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers();
}
Á *cascade08ÁÎ*cascade08ÎÏ *cascade08ÏÌ*cascade08ÌÔ *cascade08ÔÙ*cascade08Ùˆ *cascade08ˆ˜*cascade08˜¯ *cascade08¯é*cascade08éè *cascade08èê*cascade08êí *cascade08íî*cascade08î’ *cascade08’◊*cascade08◊· *cascade08·‚*cascade08‚„ *cascade08„‰*cascade08‰Â *cascade08ÂË*cascade08ËÈ *cascade08ÈÍ*cascade08ÍÌ *cascade08Ì˜*cascade08˜¯ *cascade08¯˛*cascade08˛Ä *cascade08ÄÑ*cascade08Ñá *cascade08áà*cascade08àâ *cascade08âä*cascade08äé *cascade08éè*cascade08èê *cascade08êë*cascade08ëí *cascade08íì*cascade08ìî *cascade08îï*cascade08ïó *cascade08óõ*cascade08õ∂ *cascade08∂∫*cascade08∫˘ *cascade08˘˝*cascade08˝º *cascade08º¿*cascade08¿” *cascade08”◊*cascade08◊Ÿ *cascade08Ÿ›*cascade08›Â *cascade08ÂÎ *cascade08ÎÓ*cascade08ÓÙ *cascade08Ù˙*cascade08˙˚ *cascade08˚¸*cascade08¸Å *cascade08ÅÖ*cascade08Öë *cascade08ë¿ *cascade08¿ƒ*cascade08ƒŒ *cascade08Œ’*cascade08’Ï *cascade08ÏÌ*cascade08Ì˝ *cascade08˝Ä*cascade08ÄÑ*cascade08ÑÖ *cascade08Öâ*cascade08âä *cascade08äé*cascade08éê *cascade08êí*cascade08íì *cascade08ìï*cascade08ïò *cascade08òú*cascade08ú≠ *cascade08≠∞*cascade08∞± *cascade08±≤*cascade08≤≥ *cascade08≥∂*cascade08∂∑ *cascade08∑∏ *cascade08∏ª *cascade08ªø*cascade08ø  *cascade08 À *cascade08À–*cascade08–“ *cascade08“·*cascade08·‚ *cascade08‚ *cascade08Ò*cascade08ÒÅ *cascade08ÅÑ*cascade08ÑÜ *cascade08Üä*cascade08ä• *cascade08•¶ *cascade08¶¨*cascade08¨≠ *cascade08≠∂*cascade08∂∑ *cascade08∑  *cascade08 Ã*cascade08ÃŒ *cascade08Œ– *cascade08–‡ *cascade08‡‚*cascade08‚„ *cascade08„*cascade08Ú *cascade08Ú˛*cascade08˛ˇ *cascade08ˇÇ*cascade08ÇÉ *cascade08ÉÜ*cascade08Üà *cascade08àë *cascade08ëí*cascade08í¢ *cascade08¢•*cascade08•ß *cascade08ß® *cascade08®≤*cascade08≤≥ *cascade08≥∫*cascade08∫ª *cascade08ªº*cascade08ºΩ *cascade08Ωæ*cascade08æø *cascade08ø«*cascade08«» *cascade08»·*cascade08·‚ *cascade08‚‰ *cascade08‰Ë*cascade08ËÚ *cascade08ÚÛ *cascade08ÛÉ*cascade08ÉÑ *cascade08Ñá *cascade08áâ*cascade08âñ *cascade08ñó *cascade08óù *cascade08ùü*cascade08ü§ *cascade08§• *cascade08•©*cascade08©™ *cascade08™¥*cascade08¥µ *cascade08µ∂*cascade08∂∑ *cascade08∑ª*cascade08ªº *cascade08º‘ *cascade08‘ÿ*cascade08ÿﬁ *cascade08ﬁﬂ *cascade08ﬂ‰*cascade08‰Â *cascade08Âı *cascade08ı˜*cascade08˜ã *cascade08ãç*cascade08çí *cascade08íì *cascade08ì§*cascade08§• *cascade08•ª *cascade08ªø*cascade08ø— *cascade08—“*cascade08“” *cascade08”’ *cascade08’Ÿ*cascade08Ÿ⁄ *cascade08⁄€*cascade08€‹ *cascade08‹› *cascade08›·*cascade08·„ *cascade08„Â*cascade08ÂÊ *cascade08ÊÁ*cascade08ÁÓ *cascade08ÓÔ *cascade08ÔÛ*cascade08ÛÙ *cascade08Ù¯*cascade08¯˘ *cascade08˘˙ *cascade08˙˚*cascade08˚Ñ *cascade08ÑÖ *cascade08Öä*cascade08äã *cascade08ãç *cascade08çé*cascade08éò *cascade08òö *cascade08öõ*cascade08õú *cascade08úù*cascade08ù† *cascade08†°*cascade08°§ *cascade08§•*cascade08•µ *cascade08µ∏*cascade08∏ª *cascade08ªº*cascade08ºÃ *cascade08Ãœ*cascade08œ÷ *cascade08÷ÿ *cascade08ÿ‡*cascade08‡· *cascade08·‚*cascade08‚„ *cascade08„í*cascade08íñ*cascade08ñò *cascade08òú*cascade08ú˝ *cascade08˝ˇ*cascade08ˇï *cascade08ï¥ *cascade08¥ƒ *cascade08ƒ«*cascade08«» *cascade08» *cascade08 Ã *cascade08Ã“*cascade08“‘ *cascade08‘’*cascade08’◊ *cascade08◊‡*cascade08‡‚ *cascade08‚‰*cascade08‰Í *cascade08ÍÎ*cascade08ÎÖ *cascade08Öâ*cascade08âô *cascade08ô∞ *cascade08∞¡*cascade08¡÷ *cascade08÷⁄*cascade08⁄Ê *cascade08ÊÍ*cascade08Íê *cascade08êû *cascade08ûË*cascade08ËÈ *cascade08ÈÎ *cascade08ÎÔ*cascade08Ô˙ *cascade08˙å*cascade08åé *cascade08éï*cascade08ïù *cascade08ùû*cascade08û© *cascade08©Ø*cascade08Øπ *cascade08π¡*cascade08¡… *cascade08… *cascade08 À *cascade08ÀŒ*cascade08Œœ *cascade08œ”*cascade08”÷ *cascade08÷ﬁ*cascade08ﬁò *cascade08ò† *cascade08†•*cascade08•Ø *cascade08Ø¡*cascade08¡Î *cascade08Îâ*cascade08âã *cascade08ãè*cascade08è± *cascade08±≥*cascade08≥∂ *cascade08∂¿*cascade08¿· *cascade08·Ë*cascade08Ëû  *cascade08û § *cascade08§ •  *cascade08• © *cascade08© ™  *cascade08™ ≠ *cascade08≠ ¥  *cascade08¥ ∂ *cascade08∂ ∏  *cascade08∏ ≈ *cascade08≈ Í  *cascade08Í Ù *cascade08Ù Ä! *cascade08Ä!ä!*cascade08ä!´! *cascade08´!¨! *cascade08¨!Ñ" *cascade08Ñ"â"*cascade08â"ã"*cascade08ã"ë" *cascade08ë"©"*cascade08©"À"*cascade08À"†# *cascade08†#•#*cascade08•#ß#*cascade08ß#‰# *cascade08"(9e377b08b88e30057baf56c074b3d12f1d22372328file:///c:/VBA/SGQ%201.65/scripts/import-vba-modules.ps1:file:///c:/VBA/SGQ%201.65