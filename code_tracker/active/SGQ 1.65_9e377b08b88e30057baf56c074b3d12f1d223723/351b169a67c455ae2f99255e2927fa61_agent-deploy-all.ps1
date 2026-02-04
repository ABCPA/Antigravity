�?<#
Agent deploy all-in-one
Copiez-collez l'intégralité de ce fichier dans votre agent local (PowerShell Core / pwsh).
Ce script va :
- créer un dossier "scripts" dans le dossier courant,
- y écrire deux scripts robustes : inspect-chunk.fixed.ps1 et import-single-module-and-save.fixed.ps1,
- lancer le script d'import avec les chemins par défaut (modifiable via paramètres).

Usage (sur l'agent) :
  pwsh -NoProfile -ExecutionPolicy Bypass -File .\agent-deploy-all.ps1
ou avec paramètres :
  pwsh -NoProfile -ExecutionPolicy Bypass -File .\agent-deploy-all.ps1 -ModulePath "C:\VBA\SGQ 1.65\vba-files\Module\modRibbonSGQ.bas" -ExcelFile "C:\VBA\SGQ 1.65\index.xlsm"

Note : exécutez en tant qu'utilisateur ayant les autorisations sur les fichiers et sur l'automation COM Excel (VBIDE access).
#>

param(
  [string]$ModulePath = "C:\VBA\SGQ 1.65\vba-files\Module\modRibbonSGQ.bas",
  [string]$ExcelFile = "C:\VBA\SGQ 1.65\index.xlsm",
  [string]$ScriptsFolder = ".\scripts"
)

# Crée le dossier scripts si nécessaire
New-Item -Path $ScriptsFolder -ItemType Directory -Force | Out-Null

# Contenu du script inspect-chunk.fixed.ps1 (version robuste)
$inspectContent = @'
<#
Inspect-chunk - version robuste
Affiche un extrait autour d'un pattern et la table des caractères (offset, char, code hex).
Usage:
 pwsh -NoProfile -File .\inspect-chunk.fixed.ps1 -Path "d:\path\to\file.bas" -Pattern "TrySetWorksheetScrollArea" -Length 300
#>
param(
  [Parameter(Mandatory=$true)][string]$Path,
  [string]$Pattern = "TrySetWorksheetScrollArea",
  [int]$Length = 200
)

if (-not (Test-Path $Path)) {
  Write-Error "File not found: $Path"
  exit 2
}

# Lecture binaire puis décodage UTF8 strict
$bytes = [System.IO.File]::ReadAllBytes($Path)
# Try decode as UTF8
[string]$text = [System.Text.Encoding]::UTF8.GetString($bytes)

$pos = $text.IndexOf($Pattern)
if ($pos -lt 0) {
  Write-Host "pattern not found"
  exit 1
}

$len = [math]::Min($Length, $text.Length - $pos)
$chunk = $text.Substring($pos, $len)

Write-Output '---chunk---'
Write-Output $chunk
Write-Output '---chars with codes---'

$i = 0
foreach ($c in $chunk.ToCharArray()) {
  $code = [int][char]$c
  # Print printable representation (escape CR/LF)
  $repr = switch ($code) {
    13 { '\r' }
    10 { '\n' }
    9  { '\t' }
    default { $c }
  }
  # Format with -f to avoid literal braces issues (we use single quoted format)
  Write-Output ('{0,3}: ''{1}'' (0x{2:X2})' -f $i, $repr, $code)
  $i++
  if ($i -gt $Length) { break }
}
'@

# Contenu du script import-single-module-and-save.fixed.ps1 (version robuste)
$importContent = @'
<#
Import-single-module-and-save - version robuste
- Détecte encodage / caractères invalides dans le .bas, sanitize si besoin.
- Sauvegarde backup du .bas et du workbook (temp copy).
- Ouvre Excel via COM, supprime le composant existant du même nom s'il existe, importe le module.
- Affiche Exception.Message et HResult en hex en cas d'erreur.

Usage:
 pwsh -NoProfile -ExecutionPolicy Bypass -File .\import-single-module-and-save.fixed.ps1 -ModulePath "d:\...modRibbonSGQ.bas" -ExcelFile "d:\...index.xlsm"
#>
param(
  [Parameter(Mandatory=$true)][string]$ModulePath,
  [Parameter(Mandatory=$true)][string]$ExcelFile,
  [string]$BackupRoot = ".\vba-files\backups"
)

function Sanitize-ModuleFile {
  param(
    [string]$Path
  )
  $bytes = [System.IO.File]::ReadAllBytes($Path)
  $utf8 = [System.Text.Encoding]::UTF8.GetString($bytes)
  $hasReplacement = $utf8.Contains([char]0xFFFD)

  if (-not $hasReplacement) {
    Write-Host "No UTF-8 replacement char detected. Using original file."
    return $Path
  }

  Write-Host "Replacement chars detected in UTF-8 decode. Trying Windows-1252 interpretation and sanitize..."

  $cp1252 = [System.Text.Encoding]::GetEncoding(1252).GetString($bytes)

  # Remove/control non-printable chars except CR/LF/Tab
  $chars = $cp1252.ToCharArray()
  $cleanChars = foreach ($c in $chars) {
    $code = [int][char]$c
    if ($code -eq 13 -or $code -eq 10 -or $code -eq 9) {
      $c
    } elseif ($code -lt 32) {
      ''  # drop control chars
    } else {
      $c
    }
  }
  $sanitized = -join $cleanChars

  $temp = [System.IO.Path]::GetTempFileName()
  $tempPath = [System.IO.Path]::ChangeExtension($temp, '.bas')
  # Save UTF8 without BOM (safer for VBIDE import)
  [System.IO.File]::WriteAllText($tempPath, $sanitized, [System.Text.Encoding]::UTF8)
  Write-Host "Sanitized file written to: $tempPath"
  return $tempPath
}

if (-not (Test-Path $ModulePath)) { Write-Error "ModulePath not found: $ModulePath"; exit 2 }
if (-not (Test-Path $ExcelFile)) { Write-Error "ExcelFile not found: $ExcelFile"; exit 2 }

# Backup original module file
$timestamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
$backupDir = Join-Path $BackupRoot $timestamp
New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
$backupModule = Join-Path $backupDir ([System.IO.Path]::GetFileName($ModulePath))
Copy-Item -Path $ModulePath -Destination $backupModule -Force
Write-Host "Backed up module to: $backupModule"

# sanitize if needed
$toImportPath = Sanitize-ModuleFile -Path $ModulePath

# Open Excel and import
$ComponentName = [System.IO.Path]::GetFileNameWithoutExtension($ModulePath)
try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $wb = $excel.Workbooks.Open((Resolve-Path $ExcelFile).ProviderPath)
  $vb = $wb.VBProject
  Write-Host "VBProject opened. Attempting to remove existing component named '$ComponentName' if present..."
  try {
    $existing = $vb.VBComponents.Item($ComponentName)
    if ($existing -ne $null) {
      $vb.VBComponents.Remove($existing)
      Write-Host "Removed existing VB component: $ComponentName"
    }
  } catch {
    # ignore if not present
    Write-Host "No existing component named '$ComponentName' or unable to enumerate (OK)."
  }

  try {
    $vb.VBComponents.Import($toImportPath)
    Write-Host "Imported module: $toImportPath"
  } catch {
    $ex = $_.Exception
    $hr = ''
    if ($ex -ne $null -and $ex.HResult -ne $null) { $hr = $ex.HResult.ToString('X') }
    Write-Error "Import failed: $($ex.Message) HResult=$hr"
    # Save workbook as-is to keep state
    try { $wb.Save() } catch {}
    $wb.Close($false)
    $excel.Quit()
    exit 3
  }

  # Save workbook
  $wb.Save()
  Write-Host "Workbook saved after import."

  # Close and quit
  $wb.Close($false)
  $excel.Quit()
  Write-Host "Done. Excel closed."
} catch {
  $ex = $_.Exception
  $hr = ''
  if ($ex -ne $null -and $ex.HResult -ne $null) { $hr = $ex.HResult.ToString('X') }
  Write-Error "Outer error during COM operations: $($ex.Message) HResult=$hr"
  try { $excel.Quit() } catch {}
  exit 4
} finally {
  [gc]::Collect(); [gc]::WaitForPendingFinalizers()
}

# If we created a sanitized temp file, inform user and keep it for inspection
if ($toImportPath -ne $ModulePath) {
  Write-Host "NOTE: sanitized import file used: $toImportPath"
  Write-Host "Original backed up to: $backupModule"
}
'@

# Écrit les fichiers sur disque (UTF8)
$inspectPath = Join-Path $ScriptsFolder "inspect-chunk.fixed.ps1"
$importPath = Join-Path $ScriptsFolder "import-single-module-and-save.fixed.ps1"

# Write files
$inspectContent | Out-File -FilePath $inspectPath -Encoding utf8 -Force
$importContent  | Out-File -FilePath $importPath  -Encoding utf8 -Force

Write-Host "Wrote scripts to: $inspectPath"
Write-Host "Wrote scripts to: $importPath"

# Execution de l'import (appel pwsh -File)
Write-Host "Launching import script..."
& pwsh -NoProfile -ExecutionPolicy Bypass -File $importPath -ModulePath $ModulePath -ExcelFile $ExcelFile

# Après import, suggestion: exécuter le script de compilation si vous l'avez :
Write-Host ""
Write-Host "Si l'import a réussi, lancez ensuite votre test-vbide-and-compile.ps1 pour vérifier la compilation:"
Write-Host " pwsh -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-vbide-and-compile.ps1 -WorkbookPath `"$ExcelFile`" -AttemptCompile"
� *cascade08��*cascade08�� *cascade08��*cascade08��? *cascade08"(9e377b08b88e30057baf56c074b3d12f1d22372326file:///c:/VBA/SGQ%201.65/scripts/agent-deploy-all.ps1:file:///c:/VBA/SGQ%201.65