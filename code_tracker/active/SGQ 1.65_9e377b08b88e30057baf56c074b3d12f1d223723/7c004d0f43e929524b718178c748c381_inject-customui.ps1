€'<#
.SYNOPSIS
  Injects or updates the customUI14.xml into an Excel workbook's OpenXML structure.

.DESCRIPTION
  This script treats the .xlsm file as a ZIP archive. It:
  1. Creates a backup of the workbook.
  2. Injects 'customUI/customUI14.xml' from the source file.
  3. Updates '_rels/.rels' to ensure the UI extensibility relationship exists.
  
.PARAMETER WorkbookPath
  Path to the target .xlsm file.

.PARAMETER XmlPath
  Path to the customUI14.xml source file.

.EXAMPLE
  .\inject-customui.ps1 -WorkbookPath "..\index.xlsm" -XmlPath "..\ribbons\customUI14.xml"
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,
    
    [Parameter(Mandatory = $true)]
    [string]$XmlPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression.FileSystem

function Write-Info($msg) { Write-Host "[INFO] $msg" -ForegroundColor Green }
function Write-Warn($msg) { Write-Warning $msg }

$WorkbookPath = (Resolve-Path $WorkbookPath).Path
$XmlPath = (Resolve-Path $XmlPath).Path

if (-not (Test-Path $WorkbookPath)) { throw "Workbook not found: $WorkbookPath" }
if (-not (Test-Path $XmlPath)) { throw "XML file not found: $XmlPath" }

# Backup
$backupPath = "$WorkbookPath.bak"
Copy-Item -LiteralPath $WorkbookPath -Destination $backupPath -Force
Write-Info "Backup created at: $backupPath"

try {
    # Open as Zip
    $zip = [System.IO.Compression.ZipFile]::Open($WorkbookPath, [System.IO.Compression.ZipArchiveMode]::Update)
    
    # 1. Update/Add customUI14.xml
    $entryName = "customUI/customUI14.xml"
    $entry = $zip.GetEntry($entryName)
    if ($entry) {
        $entry.Delete()
        Write-Info "Removed existing $entryName"
    }
    else {
        # Ensure customUI folder technically exists in path implies it just works in Zip
    }
    
    $newEntry = $zip.CreateEntry($entryName)
    $entryStream = $newEntry.Open()
    $fileStream = [System.IO.File]::OpenRead($XmlPath)
    $fileStream.CopyTo($entryStream)
    $fileStream.Close()
    $entryStream.Close()
    
    Write-Info "Injecting $entryName from $XmlPath"
    
    # 2. Update Relationships (_rels/.rels)
    $relsEntryName = "_rels/.rels"
    $relsEntry = $zip.GetEntry($relsEntryName)
    
    if (-not $relsEntry) {
        Write-Warn "Could not find $_rels/.rels. Is this a valid OpenXML file?"
    }
    else {
        # Read .rels
        $stream = $relsEntry.Open()
        $reader = New-Object System.IO.StreamReader($stream)
        $xmlContent = $reader.ReadToEnd()
        $reader.Close()
        $stream.Close()
        
        $xmlDoc = [xml]$xmlContent
        $nsManager = New-Object System.Xml.XmlNamespaceManager($xmlDoc.NameTable)
        $nsManager.AddNamespace("r", "http://schemas.openxmlformats.org/package/2006/relationships")
        
        # Check if relationship exists
        # Schema matching can be tricky with default namespaces, using local-name() is robust
        $relType = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
        $nodes = $xmlDoc.SelectNodes("//*[local-name()='Relationship']", $nsManager)
        
        $found = $false
        foreach ($node in $nodes) {
            if ($node.Type -eq $relType) {
                $found = $true
                # Ensure Target is correct
                if ($node.Target -ne "customUI/customUI14.xml") {
                    $node.Target = "customUI/customUI14.xml"
                    Write-Info "Updated existing relationship target."
                }
                break
            }
        }
        
        if (-not $found) {
            Write-Info "Adding new Relationship for customUI."
            $relationships = $xmlDoc.DocumentElement
            $newRel = $xmlDoc.CreateElement("Relationship", "http://schemas.openxmlformats.org/package/2006/relationships")
            
            # Generate new ID (R + random hex or simple increment)
            $newId = "R" + [Guid]::NewGuid().ToString("N").Substring(0, 8)
            $newRel.SetAttribute("Id", $newId)
            $newRel.SetAttribute("Type", $relType)
            $newRel.SetAttribute("Target", "customUI/customUI14.xml")
            
            $relationships.AppendChild($newRel) | Out-Null
        }
        
        # Write back .rels
        $relsEntry.Delete()
        $newRelsEntry = $zip.CreateEntry($relsEntryName)
        $writer = New-Object System.IO.StreamWriter($newRelsEntry.Open())
        $xmlDoc.Save($writer)
        $writer.Close()
        Write-Info "Updated $relsEntryName"
    }
    
    $zip.Dispose()
    Write-Info "Injection complete. Please reload Excel to see changes."
}
catch {
    Write-Error "Failed to inject XML: $_"
    # Restore backup
    Copy-Item -LiteralPath $backupPath -Destination $WorkbookPath -Force
    Write-Warn "Restored backup due to failure."
    throw
}
€'"(9e377b08b88e30057baf56c074b3d12f1d22372325file:///c:/VBA/SGQ%201.65/scripts/inject-customui.ps1:file:///c:/VBA/SGQ%201.65