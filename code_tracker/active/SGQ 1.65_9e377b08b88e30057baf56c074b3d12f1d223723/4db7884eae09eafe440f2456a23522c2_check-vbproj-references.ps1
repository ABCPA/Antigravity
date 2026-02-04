¿param(
    [string]$WorkbookPath = "C:\VBA\SGQ 1.65\6-Suivi annuel - Gervais Boily CPA V. 24-12.xlsm"
)

$ErrorActionPreference = 'Stop'

Write-Host "Checking VBProject references for: $WorkbookPath"
$excel = New-Object -ComObject Excel.Application
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($WorkbookPath)

foreach ($r in $wb.VBProject.References) {
    $isBroken = $false
    try { $isBroken = $r.IsBroken } catch { $isBroken = 'N/A' }
    Write-Host ("Name: {0} | GUID: {1} | FullPath: {2} | IsBroken: {3}" -f $r.Name, $r.GUID, $r.FullPath, $isBroken)
}

$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
¿*cascade08"(9e377b08b88e30057baf56c074b3d12f1d2237232=file:///c:/VBA/SGQ%201.65/scripts/check-vbproj-references.ps1:file:///c:/VBA/SGQ%201.65