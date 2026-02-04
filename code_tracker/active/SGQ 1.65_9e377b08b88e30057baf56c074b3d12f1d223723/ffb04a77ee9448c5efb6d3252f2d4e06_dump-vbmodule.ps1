•param(
    [string]$WorkbookPath = 'C:\VBA\SGQ 1.65\index.xlsm',
    [string]$ComponentName = 'Sheet6101',
    [int]$ContextLines = 6
)

$excel = New-Object -ComObject Excel.Application
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($WorkbookPath)
try {
    $vbComp = $wb.VBProject.VBComponents.Item($ComponentName)
    $cm = $vbComp.CodeModule
    $count = $cm.CountOfLines
    Write-Host "Component: $ComponentName | Lines: $count"
    # Print entire module or with context - here we print all lines
    $all = $cm.Lines(1, $count)
    $ln = 1
    foreach ($line in $all -split "`r?`n") {
        Write-Host ('[{0}] {1}' -f $ln, $line)
        $ln++
    }
}
catch {
    Write-Host 'Error: ' $_.Exception.Message
}
finally {
    if ($wb) { $wb.Close($false) }
    if ($excel) { $excel.Quit() ; [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
}
•*cascade08"(9e377b08b88e30057baf56c074b3d12f1d22372323file:///c:/VBA/SGQ%201.65/scripts/dump-vbmodule.ps1:file:///c:/VBA/SGQ%201.65