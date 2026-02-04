ï
$filesVals = @(
    "vba-files\Module\modSGQInterface.bas",
    "vba-files\Module\modSGQViews.bas"
)

foreach ($f in $filesVals) {
    if (Test-Path $f) {
        $path = (Resolve-Path $f).Path
        try {
            $text = [System.IO.File]::ReadAllText($path) # Auto-detects UTF-8 usually
            [System.IO.File]::WriteAllText($path, $text, [System.Text.Encoding]::GetEncoding(1252))
            Write-Host "Fixed encoding for: $path"
        }
        catch {
            Write-Error "Failed to fix $path : $_"
        }
    }
    else {
        Write-Warning "File not found: $f"
    }
}
ï*cascade08"(9e377b08b88e30057baf56c074b3d12f1d2237232@file:///c:/VBA/SGQ%201.65/scripts/temp_fix_critical_encoding.ps1:file:///c:/VBA/SGQ%201.65