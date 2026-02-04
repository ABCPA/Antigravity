«param(
    [string]$Path = "c:\VBA\SGQ 1.65\mcp.json"
)

Write-Host "Validating MCP Configuration at: $Path" -ForegroundColor Cyan

if (-not (Test-Path $Path)) {
    Write-Error "File not found: $Path"
    exit 1
}

try {
    $content = Get-Content -Raw -Path $Path
    $json = $content | ConvertFrom-Json
    
    if (-not $json.mcpServers) {
        Write-Warning "Key 'mcpServers' missing in JSON."
        exit 2
    }

    $servers = $json.mcpServers | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
    Write-Host "Found servers:" -ForegroundColor Green
    foreach ($s in $servers) {
        Write-Host " - $s"
    }
    
    Write-Host "Validation SUCCESS" -ForegroundColor Green
}
catch {
    Write-Error "JSON Parsing Failed: $_"
    exit 3
}
«*cascade08"(d5ad6e0eb0dae28f888b626c904ac04ceba4e28329file:///c:/VBA/SGQ%201.65/scripts/validate-mcp-config.ps1:file:///c:/VBA/SGQ%201.65