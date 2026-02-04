à
# Ensures all .bas/.cls/.frm files are encoded in Windows-1252
# This script is aggressive: it forces conversion to support VBA import requirements.

param(
  [string]$SearchPath = "vba-files"
)

$root = (Resolve-Path $SearchPath).Path
if (-not (Test-Path $root)) { exit 0 }

$extensions = @("*.bas", "*.cls", "*.frm")
$files = Get-ChildItem -Path $root -Recurse -Include $extensions -File

$enc1252 = [System.Text.Encoding]::GetEncoding(1252)
$utf8 = [System.Text.Encoding]::UTF8

foreach ($f in $files) {
  try {
    # 1. Read bytes to detect signature
    $bytes = [System.IO.File]::ReadAllBytes($f.FullName)
        
    # 2. Check for UTF-8 BOM
    $hasUtf8Bom = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
        
    # 3. Reading logic
    $text = $null
    $needsRewrite = $false
        
    if ($hasUtf8Bom) {
      # Definitely UTF-8. Read as such, will write as 1252.
      $text = [System.IO.StreamReader]::new($f.FullName, $utf8).ReadToEnd()
      $needsRewrite = $true
    }
    else {
      # No BOM. Could be ANSI or UTF-8 no-BOM.
      # Try parsing as UTF-8. If it's valid UTF-8 and differs from ANSI interpretation, we treat as UTF-8.
      try {
        # Strict decoding to fail if invalid UTF-8 sequences found
        $strictUtf8 = [System.Text.UTF8Encoding]::new($false, $true)
        $text = $strictUtf8.GetString($bytes)
                
        # If we are here, it matches UTF-8 structure.
        # Check if it contains characters that would change in strict 1252
        # (Simple heuristic: Just rewrite it as 1252 to ensure consistency)
        # But wait, if it was ALREADY 1252 with accents, strict UTF-8 might accidentally fail or succeed?
        # CP1252 byte 0xE9 ('Ã©') is invalid in UTF-8 usually. 
        # So if strict UTF-8 succeeds, it is extremely likely to be UTF-8 (or pure ASCII).
        # In both cases (UTF-8 or ASCII), we can safely write as 1252.
                
        # Check if file has mixed content? No, just rewrite.
        # Only rewrite if it's NOT just ASCII (optimization) or if we want to enforce.
        # Let's enforce.
        $needsRewrite = $true
      }
      catch {
        # Invalid UTF-8. Must be ANSI/CP1252 (or binary garbage).
        # Leave it alone? Or Ensure it is 1252?
        # If we read as 1252 and write as 1252, it's a no-op, but safe.
        # Let's Skip if invalid UTF-8, assuming it's already correct ANSI.
        $needsRewrite = $false 
      }
    }

    if ($needsRewrite) {
      # Debug: Write-Host "Normalizing to CP1252: $($f.Name)"
      [System.IO.File]::WriteAllText($f.FullName, $text, $enc1252)
    }
  }
  catch {
    Write-Warning "Could not process $($f.Name): $_"
  }
}
Write-Host "Encoding normalization complete."
à*cascade08"(9e377b08b88e30057baf56c074b3d12f1d22372325file:///c:/VBA/SGQ%201.65/scripts/ensure-encoding.ps1:file:///c:/VBA/SGQ%201.65