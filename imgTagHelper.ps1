<#
DEVEN WEEDEN 2026
- GET : grabs tags off the images in a folder and writes them to a txt file.
- SET : reads that txt file and writes the tags onto the images in a folder.
#>
param(
    [string]$Mode,
    [string]$Path,
    [string]$TagFile = "$PSScriptRoot\tags.txt"
)

while ($Mode -notin 'Get','Set') {
    $Mode = (Read-Host 'GET tags off images, or SET tags onto them? [get/set]').Trim()
}
if (-not $Path) { $Path = Read-Host "Paste folder path here" }
$Path = $Path.Trim().Trim('"')

$files = Get-ChildItem -LiteralPath $Path -File |
    Where-Object { $_.Extension -in '.jpg','.jpeg','.png' } |
    Sort-Object { [regex]::Replace($_.Name, '\d+', { $args[0].Value.PadLeft(20,'0') }) }
if (-not $files) { Write-Error "NO IMAGES"; return }

# GET
if ($Mode -eq 'Get') {
    $folder = (New-Object -ComObject Shell.Application).Namespace($Path)
    $tags = foreach ($file in $files) {
        $folder.GetDetailsOf($folder.ParseName($file.Name), 18)
    }
    Set-Content -LiteralPath $TagFile -Value $tags -Encoding UTF8
    Write-Host "Wrote $($files.Count) lines to $TagFile" -ForegroundColor Green
    return
}

# SET
$exif  = "$PSScriptRoot\exiftool.exe"
$lines = @(Get-Content -LiteralPath $TagFile -Encoding UTF8)
if ($lines.Count -ne $files.Count) {
    Write-Error "COUNT MISMATCH - $($lines.Count) lines in the txt, $($files.Count) images in the folder"
    return
}

$ok = 0
for ($i = 0; $i -lt $files.Count; $i++) {
    $tags = $lines[$i].Trim()
    if (-not $tags) { continue }
    & $exif -sep '; ' "-XMP:Subject=$tags" "-IPTC:Keywords=$tags" "-XPKeywords=$tags" `
        -overwrite_original_in_place -q $files[$i].FullName
    if ($LASTEXITCODE -eq 0) { $ok++ }
}
Write-Host "Tagged $ok images" -ForegroundColor Green
