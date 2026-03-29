# Crea un ZIP de distribució (sense node_modules)
$source  = "C:\HTML_in_Excel"
$output  = "$env:USERPROFILE\Desktop\RichCell_install.zip"

$exclude = @("node_modules", ".git")

$tmpDir = "$env:TEMP\RichCell_dist"
if (Test-Path $tmpDir) { Remove-Item $tmpDir -Recurse -Force }
New-Item -ItemType Directory -Path $tmpDir | Out-Null

Get-ChildItem -Path $source | Where-Object { $_.Name -notin $exclude } | ForEach-Object {
    Copy-Item -Path $_.FullName -Destination $tmpDir -Recurse -Force
}

if (Test-Path $output) { Remove-Item $output -Force }
Compress-Archive -Path "$tmpDir\*" -DestinationPath $output

Remove-Item $tmpDir -Recurse -Force

Write-Host ""
Write-Host "ZIP creat a: $output" -ForegroundColor Green
Write-Host "Mida: $([math]::Round((Get-Item $output).Length / 1MB, 1)) MB" -ForegroundColor Cyan
