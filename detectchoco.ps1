# Verifica se o Chocolatey está instalado
$chocoPath = "$env:ProgramData\chocolatey\choco.exe"
if (Test-Path $chocoPath) {
    Write-Host "Chocolatey instalado."
    exit 0
} else {
    exit 1
}