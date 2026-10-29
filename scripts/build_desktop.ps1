$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
Set-Location $ProjectRoot

Write-Host "Projeto: $ProjectRoot"
Write-Host "Instalando dependencias..."
pip install -r requirements.txt pyinstaller --quiet

Write-Host "Gerando executavel..."
pyinstaller "$ProjectRoot\RoteiroTurismo.spec" --noconfirm --workpath "$ProjectRoot\build" --distpath "$ProjectRoot\dist"

$exe = Join-Path $ProjectRoot "dist\RoteiroTurismo.exe"
if (Test-Path $exe) {
    Write-Host "Pronto: $exe"
} else {
    Write-Error "Falha ao gerar o executavel."
}
