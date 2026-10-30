$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
Set-Location $ProjectRoot

Write-Host "Projeto: $ProjectRoot"
Write-Host "Instalando dependencias..."
pip install -r requirements.txt pyinstaller --quiet

Write-Host "Gerando icone (.ico) a partir do logo..."
python -c "from src.app_icon import ensure_app_icon; p=ensure_app_icon(); print('Icone:', p or 'logo nao encontrado')"

Write-Host "Gerando executavel..."
pyinstaller "$ProjectRoot\RoteiroTurismo.spec" --noconfirm --workpath "$ProjectRoot\build" --distpath "$ProjectRoot\dist"

$exe = Join-Path $ProjectRoot "dist\RoteiroTurismo.exe"
if (Test-Path $exe) {
    Write-Host "Pronto: $exe"
} else {
    Write-Error "Falha ao gerar o executavel."
}
