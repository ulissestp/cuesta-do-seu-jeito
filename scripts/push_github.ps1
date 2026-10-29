# Envia main para um repositorio GitHub ja criado (vazio, sem README).
$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent
$repoName = "cuesta-do-seu-jeito"
$user = "ulissestp"

$git = "$env:LOCALAPPDATA\MinGit\cmd\git.exe"
if (-not (Test-Path $git)) {
    $git = (Get-Command git -ErrorAction SilentlyContinue).Source
}

Set-Location $root
$url = "https://github.com/$user/$repoName.git"

$remotes = & $git remote 2>$null
if ($remotes -contains "origin") {
    & $git remote set-url origin $url
} else {
    & $git remote add origin $url
}

Write-Host "Enviando para $url ..."
& "C:\Program Files\GitHub CLI\gh.exe" auth setup-git 2>$null | Out-Null
& $git push -u origin main
if ($LASTEXITCODE -ne 0) {
    Write-Host "Falha no push. Rode: gh auth setup-git"
    exit $LASTEXITCODE
}
Write-Host "OK: https://github.com/$user/$repoName"
