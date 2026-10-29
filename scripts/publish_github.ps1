# Publica o projeto no GitHub (requer: gh auth login feito uma vez)
$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent

$git = "$env:LOCALAPPDATA\MinGit\cmd\git.exe"
$gh = (Get-ChildItem "$env:LOCALAPPDATA\GitHubCLI" -Recurse -Filter gh.exe -ErrorAction SilentlyContinue |
    Select-Object -ExpandProperty FullName -First 1)

if (-not (Test-Path $git)) {
    Write-Host "Git nao encontrado. Instale Git for Windows ou rode o assistente do Cursor novamente."
    exit 1
}
if (-not $gh) {
    Write-Host "GitHub CLI (gh) nao encontrado. Instale: winget install GitHub.cli"
    exit 1
}

Set-Location $root
& $gh auth status 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-Host "Faca login no GitHub:"
    & $gh auth login
}

$repoName = "roteiro-turismo"
if (-not (Test-Path ".git")) {
    & $git init -b main
}

& $git add -A
$status = & $git status --porcelain
if ($status) {
    & $git commit -m @"
Publica o gerador de roteiro turistico (desktop + Streamlit).

Inclui motor de roteiro, exportacao PDF e importacao de inventario Excel.
"@
}

$remotes = & $git remote 2>$null
if ($remotes -notcontains "origin") {
    & $gh repo create $repoName --public --source=. --remote=origin --push
} else {
    & $git push -u origin main
}

Write-Host "Pronto. Abra: https://github.com/$(& $gh api user -q .login)/$repoName"
