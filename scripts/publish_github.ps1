# Publica o projeto no GitHub (requer: gh auth login feito uma vez)
$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent

$git = "$env:LOCALAPPDATA\MinGit\cmd\git.exe"
if (-not (Test-Path $git)) {
    $gitCmd = Get-Command git -ErrorAction SilentlyContinue
    if ($gitCmd) { $git = $gitCmd.Source }
}

$gh = $null
$systemGh = "C:\Program Files\GitHub CLI\gh.exe"
if (Test-Path $systemGh) { $gh = $systemGh }
if (-not $gh) {
    $ghCmd = Get-Command gh -ErrorAction SilentlyContinue
    if ($ghCmd) { $gh = $ghCmd.Source }
}
if (-not $gh) {
    $gh = (Get-ChildItem "$env:LOCALAPPDATA\GitHubCLI" -Recurse -Filter gh.exe -ErrorAction SilentlyContinue |
        Select-Object -ExpandProperty FullName -First 1)
}

if (-not (Test-Path $git)) {
    Write-Host "Git nao encontrado. Instale Git for Windows ou rode o assistente do Cursor novamente."
    exit 1
}
if (-not $gh) {
    Write-Host "GitHub CLI (gh) nao encontrado. Instale: winget install GitHub.cli"
    exit 1
}

Set-Location $root

$prevEap = $ErrorActionPreference
$ErrorActionPreference = "SilentlyContinue"
& $gh auth status 2>&1 | Out-Null
$authOk = ($LASTEXITCODE -eq 0)
$ErrorActionPreference = $prevEap
if (-not $authOk) {
    Write-Host ""
    Write-Host "GitHub: faca login uma vez no terminal (abre o navegador):"
    Write-Host "  & `"$gh`" auth login"
    Write-Host ""
    Write-Host "Depois rode este script de novo."
    exit 1
}

$repoName = "cuesta-do-seu-jeito"
$repoDescription = "CUESTA DO SEU JEITO - gerador de roteiro turistico (desktop + Streamlit)"
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
    & $gh repo create $repoName --public --description $repoDescription --source=. --remote=origin --push
} else {
    & $git push -u origin main
}

Write-Host "Pronto. Abra: https://github.com/$(& $gh api user -q .login)/$repoName"
