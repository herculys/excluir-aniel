# Script para executar o programa usando apenas o ambiente virtual (venv)
# Muda para o diretório do script
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ScriptDir

Write-Host "Ativando ambiente virtual..." -ForegroundColor Yellow

# Ativar ambiente virtual
try {
    & .\venv\Scripts\Activate.ps1
    Write-Host "Ambiente virtual ativado com sucesso!" -ForegroundColor Green
} catch {
    Write-Host "Erro ao ativar ambiente virtual: $_" -ForegroundColor Red
    Read-Host "Pressione Enter para sair"
    exit 1
}

Write-Host "Executando Exclusao.py..." -ForegroundColor Cyan

# Executar script usando Python do venv
try {
    & .\venv\Scripts\python.exe Exclusao.py
} catch {
    Write-Host "Erro ao executar o script: $_" -ForegroundColor Red
}

Write-Host "`nExecução concluída!" -ForegroundColor Green
Write-Host "Pressione qualquer tecla para fechar..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")