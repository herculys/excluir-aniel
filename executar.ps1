# Script para executar diretamente o programa principal
# Change to script directory (same folder as this script)
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ScriptDir

# Activate virtual environment
& .\venv\Scripts\Activate.ps1

# Run the main script
& python Exclusao.py

# Keep the window open to see results
Write-Host "Pressione qualquer tecla para continuar..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")