Set-Location -Path $PSScriptRoot

if (-not (Test-Path ".venv")) {
    Write-Host "Criando ambiente virtual..."
    py -m venv .venv
}

. .\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt

if (-not (Test-Path ".env")) {
    Copy-Item ".env.example" ".env"
    Write-Host "Arquivo .env criado. Preencha SMARTSHEET_ACCESS_TOKEN e SMARTSHEET_SHEET_ID, ou XLSX_FILE_PATH."
    Read-Host "Pressione Enter para continuar"
}

python app.py
