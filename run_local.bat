@echo off
setlocal
cd /d "%~dp0"

if not exist .venv (
  echo Criando ambiente virtual...
  py -m venv .venv
)

call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt

if not exist .env (
  copy .env.example .env
  echo.
  echo Arquivo .env criado. Preencha SMARTSHEET_ACCESS_TOKEN e SMARTSHEET_SHEET_ID, ou XLSX_FILE_PATH.
  pause
)

python app.py
