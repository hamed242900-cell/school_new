@echo off
chcp 65001 >nul
cd /d "%~dp0app"
where py >nul 2>nul && set PY=py || set PY=python
%PY% -m venv .venv
call .venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
start "" http://127.0.0.1:5000
python app.py
pause
