@echo off
REM Script para iniciar o servidor Flask automaticamente

REM Ativar o ambiente virtual
call venv\Scripts\activate.bat

REM Iniciar o servidor Flask
python app.py --host=0.0.0.0

REM Manter a janela aberta após o término
pause
