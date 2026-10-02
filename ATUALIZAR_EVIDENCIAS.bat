@echo off
cd /d "%~dp0"
py -3 atualizar_evidencias.py
if errorlevel 1 (
  echo.
  echo Nao foi possivel atualizar. Confira a mensagem acima.
)
pause
