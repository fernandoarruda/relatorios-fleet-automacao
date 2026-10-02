@echo off
cd /d "%~dp0"
py -3 atualizar_evidencias.py
if errorlevel 1 goto fim
py -3 gerar_relatorios.py %*
:fim
pause
