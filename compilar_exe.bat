@echo off
echo ========================================================
echo COMPILADOR AUTOMATICO - SAP / CARGO HEROES
echo ========================================================
echo Este script ira gerar um executavel usando o PyInstaller.
echo Aguarde, isso pode levar alguns minutos...
echo.

python build.py

echo.
echo Processo finalizado! Pressione qualquer tecla para sair.
pause >nul
