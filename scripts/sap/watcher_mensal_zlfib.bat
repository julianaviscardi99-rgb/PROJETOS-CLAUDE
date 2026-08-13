@echo off
rem Roda sempre a partir da pasta onde este arquivo esta, nao importa de onde for chamado
cd /d "%~dp0..\.."

if not exist logs mkdir logs

python scripts\sap\verificacao_mensal_zlfib.py >> logs\zlfib_mensal.log 2>&1
