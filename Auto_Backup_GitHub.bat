@echo off
setlocal

rem Roda sempre a partir da pasta onde este arquivo esta, nao importa de onde for chamado
cd /d "%~dp0"

echo ===============================
echo  Backup automatico - %date% %time%
echo ===============================

git add -A

git diff --cached --quiet
if %errorlevel%==0 (
    echo Nada novo para salvar. Projeto ja esta atualizado.
) else (
    git commit -m "Backup automatico %date% %time%"
    git push
    echo Backup enviado para o GitHub com sucesso.
)

endlocal
