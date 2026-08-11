@echo off
title Diagnostico do popup SAP
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo Python nao foi encontrado neste computador.
    echo Instale o Python em https://www.python.org/downloads/ ^(marque "Add to PATH" na instalacao^)
    echo e rode este atalho de novo.
    pause
    exit /b 1
)

python -c "import win32com" >nul 2>nul
if errorlevel 1 (
    echo Instalando dependencia necessaria ^(pywin32^)...
    python -m pip install --quiet pywin32
)

python "%~dp0diagnosticar_popup.py"

echo.
echo ============================================================
echo Copie todo o texto acima (desde "Elementos dentro do popup")
echo e cole na conversa com o Claude.
echo ============================================================
pause
