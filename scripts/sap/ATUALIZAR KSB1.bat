@echo off
title Atualizar KSB1 - Fitted Units
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

python "%~dp0atualizar_ksb1_gui.py"
