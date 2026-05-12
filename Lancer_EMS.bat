@echo off
title EMS - Emeraude Moteurs Systemes
cd /d "%~dp0"
python main.py
if errorlevel 1 (
    echo.
    echo ERREUR : Python n'est pas installe ou introuvable.
    echo Telechargez Python sur https://www.python.org
    pause
)
