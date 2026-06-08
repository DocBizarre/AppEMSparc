@echo off
REM ─────────────────────────────────────────────────────────────────────
REM  EMS API — Démarrage du serveur (Windows)
REM
REM  Place ce fichier à la racine du projet et double-clique pour
REM  démarrer l'API. Une fenêtre noire s'ouvre — la laisser ouverte.
REM
REM  L'API sera accessible sur :
REM    http://127.0.0.1:8765       (uniquement depuis ce PC)
REM  Pour accès réseau, change EMS_API_HOST en 0.0.0.0
REM ─────────────────────────────────────────────────────────────────────

cd /d "%~dp0"

REM Configuration
set EMS_API_HOST=127.0.0.1
set EMS_API_PORT=8765
REM set EMS_API_KEY=cle-secrete-pour-securiser-l-acces

echo.
echo ============================================================
echo   EMS API — Demarrage du serveur
echo   URL  : http://%EMS_API_HOST%:%EMS_API_PORT%
echo   Docs : http://%EMS_API_HOST%:%EMS_API_PORT%/docs
echo   Ctrl+C pour arreter
echo ============================================================
echo.

python -m uvicorn ems_api.main:app --host %EMS_API_HOST% --port %EMS_API_PORT%

pause
