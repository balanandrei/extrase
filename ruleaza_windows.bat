@echo off
REM Dublu-click pe acest fisier ruleaza extragerea folosind config.toml de alaturi.
REM Trebuie ca Python 3 sa fie instalat (vezi README.md).

cd /d "%~dp0"

python extract_extrase.py config.toml
if errorlevel 1 (
    echo.
    echo *** A aparut o eroare. Citeste mesajele de mai sus. ***
) else (
    echo.
    echo *** Gata! Verifica folderul de output. ***
)

echo.
pause
