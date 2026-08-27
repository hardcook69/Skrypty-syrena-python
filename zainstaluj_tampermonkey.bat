@echo off
chcp 65001 >nul
title Instalacja Tampermonkey

echo ============================================================
echo   Instalacja Tampermonkey
echo ============================================================
echo.
echo Za chwile otworzy sie Chrome Web Store.
echo Kliknij "Dodaj do Chrome" (pomin, jesli juz masz Tampermonkey).
echo.
pause

start "" "https://chromewebstore.google.com/detail/tampermonkey/dhdgffkkebhmkfjojejmpbldmpobfkfo"

echo.
echo ============================================================
echo Gotowe. Potwierdz instalacje w otwartej karcie.
echo ============================================================
pause
