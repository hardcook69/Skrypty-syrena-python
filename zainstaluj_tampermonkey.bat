@echo off
chcp 65001 >nul
title Instalacja Tampermonkey + skrypty Syrena

echo ============================================================
echo   Instalacja Tampermonkey + skrypty Syrena
echo ============================================================
echo.
echo Za chwile otworza sie kolejno karty w Chrome:
echo   1. Strona Tampermonkey w Chrome Web Store
echo        -^> kliknij "Dodaj do Chrome" (pomin, jesli juz masz)
echo   2. Kazdy skrypt .user.js
echo        -^> Tampermonkey pokaze okno "Zainstaluj" - kliknij "Zainstaluj"
echo.
echo Kazdy krok wymaga jednego kliknieca w oknie potwierdzenia -
echo to zabezpieczenie przegladarki, ktorego nie da sie ominac.
echo.
pause

REM --- 1. Tampermonkey (Chrome Web Store) ---
start "" chrome "https://chromewebstore.google.com/detail/tampermonkey/dhdgffkkebhmkfjojejmpbldmpobfkfo"

timeout /t 4 /nobreak >nul

REM --- 2. Skrypty .user.js ---
REM PODMIEN ponizsze adresy na prawdziwe linki do Twoich plikow .user.js
REM (np. surowy link z GitHuba: https://raw.githubusercontent.com/<uzytkownik>/<repo>/main/<plik>.user.js)
REM Skopiuj linie "start "" chrome ..." i dodaj kolejne dla kazdego skryptu.

start "" chrome "https://raw.githubusercontent.com/TWOJ-UZYTKOWNIK/TWOJE-REPO/main/skrypt1.user.js"
timeout /t 2 /nobreak >nul
start "" chrome "https://raw.githubusercontent.com/TWOJ-UZYTKOWNIK/TWOJE-REPO/main/skrypt2.user.js"

echo.
echo ============================================================
echo Gotowe. Potwierdz instalacje w kazdej otwartej karcie.
echo ============================================================
pause
