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
REM Zeby dodac kolejny skrypt: skopiuj linie "start "" chrome ..." ponizej
REM i podmien adres na link do kolejnego pliku .user.js w repo
REM hardcook69/Syrena-Tempermokey.

start "" chrome "https://raw.githubusercontent.com/hardcook69/Syrena-Tempermokey/main/Wyzwalacz.user.js"

echo.
echo ============================================================
echo Gotowe. Potwierdz instalacje w kazdej otwartej karcie.
echo ============================================================
pause
