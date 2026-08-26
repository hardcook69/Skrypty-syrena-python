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
REM Zeby ten link zadzialal, plik Wyzwalacz.user.js musi byc naprawde
REM w repo hardcook69/Skrypty-syrena-python na branchu main (na razie
REM tam nie jest - git push sie nie udaje, patrz uwaga w wiadomosci).
REM Zeby dodac kolejny skrypt: skopiuj linie "start "" chrome ..." ponizej
REM i podmien adres na link do kolejnego pliku .user.js.

start "" chrome "https://raw.githubusercontent.com/hardcook69/Skrypty-syrena-python/main/Wyzwalacz.user.js"

echo.
echo ============================================================
echo Gotowe. Potwierdz instalacje w kazdej otwartej karcie.
echo ============================================================
pause
