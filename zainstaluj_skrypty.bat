@echo off
chcp 65001 >nul
title Instalacja skryptow Tampermonkey - Syrena

echo ============================================================
echo   Instalacja skryptow Tampermonkey - Syrena
echo ============================================================
echo.
echo Wymaga juz zainstalowanego Tampermonkey (patrz
echo zainstaluj_tampermonkey.bat).
echo.
echo Za chwile otworza sie karty z kazdym skryptem .user.js -
echo Tampermonkey pokaze okno "Zainstaluj", kliknij "Zainstaluj"
echo w kazdej karcie.
echo.
pause

REM Zeby dodac kolejny skrypt: skopiuj linie "start "" chrome ..." ponizej
REM i podmien adres na link do kolejnego pliku .user.js w repo
REM hardcook69/Syrena-Tempermokey.

start "" "https://raw.githubusercontent.com/hardcook69/Syrena-Tempermokey/main/Wyzwalacz.user.js"
start "" "https://raw.githubusercontent.com/hardcook69/Syrena-Tempermokey/main/ObserwacjeSkrot.user.js"

echo.
echo ============================================================
echo Gotowe. Potwierdz instalacje w kazdej otwartej karcie.
echo ============================================================
pause
