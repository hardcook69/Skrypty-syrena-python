@echo off
chcp 65001 >nul
title Instalacja - Eksport spotkan do Excela

echo ============================================================
echo   Instalacja: Eksport spotkan do Excela (Tampermonkey)
echo ============================================================
echo.
echo Wymaga juz zainstalowanego Tampermonkey (patrz
echo zainstaluj_tampermonkey.bat).
echo.
echo Za chwile otworzy sie karta ze skryptem .user.js -
echo Tampermonkey pokaze okno "Zainstaluj", kliknij "Zainstaluj".
echo.
pause

REM UWAGA: wskazuje na plik w repo Skrypty-syrena-python (gdzie na pewno
REM istnieje juz teraz), NIE na Syrena-Tempermokey (tam trzeba by go
REM najpierw recznie wgrac -- to byl powod, ze poprzednia wersja tego
REM bata nie dzialala, plik pod tamtym adresem nie istnial / 404).
start "" "https://raw.githubusercontent.com/hardcook69/Skrypty-syrena-python/claude/skills-package-install-nbqe08/EksportSpotkanSkrot.user.js"

echo.
echo ============================================================
echo Gotowe. Potwierdz instalacje w otwartej karcie.
echo Przycisk "Eksport spotkan do Excela" pojawi sie w rejestrze
echo spotkan (strona /meetings), obok przycisku "Generuj raport".
echo ============================================================
pause
