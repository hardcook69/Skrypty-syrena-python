@echo off
chcp 65001 >nul
title Instalacja Tampermonkey

echo ============================================================
echo   Instalacja Tampermonkey
echo ============================================================
echo.
echo Za chwile otworza sie dwie karty - wybierz sklep pasujacy
echo do Twojej przegladarki (Chrome Web Store lub Firefox Add-ons)
echo i kliknij "Dodaj" (pomin, jesli juz masz Tampermonkey).
echo.
pause

start "" "https://chromewebstore.google.com/detail/tampermonkey/dhdgffkkebhmkfjojejmpbldmpobfkfo"
start "" "https://addons.mozilla.org/firefox/addon/tampermonkey/"

echo.
echo ============================================================
echo Gotowe. Potwierdz instalacje w karcie dla Twojej przegladarki
echo (druga karta - z niepasujacego sklepu - mozna zamknac).
echo ============================================================
pause
