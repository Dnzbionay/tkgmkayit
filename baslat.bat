@echo off
echo ========================================
echo   TKGM Arıza Kayıt Sistemi
echo ========================================
echo.
echo Sunucu başlatılıyor...
echo Tarayıcınız otomatik açılacak.
echo Kapatmak için Ctrl+C tuşlayın.
echo.

REM Tarayıcıyı aç (1 saniye bekle)
timeout /t 1 /nobreak > nul
start http://localhost:3000

REM Sunucuyu başlat
node server.js
