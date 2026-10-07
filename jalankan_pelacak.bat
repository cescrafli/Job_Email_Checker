@echo off
cd /d "%~dp0"
echo Menjalankan Pelacak Email...
python -u track_email.py
echo Proses selesai!
pause
