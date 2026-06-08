@echo off
cd /d "%~dp0"
py -3 -m pip install --upgrade pyinstaller
py -3 -m PyInstaller --onefile --windowed --name "Username-Shuffler" --icon "icon.ico" --add-data "icon.png;." --add-data "icon.ico;." --add-data "titlebar.ico;." Username-Shuffler.pyw
pause
