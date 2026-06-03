@echo off
setlocal

set "app_name=RnSApp"

pyInstaller updater.py -n updater --onefile --noconsole --windowed -y
pyInstaller main.py -n %app_name% --onedir --icon=".\assets\rns-logo-alt.ico" --noconsole --windowed -y
mkdir .\dist\%app_name%\assets
copy .\assets\* .\dist\%app_name%\assets\
copy .\dist\updater.exe .\dist\%app_name%\updater.exe
