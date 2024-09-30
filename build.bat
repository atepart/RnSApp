@echo off
setlocal

set "app_name=RnSApp"

pyInstaller main.py -n %app_name% --onedir --icon=".\assets\rns-logo-med.png" --noconsole --windowed -y
mkdir .\dist\%app_name%\assets
copy .\assets\* .\dist\%app_name%\assets\
