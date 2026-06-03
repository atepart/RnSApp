#! /bin/bash

pyInstaller updater.py -n updater --onefile --windowed --noconsole -y
pyInstaller main.py -n RnSApp --onedir --icon=assets/rns-logo-alt.ico --noconsole -y --add-data="assets:assets"

if [ -d "dist/RnSApp.app" ]; then
    cp dist/updater.app/Contents/MacOS/updater dist/RnSApp.app/Contents/MacOS/updater || cp dist/updater dist/RnSApp.app/Contents/MacOS/updater
else
    cp dist/updater dist/RnSApp/updater || cp dist/updater.app/Contents/MacOS/updater dist/RnSApp/updater
fi
