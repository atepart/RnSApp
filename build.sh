#! /bin/bash

pyInstaller main.py -n RnSApp --onedir --icon=assets/rns-logo-med.png --noconsole -y --add-data="assets:assets"
