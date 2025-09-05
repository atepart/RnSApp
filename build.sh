#! /bin/bash

pyInstaller main.py -n RnSApp --onedir --icon=assets/rns-logo-alt.ico --noconsole -y --add-data="assets:assets"
