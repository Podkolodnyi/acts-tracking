@echo off
cd /d D:\123
python -m waitress --listen=0.0.0.0:5000 app:app