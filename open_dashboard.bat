@echo off
cd /d "%~dp0"
call %USERPROFILE%\anaconda3\Scripts\activate.bat %USERPROFILE%\anaconda3\a33
python scripts\run_pipeline.py --live
start "" "Portafogliov4.xlsm"
