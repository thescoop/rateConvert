@echo off
REM Legal Aid Rate Converter GUI Launcher
REM Activates rateConvert conda environment and runs the GUI application

REM Activate the rateConvert conda environment
call C:\Users\thescoop\anaconda3\Scripts\activate.bat rateConvert

REM Run the GUI application without console window
start "" pythonw rateConvertGUI.py

REM Deactivate conda environment
call conda deactivate
