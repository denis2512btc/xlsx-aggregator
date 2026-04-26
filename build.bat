@echo off
call .venv\Scripts\activate
set ICON_ARG=
if exist assets\icon.ico (
  set ICON_ARG=--icon assets\icon.ico
)
pyinstaller --noconfirm --windowed --name "XLSX Aggregator" ^
  %ICON_ARG% ^
  --collect-all customtkinter ^
  src\main.py
echo.
echo Result: dist\XLSX Aggregator\XLSX Aggregator.exe
