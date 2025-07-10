@echo off
pyinstaller BoMinationApp.py ^
  --noconsole ^
  --add-data "Files\OCTF-1539-COST SHEET.xlsx;Files" ^
  --add-data "chromedriver.exe;." ^
  --clean
echo Build complete!
pause