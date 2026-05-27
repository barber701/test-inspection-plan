@echo off
echo Installing dependencies...
pip install -r requirements.txt
pip install pyinstaller

echo.
echo Building standalone executable...
pyinstaller ^
  --onefile ^
  --windowed ^
  --name "TestInspectionPlan" ^
  --add-data "%LOCALAPPDATA%\Programs\Python\Python312\Lib\site-packages\customtkinter;customtkinter" ^
  tip_app.py

echo.
echo Done.  Executable is in:  dist\TestInspectionPlan.exe
echo Create a desktop shortcut pointing to that file.
pause
