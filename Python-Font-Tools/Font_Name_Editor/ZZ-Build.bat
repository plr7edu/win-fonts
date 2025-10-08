@echo off

REM Define ESC character for ANSI colors
for /F %%a in ('echo prompt $E ^| cmd') do set "ESC=%%a"

echo ========================================
echo Building Font Name Editor
echo ========================================
echo.

pyinstaller --onefile --windowed --name=Font_Name_Editor --icon=icon/font_name_editor.ico --add-data=icon;icon --clean Font_Name_Editor.pyw

echo.
echo %ESC%[92m========================================
echo Build Complete!
echo ========================================%ESC%[0m
echo.
pause