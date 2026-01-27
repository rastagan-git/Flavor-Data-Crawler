@echo off
cd /d "%~dp0"

:: IMPORTANT: If your folder is named 'venv', change 'myenv' to 'venv' below
if exist myenv\Scripts\activate.bat (
    call myenv\Scripts\activate.bat
) else (
    if exist venv\Scripts\activate.bat (
        call venv\Scripts\activate.bat
    ) else (
        echo Error: Virtual environment not found.
        echo Please check if the folder is named 'myenv' or 'venv'.
        pause
        exit /b
    )
)

echo Environment activated. Running Python script...
python nist_excel_tool.py

echo.
echo Done.
pause