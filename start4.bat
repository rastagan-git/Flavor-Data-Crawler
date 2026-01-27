@echo off
cd /d "%~dp0"

echo ==========================================
echo        ChemicalBook Crawler
echo ==========================================

:: 激活环境
if exist myenv\Scripts\activate.bat (
    call myenv\Scripts\activate.bat
) else if exist venv\Scripts\activate.bat (
    call venv\Scripts\activate.bat
) else (
    echo [Error] Virtual environment not found.
    pause
    exit /b
)

echo.
echo Checking dependencies...
:: 确保依赖都装了
pip install selenium webdriver-manager pandas openpyxl

echo.
echo Starting Python Script...
python cb_spider.py

echo.
echo Process Finished.
pause