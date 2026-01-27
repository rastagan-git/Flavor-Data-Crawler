@echo off
cd /d "%~dp0"

echo ==========================================
echo        MFFI Database Spider (Selenium)
echo ==========================================

:: 激活虚拟环境
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
echo Checking and installing dependencies...
:: 自动安装 selenium 依赖 (第一次运行时需要)
pip install selenium webdriver-manager pandas openpyxl

echo.
echo Starting Spider...
python mffi_spider.py

echo.
echo Process Finished.
pause