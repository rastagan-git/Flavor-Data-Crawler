@echo off
cd /d "%~dp0"

echo [STEP 1] Activating Virtual Environment...

:: Try to activate 'myenv' directly
if exist myenv\Scripts\activate.bat call myenv\Scripts\activate.bat

:: Try to activate 'venv' directly
if exist venv\Scripts\activate.bat call venv\Scripts\activate.bat

echo.
echo [STEP 2] Checking Python File...

:: Check if the file exists before running
if not exist name_to_cas.py goto FileNotFound

echo.
echo [STEP 3] Running Python Program...
echo ----------------------------------------
python name_to_cas.py
echo ----------------------------------------
goto End

:FileNotFound
echo.
echo [ERROR] CRITICAL ERROR!
echo Python file 'name_to_cas.py' was NOT found in this folder.
echo.
echo PLEASE CHECK:
echo 1. Is the file named 'name_to_cas.py'?
echo 2. Did Windows hide the '.txt' extension? (It might be 'name_to_cas.py.txt')
echo.

:End
echo.
echo Program Finished.
pause