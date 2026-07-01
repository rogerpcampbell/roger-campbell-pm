@echo off
setlocal
cd /d "%~dp0"
set "APP_SUBDIR=app_files"

echo ===============================================
echo Roger Campbell Rail Ponds Roads Control Center - Windows Launcher
echo ===============================================
echo.

if not exist "%APP_SUBDIR%\app.py" (
    echo App files were not found in "%APP_SUBDIR%".
    echo Keep this launcher beside the "%APP_SUBDIR%" folder.
    echo.
    pause
    exit /b 1
)

cd /d "%~dp0%APP_SUBDIR%"

where py >nul 2>nul
if %errorlevel%==0 (
    set "PYTHON_CMD=py -3"
) else (
    where python >nul 2>nul
    if errorlevel 1 (
        echo Python was not found.
        echo Install Python 3.10 or newer from python.org and tick "Add python.exe to PATH" during installation.
        echo.
        pause
        exit /b 1
    )
    set "PYTHON_CMD=python"
)

if not exist ".venv\Scripts\python.exe" (
    echo Creating local virtual environment...
    %PYTHON_CMD% -m venv .venv
    if errorlevel 1 (
        echo Failed to create virtual environment.
        pause
        exit /b 1
    )
)

call ".venv\Scripts\activate.bat"

echo Installing or updating required packages...
python -m pip install --upgrade pip
pip install -r requirements.txt
if errorlevel 1 (
    echo Failed to install required packages.
    pause
    exit /b 1
)

echo.
echo Starting Streamlit app...
echo A browser window should open automatically. If not, copy the Local URL shown below.
echo.
streamlit run app.py

pause
