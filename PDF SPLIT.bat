@echo off
setlocal enabledelayedexpansion


:: --- CONFIGURATION ---
setlocal enabledelayedexpansion

set "PY_PATH=%~dp0installed\python-3.14.2-embed-amd64"
set "SCRIPT_PATH=%~dp0PO_PDF_Splitter_GUI.py"
set "TESSDATA_PREFIX=%~dp0installed\tesseract\tessdata"

echo ====================================================
echo PORTABLE PYTHON 3.14.2 SETUP (2026)
echo ====================================================

:: 1. Verify unzipped folder exists
if not exist "%PY_PATH%\python.exe" (
    echo [ERROR] python.exe not found at %PY_PATH%
    pause
    exit /b 1
)

:: 2. Enable site-packages (CRITICAL for embeddable version)
echo [1/4] Configuring Python environment...
pushd "%PY_PATH%"
for %%f in (python*._pth) do (
    findstr /x "import site" "%%f" >nul
    if errorlevel 1 (
        echo.>> "%%f"
        echo import site>> "%%f"
        echo [OK] Enabled site-packages in %%f
    )
)
popd

:: 3. Bootstrap PIP (Fix: Check module capability, not just file existence)
echo [2/4] Verifying PIP installation...
"%PY_PATH%\python.exe" -m pip --version >nul 2>&1

if %errorlevel% NEQ 0 (
    echo        PIP module not detected. Installing/Repairing...
    
    :: Force download get-pip.py
    powershell -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri 'https://bootstrap.pypa.io/get-pip.py' -OutFile '%PY_PATH%\get-pip.py'"
    
    :: Install pip
    "%PY_PATH%\python.exe" "%PY_PATH%\get-pip.py" --no-warn-script-location
    
    :: Clean up
    if exist "%PY_PATH%\get-pip.py" del "%PY_PATH%\get-pip.py"
    
    echo        [OK] PIP installed successfully.
) else (
    echo        [OK] PIP is operational.
)

:: 4. INSTALL BUILD TOOLS
echo [INFO] Installing required build tools...
"%PY_PATH%\python.exe" -m pip install --upgrade pip setuptools wheel

:: 5. Install Project Dependencies
echo [3/4] Installing dependencies from requirements.txt...
if exist "%~dp0requirements.txt" (
    "%PY_PATH%\python.exe" -m pip install -r "%~dp0requirements.txt"
) else (
    echo [SKIP] No requirements.txt found.
)

:: 6. Run the Application
echo [4/4] Starting GUI Application...
echo ----------------------------------------------------
"%PY_PATH%\python.exe" "%SCRIPT_PATH%"

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Application crashed with exit code %ERRORLEVEL%
    pause
)
exit /b
