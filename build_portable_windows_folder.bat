@echo off
setlocal
cd /d "%~dp0"
echo ======================================================
echo Building Excel Auto Grader v6 Final portable Windows folder
echo Target: Windows 11 Enterprise, 64-bit x64
echo Output: dist\Excel Auto Grader\Excel Auto Grader.exe
echo ======================================================
echo.

set "PY_CMD="
py -3 --version >nul 2>nul
if %errorlevel%==0 set "PY_CMD=py -3"
if not defined PY_CMD (
    python --version >nul 2>nul
    if %errorlevel%==0 set "PY_CMD=python"
)
if not defined PY_CMD (
    python3 --version >nul 2>nul
    if %errorlevel%==0 set "PY_CMD=python3"
)
if not defined PY_CMD (
    echo Python was not found on this build computer.
    echo Use the GitHub Actions workflow included in this package.
    echo The restricted/domain computer does NOT need Python after the portable folder is built.
    pause
    exit /b 1
)

if not exist ".venv" %PY_CMD% -m venv .venv
if %errorlevel% neq 0 goto error
call .venv\Scripts\activate.bat
if %errorlevel% neq 0 goto error
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if %errorlevel% neq 0 goto error

python -m py_compile app.py grader.py
if %errorlevel% neq 0 goto error
python -m PyInstaller --noconfirm --clean --onedir --name "Excel Auto Grader" app.py --collect-all flask --collect-all openpyxl --hidden-import et_xmlfile
if %errorlevel% neq 0 goto error

if exist "dist\Excel Auto Grader\Excel Auto Grader.exe" (
    copy /Y rubric_template.xlsx "dist\Excel Auto Grader\rubric_template.xlsx" >nul
    copy /Y PE2_Ver_A_rubric_v6.xlsx "dist\Excel Auto Grader\PE2_Ver_A_rubric_v6.xlsx" >nul
    copy /Y README.md "dist\Excel Auto Grader\README.md" >nul
    echo.
    echo Build complete: %cd%\dist\Excel Auto Grader
    pause
    exit /b 0
)

echo Expected executable was not found.
pause
exit /b 1

:error
echo Build failed. Copy the messages above for troubleshooting.
pause
exit /b 1
