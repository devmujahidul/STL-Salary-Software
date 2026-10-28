@echo off
setlocal

title STL Salary - Attendance Converter

rem Run from the folder this .bat file lives in, no matter where it was
rem started from (double-click uses its folder, but "Run as administrator"
rem starts in system32 - the next line fixes that).
cd /d "%~dp0"

echo ==================================================
echo    STL SALARY - ATTENDANCE CONVERTER
echo ==================================================
echo    Developer : Mujahidul Islam
echo    Role      : MERN Stack Developer
echo    Formerly  : IT Officer (IT Engineering)
echo    Website   : www.mujahidul.dev
echo    Phone     : +88016088-64209
echo ==================================================
echo.

rem Pick a Python launcher: prefer "python", fall back to "py".
set "PYTHON="
set "PYVER="
where python >nul 2>nul && set "PYTHON=python"
if not defined PYTHON where py >nul 2>nul && set "PYTHON=py"

if not defined PYTHON (
    echo [ERROR] Python was not found on this computer.
    echo         Install Python 3 and make sure "python" is on the PATH,
    echo         then run this file again.
    echo.
    pause
    exit /b 1
)

for /f "tokens=*" %%v in ('%PYTHON% --version 2^>^&1') do set "PYVER=%%v"
echo Using   : %PYVER%

rem The macro (.xlsm) build needs the "pywin32" package. On a fresh PC it is
rem usually missing - install it automatically once, right here.
%PYTHON% -c "import win32com.client" >nul 2>nul
if errorlevel 1 (
    echo Setting up: installing Python package "pywin32" ...
    %PYTHON% -m pip install --quiet pywin32 2>nul
    %PYTHON% -c "import win32com.client" >nul 2>nul
    if errorlevel 1 (
        echo   [WARN] "pywin32" is still missing - the .xlsm macros need it.
        echo   Run this manually:  %PYTHON% -m pip install pywin32
        echo   The .xlsx is built even without it.
        echo.
    )
)

echo Command : %PYTHON% convert_attendance.py %*
echo --------------------------------------------------
echo.

%PYTHON% convert_attendance.py %*
set "RC=%ERRORLEVEL%"

echo.
echo --------------------------------------------------
if not "%RC%"=="0" (
    echo [FAILED] The converter stopped with exit code %RC%.
    echo          Read the messages above to fix the problem.
) else (
    echo [DONE] Conversion finished. Newest files in the output folder:
    echo.
    dir /b /o-d "output\Salary_Attendance_*" 2>nul
    echo.
    echo Open the .xlsm in Excel and click "Enable Content" if asked.
    echo If only a .xlsx appears, macros were not added - the .xlsx still
    echo filters from the SubSection Filter column.
)
echo.
pause
exit /b %RC%
