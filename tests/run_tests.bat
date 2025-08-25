@echo off
setlocal enabledelayedexpansion

where python >NUL 2>&1
if errorlevel 1 (
  echo Python not found on PATH.
  exit /b 1
)

echo Upgrading pip and installing pytest (if needed)...
python -m pip install --upgrade pip >NUL 2>&1
python -m pip install pytest >NUL 2>&1

echo Running tests...
python -m pytest -q tests
set ERR=%ERRORLEVEL%

if %ERR% neq 0 (
  echo.
  echo Some tests failed. See output above.
) else (
  echo.
  echo All tests passed.
)
exit /b %ERR%
