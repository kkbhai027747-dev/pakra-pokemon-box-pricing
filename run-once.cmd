@echo off
setlocal
set "PYTHONUTF8=1"
set "POKEMON_PYTHON=%~dp0.venv\Scripts\python.exe"
if not exist "%POKEMON_PYTHON%" set "POKEMON_PYTHON=python"
"%POKEMON_PYTHON%" "%~dp0collect.py" %*
set "POKEMON_EXIT=%ERRORLEVEL%"
pause
exit /b %POKEMON_EXIT%
