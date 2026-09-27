@echo off
call "%~dp0run-once.cmd" --daily %*
exit /b %ERRORLEVEL%
