@echo off
setlocal

rem main function
call :test1
call :test2
exit /b 0

rem function test1
:test1
echo test1
exit /b 0

rem function test2
:test2
echo test2
exit /b 0