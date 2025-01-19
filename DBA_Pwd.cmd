@echo off
SET Programma=%~n0

if "%1" == "" (
call :usage
goto :fine
)

if "%2" == "" (
call :usage
goto :fine
)

If "%3" == "" (
	cscript //nologo %Programma%.vbs "MyPWD=%1" "ACTION=%2" 
) else (
	cscript //nologo %Programma%.vbs "MyPWD=%1" "ACTION=%2" "BaseCode=%3" 
)
if not %ERRORLEVEL% == 0 (goto :usage)
goto :fine


:usage
echo %Programma%.cmd 
echo.
echo usage: %Programma%.cmd "string Password" ("D" for Decript: "E" For Encript) [BaseCode]
echo.
echo        examples
echo        	Encript password
echo        		%Programma%.cmd MyPasswrod E [BaseCode]
echo.
echo        	Encript password
echo        		%Programma%.cmd 12344487523572345235 E [BaseCode]
echo.
echo        	Decript password
echo        		%Programma%.cmd EncryptMyPasswrod D [BaseCode]
echo.
echo        	Decript password
echo        		%Programma%.cmd 12344487523572345235 D [BaseCode]
goto :fine

:fine