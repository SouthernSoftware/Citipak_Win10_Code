@echo off

cls
echo Southern Software Fund Accounting System version 8.11
echo Copyright 1996 Southern Software Systems, Inc.
echo.

if z%1 == z goto SYNTAX
if z%2 == z goto SYNTAX
if not exist %1%2\glsetup.dat goto PathError

echo Ready to install update.  Press Ctrl+Break to cancel or,
pause
cls
echo Installing update. . .
GL82 -o %1%2
cls
echo Update complete.
echo Don't forget to change the settings as described in the instructions!
goto ExitBatch

:PathError
echo   ERROR  :  %1%2
echo             Invalid path to the active citipak directory
echo.

:SYNTAX
echo   SYNTAX :  glupdate (Drive:) (\Directory)
echo.  Note   :  Include a space between the drive and path specifications .
echo.
echo   Example:  glupdate c: \citipak
echo             Would install to:  C:\CITIPAK
echo.

:ExitBatch
