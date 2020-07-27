@echo off

cls
echo.
echo Southern Software Fund Accounting System version 8.11
echo Copyright 1996 Southern Software Systems, Inc.
echo.

if z%1 == z goto SYNTAX
if z%2 == z goto SYNTAX
REM if not exist %1%2\nul goto PathError
echo OK to install to %1%2?
echo.
echo If you want to Cancel the installation, Press Ctrl-Break NOW or,
pause

cls
echo Installing System. . .
ap81 %1%2
   
cls
echo G/L Installation complete. 
goto ExitBatch

:PathError
echo     ERROR:   Invalid Drive/Path specification in:  %1%2
echo.
:SYNTAX
echo   SYNTAX:   INSTALL (Drive:) (\Directory)
echo.
echo   Example:  install c: \citipak 
echo.
echo   Would install to:  C:\CITIPAK
echo.
echo   NOTE:   One space is required after the (Drive:).
echo.
echo.
:ExitBatch
