cls

@echo off
echo Ready to create Citipak 8.2 installation files.
echo Press Ctrl+Break to abort or,
pause

cls
echo Cleaning up old files...
del gl82.exe
echo Setting up GL
pkzip gl82 gl*.exe bgt*.exe gl*.q?l cmndlg.q?l 
zip2exe gl82
del gl82.zip
echo.
echo.
cls
echo Citipak GL installation file created.

