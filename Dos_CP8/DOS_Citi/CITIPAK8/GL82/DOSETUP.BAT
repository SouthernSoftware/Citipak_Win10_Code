cls

@echo off
echo Ready to create Citipak 8.2 installation files.
echo Press Ctrl+Break to abort or,
pause

cls
echo Cleaning up old files...
del gl82.exe
DEL ap82.exe
echo.
echo.

echo Setting up GL
pkzip gl82 cpicons.dll cpmenu.exe gl*.exe bgt*.exe gl*.q?l cmndlg.q?l 
zip2exe gl82
del gl82.zip
echo.
echo.

echo Setting up AP
pkzip ap82 ap*.exe ap*.q?l apcheck.msk chkrec.exe
zip2exe ap82
del ap82.zip

cls
echo Citipak installation files created.

