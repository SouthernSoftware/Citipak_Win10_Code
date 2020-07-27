cls

@echo off
echo Ready to create Citipak 8.2 A/P installation files.
echo Press Ctrl+Break to abort or,
pause

cls
del ap82.exe
echo.
echo.

echo Setting up AP
pkzip ap82 ap*.exe ap*.q?l apcheck.msk chkrec.exe
zip2exe ap82
del ap82.zip

cls
echo Citipak installation files created.

