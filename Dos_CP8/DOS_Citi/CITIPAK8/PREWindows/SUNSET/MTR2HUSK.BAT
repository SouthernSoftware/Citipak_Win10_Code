@echo off
cls
echo Make sure the HUSKY is connected and ready to
echo transfer files. The HUSKY should be at the
echo "Husky File Transfer Utility" screen. You can
echo get to the "Utility" by typing the "H" key
echo and pressing the "Yes" key when the C: prompt
echo is being displayed.
echo.
echo.
pause
cls
cmdhcom /port=1 /irq=4 /addr=3F8 /baud=19200 /tx=READMTR.EXE /abort
