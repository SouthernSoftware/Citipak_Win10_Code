rem CALL c:\SETQBX

del *.exe /q
del *.obj /q

for %%1 in (*.bas) do (echo %%1^bc %%1 /lr/o/ot/ah/g2;)

bc NCPFPRNT /lr/o/ot/g2/X;
bc comnaux /lr/o/ot/g2/ah;

LINK @ARMENU.RSP
LINK @ARBILLIN.RSP
LINK @ARSETBAL.RSP
LINK @ARCODE.RSP
LINK @ARCUST.RSP
LINK @ARPAYMEN.RSP
LINK @ARREPORT.RSP
LINK @ARSPEC.RSP
LINK @ARWACK.RSP
LINK @ARADJUST.RSP
rem del *.obj
