CALL c:\SETQBX
for %%1 in (*.bas) do bc %%1 /o/ot/ah/g2/x;

bc NCPFPRNT /o/ot/g2/X;
bc comnaux /o/ot/g2/ah;
bc cmmenu /o/ot/g2/ah;

LINK @CMMENU.RSP
del *.obj
