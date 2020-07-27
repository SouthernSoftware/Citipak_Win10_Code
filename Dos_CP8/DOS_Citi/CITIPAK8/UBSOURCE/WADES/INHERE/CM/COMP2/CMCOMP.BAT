setlocal
echo on
set include=f:\citipak8\ubsource\wades\incs
@echo off
pause
for %%1 in (*.bas) do bc %%1 /o/ot/ah/g2/x;

bc NCPFPRNT /o/ot/g2/X;
bc comnaux /o/ot/g2/ah;

LINK @CMMENU.RSP
link @cmsendgl.rsp
del *.obj
