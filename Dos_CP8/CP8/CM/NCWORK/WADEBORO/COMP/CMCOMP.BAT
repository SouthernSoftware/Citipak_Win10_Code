set include=D:\CP8\UB\WORK\WADEBORO\INCS\

rem CALL c:\SETQBX
for %%1 in (*.bas) do (echo %%1^bc %%1 /o/ot/ah/g2;)

bc NCPFPRNT /o/ot/g2/X;
bc comnaux /o/ot/g2/ah;
bc ubopendr/o/ot/g2/x;

LINK @CMMENU.RSP
rem link @cmsendgl.rsp
del *.obj
