CALL c:\SETQBX
for %%1 in (*.bas) do bc %%1 /o/ot/ah/g2/x;

bc NCPFPRNT /o/ot/g2/X;
bc comnaux /o/ot/g2/ah;

LINK @TBCUSTRE.RSP
LINK @TBREP11.RSP
LINK @TBREP12.RSP
LINK @TBREP13.RSP
LINK @TBREP14.RSP
LINK @TBREP15.RSP
link @TBREP16.RSP
link @TBBILLRP.RSP
del *.obj
