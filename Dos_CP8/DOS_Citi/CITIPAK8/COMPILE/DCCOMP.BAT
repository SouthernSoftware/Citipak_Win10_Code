for %%1 in (*.bas) do bc %%1 /o/ot/ah/fs;

bc NCPFPRNT /o/ot/g2/X/fs;
bc comnaux /o/ot/g2/ah/fs;

LINK @DCMENU.RSP
LINK @DCCODE.RSP
LINK @DCCUST.RSP
LINK @DCPAYMEN.RSP
LINK @DCREPORT.RSP
del *.obj
