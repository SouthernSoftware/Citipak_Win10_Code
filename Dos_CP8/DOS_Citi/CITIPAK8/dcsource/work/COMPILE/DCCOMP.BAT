del dc*.exe
del *.obj
for %%1 in (*.bas) do (echo %1^bc %%1 /lr/o/ot/ah;)

bc NCPFPRNT /lr/o/ot/g2/X/AH;
bc comnaux /lr/o/ot/g2/ah;

LINK @DCMENU.RSP
LINK @DCCODE.RSP
LINK @DCCUST.RSP
LINK @DCPAY.RSP
LINK @DCREPORT.RSP
