ren NCPFPRNT.BAS NCPFPRNT.
for %%1 in (*.bas) do bc %%1 /o/ot/ah/g2/x;

bc NCPFPRNT /o/ot/g2/X;
bc UBMISC /o/ot/g2/ah;
bc comnaux /o/ot/g2/ah;
bc ubmtread /o/ot/g2/ah;
bc ubcust /o/ot/g2/ah;
bc ublocat /o/ot/g2/ah;
bc ublocat /o/ot/g2/ah;
bc ubmenu /o/ot/g2/ah;
bc ubfinbil /o/ot/g2/ah;
bc formedit /d/o/ot/g2/ah;

LINK @GLEXP.RSP
link @BGEXP.RSP
DEL *.obj
