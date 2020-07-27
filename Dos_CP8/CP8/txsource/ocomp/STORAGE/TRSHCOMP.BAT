ren NCPFPRNT.BAS NCPFPRNT.
for %%1 in (*.bas) do bc %%1 /o/ot/g2;
ren NCPFPRNT. NCPFPRNT.BAS
bc NCPFPRNT /o/ot/g2/X;
bc UBMISC /o/ot/g2/ah;
bc comnaux /o/ot/g2/ah;
bc ubmtread /o/ot/g2/ah;
bc ubcust /o/ot/g2/ah;
BC UBSETFIN /O/OT/G2/AH;
BC UBCUSTRE /O/OT/G2/AH;
bc ubmenu /o/ot/g2;
bc ubzap /o/ot/g2/AH;
rem bc ubelspec /o/ot/g2;

bc formedit /o/ot/g2/ah;

LINK @TRASHRPT.RSP

DEL *.obj
