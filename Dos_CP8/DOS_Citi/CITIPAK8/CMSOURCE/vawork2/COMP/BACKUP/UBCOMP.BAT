ren NCPFPRNT.BAS NCPFPRNT.
for %%1 in (*.bas) do (bc %%1 /o/ot/g2;)
ren NCPFPRNT. NCPFPRNT.BAS
bc NCPFPRNT /o/ot/g2/X;
bc comnaux /o/ot/g2/ah;
bc ubbillin /o/ot/g2/ah;
rem bc formedit /d/o/ot/g2/ah;

LINK @UBMENU.RSP
LINK @UBMISC.RSP
LINK @UBSTATRE.RSP
LINK @UBCUSTRE.RSP
LINK @UBBILLIN.RSP
LINK @UBFINBIL.RSP
LINK @UBPRNBIL.RSP
rem LINK @UBPRNFIN.RSP
LINK @UBPAYMNT.RSP
LINK @UBDRAFT.RSP
LINK @UBMTREAD.RSP
LINK @UBPENLTY.RSP


DEL *.obj
