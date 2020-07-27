echo on
set include=d:\cp8\ub\work\wadeboro\incs
set
rem quit
rem pause

ren NCPFPRNT.BAS NCPFPRNT.
for %%1 in (*.bas) do (echo %%1^bc %%1 /o/ot/g2;)
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
bc formedit /o/ot/g2/ah;

LINK @UBMENU.RSP
LINK @UBMISC.RSP
LINK @UBADJUST.RSP
LINK @UBSETUP.RSP
LINK @UBSTATRE.RSP
LINK @UBCUSTRE.RSP
LINK @UBWRKORD.RSP
LINK @UBBILLIN.RSP
LINK @UBZAP.RSP
LINK @UBFINBIL.RSP
LINK @UBPRNBIL.RSP
LINK @UBLNPRNT.RSP
LINK @UBPRNFIN.RSP
LINK @UBPAYMNT.RSP
LINK @UBDRAFT.RSP
LINK @UBMTREAD.RSP
LINK @UBPENLTY.RSP
LINK @UBMAILLB.RSP

rem DEL *.obj
