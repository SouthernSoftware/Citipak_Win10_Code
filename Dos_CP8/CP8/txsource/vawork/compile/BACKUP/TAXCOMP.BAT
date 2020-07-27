ren NCPFPRNT.BAS NCPFPRNT.HLD
for %%1 in (*.bas) do bc %%1 /o/ot/g2/ah;
ren NCPFPRNT.HLD NCPFPRNT.BAS
bc NCPFPRNT /o/ot/g2/X;

LINK @TAXMENU.RSP
link @taxsetup.rsp
link @taxreprt.rsp
link @txprntpb.rsp
link @txprntrb.rsp
link @taxppbil.rsp
link @taxrebil.rsp
link @taxint.rsp
link @taxpaymn.rsp

rem DEL *.obj
