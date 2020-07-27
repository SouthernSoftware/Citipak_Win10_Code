ren NCPFPRNT.BAS NCPFPRNT.HLD
for %%1 in (*.bas) do bc %%1 /o/ot/g2/ah;
ren NCPFPRNT.HLD NCPFPRNT.BAS
bc NCPFPRNT /o/ot/g2/X;

LINK @TAXMENU.RSP
rem DEL *.obj
