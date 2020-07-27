ren NCPFPRNT.BAS NCPFPRNT.HLD
for %%1 in (*.bas) do bc %%1 /lr/o/ot/g2/ah;
ren NCPFPRNT.HLD NCPFPRNT.BAS
bc NCPFPRNT /lr/o/ot/g2/X;

LINK @CAMENU.RSP
link @CAINV.rsp
link @CAREPORT.rsp
link @CAINT.rsp
link @CABEG.rsp
link @CAPAY.rsp
link @CAADJUST.rsp
link @CACUST.rsp
rem DEL *.obj
