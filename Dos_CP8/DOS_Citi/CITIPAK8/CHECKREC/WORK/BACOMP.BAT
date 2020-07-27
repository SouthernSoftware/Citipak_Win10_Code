cls

ren glrpt.bas glrpt.src
ren glrpt2.bas glrpt2.src
ren ncpfprnt.bas ncpfprnt.src
ren glcopyrg.bas glcopyrg.src

for %%1 in (*.bas) do bc %%1 /o/ot/ah/g2; 

ren glcopyrg.src glcopyrg.bas
ren ncpfprnt.src ncpfprnt.bas
ren glrpt.src glrpt.bas
ren glrpt2.src glrpt2.bas

bc glcopyrg /o/ot/g2/e/v;
bc glrpt /ah/o/ot/g2/e;
bc glrpt2 /ah/o/ot/g2/e;
bc ncpfprnt /o/ot/g2/x;

LINK @CPMENU.RSP
LINK @GLMENU.RSP
LINK @BGTEDIT.RSP
LINK @BGTPREP.RSP
LINK @GLGJED.RSP
LINK @GLMAINT.RSP
LINK @GLSYS.RSP
LINK @GLFUND.RSP
LINK @GLACCT.RSP
LINK @GLDEPT.RSP
LINK @GLBANK.RSP
LINK @GLRPT.RSP
LINK @GLCOPYRG.RSP
LINK @GLCLOSE.RSP
LINK @GLUTIL.RSP
LINK @APMENU.RSP
LINK @APV.RSP
LINK @APPO.RSP
LINK @APINV.RSP
LINK @APCHECK.RSP
LINK @APRPT1.RSP
Link @APLINK.RSP
Link @AP1099.rsp
