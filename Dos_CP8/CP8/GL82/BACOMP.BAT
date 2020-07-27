cls
set include=p:\incfiles
set lib=p:\lib

ren glrpt.bas glrpt.src
ren glrpt2.bas glrpt2.src
ren glrptrd.bas glrptrd.src
ren ncpfprnt.bas ncpfprnt.src
ren glcopyrg.bas glcopyrg.src

for %%1 in (*.bas) do p:\binb\bc %%1 /o/ot/ah/g2; 

ren glcopyrg.src glcopyrg.bas
ren ncpfprnt.src ncpfprnt.bas
ren glrpt.src glrpt.bas
ren glrptrd.src glrptrd.bas
ren glrpt2.src glrpt2.bas

p:\binb\bc glcopyrg /o/ot/g2/e/v;
p:\binb\bc glrpt /ah/o/ot/g2/e;
p:\binb\bc glrpt2 /ah/o/ot/g2/e;
p:\binb\bc glrptrd /ah/o/ot/g2/e;
p:\binb\bc ncpfprnt /o/ot/g2/x;

rem p:\binb\LINK @CPMENU.RSP
p:\binb\LINK @GLMENU.RSP
p:\binb\LINK @BGTEDIT.RSP
p:\binb\LINK @BGTPREP.RSP
p:\binb\LINK @GLGJED.RSP
p:\binb\LINK @GLMAINT.RSP
p:\binb\LINK @GLSYS.RSP
p:\binb\LINK @GLFUND.RSP
p:\binb\LINK @GLACCT.RSP
p:\binb\LINK @GLDEPT.RSP
p:\binb\LINK @GLBANK.RSP
p:\binb\LINK @GLRPT.RSP
rem p:\binb\LINK @GLRPTRD.RSP
p:\binb\LINK @GLCOPYRG.RSP
p:\binb\LINK @GLCLOSE.RSP
p:\binb\LINK @GLUTIL.RSP
rem Pause
p:\binb\LINK @GLGETTR.RSP
rem Pause
p:\binb\LINK @APMENU.RSP
p:\binb\LINK @APV.RSP
p:\binb\LINK @APPO.RSP
p:\binb\LINK @APINV.RSP
p:\binb\LINK @APCHECK.RSP
p:\binb\LINK @APRPT1.RSP
rem p:\binb\LINK @APRPTRD.RSP
p:\binb\Link @APLINK.RSP
p:\binb\Link @AP1099.rsp
rem p:\binb\Link @CPRDMENU.rsp
