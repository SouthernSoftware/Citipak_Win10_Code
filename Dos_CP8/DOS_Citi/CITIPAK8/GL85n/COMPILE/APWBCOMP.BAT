cls

ren glrpt.bas glrpt.src
ren glrpt2.bas glrpt2.src
ren glrptrd.bas glrptrd.src
ren ncpfprnt.bas ncpfprnt.src
ren glcopyrg.bas glcopyrg.src

for %%1 in (*.bas) do bc %%1 /o/ot/ah/g2; 

ren glcopyrg.src glcopyrg.bas
ren ncpfprnt.src ncpfprnt.bas
ren glrpt.src glrpt.bas
ren glrptrd.src glrptrd.bas
ren glrpt2.src glrpt2.bas

bc ncpfprnt /o/ot/g2/x;

LINK @APPOWB.RSP
LINK @APPO.RSP
