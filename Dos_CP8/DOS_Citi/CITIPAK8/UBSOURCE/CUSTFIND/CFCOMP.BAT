ren ncpfprnt.bas ncpfprnt.hld
ren cfaux01.bas cfaux01.hld

for %%1 in (*.bas) do bc %%1 /o/ot/g2;

ren ncpfprnt.hld ncpfprnt.bas
ren cfaux01.hld cfaux01.bas

bc cfaux01 /o/ot/g2/x;
bc ncpfprnt /o/ot/g2/x;

LINK @CustFind.RSP

