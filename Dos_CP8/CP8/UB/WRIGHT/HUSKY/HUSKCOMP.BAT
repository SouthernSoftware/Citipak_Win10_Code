setlocal
set include=d:\cp8\ub\wright\incs
for %%1 in (*.bas) do bc %%1 /fpa/o/ot/g2;
link @HUSKREAD.RSP
lzeshell huskread
rem DEL *.obj
