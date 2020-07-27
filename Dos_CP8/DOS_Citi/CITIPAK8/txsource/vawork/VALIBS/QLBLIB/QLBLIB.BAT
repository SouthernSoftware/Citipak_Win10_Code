REN ..\source\NCPFPRNT.BAS ..\source\NCPFPRNT.
FOR %%1 IN (..\source\*.BAS) DO BC %%1 /O/FS;
REN ..\source\NCPFPRNT. ..\source\NCPFPRNT.BAS
BC ..\source\NCPFPRNT.BAS  /O/FS/X;
lib @taxqlb.rsp
del *.obj
rem wmakeqlb tax.lst,,nul,pro7 taxqlb QBX wff_ii basic7;
wmakeqlb tax.lst,,nul,pro7 taxqlb basic7;
del /q taxqlb.lib

REM dir \qbx\ub\ub.qlb
rem move ub.qlb ..\..\work2\
rem dir \qbx\ub\ub.qlb
