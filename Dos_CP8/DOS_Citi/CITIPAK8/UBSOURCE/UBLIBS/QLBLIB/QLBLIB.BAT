REN ..\source\NCPFPRNT.BAS ..\source\NCPFPRNT.
FOR %%1 IN (..\source\*.BAS) DO BC %%1 /O/FS;
REN ..\source\NCPFPRNT. ..\source\NCPFPRNT.BAS
BC ..\source\NCPFPRNT.BAS  /O/FS/X;
lib @Ubqlb.rsp
del *.obj
wmakeqlb ub.lst,,nul,vtoolf pro7 ubqlb QBX basic7;
del /q ubqlb.lib
move ub.qlb ..\..\UB\
