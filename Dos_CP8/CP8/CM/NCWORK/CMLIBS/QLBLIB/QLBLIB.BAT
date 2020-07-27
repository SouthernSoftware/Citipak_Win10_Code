rem set include=d:\cp8\ub\work\wadeboro\incs
REN ..\source\NCPFPRNT.BAS ..\source\NCPFPRNT.
ren ..\source\UBOPENDR.BAS ..\source\UBOPENDR.
FOR %%1 IN (..\source\*.BAS) DO BC %%1 /O/FS/AH;
REN ..\source\NCPFPRNT. ..\source\NCPFPRNT.BAS
ren ..\source\UBOPENDR. ..\source\UBOPENDR.BAS
BC ..\source\NCPFPRNT.BAS  /O/FS/X/AH;
BC ..\source\UBOPENDR.BAS  /O/FS/E/AH;

lib @Ubqlb.rsp
del *.obj
wmakeqlb ub.lst,,nul,vtoolf pro7 ubqlb QBX basic7;
del /q ubqlb.lib
move ub.qlb ..\..\work\
