FOR %%1 IN (..\source\*.BAS) DO BC %%1 /O/FS;
BC ..\source\ncpfprnt /X/O/FS;
lib @crqlb.rsp
del *.obj
wmakeqlb checkrec.lst,,nul,vtoolf pro7 crqlb QBX basic7;
del /q crqlb.lib
move checkrec.qlb ..\..\work\
