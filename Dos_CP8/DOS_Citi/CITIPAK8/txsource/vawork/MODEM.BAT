@echo off
if %1&==& goto end
goto %1
:ON
:on
:On
copy p:\mshare.txt P:\share.mod
del  p:\noshare.mod
goto end

:OFF
:off
:Off
copy p:\mshare.txt P:\noshare.mod
del  p:\share.mod
goto end

:end
