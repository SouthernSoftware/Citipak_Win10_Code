ren prfprint.bas prfprint.
ren Npost2gl.bas Npost2gl.
for %%a in (*.bas) do (Echo BC %%a /o /ot /g2 ^echo.^bc %%a /o/ot/g2;)
ren Npost2gl. Npost2gl.bas
ren prfprint. prfprint.bas
bc Npost2gl.bas /x/o/ot/g2;
bc prfprint.bas /x/o/ot/g2;
bc prprint1.bas /ah/o/ot/g2;
bc split01.bas /s/o/ot/g2;
link @dirlink.rsp
