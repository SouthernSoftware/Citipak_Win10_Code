ren NCPFPRNT.BAS NCPFPRNT.HLD
for %%1 in (*.bas) do bc %%1 /o/ot/g2/ah;
ren NCPFPRNT.HLD NCPFPRNT.BAS
bc NCPFPRNT /o/ot/g2/X;
bc TAXCPYRG /o/ot/g2/X;

LINK @TAXMENU.RSP
link @taxsetup.rsp
link @taxreprt.rsp
link @taxprntb.rsp
link @taxbill.rsp
link @taxint.rsp
link @taxSintR.rsp
link @taxpaymn.rsp
link @taxconv.rsp
link @Taxlnote.rsp
link @Taxdnote.rsp
link @TaxCPYRG.rsp
link @TaxADVRT.rsp

rem DEL *.obj
