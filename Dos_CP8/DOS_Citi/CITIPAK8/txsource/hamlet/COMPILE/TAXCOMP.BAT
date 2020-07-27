del *.exe *.obj
ren NCPFPRNT.BAS NCPFPRNT.HLD
for %%1 in (*.bas) do bc %%1 /lr/o/ot/g2/ah;
ren NCPFPRNT.HLD NCPFPRNT.BAS
bc NCPFPRNT /lr/o/ot/g2/X;
bc TAXCPYRG /lr/o/ot/g2/X;

LINK @TAXMENU.RSP
link @taxsetup.rsp
link @taxreprt.rsp
link @taxprntb.rsp
link @taxbill.rsp
link @taxint.rsp
link @taxpaymn.rsp
link @taxconv.rsp
link @Taxlnote.rsp
link @TaxCPYRG.rsp
rem DEL *.obj
