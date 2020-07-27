TBal# = Round#(UBCustRec(1).CurrBalance# + UBCustRec(1).PrevBalance#)
TPen# = Round#(UBCustRec(1).CurrRevAmts(4) + UBCustRec(1).CurrRevAmts(5) + UBCustRec(1).CurrRevAmts(6) + UBCustRec(1).CurrRevAmts(10))
BalLessPen# = Round#(TBal# - TPen#)

PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt,

PRINT #UBRpt, TAB(24); Location$
PRINT #UBRpt,
PRINT #UBRpt, TAB(55); USING "#####.##"; BalLessPen#
PRINT #UBRpt, TAB(16); USING "####.##"; TPen#;
PRINT #UBRpt, TAB(47); USING "#####.##"; TBal#
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt, TAB(40); LEFT$(UBCustRec(1).CustName, 30)
PRINT #UBRpt, TAB(40); LEFT$(UBCustRec(1).Addr1, 30)
PRINT #UBRpt, TAB(40); LEFT$(UBCustRec(1).Addr2, 30)
PRINT #UBRpt, TAB(40); LEFT$(UBCustRec(1).City, 20); " "; UBCustRec(1).State; " "; Zip$
PRINT #UBRpt,

