TotalAmt# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
Rev6# = UBCustRec(1).CurrRevAmts(6)

PRINT #UBRpt, "~"
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt, TAB(68); NDate$
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt,                       'line 10
PRINT #UBRpt, TAB(45); LongPDate$
PRINT #UBRpt, "   11:00 am."
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt, TAB(68); AcctNo&
PRINT #UBRpt,
PRINT #UBRpt, TAB(20); UBCustRec(1).CustName; TAB(68); USING "$$,######.##"; Round#(TotalAmt# - Rev6#)
PRINT #UBRpt, TAB(20); UBCustRec(1).Addr1
PRINT #UBRpt, TAB(20); UBCustRec(1).Addr2; TAB(68); USING "$$,######.##"; Rev6#
PRINT #UBRpt, TAB(20); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; UBCustRec(1).ZipCode
PRINT #UBRpt, TAB(68); USING "$$,######.##"; TotalAmt#
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt, "~"

