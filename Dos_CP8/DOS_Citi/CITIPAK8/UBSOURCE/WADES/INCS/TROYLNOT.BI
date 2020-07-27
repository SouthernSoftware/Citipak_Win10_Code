    TotalAmt# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
PRINT #UBRpt, "~"
PRINT #UBRpt, TAB(8); AcctNo&; TAB(65); USING "#####.##"; TotalAmt#
PRINT #UBRpt,
PRINT #UBRpt, TAB(14); NDate$; TAB(58); "ACCT #"; AcctNo&
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt, TAB(11); PDate$
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt, TAB(40); USING "#####.##"; TotalAmt#
PRINT #UBRpt,
PRINT #UBRpt, TAB(8); UBCustRec(1).CustName
PRINT #UBRpt, TAB(8); UBCustRec(1).Addr1
PRINT #UBRpt, TAB(8); UBCustRec(1).Addr2
PRINT #UBRpt, TAB(8); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; UBCustRec(1).ZipCode
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt, "~"

