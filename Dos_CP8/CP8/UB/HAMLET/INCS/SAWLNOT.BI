    Zip$ = QPTrim$(UBCustRec(1).ZipCode)
    Previous# = UBCustRec(1).PrevBalance
    Current# = UBCustRec(1).CurrBalance
    TotalAmt# = Round#(Previous# + Current#)

    PRINT #UBRpt, "~"
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(39); LEFT$(UBCustRec(1).CustName, 29)

    PRINT #UBRpt, TAB(6); "PAST DUE/FINAL NOTICE"; TAB(39); UBCustRec(1).Addr1
    PRINT #UBRpt, TAB(39); UBCustRec(1).Addr2
    PRINT #UBRpt, TAB(39); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State
    PRINT #UBRpt, TAB(48); LEFT$(Zip$, 5); "-"; RIGHT$(Zip$, 4)
    PRINT #UBRpt, TAB(2); "ACCT# "; USING "#####"; AcctNo&;
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(39); UBCustRec(1).Book; TAB(43); USING "#####"; AcctNo&;
    PRINT #UBRpt, TAB(50); LEFT$(PDate$, 2); TAB(53); MID$(PDate$, 4, 2); TAB(57); RIGHT$(PDate$, 2)
    PRINT #UBRpt,

    PRINT #UBRpt, TAB(3); LEFT$(METERREADDATE$, 2); TAB(6); MID$(METERREADDATE$, 3, 2);
    PRINT #UBRpt, TAB(12); USING "#####.##"; TotalAmt#;

     PRINT #UBRpt, TAB(29); USING "#####.##"; TotalAmt#;
     PRINT #UBRpt, TAB(41); USING "#####.##"; TotalAmt#;
     PRINT #UBRpt, TAB(52); USING "#####.##"; TotalAmt#
     PRINT #UBRpt,
     PRINT #UBRpt,
     PRINT #UBRpt, TAB(12); USING "#####.##"; TotalAmt#;
     PRINT #UBRpt, TAB(29); USING "#####.##"; TotalAmt#;
     PRINT #UBRpt, TAB(41); USING "#####.##"; TotalAmt#;
     PRINT #UBRpt, TAB(52); USING "#####.##"; TotalAmt#
     PRINT #UBRpt, NoticeInfo.MsgLine1

