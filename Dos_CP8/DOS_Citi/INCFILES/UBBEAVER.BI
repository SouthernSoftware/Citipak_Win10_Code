'
    STAmt# = 0
    DuDate$ = LEFT$(PastDueDate$, 2) + "/" + MID$(PastDueDate$, 4, 2) + "/" + RIGHT$(PastDueDate$, 2)
    ThDate$ = LEFT$(DateRead$, 2) + "/" + MID$(DateRead$, 4, 2) + "/" + RIGHT$(DateRead$, 2)
    IF UBBillRec(1).CurRead(1) >= 0 AND UBBillRec(1).PrevRead(1) >= 0 THEN
      UsageAmt& = UBBillRec(1).CurRead(1) - UBBillRec(1).PrevRead(1)
      IF UsageAmt& < 0 THEN
        MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(1))) - 1)
        UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(1)) + UBBillRec(1).CurRead(1)
      END IF
    ELSE
      UsageAmt& = 0
    END IF

    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    TotalAmt# = Round(UBBillRec(1).TransAmt + Previous#)

    IF TotalAmt# > 0 THEN
      TenPct# = Round#(UBBillRec(1).TransAmt * .1)
    ELSE
      TenPct# = 0
    END IF
    
    PRINT #UBRpt, "~"; USING "#####"; FBillNO& + PrintedCnt
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(2); LEFT$(UBCustRec(1).ServAddr, 26)
    PRINT #UBRpt, TAB(2); USING "#########"; UBBillRec(1).CurRead(1);
    PRINT #UBRpt, TAB(12); USING "#########"; UBBillRec(1).PrevRead(1);
    PRINT #UBRpt, TAB(22); USING "########"; UsageAmt&;
    IF UsageAmt& > 0 THEN
      PRINT #UBRpt, "000"
    ELSE
      PRINT #UBRpt,
    END IF
    PRINT #UBRpt, TAB(40); ThDate$
    PRINT #UBRpt,
    PRINT #UBRpt, USING "#####"; CustAcctNo&;
    PRINT #UBRpt, TAB(16); DuDate$; "  "; ThDate$; TAB(40); USING "#####"; CustAcctNo&
    PRINT #UBRpt,
    IF UBBillRec(1).RevAmt(1) <> 0 THEN
      STAmt# = Round#(STAmt# + UBBillRec(1).RevAmt(1))
      PRINT #UBRpt, TAB(1); UBSetUpRec(1).Revenues(1).RevName;
      PRINT #UBRpt, TAB(16); USING "######.##"; UBBillRec(1).RevAmt(1); TAB(25); STAmt#;
    END IF
    PRINT #UBRpt, TAB(36); LEFT$(UBCustRec(1).CustName, 25)

    IF UBBillRec(1).RevAmt(2) <> 0 THEN
      STAmt# = Round#(STAmt# + UBBillRec(1).RevAmt(2))
      PRINT #UBRpt, TAB(1); UBSetUpRec(1).Revenues(2).RevName;
      PRINT #UBRpt, TAB(16); USING "######.##"; UBBillRec(1).RevAmt(2); TAB(25); STAmt#;
    END IF
    PRINT #UBRpt, TAB(36); LEFT$(UBCustRec(1).Addr1, 25)

    IF UBBillRec(1).RevAmt(3) <> 0 THEN
      STAmt# = Round#(STAmt# + UBBillRec(1).RevAmt(3))
      PRINT #UBRpt, TAB(1); UBSetUpRec(1).Revenues(3).RevName;
      PRINT #UBRpt, TAB(16); USING "######.##"; UBBillRec(1).RevAmt(3); TAB(25); STAmt#;
    END IF
    PRINT #UBRpt, TAB(36); LEFT$(UBCustRec(1).Addr2, 25)

    IF UBBillRec(1).RevAmt(7) <> 0 THEN
      STAmt# = Round#(STAmt# + UBBillRec(1).RevAmt(7))
      PRINT #UBRpt, TAB(1); UBSetUpRec(1).Revenues(7).RevName;
      PRINT #UBRpt, TAB(16); USING "######.##"; UBBillRec(1).RevAmt(7); TAB(25); STAmt#;
    END IF

'    IF UBBillRec(1).RevAmt(4) <> 0 THEN
'      STAmt# = Round#(STAmt# + UBBillRec(1).RevAmt(4))
'      PRINT #UBRpt, TAB(1); UBSetUpRec(1).Revenues(4).RevName;
'      PRINT #UBRpt, TAB(16); USING "######.##"; UBBillRec(1).RevAmt(4); TAB(25); STAmt#;
'    END IF
    PRINT #UBRpt, TAB(36); QPTrim$(UBCustRec(1).City); " "; UBCustRec(1).State; " "; UBCustRec(1).ZipCode

    IF UBBillRec(1).RevAmt(5) <> 0 THEN
      STAmt# = Round#(STAmt# + UBBillRec(1).RevAmt(5))
      PRINT #UBRpt, TAB(1); UBSetUpRec(1).Revenues(5).RevName;
      PRINT #UBRpt, TAB(16); USING "######.##"; UBBillRec(1).RevAmt(5); TAB(25); STAmt#;
    END IF
    PRINT #UBRpt, TAB(36); STRING$(22, "-")
    

'    IF UBBillRec(1).RevAmt(6) <> 0 THEN
'      PRINT #UBRpt, tab(1); UBSetUpRec(1).Revenues(6).RevName;
'      PRINT #UBRpt, tab(16); USING "######.##"; UBBillRec(1).RevAmt(6)
'    ELSE
'      PRINT #UBRpt, ""
'    END IF

    IF Previous# <> 0 THEN
      STAmt# = Round#(STAmt# + Previous#)
      PRINT #UBRpt, TAB(1); "Previous:"; TAB(16); USING "######.##"; Previous#; TAB(25); STAmt#;
    'ELSE
    '  PRINT #UBRpt,
    END IF
    PRINT #UBRpt, TAB(36); Message$


    IF FinalFlag AND CDeposit# THEN
      PRINT #UBRpt, TAB(4); "Deposit:"; TAB(16); USING "######.##"; -UBCustRec(1).DepositAmt;
      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
      PRINT #UBRpt, 'TAB(50); USING "######.##"; TotalAmt#
    ELSE
      PRINT #UBRpt, 'TAB(50); USING "######.##"; TotalAmt#
    END IF
    PRINT #UBRpt,
    PRINT #UBRpt,
    'PRINT #UBRpt, TAB(5); CustAcctNo&; TAB(25); USING "######.##"; TotalAmt#;
    PRINT #UBRpt, USING "######.##"; TAB(15); Round#(TotalAmt# + TenPct#); TAB(25); TotalAmt#;
    PRINT #UBRpt, USING "######.##"; TAB(37); Round#(TotalAmt# + TenPct#); TAB(50); TotalAmt#
    'PRINT #UBRpt, TAB(37); CustAcctNo&; TAB(50); USING "######.##"; TotalAmt#
    PRINT #UBRpt,
    PRINT #UBRpt, "~"

