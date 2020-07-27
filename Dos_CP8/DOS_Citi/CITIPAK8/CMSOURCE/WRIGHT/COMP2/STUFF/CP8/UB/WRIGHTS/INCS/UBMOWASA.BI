
    'CustMsg$ = QPTrim$(UBCustRec(1).BILLCMNT)
    FoundAMtr = False

    FOR mChk = 1 TO 7
      IF UBBillRec(1).MtrTypes(mChk) > 0 THEN
        FoundAMtr = True
        EXIT FOR
      END IF
    NEXT

    IF FoundAMtr = False THEN
      'if no metered services then adjust read dates to billdate
      'and billdate - 30
      DateRead$ = Num2Date$(UBBillRec(1).BillDate)
      PrevDate$ = Num2Date$(UBBillRec(1).BillDate - 30)
    END IF

    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)

    IF FinalFlag AND CDeposit# THEN
      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
    END IF

    IF TotalAmt# > 0 THEN
      TenPct# = 0
      IF DaysINRead < 1 THEN DaysINRead = 1
      AvgCst# = Round#(TotalAmt# / DaysINRead)
    ELSE
      TenPct# = 0
      AvgCst# = 0
    END IF
   
    PRINT #UBRpt, TAB(5); USING "########"; FBillNO& + PrintedCnt
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(30); LEFT$(UBCustRec(1).CustName, 29)
    PRINT #UBRpt, TAB(30); UBCustRec(1).Addr1
    PRINT #UBRpt, TAB(30); UBCustRec(1).Addr2
    PRINT #UBRpt, TAB(30); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; LEFT$(UBCustRec(1).ZipCode, 5)
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(5); Num2Date$(UBBillRec(1).BillDate); TAB(17); Num2Date$(UBBillRec(1).PastDueDate);
    PRINT #UBRpt, TAB(44); USING "$#####.##"; TotalAmt#
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(14); USING "######"; UBBillRec(1).CustAcctNo;
    IF DidADraftFlag THEN
      PRINT #UBRpt, TAB(38); "Draft Date: "; Num2Date$(UBBillRec(1).DraftDate)
    ELSE
      PRINT #UBRpt,
    END IF
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    TempName$ = QPTrim$(UBCustRec(1).CustName)
    PRINT #UBRpt, LEFT$(TempName$, 23); TAB(26); LEFT$(UBCustRec(1).ServAddr, 23); TAB(54); USING "######"; UBBillRec(1).CustAcctNo
    PRINT #UBRpt,
    PRINT #UBRpt,
    
    IF UBBillRec(1).RevAmt(1) <> 0 THEN
      UsageAmt& = UBBillRec(1).CurRead(1) - UBBillRec(1).PrevRead(1)
      IF UsageAmt& < 0 THEN
        MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(1))) - 1)
        UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(1)) + UBBillRec(1).CurRead(1)
      END IF
      IF UBCustRec(1).LocMeters(1).MTRUnit = "C" THEN
        UsageAmt& = UsageAmt& * 7.481
      END IF
      PRINT #UBRpt, TAB(1); USING "##########"; UBBillRec(1).CurRead(1);
      PRINT #UBRpt, TAB(11); USING "##########"; UBBillRec(1).PrevRead(1);
      PRINT #UBRpt, TAB(24); USING "#######"; UsageAmt&;
      PRINT #UBRpt, TAB(33); LEFT$(UBSetUpRec(1).Revenues(1).RevName, 3);
      PRINT #UBRpt, TAB(37); USING "######.##"; UBBillRec(1).RevAmt(1)
    ELSE
      PRINT #UBRpt,
    END IF

    IF UBBillRec(1).RevAmt(2) <> 0 THEN
      UsageAmt& = UBBillRec(1).CurRead(1) - UBBillRec(1).PrevRead(1)
      IF UsageAmt& < 0 THEN
        MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(1))) - 1)
        UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(1)) + UBBillRec(1).CurRead(1)
      END IF
      IF UBCustRec(1).LocMeters(1).MTRUnit = "C" THEN
        UsageAmt& = UsageAmt& * 7.481
      END IF
      PRINT #UBRpt, TAB(1); USING "##########"; UBBillRec(1).CurRead(1);
      PRINT #UBRpt, TAB(11); USING "##########"; UBBillRec(1).PrevRead(1);
      PRINT #UBRpt, TAB(24); USING "#######"; UsageAmt&;
      PRINT #UBRpt, TAB(33); LEFT$(UBSetUpRec(1).Revenues(2).RevName, 3);
      PRINT #UBRpt, TAB(37); USING "######.##"; UBBillRec(1).RevAmt(2)
    ELSE
      PRINT #UBRpt,
    END IF

    FOR RevCnt = 3 TO 4
      IF UBBillRec(1).RevAmt(RevCnt) <> 0 THEN
        PRINT #UBRpt, TAB(33); LEFT$(UBSetUpRec(1).Revenues(RevCnt).RevName, 3);
        PRINT #UBRpt, TAB(37); USING "######.##"; UBBillRec(1).RevAmt(RevCnt)
      ELSE
        PRINT #UBRpt,
      END IF
    NEXT

    IF UBBillRec(1).RevAmt(5) <> 0 THEN
      UsageAmt& = UBBillRec(1).CurRead(2) - UBBillRec(1).PrevRead(2)
      IF UsageAmt& < 0 THEN
        MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(2))) - 1)
        UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(2)) + UBBillRec(1).CurRead(2)
      END IF
      IF UBCustRec(1).LocMeters(2).MTRUnit = "C" THEN
        UsageAmt& = UsageAmt& * 7.481
      END IF
      PRINT #UBRpt, TAB(1); USING "##########"; UBBillRec(1).CurRead(2);
      PRINT #UBRpt, TAB(12); USING "##########"; UBBillRec(1).PrevRead(2);
      PRINT #UBRpt, TAB(24); USING "#######"; UsageAmt&;
      PRINT #UBRpt, TAB(33); LEFT$(UBSetUpRec(1).Revenues(5).RevName, 3);
      PRINT #UBRpt, TAB(37); USING "######.##"; UBBillRec(1).RevAmt(5)
    ELSE
      PRINT #UBRpt,
    END IF

    FOR RevCnt = 6 TO 7
      IF UBBillRec(1).RevAmt(RevCnt) <> 0 THEN
        PRINT #UBRpt, TAB(33); LEFT$(UBSetUpRec(1).Revenues(RevCnt).RevName, 3);
        PRINT #UBRpt, TAB(37); USING "######.##"; UBBillRec(1).RevAmt(RevCnt)
      ELSE
        PRINT #UBRpt,
      END IF
    NEXT

    IF Previous# <> 0 THEN
      PRINT #UBRpt, "  Previous:"; TAB(37); USING "######.##"; Previous#
    ELSE
      PRINT #UBRpt,
    END IF

    IF FinalFlag AND CDeposit# THEN
      PRINT #UBRpt, "  Deposit:  "; TAB(37); USING "######.##"; -UBCustRec(1).DepositAmt
    ELSE
      PRINT #UBRpt,
    END IF

    PRINT #UBRpt, Message$
    PRINT #UBRpt,

    PRINT #UBRpt, LEFT$(BillDate$, 6); RIGHT$(BillDate$, 2);
    PRINT #UBRpt, TAB(10); LEFT$(PastDate$, 6); RIGHT$(PastDate$, 2);
    PRINT #UBRpt, TAB(19); LEFT$(DateRead$, 6); RIGHT$(DateRead$, 2);
    PRINT #UBRpt, TAB(28); USING "##"; DaysINRead;
    PRINT #UBRpt, TAB(30); USING "###.##"; AvgCst#;
    PRINT #UBRpt, TAB(36); USING "$$#####.##"; TotalAmt#
    PRINT #UBRpt,
    PRINT #UBRpt, "~"


