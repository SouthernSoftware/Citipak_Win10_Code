'Appalachia VA Bill format 01-26-98

    Loca$ = UBCustRec(1).Book + "-" + UBCustRec(1).SeqNumb
    CustMsg$ = QPTrim$(UBCustRec(1).BILLCMNT)

    IF NOT LPIFlag THEN
      LPIFlag = -2
      PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(58);
      'PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(77);
      ' put printer in     8 lpi
    END IF

    FoundAMtr = False
    FOR mChk = 1 TO 7
      IF UBBillRec(1).MtrTypes(mChk) > 0 THEN
        FoundAMtr = True
        UsageAmt& = UBBillRec(1).CurRead(mChk) - UBBillRec(1).PrevRead(mChk)
        CRead# = UBBillRec(1).CurRead(mChk)
        PRead# = UBBillRec(1).PrevRead(mChk)
        IF UsageAmt& < 0 THEN
          MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(mChk))) - 1)
          UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(mChk)) + UBBillRec(1).CurRead(mChk)
        END IF
        EXIT FOR
      END IF
    NEXT

    IF FoundAMtr = False THEN
      'if no metered services then adjust read dates to billdate
      'and billdate - 30
      DateRead$ = Num2Date$(UBBillRec(1).BillDate)
      PrevDate$ = Num2Date$(UBBillRec(1).BillDate - 30)
    END IF

    AcctNum = UBBillRec(1).CustAcctNo

    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)

    IF TotalAmt# > 0 THEN
      TenPct# = 0
      'TenPct# = Round#(UBBillRec(1).TransAmt * .1)
    ELSE
      TenPct# = 0
    END IF

    'AcctNum = AcctNo&
    Acct$ = QPTrim$(STR$(AcctNum))
    SELECT CASE AcctNum
    CASE IS < 10
      Acct$ = "00" + Acct$
    CASE IS < 100
      Acct$ = "0" + Acct$
    END SELECT
    AcctLen = LEN(Acct$)

    PMth$ = LEFT$(PrevDate$, 2)
    PDay$ = MID$(PrevDate$, 4, 2)
    TMth$ = LEFT$(DateRead$, 2)
    TDay$ = MID$(DateRead$, 4, 2)
    T$ = Num2Date$(UBBillRec(1).PastDueDate)
    DMth$ = LEFT$(T$, 2)
    DDay$ = MID$(T$, 4, 2)
    DYer$ = RIGHT$(T$, 4)

    PRINT #UBRpt, TAB(27); USING "####"; FBillNO& + PrintedCnt
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, ""  '11
    PRINT #UBRpt, TAB(2); PMth$; TAB(8); PDay$; TAB(14); TMth$; TAB(20); TDay$;
    PRINT #UBRpt, TAB(26); DMth$; TAB(32); DDay$; TAB(38); DYer$; TAB(44); "PRE-SORT"; "  "; Loca$
    PRINT #UBRpt, TAB(54); USING "#########"; AcctNum
    PRINT #UBRpt, TAB(44); LEFT$(UBCustRec(1).CustName, 23)
    PRINT #UBRpt, TAB(44); LEFT$(UBCustRec(1).Addr1, 23)
    PRINT #UBRpt, TAB(44); LEFT$(UBCustRec(1).Addr2, 23)
    PRINT #UBRpt, TAB(44); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; LEFT$(UBCustRec(1).ZipCode, 5)

    IF UBBillRec(1).RevAmt(1) <> 0 THEN
      IF FoundAMtr THEN
        PRINT #UBRpt, USING "#########"; PRead#;
        PRINT #UBRpt, TAB(10); USING "#########"; CRead#; UsageAmt&;
      END IF
      PRINT #UBRpt, TAB(29); LEFT$(UBSetUpRec(1).Revenues(1).RevName, 3);
      PRINT #UBRpt, TAB(33); USING "#####.##"; UBBillRec(1).RevAmt(1)
    ELSE
      PRINT #UBRpt, 'line 17
    END IF

    IF UBBillRec(1).RevAmt(2) <> 0 THEN
      IF FoundAMtr THEN
        PRINT #UBRpt, USING "#########"; PRead#;
        PRINT #UBRpt, TAB(10); USING "#########"; CRead#; UsageAmt&;
      END IF
      PRINT #UBRpt, TAB(29); LEFT$(UBSetUpRec(1).Revenues(2).RevName, 3);
      PRINT #UBRpt, TAB(33); USING "#####.##"; UBBillRec(1).RevAmt(2)
    ELSE
      PRINT #UBRpt, 'line 18
    END IF

    FOR RevCnt = 3 TO 7
      IF UBBillRec(1).RevAmt(RevCnt) <> 0 THEN
        PRINT #UBRpt, TAB(29); LEFT$(UBSetUpRec(1).Revenues(RevCnt).RevName, 3);
        PRINT #UBRpt, TAB(33); USING "#####.##"; UBBillRec(1).RevAmt(RevCnt)
      ELSE
        PRINT #UBRpt,
      END IF
    NEXT
    
    'PRINT #UBRpt,   '"line 23"

    IF Previous# <> 0 THEN
      PRINT #UBRpt, TAB(25); USING "BAL FWD #####.##"; Previous#
    ELSE
      PRINT #UBRpt,
    END IF

    PRINT #UBRpt,   '25
    PRINT #UBRpt,   '26
    PRINT #UBRpt, TAB(17); Loca$; TAB(29); USING "$$,######.##"; TotalAmt#; TAB(44); TotalAmt#

    'IF FinalFlag AND CDeposit# THEN
    '  PRINT #UBRpt, "                   Deposit:  "; USING "$$,######.##"; -UBCustRec(1).DepositAmt
    '  TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
    'ELSE
    '  PRINT #UBRpt,
    'END IF
    
    'IF TotalAmt# < 0 AND FinalFlag THEN
    '  PRINT #UBRpt, "                Refund Due:  "; USING "$$,######.##"; ABS(TotalAmt#)
    'ELSE
    '  'STOP
    '  PRINT #UBRpt, "                     Total:  "; USING "$$,######.##"; TotalAmt#
    'END IF

    PRINT #UBRpt, '27
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, "  "; CustMsg$
    PRINT #UBRpt, "  "; Message$
    PRINT #UBRpt, "  "; Msg2$
    PRINT #UBRpt, "~"

