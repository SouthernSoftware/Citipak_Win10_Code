'New Utility Bill format 10-28-96 BAR CODE PRINTABLE
'MUST SHOW BOTH METERS OR, TOTAL CONSUMPTION ON THIS BILL

    CustMsg$ = QPTrim$(UBCustRec(1).BILLCMNT)

    IF NOT LPIFlag THEN
      LPIFlag = -2
      PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(58);
      'PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(77);
      ' put printer in     8 lpi             12 cpi  oki mode
    END IF
                                                      ':  M
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

    AcctNum = UBBillRec(1).CustAcctNo
    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)

    IF FinalFlag AND CDeposit# THEN
      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
    END IF

    IF TotalAmt# > 0 THEN
      TenPct# = Round#(UBBillRec(1).TransAmt * .1)
    ELSE
      TenPct# = 0
    END IF

    AcctNum = UBBillRec(1).CustAcctNo
    Acct$ = QPTrim$(STR$(AcctNum))
    SELECT CASE AcctNum
    CASE IS < 10
      Acct$ = "00" + Acct$
    CASE IS < 100
      Acct$ = "0" + Acct$
    END SELECT
    AcctLen = LEN(Acct$)

    Acct$ = QPTrim$(UBCustRec(1).ZipCode)
    AcctLen = LEN(Acct$)

    PRINT #UBRpt, "~"; TAB(50); USING "########"; FBillNO& + PrintedCnt
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(18); PrevDate$; TAB(33); DateRead$;
     'Only Print Days if Greater than 0
     IF DaysINRead > 0 THEN
       PRINT #UBRpt, USING "####"; DaysINRead
     ELSE
       PRINT #UBRpt, ""
     END IF

    PRINT #UBRpt,
    PRINT #UBRpt,

    PCnt = 0
    FOR WRevCnt = 1 TO 7
      PCnt = PCnt + 1
      IF UBBillRec(1).RevAmt(WRevCnt) <> 0 THEN
        PRINT #UBRpt, LEFT$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 3);
        IF UBBillRec(1).CurRead(WRevCnt) > 0 THEN
          UsageAmt& = UBBillRec(1).CurRead(WRevCnt) - UBBillRec(1).PrevRead(WRevCnt)
          IF UsageAmt& < 0 THEN
            MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(WRevCnt))) - 1)
            UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(WRevCnt)) + UBBillRec(1).CurRead(WRevCnt)
          END IF
          PRINT #UBRpt, TAB(4); USING "#########"; UBBillRec(1).PrevRead(WRevCnt);
          PRINT #UBRpt, TAB(14); USING "#########"; UBBillRec(1).CurRead(WRevCnt);
          PRINT #UBRpt, TAB(25); USING "#######"; UsageAmt&;
        END IF
        PRINT #UBRpt, TAB(33); USING "#####.##"; UBBillRec(1).RevAmt(WRevCnt);
      END IF
      SELECT CASE PCnt
      CASE 1
        PRINT #UBRpt, TAB(44); USING "##########"; UBBillRec(1).CustAcctNo
      CASE 5
        PRINT #UBRpt, TAB(49); LEFT$(UBCustRec(1).ServAddr, 26)
      CASE ELSE
        PRINT #UBRpt,
      END SELECT
    NEXT

    PRINT #UBRpt,
    PRINT #UBRpt, TAB(45); Num2Date$(UBBillRec(1).BillDate); TAB(60); Num2Date$(UBBillRec(1).PastDueDate)

    IF TotalTax# > 0 THEN
      PRINT #UBRpt, TAB(14); USING "     TAX:       $$,######.##"; TotalTax#
    ELSE
      PRINT #UBRpt, ""
    END IF

    IF Previous# <> 0 THEN
      PRINT #UBRpt, TAB(14); USING "Previous:       $$,######.##"; Previous#
    ELSE
      PRINT #UBRpt,
    END IF

    PRINT #UBRpt, TAB(14); USING " Current:       $$,######.##"; UBBillRec(1).TransAmt
    PRINT #UBRpt,

    IF FinalFlag AND CDeposit# THEN
      PRINT #UBRpt, TAB(14); USING " Deposit:       $$,######.##"; -UBCustRec(1).DepositAmt
    ELSE
      PRINT #UBRpt,
    END IF
    PRINT #UBRpt,

    IF TotalAmt# < 0 AND FinalFlag THEN
      PRINT #UBRpt, TAB(34); USING "Refund: $$,######.##"; ABS(TotalAmt#)
    ELSE
      PRINT #UBRpt, TAB(34); USING " Total: $$,######.##"; TotalAmt#
    END IF

    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,

    PRINT #UBRpt, USING "##########"; UBBillRec(1).CustAcctNo;
    PRINT #UBRpt, TAB(15); CustMsg$
    PRINT #UBRpt, TAB(15); Message$
    PRINT #UBRpt, TAB(15); Msg2$; TAB(47); LEFT$(UBCustRec(1).CustName, 29)
    PRINT #UBRpt, USING "#######.##"; TotalAmt#;
    PRINT #UBRpt, TAB(15); Msg3$; TAB(47); UBCustRec(1).Addr1
    PRINT #UBRpt, TAB(15); Msg4$; TAB(47); UBCustRec(1).Addr2
    PRINT #UBRpt, TAB(47); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; LEFT$(UBCustRec(1).ZipCode, 5)

    'PRINT #UBRpt, STRING$(20, " "); CHR$(27); CHR$(16); "A";
    'PRINT #UBRpt, CHR$(8);
    'PRINT #UBRpt, CHR$(2); CHR$(0);
    'PRINT #UBRpt, CHR$(0); CHR$(2);
    'PRINT #UBRpt, CHR$(1); CHR$(1);
    'PRINT #UBRpt, CHR$(1); CHR$(1);
    'PRINT #UBRpt, CHR$(27); CHR$(16); "B"; CHR$(AcctLen); Acct$

    IF FinalFlag THEN
      PRINT #UBRpt, USING "#######.##"; Round#(TotalAmt#)
    ELSE
      PRINT #UBRpt, USING "#######.##"; Round#(TotalAmt# + TenPct#)
    END IF
    IF DidADraftFlag THEN
      PRINT #UBRpt, TAB(34); "Your account well be drafted."
    ELSE
      PRINT #UBRpt,
    END IF
    PRINT #UBRpt,




'  This is George Town
'073198 LAST CHANGE
'    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
'    TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)
'    TenPercentAmount# = Round#(UBBillRec(1).TransAmt * .1)
'    WRevCnt = 1
'    UsageAmt& = UBBillRec(1).CurRead(WRevCnt) - UBBillRec(1).PrevRead(WRevCnt)
'    IF UsageAmt& < 0 THEN
'      MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(WRevCnt))) - 1)
'      UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(WRevCnt)) + UBBillRec(1).CurRead(WRevCnt)
'    END IF
'    IF TotalAmt# <= 0 THEN
'      TenPercentAmount# = 0
'    END IF
'     PRINT #UBRpt, "~"
'     PRINT #UBRpt,
'     PRINT #UBRpt,
'     PRINT #UBRpt,
'     PRINT #UBRpt,
'     PRINT #UBRpt, TAB(2); LEFT$(UBCustRec(1).ServAddr, 25)   'line 5 is serv address
'     PRINT #UBRpt,
'     PRINT #UBRpt, TAB(2); LEFT$(BillDate$, 6); RIGHT$(BillDate$, 2); ' line 7 is billdate and meter read date
'     PRINT #UBRpt, TAB(11); LEFT$(DateRead$, 6); RIGHT$(DateRead$, 2); USING "#########"; UBBillRec(1).CustAcctNo
'     PRINT #UBRpt,
'     PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).CustName, 25)
'     PRINT #UBRpt, TAB(2); USING "########"; UBBillRec(1).PrevRead(WRevCnt);
'     PRINT #UBRpt, TAB(12); USING "#########"; UBBillRec(1).CurRead(WRevCnt);
'     PRINT #UBRpt, TAB(22); USING "########"; UsageAmt&;
'     PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).Addr1, 25)
'     PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).Addr2, 25)
'     IF UBBillRec(1).RevAmt(1) <> 0 THEN
'       PRINT #UBRpt, " "; LEFT$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 14);
'       PRINT #UBRpt, TAB(22); USING "#####.##"; UBBillRec(1).RevAmt(WRevCnt);
'     END IF
'     PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; LEFT$(UBCustRec(1).ZipCode, 5)
'     WRevCnt = 2
'     IF UBBillRec(1).RevAmt(2) <> 0 THEN
'       PRINT #UBRpt, " "; LEFT$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 14);
'       PRINT #UBRpt, TAB(22); USING "#####.##"; UBBillRec(1).RevAmt(WRevCnt)
'     ELSE
'       PRINT #UBRpt,
'     END IF
'     WRevCnt = 3
'     IF UBBillRec(1).RevAmt(3) <> 0 THEN
'       PRINT #UBRpt, " "; LEFT$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 14);
'       PRINT #UBRpt, TAB(22); USING "#####.##"; UBBillRec(1).RevAmt(WRevCnt);
'     END IF
'     IF DidADraftFlag THEN
'       PRINT #UBRpt, TAB(34); "Account Drafted"
'       'STOP
'     ELSE
'       PRINT #UBRpt,
'     END IF
'     WRevCnt = 4
'     IF UBBillRec(1).RevAmt(4) <> 0 THEN
'       PRINT #UBRpt, " "; LEFT$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 14);
'       PRINT #UBRpt, TAB(22); USING "#####.##"; UBBillRec(1).RevAmt(WRevCnt)
'     ELSE
'       PRINT #UBRpt,
'     END IF
'    IF Previous# <> 0 THEN
'      PRINT #UBRpt, " Balance Forward     "; ; USING "#####.##"; Previous#;
'    END IF
'    PRINT #UBRpt, TAB(34); LEFT$(PastDueDate$, 6); RIGHT$(PastDueDate$, 2); TAB(47); USING "#####"; UBBillRec(1).CustAcctNo
'    IF FinalFlag AND CDeposit# THEN
'      PRINT #UBRpt, "         Deposit     "; USING "#####.##"; -UBCustRec(1).DepositAmt
'      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
'    ELSE
'      PRINT #UBRpt,
'    END IF
'    PRINT #UBRpt,
'    PRINT #UBRpt, TAB(2); LEFT$(PastDueDate$, 6); RIGHT$(PastDueDate$, 2);
'    IF TotalAmt# < 0 THEN
'      PRINT #UBRpt, TAB(15); USING "#####.##"; TotalAmt#;
'    ELSE
'      PRINT #UBRpt, TAB(18); USING "####.##"; Round#(TotalAmt# + TenPercentAmount#);
'    END IF
'    PRINT #UBRpt, TAB(25); USING "####.##"; TotalAmt#;
'    IF TotalAmt# < 0 AND FinalFlag THEN
'      PRINT #UBRpt, TAB(40); USING "#####.##"; TotalAmt#;
'    ELSE
'      PRINT #UBRpt, TAB(40); USING "#####.##"; Round#(TotalAmt# + TenPercentAmount#);
'    END IF
'    PRINT #UBRpt, TAB(48); USING "#####.##"; TotalAmt#
'    IF TotalAmt# < 0 AND FinalFlag THEN
'      PRINT #UBRpt, "       Refund Due     "; USING "$$####.##"; ABS(TotalAmt#)
'    ELSE
'      PRINT #UBRpt,
'    END IF
'    PRINT #UBRpt, TAB(2); QPTrim$(Message$)
'    PRINT #UBRpt, TAB(2); QPTrim$(Msg2$); TAB(34); USING "########"; (FBillNO& + PrintedCnt)
'    PRINT #UBRpt, "~"

