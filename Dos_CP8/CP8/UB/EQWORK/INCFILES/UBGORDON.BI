'' Gordonsville VA Utility Card Format
'  '050797 LAST CHANGE
'
'    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
'    TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)
'    WRevCnt = 1
'
'    UsageAmt& = UBBillRec(1).CurRead(WRevCnt) - UBBillRec(1).PrevRead(WRevCnt)
'
'    IF UsageAmt& < 0 THEN
'      MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(WRevCnt))) - 1)
'      UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(WRevCnt)) + UBBillRec(1).CurRead(WRevCnt)
'    END IF
'
'    SELECT CASE UBCustRec(1).LocMeters(WRevCnt).MTRMulti
'    CASE 10
'      Zero$ = "0"
'    CASE 100
'      Zero$ = "00"
'    CASE 1000
'      Zero$ = "000"
'    CASE ELSE
'      Zero$ = ""
'    END SELECT
'
'    PRINT #UBRpt, "~"
'    PRINT #UBRpt,
'    PRINT #UBRpt,
'    PRINT #UBRpt,
'    PRINT #UBRpt,
'    PRINT #UBRpt, TAB(2); LEFT$(UBCustRec(1).ServAddr, 25)   'line 5 is serv address
'    PRINT #UBRpt,
'    PRINT #UBRpt, TAB(2); LEFT$(BillDate$, 6); RIGHT$(BillDate$, 2); ' line 7 is billdate and meter read date
'    PRINT #UBRpt, TAB(11); LEFT$(DateRead$, 6); RIGHT$(DateRead$, 2); USING "#########"; UBBillRec(1).CustAcctNo
'    PRINT #UBRpt,
'    PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).CustName, 25)
'    PRINT #UBRpt, TAB(1); USING "########"; UBBillRec(1).PrevRead(WRevCnt);
'    PRINT #UBRpt, Zero$;
'    PRINT #UBRpt, TAB(13); USING "########"; UBBillRec(1).CurRead(WRevCnt);
'    PRINT #UBRpt, Zero$;
'    PRINT #UBRpt, TAB(23); USING "######"; UsageAmt&;
'    PRINT #UBRpt, Zero$;
'    PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).Addr1, 25)
'    PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).Addr2, 25)
'
'    IF UBBillRec(1).RevAmt(1) <> 0 THEN
'       PRINT #UBRpt, " "; LEFT$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 14);
'       PRINT #UBRpt, TAB(22); USING "#####.##"; UBBillRec(1).RevAmt(WRevCnt);
'    END IF
'    PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; LEFT$(UBCustRec(1).ZipCode, 5)
'
'    WRevCnt = 2
'    IF UBBillRec(1).RevAmt(2) <> 0 THEN
'      PRINT #UBRpt, " "; LEFT$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 14);
'      PRINT #UBRpt, TAB(22); USING "#####.##"; UBBillRec(1).RevAmt(WRevCnt)
'    ELSE
'      PRINT #UBRpt,
'    END IF
'
'    WRevCnt = 3
'    IF UBBillRec(1).RevAmt(3) <> 0 THEN
'      PRINT #UBRpt, " "; LEFT$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 14);
'      PRINT #UBRpt, TAB(22); USING "#####.##"; UBBillRec(1).RevAmt(WRevCnt)
'    ELSE
'      PRINT #UBRpt,
'    END IF
'
'    WRevCnt = 4
'    IF UBBillRec(1).RevAmt(4) <> 0 THEN
'      PRINT #UBRpt, " "; LEFT$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 14);
'      PRINT #UBRpt, TAB(22); USING "#####.##"; UBBillRec(1).RevAmt(WRevCnt)
'    ELSE
'      PRINT #UBRpt,
'    END IF
'
'    IF Previous# <> 0 THEN
'      PRINT #UBRpt, " Balance Forward     "; USING "#####.##"; Previous#;
'    END IF
'
'    PRINT #UBRpt, TAB(34); LEFT$(PastDueDate$, 6); RIGHT$(PastDueDate$, 2);
'    PRINT #UBRpt, TAB(47); USING "#####"; UBBillRec(1).CustAcctNo
'
'    IF FinalFlag AND CDeposit# THEN
 '     PRINT #UBRpt, "         Deposit     "; USING "#####.##"; -UBCustRec(1).DepositAmt
'      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
'    ELSE
 '     PRINT #UBRpt,
'    END IF
'
'    PRINT #UBRpt,
'
'    PRINT #UBRpt, TAB(2); LEFT$(BillDate$, 6); RIGHT$(BillDate$, 2);
'    PRINT #UBRpt, TAB(25); USING "####.##"; TotalAmt#; TAB(50); TotalAmt#
'
'    IF TotalAmt# < 0 AND FinalFlag THEN
'      PRINT #UBRpt, "       Refund Due     "; USING "$$####.##"; ABS(TotalAmt#)
'    ELSE
'      PRINT #UBRpt,
'    END IF
'
'    PRINT #UBRpt,
'    PRINT #UBRpt, TAB(34); Message$
'    PRINT #UBRpt, TAB(34); USING "########"; FBillNO& + PrintedCnt
'
'
'
'MUST SHOW BOTH METERS OR, TOTAL CONSUMPTION ON THIS BILL
    CustMsg$ = QPTrim$(UBCustRec(1).BILLCMNT)
    IF NOT LPIFlag THEN
      LPIFlag = -2
      'PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(58);
      PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(77);
      ' put printer in     8 lpi             12 cpi  oki mode
    END IF

    AcctNum = UBBillRec(1).CustAcctNo
    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)
'    IF AcctNum = 3 THEN STOP
    IF TotalAmt# > 0 THEN
      TenPct# = 20  'Round#(UBBillRec(1).TransAmt + 20)
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

    'IF UBCustRec(1).PrevBalance <> 0 OR UBCustRec(1).CurrBalance <> 0 THEN
    'ELSE
    '  Previous# = 0
    'END IF

    PRINT #UBRpt, TAB(50); USING "########"; FBillNO& + PrintedCnt
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, USING "##########"; UBBillRec(1).CustAcctNo;
    PRINT #UBRpt, TAB(15); LEFT$(UBCustRec(1).ServAddr, 19); TAB(50); USING "########"; UBBillRec(1).CustAcctNo;
    PRINT #UBRpt, TAB(62); Num2Date$(UBBillRec(1).BillDate)
    PRINT #UBRpt,

    PRINT #UBRpt, TAB(50); Num2Date$(UBBillRec(1).PastDueDate); TAB(64); USING "#####.##"; TotalAmt#
    PRINT #UBRpt, TAB(3); Num2Date$(UBBillRec(1).BillDate); TAB(15); PrevDate$; TAB(26); DateRead$;
     'Only Print Days if Greater than 0
     IF DaysINRead > 0 THEN
       PRINT #UBRpt, TAB(40); USING "####"; DaysINRead
     ELSE
       PRINT #UBRpt, ""
     END IF

     'IF Lilesville > 0 THEN
       PRINT #UBRpt, TAB(50); Num2Date$(UBBillRec(1).PastDueDate); TAB(64); USING "#####.##"; TotalAmt# + TenPct#
     'ELSE
     '  PRINT #UBRpt, ""
     'END IF

    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,

    'PRINT #UBRpt, STRING$(50, " "); CHR$(27); CHR$(16); "A";
    'PRINT #UBRpt, CHR$(8);
    'PRINT #UBRpt, CHR$(2); CHR$(0);
    'PRINT #UBRpt, CHR$(0); CHR$(2);
    'PRINT #UBRpt, CHR$(1); CHR$(1);
    'PRINT #UBRpt, CHR$(1); CHR$(2);
    'PRINT #UBRpt, CHR$(27); CHR$(16); "B"; CHR$(AcctLen); Acct$

    PCnt = 0
    FOR WRevCnt = 1 TO 7
      PCnt = PCnt + 1
      IF UBBillRec(1).RevAmt(WRevCnt) <> 0 THEN
        PRINT #UBRpt, " "; LEFT$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 5);
        IF UBBillRec(1).CurRead(WRevCnt) > 0 THEN
          UsageAmt& = UBBillRec(1).CurRead(WRevCnt) - UBBillRec(1).PrevRead(WRevCnt)
          IF UsageAmt& < 0 THEN
            MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(WRevCnt))) - 1)
            UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(WRevCnt)) + UBBillRec(1).CurRead(WRevCnt)
          END IF
          PRINT #UBRpt, TAB(7); USING "##########"; UBBillRec(1).CurRead(WRevCnt);
          PRINT #UBRpt, TAB(17); USING "##########"; UBBillRec(1).PrevRead(WRevCnt);
          PRINT #UBRpt, TAB(28); USING "#######"; UsageAmt&;
        END IF
        PRINT #UBRpt, TAB(36); USING "#####.##"; UBBillRec(1).RevAmt(WRevCnt);
      END IF
      SELECT CASE PCnt
      CASE 4
        PRINT #UBRpt, TAB(47); LEFT$(UBCustRec(1).CustName, 29)
      CASE 5
        PRINT #UBRpt, TAB(47); UBCustRec(1).Addr1
      CASE 6
        PRINT #UBRpt, TAB(47); UBCustRec(1).Addr2
      CASE 7
        PRINT #UBRpt, TAB(47); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; LEFT$(UBCustRec(1).ZipCode, 5)
      CASE ELSE
        PRINT #UBRpt,
      END SELECT
    NEXT

    PRINT #UBRpt, ""

    IF Previous# <> 0 THEN
      PRINT #UBRpt, "                  Previous:  "; USING "$$,######.##"; Previous#
    ELSE
      PRINT #UBRpt,
    END IF
    PRINT #UBRpt, "                   Current:  "; USING "$$,######.##"; UBBillRec(1).TransAmt
    PRINT #UBRpt, "                           --------------"

    IF FinalFlag AND CDeposit# THEN
      PRINT #UBRpt, "                   Deposit:  "; USING "$$,######.##"; -UBCustRec(1).DepositAmt
      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
    ELSE
      PRINT #UBRpt,
    END IF

    IF TotalAmt# < 0 AND FinalFlag THEN
      PRINT #UBRpt, "                Refund Due:  "; USING "$$,######.##"; ABS(TotalAmt#)
    ELSE
      'STOP
      PRINT #UBRpt, "                     Total:  "; USING "$$,######.##"; TotalAmt#
    END IF

    PRINT #UBRpt,
    PRINT #UBRpt, "  "; CustMsg$
    PRINT #UBRpt, "  "; Message$
    PRINT #UBRpt,
    PRINT #UBRpt, "~"

