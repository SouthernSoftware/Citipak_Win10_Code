'New Utility Bill format 10-28-96 BAR CODE PRINTABLE
'
'MUST SHOW BOTH METERS OR, TOTAL CONSUMPTION ON THIS BILL

    IF NOT LPIFlag THEN
      LPIFlag = -2
      'PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(58);
      PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(77);
      ' put printer in     8 lpi             12 cpi  oki mode
    END IF

    AcctNum = UBBillRec(1).CustAcctNo
    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)

    IF TotalAmt# > 0 THEN
      TenPct# = Round#(UBBillRec(1).TransAmt * .1)
    END IF

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

    PRINT #UBRpt, TAB(50); Num2Date$(UBBillRec(1).PastDueDate); TAB(64); USING "#####.##"; ABS(TotalAmt#)
    PRINT #UBRpt, TAB(3); Num2Date$(UBBillRec(1).BillDate); TAB(15); PrevDate$; TAB(26); DateRead$;
     'Only Print Days if Greater than 0
     IF DaysINRead > 0 THEN
       PRINT #UBRpt, TAB(40); USING "####"; DaysINRead
     ELSE
       PRINT #UBRpt, ""
     END IF

    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,

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
          PRINT #UBRpt, TAB(18); USING "##########"; UBBillRec(1).PrevRead(WRevCnt);
          PRINT #UBRpt, TAB(28); USING "######"; UsageAmt&;
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
      PRINT #UBRpt, "                     Total:  "; USING "$$,######.##"; TotalAmt#
    END IF
    PRINT #UBRpt, ""
    IF Previous# > 10 THEN
      PRINT #UBRpt, TAB(2); "Disconnect Notice"
      PRINT #UBRpt, TAB(2); "Total Payment due by "; PastDueDate$
      PRINT #UBRpt, TAB(2); "Reconnect Fee $40.00"
    ELSE
      PRINT #UBRpt,
      PRINT #UBRpt, TAB(2); "Payment due by "; PastDueDate$
      PRINT #UBRpt, TAB(2); Message$
    END IF
    PRINT #UBRpt,
    PRINT #UBRpt,
    

'    PRINT #UBRpt, "  "; Message$
'    PRINT #UBRpt,
'    PRINT #UBRpt, "~"


'050797 Tuskaseigee NC Water bill format

'    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
'    TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)
'
'    TheRest# = 0
'    FOR ThisRevCnt = 3 TO 15
'      IF UBBillRec(1).RevAmt(ThisRevCnt) <> 0 THEN
'        TheRest# = Round#(TheRest# + UBBillRec(1).RevAmt(ThisRevCnt))
'      END IF
'    NEXT
'
'    WRevCnt = 1
'
'    UsageAmt& = UBBillRec(1).CurRead(WRevCnt) - UBBillRec(1).PrevRead(WRevCnt)
'    IF UsageAmt& < 0 THEN
'      MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(WRevCnt))) - 1)
'      UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(WRevCnt)) + UBBillRec(1).CurRead(WRevCnt)
'    END IF
'
'    PRINT #UBRpt, "~"; TAB(26); USING "########"; FBillNO& + PrintedCnt
'    PRINT #UBRpt, ""
'    PRINT #UBRpt, ""
'    PRINT #UBRpt, ""
'    PRINT #UBRpt, ""
'    PRINT #UBRpt, ""
'    PRINT #UBRpt, ""
'    PRINT #UBRpt, ""
'    PRINT #UBRpt, TAB(3); USING "########"; UBBillRec(1).PrevRead(1); TAB(15); UBBillRec(1).CurRead(1); TAB(25); UsageAmt&
'    PRINT #UBRpt, TAB(37); LEFT$(UBCustRec(1).CustName, 25)
'    PRINT #UBRpt, TAB(37); LEFT$(UBCustRec(1).Addr1, 25)
'    PRINT #UBRpt, TAB(37); LEFT$(UBCustRec(1).Addr2, 25)
'    PRINT #UBRpt, TAB(37); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; LEFT$(UBCustRec(1).ZipCode, 5)
'
'    PRINT #UBRpt, TAB(1); USING "####.##"; UBBillRec(1).RevAmt(1);
'    PRINT #UBRpt, TAB(11); USING "####.##"; UBBillRec(1).RevAmt(2);
'    PRINT #UBRpt, TAB(19); USING "####.##"; TheRest#;
'    PRINT #UBRpt, TAB(28); USING "####.##"; Previous#
'
'    PRINT #UBRpt,
'    PRINT #UBRpt,
'    PRINT #UBRpt,
'    PRINT #UBRpt,
'
'    PRINT #UBRpt, TAB(3); LEFT$(BillDate$, 6); RIGHT$(BillDate$, 2);
'    PRINT #UBRpt, TAB(14); USING "########"; UBBillRec(1).CustAcctNo;
'    PRINT #UBRpt, TAB(27); USING "#####.##"; TotalAmt#;
'    PRINT #UBRpt, TAB(37); LEFT$(BillDate$, 6); RIGHT$(BillDate$, 2)
'    PRINT #UBRpt,
'    PRINT #UBRpt,
'
'    IF Previous# > 10 THEN
'      PRINT #UBRpt, TAB(2); "Disconnect Notice"; TAB(35); USING "########"; UBBillRec(1).CustAcctNo;
'      PRINT #UBRpt, TAB(47); USING "#####,#.##"; TotalAmt#
'      PRINT #UBRpt, TAB(2); "Total Payment due by "; PastDueDate$
'      PRINT #UBRpt, TAB(2); "Reconnect Fee $40.00"
'    ELSE
'      PRINT #UBRpt, TAB(2); TAB(35); USING "########"; UBBillRec(1).CustAcctNo;
'      PRINT #UBRpt, TAB(47); USING "#####,#.##"; TotalAmt#
'      PRINT #UBRpt, TAB(2); "Payment due by "; PastDueDate$
'      PRINT #UBRpt, TAB(2); Message$
'    END IF

