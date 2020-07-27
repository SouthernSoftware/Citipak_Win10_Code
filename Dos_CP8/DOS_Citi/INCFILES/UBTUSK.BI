'New Utility Bill format 10-28-96 BAR CODE PRINTABLE
'
'MUST SHOW BOTH METERS OR, TOTAL CONSUMPTION ON THIS BILL

    IF NOT LPIFlag THEN
      LPIFlag = -2
      'PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(58);
      PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(77);
      ' put printer in     8 lpi             12 cpi  oki mode
    END IF
'***********************************************************************
    WFoundMtr = False

    FOR mChk = 1 TO 7
      IF UBBillRec(1).MtrTypes(mChk) > 0 THEN
        SELECT CASE UBBillRec(1).MtrTypes(mChk)
        CASE 1, 2, 3
          WCurrRead& = UBBillRec(1).CurRead(mChk)
          WPrevRead& = UBBillRec(1).PrevRead(mChk)
          WUsageAmt& = WCurrRead& - WPrevRead&
          IF WUsageAmt& < 0 THEN
            MaxMeterAmt& = 10& ^ (LEN(STR$(WPrevRead&)) - 1)
            WUsageAmt& = (MaxMeterAmt& - WPrevRead&) + WCurrRead&
          END IF
          WFoundMtr = True
        END SELECT
      END IF
    NEXT
    
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

'**********************************************************************
    
    AcctNum = UBBillRec(1).CustAcctNo
    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)

    IF TotalAmt# > 0 THEN
      LessAmt# = Round#(UBCustRec(1).CurrRevAmts(7) + UBCustRec(1).PrevRevAmts(7))
      IF LessAmt# >= 0 THEN
        PenAmt# = Round#(TotalAmt# - LessAmt#)
      ELSE
        PenAmt# = Round#(TotalAmt# + LessAmt#)
      END IF

      'PenAmt# = Round(PenAmt# + UBBillRec(1).RevAmt(1) + UBBillRec(1).RevAmt(5))
      IF PenAmt# > 0 THEN
        TenPct# = Round#(PenAmt# * .1)
      ELSE
        TenPct# = 0
      END IF
      IF TenPct# > 0 THEN
        IF TenPct# < 10 THEN
          TenPct# = 10
        END IF
      END IF
    ELSE
      TenPct# = 0
    END IF

'    IF TotalAmt# > 0 THEN
'      TenPct# = Round#(UBBillRec(1).TransAmt * .1)
'    END IF

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

    PRINT #UBRpt, TAB(50); Num2Date$(UBBillRec(1).PastDueDate);
    IF FinalFlag THEN
      PRINT #UBRpt, TAB(64); USING "#####.##"; Round#(TotalAmt#)
    ELSE
      PRINT #UBRpt, TAB(64); USING "#####.##"; Round#(TotalAmt# + TenPct#)
    END IF
    
    'PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
'****************************************
    PCnt = 0
    FOR WRevCnt = 1 TO 7
      PCnt = PCnt + 1
      IF UBBillRec(1).RevAmt(WRevCnt) <> 0 THEN
        PRINT #UBRpt, " "; LEFT$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 5);
        SELECT CASE PCnt
        CASE 1, 2
          IF WFoundMtr THEN
            PRINT #UBRpt, TAB(7); USING "##########"; WCurrRead&;
            PRINT #UBRpt, TAB(17); USING "##########"; WPrevRead&;
            PRINT #UBRpt, TAB(28); USING "#######"; WUsageAmt&;
          END IF
        END SELECT
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
    PRINT #UBRpt, TAB(2); Message$
    PRINT #UBRpt, TAB(2); Msg2$
    PRINT #UBRpt, TAB(2); Msg3$
    PRINT #UBRpt, TAB(2); Msg4$
    PRINT #UBRpt, "~"

