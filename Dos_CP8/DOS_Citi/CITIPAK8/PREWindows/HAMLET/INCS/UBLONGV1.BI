'New Longview Utility Bill format 10-28-96
'
'MUST SHOW BOTH METERS OR, TOTAL CONSUMPTION ON THIS BILL

    IF NOT LPIFlag THEN
'     STOP
      LPIFlag = -2
      PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(58);
      ' put printer in     8 lpi             12 cpi
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
    
    PRINT #UBRpt, TAB(50); USING "########"; FBillNO& + PrintedCnt
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, USING "##########"; UBBillRec(1).CustAcctNo;
    PRINT #UBRpt, TAB(15); LEFT$(UBLocaRec(1).ServAddr, 19); TAB(50); USING "########"; UBBillRec(1).CustAcctNo;
    PRINT #UBRpt, TAB(63); Num2Date$(UBBillRec(1).BillDate)
    PRINT #UBRpt,

    IF UBCustRec(1).PrevBalance <> 0 OR UBCustRec(1).CurrBalance <> 0 THEN
      PTotalAmt# = Round(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    ELSE
      PTotalAmt# = 0
    END IF
    TotalAmt# = Round(PTotalAmt# + UBBillRec(1).TransAmt)

    PRINT #UBRpt, TAB(50); Num2Date$(UBBillRec(1).PastDueDate); TAB(65); USING "#####.##"; TotalAmt#
    PRINT #UBRpt, TAB(3); Num2Date$(UBBillRec(1).BillDate); TAB(15); PrevDate$; TAB(26); DateRead$; TAB(40); USING "####"; DaysINRead
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(50); Num2Date$(UBBillRec(1).PastDueDate); TAB(65); USING "#####.##"; Round(TotalAmt# + 5)
    PRINT #UBRpt,

    PRINT #UBRpt, STRING$(50, " "); CHR$(27); CHR$(16); "A";
    PRINT #UBRpt, CHR$(8);
    PRINT #UBRpt, CHR$(2); CHR$(0);
    PRINT #UBRpt, CHR$(0); CHR$(2);
    PRINT #UBRpt, CHR$(1); CHR$(1);
    PRINT #UBRpt, CHR$(1); CHR$(2);
    PRINT #UBRpt, CHR$(27); CHR$(16); "B"; CHR$(AcctLen); Acct$

    'PTotalAmt# = 0
    PCnt = 0
    FOR WRevCnt = 1 TO 7
      PCnt = PCnt + 1
      'IF WRevCnt < 8 THEN
      '  SELECT CASE UBLocaRec(1).LocMeters(WRevCnt).MTRMulti
      '  CASE 100
      '    Zero$ = "00"
      '  CASE 1000
      '    Zero$ = "000"
      '  CASE ELSE
      '    Zero$ = ""
      '  END SELECT
      'END IF
      IF UBBillRec(1).RevAmt(WRevCnt) > 0 THEN
        PRINT #UBRpt, " "; LEFT$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 5);
        IF UBBillRec(1).CurRead(WRevCnt) > 0 THEN
          UsageAmt& = UBBillRec(1).CurRead(WRevCnt) - UBBillRec(1).PrevRead(WRevCnt)
          IF UsageAmt& < 0 THEN
            MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(WRevCnt))) - 1)
            UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(WRevCnt)) + UBBillRec(1).CurRead(WRevCnt)
          END IF
          PRINT #UBRpt, TAB(7); USING "#######"; UBBillRec(1).CurRead(WRevCnt);
          PRINT #UBRpt, TAB(16); USING "#######"; UBBillRec(1).PrevRead(WRevCnt);
          PRINT #UBRpt, TAB(26); USING "######"; UsageAmt&;
          'PRINT #UBRpt, Zero$;
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
        PRINT #UBRpt, TAB(47); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; UBCustRec(1).ZipCode
      CASE ELSE
        PRINT #UBRpt,
      END SELECT
    NEXT
    
    PRINT #UBRpt, ""

    IF FinalFlag AND UBCustRec(1).DepositAmt = UBBillRec(1).TaxAmt(15) THEN
      DepositAmt# = UBCustRec(1).DepositAmt
      PRINT #UBRpt, "                   Deposit:  "; USING "$$,######.##"; UBCustRec(1).DepositAmt
    ELSE
      DepositAmt# = 0
      PRINT #UBRpt,
    END IF
'******** long view
    IF UBCustRec(1).PrevBalance <> 0 OR UBCustRec(1).CurrBalance <> 0 THEN
      IF PTotalAmt# <> 0 THEN
        PRINT #UBRpt, "                  Previous:  "; USING "$$,######.##"; PTotalAmt#
      ELSE
        PRINT #UBRpt,
      END IF
      PRINT #UBRpt, "                   Current:  "; USING "$$,######.##"; UBBillRec(1).TransAmt
      PRINT #UBRpt, "                           --------------"
    ELSE
      PRINT #UBRpt,
      PRINT #UBRpt, "                   Current:  "; USING "$$,######.##"; UBBillRec(1).TransAmt
      PRINT #UBRpt, "                           --------------"
    END IF
    TotalAmt# = Round((PTotalAmt# + UBBillRec(1).TransAmt) - DepositAmt#)
    IF TotalAmt# < 0 AND FinalFlag THEN
      PRINT #UBRpt, "                Refund Due:  "; USING "$$,######.##"; ABS(TotalAmt#)
    ELSE
      PRINT #UBRpt, "                     Total:  "; USING "$$,######.##"; TotalAmt#
    END IF
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, Message$
    PRINT #UBRpt,
    PRINT #UBRpt,

