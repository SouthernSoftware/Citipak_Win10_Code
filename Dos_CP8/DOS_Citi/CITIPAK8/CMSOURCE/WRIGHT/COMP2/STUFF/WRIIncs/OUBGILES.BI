'Giles  03-21-97
'6 lpi

    'IF NOT LPIFlag THEN
    '  LPIFlag = -2
    '  PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(58);
    '  'PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(77);
    '  ' put printer in     8 lpi             12 cpi  oki mode
    'END IF

    AcctNum = UBBillRec(1).CustAcctNo
    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)
    
    PRINT #UBRpt, USING "########"; FBillNO& + PrintedCnt
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    '5
    PRINT #UBRpt, TAB(8); LEFT$(UBCustRec(1).ServAddr, 25)
    PRINT #UBRpt,
    '7
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, Num2Date$(UBBillRec(1).BillDate); TAB(12); PrevDate$; TAB(23); DateRead$; TAB(34); Num2Date$(UBBillRec(1).PastDueDate)
    '10

    PRINT #UBRpt, TAB(47); LEFT$(UBCustRec(1).CustName, 29)
    PRINT #UBRpt, TAB(47); UBCustRec(1).Addr1
    PCnt = 0
    FOR WRevCnt = 1 TO 3
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
      CASE 1
        PRINT #UBRpt, TAB(47); UBCustRec(1).Addr2
      CASE 2
        PRINT #UBRpt, TAB(47); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; LEFT$(UBCustRec(1).ZipCode, 5)
      CASE ELSE
        PRINT #UBRpt,
      END SELECT
    NEXT
    IF Previous# <> 0 THEN
      PRINT #UBRpt, " Previous:"; TAB(33); USING "$$,#####.##"; Previous#
    ELSE
      PRINT #UBRpt, ""
    END IF
    PRINT #UBRpt, ""
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(7); USING "######"; UBBillRec(1).CustAcctNo;
    PRINT #UBRpt, TAB(31); USING "$$#####.##"; ABS(TotalAmt#);
    PRINT #UBRpt, TAB(46); USING "######"; UBBillRec(1).CustAcctNo;
    PRINT #UBRpt, TAB(57); USING "$$#####.##"; ABS(TotalAmt#)
    'IF FinalFlag AND CDeposit# THEN
    '  PRINT #UBRpt, "                   Deposit:  "; USING "$$,######.##"; -UBCustRec(1).DepositAmt
    '  TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
    'ELSE
    '  PRINT #UBRpt,
    'END IF

    'IF TotalAmt# < 0 AND FinalFlag THEN
    '  PRINT #UBRpt, "                Refund Due:  "; USING "$$,######.##"; ABS(TotalAmt#)
    'ELSE
    '  PRINT #UBRpt, "                     Total:  "; USING "$$,######.##"; TotalAmt#
    'END IF

    PRINT #UBRpt,
    PRINT #UBRpt, "  "; Message$
    PRINT #UBRpt,

