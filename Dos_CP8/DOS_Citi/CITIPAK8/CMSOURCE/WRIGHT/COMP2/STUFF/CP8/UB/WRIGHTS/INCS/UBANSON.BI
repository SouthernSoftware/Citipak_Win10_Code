'Ansonville NC Bill format 06-24-97

    AcctNum = UBBillRec(1).CustAcctNo
    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)

    TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)
    
    PRINT #UBRpt, TAB(32); USING "########"; FBillNO& + PrintedCnt
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, USING "########"; UBBillRec(1).CurRead(1); TAB(13); UBBillRec(1).PrevRead(1);
    UsageAmt& = UBBillRec(1).CurRead(1) - UBBillRec(1).PrevRead(1)
    IF UsageAmt& < 0 THEN
      MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(1))) - 1)
      UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(1)) + UBBillRec(1).CurRead(1)
    END IF
    PRINT #UBRpt, TAB(25); USING "########"; UsageAmt&

    PRINT #UBRpt, TAB(35); LEFT$(UBCustRec(1).CustName, 29)
    PRINT #UBRpt, TAB(35); UBCustRec(1).Addr1

    PCnt = 0
    FOR WRevCnt = 1 TO 5
      PCnt = PCnt + 1
      IF UBBillRec(1).RevAmt(WRevCnt) <> 0 THEN
        PRINT #UBRpt, " "; LEFT$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 5);
        PRINT #UBRpt, TAB(25); USING "#####.##"; UBBillRec(1).RevAmt(WRevCnt);
      END IF
      SELECT CASE PCnt
      CASE 1
        PRINT #UBRpt, TAB(35); UBCustRec(1).Addr2
      CASE 2
        PRINT #UBRpt, TAB(35); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; LEFT$(UBCustRec(1).ZipCode, 5)
      CASE ELSE
        PRINT #UBRpt,
      END SELECT
    NEXT

    IF Previous# <> 0 THEN
      PRINT #UBRpt, " Previous:"; TAB(25); USING "#####.##"; Previous#
    ELSE
      PRINT #UBRpt,
    END IF
    IF FinalFlag AND CDeposit# THEN
      PRINT #UBRpt, " Deposit:"; TAB(25); USING "#####.##"; -UBCustRec(1).DepositAmt
      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
    ELSE
      PRINT #UBRpt,
    END IF
    PRINT #UBRpt, ""
    PRINT #UBRpt, ""
    PRINT #UBRpt, ""
    PRINT #UBRpt, TAB(25); USING "#####.##"; TotalAmt#; TAB(50); TotalAmt#
    PRINT #UBRpt, TAB(25); USING "########"; AcctNum; TAB(50); AcctNum
    PRINT #UBRpt, TAB(23); BillDate$; TAB(48); BillDate$

    'IF TotalAmt# < 0 AND FinalFlag THEN
    '  PRINT #UBRpt, "                Refund Due:  "; USING "$$,######.##"; ABS(TotalAmt#)
    'ELSE
    '  PRINT #UBRpt, "                     Total:  "; USING "$$,######.##"; TotalAmt#
    'END IF

    PRINT #UBRpt,
    PRINT #UBRpt, "~"; Message$

