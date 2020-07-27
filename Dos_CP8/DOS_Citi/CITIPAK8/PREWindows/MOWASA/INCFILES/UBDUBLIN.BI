'062597 Dublin NC

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
    
    PRINT #UBRpt, "~"; TAB(34); USING "#####"; FBillNO& + PrintedCnt
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(3); UBCustRec(1).ServAddr
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(2); LEFT$(BillDate$, 6); RIGHT$(BillDate$, 2); TAB(11); LEFT$(DateRead$, 6); RIGHT$(DateRead$, 2); TAB(20); USING "########"; CustAcctNo&
    PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).CustName, 25)
    PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).Addr1, 25)
    PRINT #UBRpt, TAB(2); USING "#########"; UBBillRec(1).PrevRead(1);
    PRINT #UBRpt, TAB(12); USING "#########"; UBBillRec(1).CurRead(1);
    PRINT #UBRpt, TAB(22); USING "########"; UsageAmt&;
    PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).Addr2, 25)
    PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; UBCustRec(1).ZipCode

    IF UBBillRec(1).RevAmt(1) <> 0 THEN
      PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(1).RevName; TAB(23); USING "#####.##"; UBBillRec(1).RevAmt(1);
    END IF
    PRINT #UBRpt,

    IF UBBillRec(1).RevAmt(2) <> 0 THEN
      PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(2).RevName; TAB(23); USING "#####.##"; UBBillRec(1).RevAmt(2);
    END IF
    PRINT #UBRpt,

    IF UBBillRec(1).RevAmt(3) <> 0 THEN
      PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(3).RevName; TAB(23); USING "#####.##"; UBBillRec(1).RevAmt(3);
    END IF
    PRINT #UBRpt,

    IF UBBillRec(1).RevAmt(4) <> 0 THEN
      PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(4).RevName; TAB(23); USING "#####.##"; UBBillRec(1).RevAmt(4);
    END IF

    IF UBBillRec(1).RevAmt(5) <> 0 THEN
      PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(5).RevName; TAB(23); USING "#####.##"; UBBillRec(1).RevAmt(5);
    END IF
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(34); LEFT$(PastDueDate$, 6); RIGHT$(BillDate$, 2); TAB(44); USING "########"; CustAcctNo&

    'insert 6 here

    IF Previous# <> 0 THEN
      PRINT #UBRpt, TAB(3); "Previous:"; TAB(23); USING "#####.##"; Previous#
    ELSE
      PRINT #UBRpt, ""
    END IF

    IF FinalFlag AND CDeposit# THEN
      PRINT #UBRpt, TAB(4); "Deposit:"; TAB(20); USING "#####.##"; -UBCustRec(1).DepositAmt
      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
    ELSE
      PRINT #UBRpt,
    END IF

    PRINT #UBRpt, TAB(2); USING "#####.##"; TotalAmt#; TAB(16); TotalAmt# + 5; TAB(24); TotalAmt#;
    PRINT #UBRpt, USING "#####.##"; TAB(41); TotalAmt# + 5; TAB(49); TotalAmt#
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, Message$
    PRINT #UBRpt, "~"


'    IF UBBillRec(1).RevAmt(6) <> 0 THEN
'      PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(6).RevName;
'      PRINT #UBRpt, TAB(20); USING "#####.##"; UBBillRec(1).RevAmt(6)
'    ELSE
'      PRINT #UBRpt, ""
'    END IF

