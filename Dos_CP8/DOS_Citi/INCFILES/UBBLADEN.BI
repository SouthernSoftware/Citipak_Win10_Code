'04-27-98

    CutOff$ = QPTrim$(Message$)

    'Loca$ = UBCustRec(1).Book + "-" + UBCustRec(1).SeqNumb

    Z$ = "00"

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

    IF TotalAmt# > 0 THEN
      FiveBucks = 5
    ELSE
      FiveBucks = 0
    END IF

    'if FinalFlag

    'PAmt# = Round#(TotalAmt# + FiveBucks)

    PRINT #UBRpt, "~"
    PRINT #UBRpt, USING "########"; CustAcctNo&; TAB(40); FBillNO& + PrintedCnt
    PRINT #UBRpt,

    IF UBBillRec(1).RevAmt(1) <> 0 THEN
      PRINT #UBRpt, TAB(2); UBSetUpRec(1).Revenues(1).RevName; TAB(22); USING "######.##"; UBBillRec(1).RevAmt(1)
    ELSE
      PRINT #UBRpt,
    END IF
    IF UBBillRec(1).RevAmt(2) <> 0 THEN
      PRINT #UBRpt, TAB(2); UBSetUpRec(1).Revenues(2).RevName; TAB(22); USING "######.##"; UBBillRec(1).RevAmt(2)
    ELSE
      PRINT #UBRpt,
    END IF

    IF UBBillRec(1).RevAmt(3) <> 0 THEN
      PRINT #UBRpt, TAB(2); UBSetUpRec(1).Revenues(3).RevName; TAB(22); USING "######.##"; UBBillRec(1).RevAmt(3)
    ELSE
      PRINT #UBRpt,
    END IF

    IF UBBillRec(1).RevAmt(4) <> 0 THEN
      PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(4).RevName;
      PRINT #UBRpt, TAB(22); USING "######.##"; UBBillRec(1).RevAmt(4);
    'ELSE
    '  PRINT #UBRpt,
    END IF
    PRINT #UBRpt, TAB(33); LEFT$(UBCustRec(1).ServAddr, 23)

    IF Previous# <> 0 THEN
      PRINT #UBRpt, " Previous:"; TAB(22); USING "######.##"; Previous#
    ELSE
      PRINT #UBRpt, ""
    END IF

    IF FinalFlag AND CDeposit# THEN
      PRINT #UBRpt, "  Deposit:"; TAB(22); USING "######.##"; -UBCustRec(1).DepositAmt;
      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
    'ELSE
    '  PRINT #UBRpt,
    END IF

    'IF TotalAmt# < 0 AND FinalFlag THEN
    '  PRINT #UBRpt, "                Refund Due:  "; USING "$$,######.##"; ABS(TotalAmt#)
    'ELSE
    '  PRINT #UBRpt, "                     Total:  "; USING "$$,######.##"; TotalAmt#
    'END IF

    PRINT #UBRpt, TAB(33); LEFT$(BillDate$, 6); RIGHT$(BillDate$, 2);
    PRINT #UBRpt, TAB(42); LEFT$(PastDate$, 6); RIGHT$(PastDate$, 2);
    PRINT #UBRpt, TAB(51); LEFT$(DateRead$, 6); RIGHT$(DateRead$, 2)
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(22); USING "######.##"; TotalAmt#;
    PRINT #UBRpt, TAB(32); USING "#######"; UBBillRec(1).PrevRead(1);
    PRINT #UBRpt, Z$;
    PRINT #UBRpt, TAB(41); USING "#######"; UBBillRec(1).CurRead(1);
    PRINT #UBRpt, Z$;
    PRINT #UBRpt, TAB(50); USING "######"; UsageAmt&;
    PRINT #UBRpt, Z$
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(10); USING "########"; CustAcctNo&;
    PRINT #UBRpt, TAB(22); USING "######.##"; TotalAmt#;
    PRINT #UBRpt, TAB(33); LEFT$(BillDate$, 6); RIGHT$(BillDate$, 2);
    PRINT #UBRpt, TAB(42); LEFT$(PastDueDate$, 6); RIGHT$(PastDueDate$, 2);
    PRINT #UBRpt, TAB(51); CutOff$

    PRINT #UBRpt,
    PRINT #UBRpt, TAB(31); LEFT$(UBCustRec(1).CustName, 25)
    PRINT #UBRpt, TAB(31); LEFT$(UBCustRec(1).Addr1, 25)
    PRINT #UBRpt, TAB(31); LEFT$(UBCustRec(1).Addr2, 25)
    PRINT #UBRpt, TAB(31); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; UBCustRec(1).ZipCode
    PRINT #UBRpt,
    PRINT #UBRpt, Msg2$
    PRINT #UBRpt, Msg3$
    IF DidADraftFlag THEN
      PRINT #UBRpt, "DRAFT NOTICE DO NOT PAY"
    ELSE
      PRINT #UBRpt,
    END IF
    PRINT #UBRpt, "~"






