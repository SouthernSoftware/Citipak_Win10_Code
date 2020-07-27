'102898 Salem SC

    NMonth$ = LEFT$(MakeMonth$(BillDate$), 3) + "."

    Loc$ = UBCustRec(1).Book + "-" + UBCustRec(1).SeqNumb
    CustMsg$ = QPTrim$(UBCustRec(1).BILLCMNT)
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
    
    PRINT #UBRpt, "~"; USING "####"; FBillNO& + PrintedCnt
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(3); LEFT$(DateRead$, 2); TAB(8); MID$(DateRead$, 4, 2); TAB(13); RIGHT$(DateRead$, 2);
    PRINT #UBRpt, TAB(17); LEFT$(PastDueDate$, 2); TAB(22); MID$(PastDueDate$, 4, 2); TAB(27); RIGHT$(PastDueDate$, 2);
    PRINT #UBRpt, TAB(35); NMonth$
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(1); USING "#########"; UBBillRec(1).PrevRead(1);
    PRINT #UBRpt, TAB(11); USING "#########"; UBBillRec(1).CurRead(1);
    PRINT #UBRpt, TAB(22); USING "########"; UsageAmt&;
    PRINT #UBRpt, TAB(35); LEFT$(UBCustRec(1).ServAddr, 24)
    PRINT #UBRpt,
    PRINT #UBRpt,

    IF UBBillRec(1).RevAmt(1) <> 0 THEN
      PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(1).RevName; TAB(20); USING "######.##"; UBBillRec(1).RevAmt(1);
    END IF
    PRINT #UBRpt, TAB(34); USING "####"; CustAcctNo&;
    PRINT #UBRpt, " "; Loc$;
    PRINT #UBRpt, TAB(49); USING "######.##"; TotalAmt#
    
    IF UBBillRec(1).RevAmt(2) <> 0 THEN
      PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(2).RevName; TAB(20); USING "######.##"; UBBillRec(1).RevAmt(2)
    ELSE
      PRINT #UBRpt,
    END IF

    IF UBBillRec(1).RevAmt(3) <> 0 THEN
      PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(3).RevName;
      PRINT #UBRpt, TAB(20); USING "######.##"; UBBillRec(1).RevAmt(3);
    END IF
    PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).CustName, 25)
     
    IF UBBillRec(1).RevAmt(4) <> 0 THEN
      PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(4).RevName;
      PRINT #UBRpt, TAB(20); USING "######.##"; UBBillRec(1).RevAmt(4);
    END IF
    PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).Addr1, 25)

    IF UBBillRec(1).RevAmt(5) <> 0 THEN
      PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(5).RevName;
      PRINT #UBRpt, TAB(20); USING "######.##"; UBBillRec(1).RevAmt(5);
    END IF
    PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).Addr2, 25)

    IF Previous# <> 0 THEN
      PRINT #UBRpt, TAB(3); "Previous:"; TAB(20); USING "######.##"; Previous#;
    END IF
    PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; UBCustRec(1).ZipCode

    IF FinalFlag AND CDeposit# THEN
      PRINT #UBRpt, TAB(4); "Deposit:"; TAB(20); USING "######.##"; -UBCustRec(1).DepositAmt;
      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
    END IF
    PRINT #UBRpt, 'TAB(34); UBCustRec(1).PostRte
    PRINT #UBRpt,
    PRINT #UBRpt, USING "######"; CustAcctNo&;
    PRINT #UBRpt, "  "; Loc$;
    PRINT #UBRpt, TAB(20); USING "######.##"; TotalAmt#;
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(33); Message$
    PRINT #UBRpt, TAB(33); Msg2$
    PRINT #UBRpt, TAB(33); Msg3$
    PRINT #UBRpt, "~"

