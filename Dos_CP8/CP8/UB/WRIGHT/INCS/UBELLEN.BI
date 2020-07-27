    REM  This is Ellenboro, NC Utility Card Format
    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    TotalAmt# = Round(UBBillRec(1).TransAmt + Previous#)
    IF (UBBillRec(1).ApplyDepFlag = "Y") AND (CDeposit# > 0) THEN
      TotalAmt# = Round#(TotalAmt# - CDeposit#)
    END IF
    UsageAmt& = UBBillRec(1).CurRead(1) - UBBillRec(1).PrevRead(1)
    IF UsageAmt& < 0 THEN
      MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(1))) - 1)
      UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(1)) + UBBillRec(1).CurRead(1)
    END IF
    PRINT #UBRpt, USING "########"; FBillNO& + PrintedCnt
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(2); USING "########"; UBBillRec(1).PrevRead(1);
    PRINT #UBRpt, TAB(13); USING "########"; UBBillRec(1).CurRead(1);
    PRINT #UBRpt, TAB(23); USING "########"; UsageAmt&;
    PRINT #UBRpt, TAB(37); LEFT$(UBCustRec(1).CustName, 25)
    PRINT #UBRpt, TAB(37); LEFT$(UBCustRec(1).Addr1, 25)
    PRINT #UBRpt, TAB(37); LEFT$(UBCustRec(1).Addr2, 25)
    PRINT #UBRpt, TAB(37); LEFT$(UBCustRec(1).CITY, 14); " "; UBCustRec(1).State; " "; LEFT$(UBCustRec(1).ZipCode, 5); "-"; MID$(UBCustRec(1).ZipCode, 6, 4)
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(2); USING "####.##"; UBBillRec(1).RevAmt(1);
    PRINT #UBRpt, TAB(20); USING "####.##"; UBBillRec(1).RevAmt(2) + UBBillRec(1).RevAmt(3);
    IF Previous# > 0 THEN
      PRINT #UBRpt, TAB(28); USING "####.##"; Previous#;
    END IF
    PRINT #UBRpt, TAB(37); "Acct#: "; CustAcctNo&
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(3); LEFT$(BILLDATE$, 6) + RIGHT$(BILLDATE$, 2);
    PRINT #UBRpt, TAB(15); USING "######"; CustAcctNo&;
    PRINT #UBRpt, TAB(27); USING "$####.##"; TotalAmt#;
    PRINT #UBRpt, TAB(37); LEFT$(BILLDATE$, 6) + RIGHT$(BILLDATE$, 2)
    IF CDeposit# > 0 THEN
      PRINT #UBRpt, TAB(50); "Less Deposit"
      PRINT #UBRpt, TAB(50); USING "####.##"; CDeposit#
      PRINT #UBRpt,
    ELSE
      PRINT #UBRpt,
      PRINT #UBRpt,
      PRINT #UBRpt,
    END IF
    PRINT #UBRpt, TAB(37); USING "######"; CustAcctNo&;
    PRINT #UBRpt, TAB(49); USING "$####.##"; TotalAmt#
    PRINT #UBRpt,

