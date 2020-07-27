
    REM  Concord, NC Card Format

    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    TotalAmt# = Round(UBBillRec(1).TransAmt + Previous#)
    UsageAmt& = UBBillRec(1).CurRead(1) - UBBillRec(1).PrevRead(1)
    'STOP
    IF UsageAmt& < 0 THEN
       MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(1))) - 1)
       UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(1)) + UBBillRec(1).CurRead(1)
    END IF

    PRINT #UBRpt, CHR$(27); CHR$(48);    'CHR$(27); CHR$(58);
    PRINT #UBRpt, "TOP"
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(35); LEFT$(UBCustRec(1).CustName, 25)
    PRINT #UBRpt, TAB(35); LEFT$(UBCustRec(1).Addr1, 25)
    PRINT #UBRpt, TAB(35); LEFT$(UBCustRec(1).Addr2, 25)
    PRINT #UBRpt, TAB(2); USING "########"; UBBillRec(1).PrevRead(1); TAB(13); UBBillRec(1).CurRead(1); TAB(23); UsageAmt&;
    PRINT #UBRpt, TAB(35); LEFT$(UBCustRec(1).CITY, 14); " "; UBCustRec(1).State
    PRINT #UBRpt, TAB(35); LEFT$(UBCustRec(1).ZipCode, 5); "-"; MID$(UBCustRec(1).ZipCode, 6, 4)
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(2); USING "####.##"; UBBillRec(1).RevAmt(1);
    TheRest# = 0
    FOR FCnt = 2 TO 10
      TheRest# = Round#(TheRest# + UBBillRec(1).RevAmt(FCnt))
    NEXT
    PRINT #UBRpt, TAB(17); USING "####.##"; TheRest#;

    IF Previous# > 0 THEN
      PRINT #UBRpt, TAB(25); USING "####.##"; Previous#;
    END IF
      PRINT #UBRpt, TAB(37); "Acct#: "; CustAcctNo&
    PRINT #UBRpt,

    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(3); LEFT$(BILLDATE$, 6) + RIGHT$(BILLDATE$, 2);
    PRINT #UBRpt, TAB(15); USING "######"; CustAcctNo&;
    PRINT #UBRpt, TAB(25); USING "$####.##"; TotalAmt#;
    PRINT #UBRpt, TAB(35); LEFT$(BILLDATE$, 6) + RIGHT$(BILLDATE$, 2)
    IF CDeposit# > 0 THEN
      PRINT #UBRpt, TAB(50); "Less Deposit"
      PRINT #UBRpt, TAB(50); USING "####.##"; CDeposit#
      PRINT #UBRpt,
    ELSE
      PRINT #UBRpt,
      PRINT #UBRpt,
      PRINT #UBRpt,
    END IF
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(37); USING "######"; CustAcctNo&;
    PRINT #UBRpt, TAB(49); USING "$####.##"; TotalAmt#
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, Message$

