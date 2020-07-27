'04-27-98

    CutOff$ = QPTrim$(Message$)

    'Loca$ = UBCustRec(1).Book + "-" + UBCustRec(1).SeqNumb


    TotalAmt# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    FiveBucks = 0

    PRINT #UBRpt, "~"
    PRINT #UBRpt, USING "########"; AcctNo&
    PRINT #UBRpt,

    'IF UBCustRec(1).CurrRevAmts(1) <> 0 THEN
    '  PRINT #UBRpt, TAB(2); UBSetUpRec(1).Revenues(1).RevName; TAB(22); USING "######.##"; UBCustRec(1).CurrRevAmts(1)
    'ELSE
      PRINT #UBRpt,
    'END IF

    'IF UBCustRec(1).CurrRevAmts(2) <> 0 THEN
    '  PRINT #UBRpt, TAB(2); UBSetUpRec(1).Revenues(2).RevName; TAB(22); USING "######.##"; UBCustRec(1).CurrRevAmts(2)
    'ELSE
      PRINT #UBRpt,
    'END IF

    'IF UBCustRec(1).CurrRevAmts(3) <> 0 THEN
    '  PRINT #UBRpt, TAB(2); UBSetUpRec(1).Revenues(3).RevName; TAB(22); USING "######.##"; UBCustRec(1).CurrRevAmts(3)
    'ELSE
      PRINT #UBRpt,
    'END IF

    'IF UBCustRec(1).CurrRevAmts(4) <> 0 THEN
    '  PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(4).RevName;
    '  PRINT #UBRpt, TAB(22); USING "######.##"; UBCustRec(1).CurrRevAmts(4);
    'END IF

    PRINT #UBRpt, TAB(33); LEFT$(UBCustRec(1).ServAddr, 23)

    PRINT #UBRpt, ""
    PRINT #UBRpt, TAB(33); LEFT$(NDate$, 6); RIGHT$(NDate$, 2)
    PRINT #UBRpt, ""

    PRINT #UBRpt, TAB(22); USING "######.##"; TotalAmt#
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(10); USING "########"; AcctNo&;
    PRINT #UBRpt, TAB(22); USING "######.##"; TotalAmt#;
    PRINT #UBRpt, TAB(33); LEFT$(NDate$, 6); RIGHT$(NDate$, 2);
    PRINT #UBRpt, TAB(42); LEFT$(PDate$, 6); RIGHT$(PDate$, 2);
    PRINT #UBRpt, TAB(51); LEFT$(PDate$, 6); RIGHT$(PDate$, 2)

    PRINT #UBRpt,
    PRINT #UBRpt, TAB(31); LEFT$(UBCustRec(1).CustName, 25)
    PRINT #UBRpt, TAB(31); LEFT$(UBCustRec(1).Addr1, 25)
    PRINT #UBRpt, TAB(31); LEFT$(UBCustRec(1).Addr2, 25)
    PRINT #UBRpt, TAB(31); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; UBCustRec(1).ZipCode
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, "~"






