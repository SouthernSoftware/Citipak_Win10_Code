'Plymouth

    CustMsg$ = QPTrim$(UBCustRec(1).BILLCMNT)

    IF NOT LPIFlag THEN
      LPIFlag = -2
      PRINT #UBRpt, CHR$(27); CHR$(48); 'CHR$(27); CHR$(58);
      'PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(77);
      ' put printer in     8 lpi             12 cpi
    END IF

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
      Bucks2# = 2
    ELSE
      Bucks2# = 0
    END IF

    PRINT #UBRpt, "~"; USING "#####"; FBillNO& + PrintedCnt
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, 'TAB(3); ; TAB(8); MID$(DateRead$, 4, 2); TAB(13); RIGHT$(DateRead$, 2);
    PRINT #UBRpt, TAB(2); LEFT$(UBCustRec(1).ServAddr, 24)
    'line 6
    PRINT #UBRpt, 'TAB(17); LEFT$(PastDueDate$, 2); TAB(22); MID$(PastDueDate$, 4, 2); TAB(27); RIGHT$(PastDueDate$, 2);
    PRINT #UBRpt, 'TAB(40); PastDueDate$
    PRINT #UBRpt, LEFT$(BillDate$, 6); RIGHT$(BillDate$, 2); " "; LEFT$(DateRead$, 6); RIGHT$(DateRead$, 2); USING "###########"; CustAcctNo&
    'line 9
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(2); USING "#########"; UBBillRec(1).PrevRead(1);
    PRINT #UBRpt, TAB(12); USING "#########"; UBBillRec(1).CurRead(1);
    PRINT #UBRpt, TAB(22); USING "########"; UsageAmt&
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(34); LEFT$(PastDueDate$, 6); RIGHT$(PastDueDate$, 2); USING "###############"; CustAcctNo&

    IF UBBillRec(1).RevAmt(1) <> 0 THEN
      PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(1).RevName; TAB(20); USING "######.##"; UBBillRec(1).RevAmt(1)
    ELSE
      PRINT #UBRpt,
    END IF

    IF UBBillRec(1).RevAmt(2) <> 0 THEN
      PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(2).RevName; TAB(20); USING "######.##"; UBBillRec(1).RevAmt(2)
    ELSE
      PRINT #UBRpt,
    END IF
    
    IF UBBillRec(1).RevAmt(3) <> 0 THEN
      PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(3).RevName;
      PRINT #UBRpt, TAB(20); USING "######.##"; UBBillRec(1).RevAmt(3)
    ELSE
      PRINT #UBRpt,
    END IF
    
    IF UBBillRec(1).RevAmt(4) <> 0 THEN
      PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(4).RevName;
      PRINT #UBRpt, TAB(20); USING "######.##"; UBBillRec(1).RevAmt(4);
    END IF
    PRINT #UBRpt, USING "#####.##"; TAB(42); Round#(TotalAmt# + Bucks2#); TAB(51); TotalAmt#

    IF UBBillRec(1).RevAmt(5) <> 0 THEN
      PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(5).RevName;
      PRINT #UBRpt, TAB(20); USING "######.##"; UBBillRec(1).RevAmt(5)
    ELSE
      PRINT #UBRpt,
    END IF
    
    IF Previous# <> 0 THEN
      PRINT #UBRpt, TAB(3); "Previous:"; TAB(20); USING "######.##"; Previous#;
    END IF
    PRINT #UBRpt,

    PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).CustName, 25)
    PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).Addr1, 25)
    PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).Addr2, 25)
    PRINT #UBRpt, USING "#####.##"; UBBillRec(1).TransAmt; TAB(16); Round#(TotalAmt# + Bucks2#); TotalAmt#;
    PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; UBCustRec(1).ZipCode
    PRINT #UBRpt,

      PRINT #UBRpt,

    PRINT #UBRpt, " "; Message$
    PRINT #UBRpt, " "; CustMsg$
    IF (UBCustRec(1).USEDRAFT = "Y") OR (LEN(QPTrim$(UBCustRec(1).BANKNAME)) > 0) THEN
      PRINT #UBRpt, " BANK DRAFT NOTICE. DO NOT PAY."
      PRINT #UBRpt, " DRAFT DATE: "; Num2Date$(UBBillRec(1).DraftDate)
    ELSE
      PRINT #UBRpt,
      PRINT #UBRpt,
    END IF
    PRINT #UBRpt, '"FAIR HOUSING DISCRIMINATION IS ILLEGAL IN N.C. IF YOU"
    PRINT #UBRpt, '"FEEL YOUR RIGHTS HAVE BEEN VIOLATED CALL TOWN OF PLYMOUTH"
    PRINT #UBRpt, '"(252)793-9101"
    PRINT #UBRpt, "~"

