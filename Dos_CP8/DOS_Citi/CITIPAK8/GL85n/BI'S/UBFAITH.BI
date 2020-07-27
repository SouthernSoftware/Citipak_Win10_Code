'04-24-98 Faith

    Loca$ = UBCustRec(1).Book + "-" + UBCustRec(1).SeqNumb

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

    PAmt# = Round#(TotalAmt# + FiveBucks)

    PRINT #UBRpt, TAB(2); USING "########"; FBillNO& + PrintedCnt
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(1); LEFT$(BillDate$, 6); RIGHT$(BillDate$, 2); TAB(13); Loca$
    PRINT #UBRpt, TAB(37); LEFT$(UBCustRec(1).ServAddr, 23)
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(2); USING "#########"; UBBillRec(1).PrevRead(1);
    PRINT #UBRpt, TAB(12); USING "#########"; UBBillRec(1).CurRead(1);
    PRINT #UBRpt, TAB(25); USING "########"; UsageAmt&;
    PRINT #UBRpt, TAB(37); LEFT$(UBCustRec(1).CustName, 25)
    PRINT #UBRpt, TAB(37); LEFT$(UBCustRec(1).Addr1, 25)

    IF UBBillRec(1).RevAmt(1) <> 0 THEN
      PRINT #UBRpt, TAB(2); UBSetUpRec(1).Revenues(1).RevName; TAB(27); USING "######.##"; UBBillRec(1).RevAmt(1);
    END IF
    PRINT #UBRpt, TAB(37); LEFT$(UBCustRec(1).Addr2, 25)

    IF UBBillRec(1).RevAmt(2) <> 0 THEN
      PRINT #UBRpt, TAB(2); UBSetUpRec(1).Revenues(2).RevName; TAB(27); USING "######.##"; UBBillRec(1).RevAmt(2);
    END IF
    PRINT #UBRpt, TAB(37); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; UBCustRec(1).ZipCode

    IF UBBillRec(1).RevAmt(3) <> 0 THEN
      PRINT #UBRpt, TAB(2); UBSetUpRec(1).Revenues(3).RevName;
      PRINT #UBRpt, TAB(27); USING "######.##"; UBBillRec(1).RevAmt(3);
    END IF
    PRINT #UBRpt,

    IF Previous# <> 0 THEN
      PRINT #UBRpt, TAB(2); "Previous:"; TAB(27); USING "######.##"; Previous#
    ELSE
      PRINT #UBRpt, ""
    END IF

    'IF UBBillRec(1).RevAmt(4) <> 0 THEN
    '  PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(4).RevName;
    '  PRINT #UBRpt, TAB(20); USING "######.##"; UBBillRec(1).RevAmt(4);
    'END IF
    'PRINT #UBRpt,

    PRINT #UBRpt,

    PRINT #UBRpt, TAB(40); Loca$; TAB(53); LEFT$(PastDueDate$, 6); RIGHT$(PastDueDate$, 2)
    PRINT #UBRpt, TAB(9); LEFT$(PastDueDate$, 6); RIGHT$(PastDueDate$, 2)

    PRINT #UBRpt, TAB(8); USING "######.##"; PAmt#; TAB(27); TotalAmt#; TAB(40); TotalAmt#; TAB(52); PAmt#

    PRINT #UBRpt,
    PRINT #UBRpt, "~"






