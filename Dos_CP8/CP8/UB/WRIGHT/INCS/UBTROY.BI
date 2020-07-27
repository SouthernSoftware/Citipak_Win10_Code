
    Zero$ = "0"
    'Multi& = 100

    IF UBBillRec(1).CurRead(1) >= 0 AND UBBillRec(1).PrevRead(1) >= 0 THEN
     Multi& = UBCustRec(1).LocMeters(1).MtrMulti
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


    PRINT #UBRpt, "~"; USING "##########"; FBillNO& + PrintedCnt
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(30); USING "######"; CustAcctNo&
    PRINT #UBRpt,
    PRINT #UBRpt,

'******
    IF UBBillRec(1).RevAmt(1) <> 0 THEN
      PRINT #UBRpt, TAB(1); LEFT$(UBSetUpRec(1).Revenues(1).RevName, 2);
      PRINT #UBRpt, TAB(4); USING "########"; UBBillRec(1).CurRead(1) * Multi&;
      PRINT #UBRpt, TAB(13); USING "########"; UBBillRec(1).PrevRead(1) * Multi&;
      PRINT #UBRpt, TAB(22); USING "########"; UsageAmt& * Multi&;
      PRINT #UBRpt, TAB(31); USING "####.##"; UBBillRec(1).RevAmt(1)
    ELSE
      PRINT #UBRpt,
    END IF
    
    IF UBBillRec(1).RevAmt(2) <> 0 THEN
      PRINT #UBRpt, TAB(1); LEFT$(UBSetUpRec(1).Revenues(2).RevName, 2);
      PRINT #UBRpt, TAB(4); USING "########"; UBBillRec(1).CurRead(1) * Multi&;
      PRINT #UBRpt, TAB(13); USING "########"; UBBillRec(1).PrevRead(1) * Multi&;

      PRINT #UBRpt, TAB(22); USING "########"; UsageAmt& * Multi&;
      PRINT #UBRpt, TAB(31); USING "####.##"; UBBillRec(1).RevAmt(2)
    ELSE
      PRINT #UBRpt,
    END IF

    IF UBBillRec(1).RevAmt(4) <> 0 THEN
      PRINT #UBRpt, TAB(1); LEFT$(UBSetUpRec(1).Revenues(4).RevName, 2);
      PRINT #UBRpt, TAB(31); USING "####.##"; UBBillRec(1).RevAmt(4)
    ELSE
      PRINT #UBRpt,
    END IF

    PRINT #UBRpt,

    PRINT #UBRpt, TAB(40); LEFT$(UBCustRec(1).CustName, 20)
    PRINT #UBRpt, TAB(1); USING "####.##"; UBBillRec(1).TransAmt; TAB(19); Previous#; TAB(31); TotalAmt#;
    PRINT #UBRpt, TAB(40); LEFT$(UBCustRec(1).Addr1, 20)
    PRINT #UBRpt, TAB(40); LEFT$(UBCustRec(1).Addr2, 25)
    PRINT #UBRpt, TAB(40); LEFT$(UBCustRec(1).City, 9); " "; UBCustRec(1).State; " "; UBCustRec(1).ZipCode
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, LEFT$(PastDate$, 6); RIGHT$(PastDate$, 2); "   "; LEFT$(DateRead$, 6); RIGHT$(DateRead$, 2); TAB(42); USING "######"; CustAcctNo&;
    PRINT #UBRpt, TAB(53); USING "####.##"; TotalAmt#
    PRINT #UBRpt,
    PRINT #UBRpt, Message$

