
    PrevRead& = 0
    CurrRead& = 0
    UsageAmt& = 0

    FOR zz = 1 TO 7
      IF UBCustRec(1).LocMeters(zz).CurRead > 0 THEN
        DateRead$ = Num2Date$(UBCustRec(1).LocMeters(zz).CurDate)
        PrevRead& = UBCustRec(1).LocMeters(zz).PrevRead
        CurrRead& = UBCustRec(1).LocMeters(zz).CurRead
        UsageAmt& = CurrRead& - PrevRead&
        IF UsageAmt& < 0 THEN
          MaxMeterAmt& = 10& ^ (LEN(STR$(PrevRead&)) - 1)
          UsageAmt& = (MaxMeterAmt& - PrevRead&) + CurrRead&
        END IF
        EXIT FOR
      END IF
    NEXT

    Previous# = UBCustRec(1).PrevBalance
    Current# = UBCustRec(1).CurrBalance
    TotalAmt# = Round#(Previous# + Current#)
    
    PRINT #UBRpt, "~"
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(3); LEFT$(DateRead$, 2); TAB(8); MID$(DateRead$, 4, 2); TAB(13); RIGHT$(DateRead$, 2);
    PRINT #UBRpt, TAB(17); LEFT$(PDate$, 2); TAB(22); MID$(PDate$, 4, 2); TAB(27); RIGHT$(PDate$, 2);
    PRINT #UBRpt, TAB(40); PDate$
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(2); USING "#########"; PrevRead&;
    PRINT #UBRpt, TAB(12); USING "#########"; CurrRead&;
    PRINT #UBRpt, TAB(22); USING "########"; UsageAmt&;
    PRINT #UBRpt, TAB(35); LEFT$(UBCustRec(1).ServAddr, 24)
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).CustName, 25)
    IF UBCustRec(1).CurrRevAmts(1) <> 0 THEN
      PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(1).RevName; TAB(20); USING "######.##"; UBCustRec(1).CurrRevAmts(1);
    END IF
    PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).Addr1, 25)

    IF UBCustRec(1).CurrRevAmts(2) <> 0 THEN
      PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(2).RevName; TAB(20); USING "######.##"; UBCustRec(1).CurrRevAmts(2);
    END IF
    PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).Addr2, 25)

    IF UBCustRec(1).CurrRevAmts(3) <> 0 THEN
      PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(3).RevName;
      PRINT #UBRpt, TAB(20); USING "######.##"; UBCustRec(1).CurrRevAmts(3);
    END IF
    PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; UBCustRec(1).ZipCode

    IF UBCustRec(1).CurrRevAmts(4) <> 0 THEN
      PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(4).RevName;
      PRINT #UBRpt, TAB(20); USING "######.##"; UBCustRec(1).CurrRevAmts(4);
    END IF
    PRINT #UBRpt, TAB(34); STRING$(24, "-")

    IF UBCustRec(1).CurrRevAmts(5) <> 0 THEN
      PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(5).RevName;
      PRINT #UBRpt, TAB(20); USING "######.##"; UBCustRec(1).CurrRevAmts(5);
    END IF
    PRINT #UBRpt, TAB(34); "LATE NOTICE"
    IF Previous# <> 0 THEN
      PRINT #UBRpt, TAB(3); "Previous:"; TAB(20); USING "######.##"; Previous#;
    'ELSE
    '  PRINT #UBRpt, ""
    END IF
    PRINT #UBRpt,

    PRINT #UBRpt,

    PRINT #UBRpt,
    PRINT #UBRpt, TAB(5); AcctNo&; TAB(20); USING "######.##"; TotalAmt#;
    PRINT #UBRpt, TAB(34); AcctNo&; USING "#####.##"; TAB(42); TotalAmt#; TAB(50); TotalAmt#
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(2); NoticeInfo.MsgLine1

    PRINT #UBRpt, TAB(2); NoticeInfo.MsgLine2

    PRINT #UBRpt, "~"

