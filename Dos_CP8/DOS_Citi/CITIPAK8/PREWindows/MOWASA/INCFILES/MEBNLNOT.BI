
    PrevRead& = 0
    CurrRead& = 0
    UsageAmt& = 0

    FOR zz = 1 TO 7
      IF UBCustRec(1).LocMeters(zz).CurRead > 0 THEN
        DateRead$ = Num2Date$(UBCustRec(1).LocMeters(zz).CurDate)
        PrevRead& = UBCustRec(1).LocMeters(zz).PrevRead
        CurrRead& = UBCustRec(1).LocMeters(zz).CurRead
        IF PrevRead& < 0 THEN PrevRead& = 0
        IF CurrRead& < 0 THEN CurrRead& = 0
        UsageAmt& = CurrRead& - PrevRead&
        MtrMulti# = UBCustRec(1).LocMeters(zz).MtrMulti
        IF MtrMulti# = 0 THEN MtrMulti# = 1
        IF UsageAmt& < 0 THEN
          MaxMeterAmt& = 10& ^ (LEN(STR$(PrevRead&)) - 1)
          UsageAmt& = (MaxMeterAmt& - PrevRead&) + CurrRead&
        END IF
        UsageAmt& = UsageAmt& * MtrMulti#
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
    PRINT #UBRpt, TAB(34); LEFT$(NDate$, 6); RIGHT$(NDate$, 2)
    PRINT #UBRpt, TAB(2); LEFT$(UBCustRec(1).ServAddr, 25)   'line 5 is serv address
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(2); LEFT$(NDate$, 6); RIGHT$(NDate$, 2);
    PRINT #UBRpt, TAB(11); LEFT$(DateRead$, 6); RIGHT$(DateRead$, 2); USING "#########"; AcctNo&
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).CustName, 25)
    PRINT #UBRpt, TAB(2); USING "########"; PrevRead&;
    PRINT #UBRpt, TAB(12); USING "#########"; CurrRead&;
    PRINT #UBRpt, TAB(22); USING "########"; UsageAmt&;
    PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).Addr1, 25)
    PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).Addr2, 25)

    IF UBCustRec(1).CurrRevAmts(1) <> 0 THEN
       PRINT #UBRpt, " "; LEFT$(UBSetUpRec(1).Revenues(1).RevName, 14);
       PRINT #UBRpt, TAB(22); USING "#####.##"; UBCustRec(1).CurrRevAmts(1);
    END IF
    PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; LEFT$(UBCustRec(1).ZipCode, 5)

    IF UBCustRec(1).CurrRevAmts(2) <> 0 THEN
      PRINT #UBRpt, " "; LEFT$(UBSetUpRec(1).Revenues(2).RevName, 14);
      PRINT #UBRpt, TAB(22); USING "#####.##"; UBCustRec(1).CurrRevAmts(2)
    ELSE
      PRINT #UBRpt,
    END IF

    IF UBCustRec(1).CurrRevAmts(3) <> 0 THEN
      PRINT #UBRpt, " "; LEFT$(UBSetUpRec(1).Revenues(3).RevName, 14);
      PRINT #UBRpt, TAB(22); USING "#####.##"; UBCustRec(1).CurrRevAmts(3);
    END IF

    PRINT #UBRpt, TAB(34); "LOC: "; UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb

    IF UBCustRec(1).CurrRevAmts(4) <> 0 THEN
      PRINT #UBRpt, " "; LEFT$(UBSetUpRec(1).Revenues(4).RevName, 14);
      PRINT #UBRpt, TAB(22); USING "#####.##"; UBCustRec(1).CurrRevAmts(4)
    ELSE
      PRINT #UBRpt,
    END IF

    IF Previous# <> 0 THEN
      PRINT #UBRpt, " Balance Forward     "; USING "#####.##"; Previous#;
    END IF

    PRINT #UBRpt, TAB(47); USING "#####"; AcctNo&
    PRINT #UBRpt,

    PRINT #UBRpt,
    PRINT #UBRpt, TAB(2); LEFT$(NDate$, 6); RIGHT$(NDate$, 2);
    PRINT #UBRpt, TAB(24); USING "#####.##"; TotalAmt#; TAB(49); TotalAmt#
    PRINT #UBRpt,
    PRINT #UBRpt, " LATE NOTICE!"
    PRINT #UBRpt, TAB(2); NoticeInfo.MsgLine1
    PRINT #UBRpt, "~"


'-----------






