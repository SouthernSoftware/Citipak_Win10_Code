'Troutman
'Last modifed 02-20-97

    'PRINT #UBRpt, "TOP LINE                                  TOP LINE"
    PRINT #UBRpt, TAB(40); "#"; USING "#########"; PrintedCnt
    FOR LL = 1 TO 3:
      PRINT #UBRpt,
    NEXT LL

    PRINT #UBRpt, TAB(4); LEFT$(DateRead$, 2); TAB(8); MID$(DateRead$, 4, 2); TAB(13); RIGHT$(DateRead$, 2);
    PRINT #UBRpt, TAB(18); LEFT$(PastDueDate$, 2); TAB(22); MID$(PastDueDate$, 4, 2); TAB(27); RIGHT$(PastDueDate$, 2)
    PRINT #UBRpt, TAB(40); PastDueDate$
    PRINT #UBRpt,
    PRINT #UBRpt,

    IF UBCustRec(1).LocMeters(1).MtrType <> " " THEN
      PRINT #UBRpt, TAB(3); USING "########"; UBBillRec(1).PrevRead(1);
      PRINT #UBRpt, TAB(12); USING "########"; UBBillRec(1).CurRead(1);
      UsageAmt& = UBBillRec(1).CurRead(1) - UBBillRec(1).PrevRead(1)
      IF UsageAmt& < 0 THEN
        MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(1))) - 1)
        UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(1)) + UBBillRec(1).CurRead(1)
      END IF
      PRINT #UBRpt, TAB(22); USING "########"; UsageAmt&;
      PRINT #UBRpt, TAB(35); LEFT$(UBCustRec(1).SERVADDR, 25)
    ELSE
      PRINT #UBRpt, TAB(35); LEFT$(UBCustRec(1).SERVADDR, 25)
    END IF

    IF UBCustRec(1).LocMeters(2).MtrType <> " " THEN
      PRINT #UBRpt, TAB(3); USING "########"; UBBillRec(1).PrevRead(2);
      PRINT #UBRpt, TAB(12); USING "########"; UBBillRec(1).CurRead(2);
      UsageAmt& = UBBillRec(1).CurRead(2) - UBBillRec(1).PrevRead(2)
      IF UsageAmt& < 0 THEN
        MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(2))) - 1)
        UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(2)) + UBBillRec(1).CurRead(2)
      END IF
      PRINT #UBRpt, TAB(22); USING "########"; UsageAmt&
    END IF
    
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(33); LEFT$(UBCustRec(1).CustName, 29)
    PRINT #UBRpt, TAB(3); QPTrim$(UBSetUpRec(1).Revenues(1).RevName); TAB(21); USING "$####.##"; UBBillRec(1).RevAmt(1);
    PRINT #UBRpt, TAB(33); UBCustRec(1).ADDR1
    PRINT #UBRpt, TAB(3); QPTrim$(UBSetUpRec(1).Revenues(2).RevName); TAB(21); USING "$####.##"; UBBillRec(1).RevAmt(2);
    PRINT #UBRpt, TAB(33); UBCustRec(1).Addr2
    PRINT #UBRpt, TAB(3); QPTrim$(UBSetUpRec(1).Revenues(3).RevName); TAB(21); USING "$####.##"; UBBillRec(1).RevAmt(3);
    PRINT #UBRpt, TAB(33); LEFT$(UBCustRec(1).CITY, 14); " "; UBCustRec(1).STATE; " "; UBCustRec(1).ZIPCODE
    FOR LL = 4 TO 5
      IF UBBillRec(1).RevAmt(LL) > 0 THEN
        PRINT #UBRpt, TAB(3); QPTrim$(UBSetUpRec(1).Revenues(LL).RevName); TAB(21); USING "$####.##"; UBBillRec(1).RevAmt(LL)
      ELSE
        PRINT #UBRpt,
      END IF
    NEXT
    Deposit# = UBCustRec(1).DepositAmt
    PRINT #UBRpt,
    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    'TotalDue# = Round#((UBBillRec(1).TransAmt + Previous#) - Deposit#)
    TotalDue# = Round#(UBBillRec(1).TransAmt + Previous#)
    PastAmt# = Round#(TotalDue# + Round#(TotalDue# * .1))
    PRINT #UBRpt, "    Previous:"; TAB(20); USING "$#####.##"; Previous#
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(4); UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb; TAB(20); USING "$#####.##"; TotalDue#;
    PRINT #UBRpt, TAB(36); UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb; TAB(50); USING "$#####.##"; TotalDue#
'    PRINT #UBRpt, TAB(4); UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb; TAB(20); USING "$#####.##"; Round#((UBBillRec(1).TransAmt + Previous#) - Deposit#);
'    PRINT #UBRpt, TAB(36); UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb; TAB(50); USING "$#####.##"; Round#((UBBillRec(1).TransAmt + Previous#) - Deposit#)

    PRINT #UBRpt, "  After 15th"; TAB(20); USING "$#####.##"; PastAmt#
    PRINT #UBRpt,
    PRINT #UBRpt, Message$


