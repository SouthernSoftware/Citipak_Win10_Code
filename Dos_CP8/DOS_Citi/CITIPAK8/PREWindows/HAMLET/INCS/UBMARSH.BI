    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(3); LEFT$(DateRead$, 2); TAB(8); MID$(DateRead$, 4, 2); TAB(13); RIGHT$(DateRead$, 2);
    PRINT #UBRpt, TAB(17); LEFT$(PastDueDate$, 2); TAB(22); MID$(PastDueDate$, 4, 2); TAB(27); RIGHT$(PastDueDate$, 2)
    PRINT #UBRpt, TAB(41); PastDueDate$
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(3); USING "########"; UBBillRec(1).PrevRead(1);
    PRINT #UBRpt, TAB(12); USING "########"; UBBillRec(1).CurRead(1);
    UsageAmt& = UBBillRec(1).CurRead(1) - UBBillRec(1).PrevRead(1)
    IF UsageAmt& < 0 THEN
      MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(1))) - 1)
      UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(1)) + UBBillRec(1).CurRead(1)
    END IF
     PRINT #UBRpt, TAB(22); USING "########"; UsageAmt&;
      PRINT #UBRpt, TAB(35); LEFT$(UBCustRec(1).ServAddr, 24)
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).CustName, 25)
    PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(1).RevName; TAB(20); USING "$#####.##"; UBBillRec(1).RevAmt(1);
     PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).Addr1, 25)
    PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(2).RevName; TAB(20); USING "$#####.##"; UBBillRec(1).RevAmt(2);
     PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).Addr2, 25)
    PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(3).RevName;
     PRINT #UBRpt, TAB(20); USING "$#####.##"; UBBillRec(1).RevAmt(3);
     PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; UBCustRec(1).ZipCode
    PCnt = 0
    FOR WRevCnt = 1 TO 3
      IF UBBillRec(1).RevAmt(WRevCnt + 3) > 0 THEN
        PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(WRevCnt + 3).RevName; TAB(20); USING "$#####.##"; UBBillRec(1).RevAmt(WRevCnt + 3)
        PCnt = PCnt + 1
      END IF
    NEXT

    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    TotalAmt# = Round(UBBillRec(1).TransAmt + Previous#)

    IF Previous# <> 0 THEN
      PRINT #UBRpt, TAB(3); "Previous: "; TAB(20); USING "$#####.##"; Previous#
    ELSE
      'TotalAmt# = UBBillRec(1).TransAmt
      PRINT #UBRpt, ""
    END IF
    
    IF ApplyDepFlag$ = "Y" THEN
      PRINT #UBRpt, " Deposit: "; TAB(19); USING "$####.##"; -UBCustRec(1).DepositAmt;
      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
    END IF

    PRINT #UBRpt,

    FOR LL = PCnt + 1 TO 3
      PRINT #UBRpt, ""
    NEXT

    IF FinalFlag THEN
      IF TotalAmt# < 0 THEN
        PRINT #UBRpt, TAB(1); "Refund Due:"; TAB(18); USING "$#####.##"; TotalAmt#;
         PRINT #UBRpt, TAB(33); "Refund Due:"; TAB(48); USING "$#####.##"; TotalAmt#
      ELSE
        PRINT #UBRpt, TAB(1); "Balance Due:"; TAB(18); USING "$#####.##"; TotalAmt#;
         PRINT #UBRpt, TAB(33); "Balance Due:"; TAB(48); USING "$#####.##"; TotalAmt#
      END IF
    ELSE
      PRINT #UBRpt, TAB(5); CustAcctNo&; TAB(20); USING "$#####.##"; TotalAmt#;
       PRINT #UBRpt, TAB(37); CustAcctNo&; TAB(50); USING "$#####.##"; TotalAmt#
    END IF

    PRINT #UBRpt,
    PRINT #UBRpt, Message$
    PRINT #UBRpt, TAB(34); UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb; TAB(50); "#"; PrintedCnt

    

