'051697 Montross, Mineral Springs
'062497 last modified

    IF UBBillRec(1).CurRead(1) >= 0 AND UBBillRec(1).PrevRead(1) >= 0 THEN
      CurRead# = 100# * UBBillRec(1).CurRead(1)
      PreRead# = 100# * UBBillRec(1).PrevRead(1)
      UsageAmt& = CurRead# - PreRead#
      IF UsageAmt& < 0 THEN
        MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(1))) - 1)
        UsageAmt& = (MaxMeterAmt& - PreRead# + CurRead#)
      END IF
    ELSE
      UsageAmt& = 0
    END IF

    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)
   
    IF TotalAmt# > 0 THEN
      'TenPct# = 0
      'TenPct# = 5

      TenPct# = Round#(TotalAmt# * .1)
    ELSE
      TenPct# = 0
    END IF
    
    PRINT #UBRpt, TAB(2); USING "########"; FBillNO& + PrintedCnt
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(3); LEFT$(DateRead$, 2); TAB(8); MID$(DateRead$, 4, 2); TAB(13); RIGHT$(DateRead$, 2);
    PRINT #UBRpt, TAB(17); LEFT$(PastDueDate$, 2); TAB(22); MID$(PastDueDate$, 4, 2); TAB(27); RIGHT$(PastDueDate$, 2);
    PRINT #UBRpt, TAB(40); PastDueDate$
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(2); USING "#########"; PreRead#;
    PRINT #UBRpt, TAB(12); USING "#########"; CurRead#;
    PRINT #UBRpt, TAB(22); USING "########"; UsageAmt&;
    PRINT #UBRpt, TAB(35); LEFT$(UBCustRec(1).ServAddr, 24)
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).CustName, 25)

    IF UBBillRec(1).RevAmt(1) <> 0 THEN
      PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(1).RevName; TAB(20); USING "######.##"; UBBillRec(1).RevAmt(1);
    END IF
    PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).Addr1, 25)

    IF UBBillRec(1).RevAmt(2) <> 0 THEN
      PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(2).RevName; TAB(20); USING "######.##"; UBBillRec(1).RevAmt(2);
    END IF
    PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).Addr2, 25)

    IF UBBillRec(1).RevAmt(3) <> 0 THEN
      PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(3).RevName;
      PRINT #UBRpt, TAB(20); USING "######.##"; UBBillRec(1).RevAmt(3);
    END IF
    PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; UBCustRec(1).ZipCode

    IF UBBillRec(1).RevAmt(4) <> 0 THEN
      PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(4).RevName;
      PRINT #UBRpt, TAB(20); USING "######.##"; UBBillRec(1).RevAmt(4);
    END IF
    PRINT #UBRpt, TAB(34); STRING$(24, "-")

    IF UBBillRec(1).RevAmt(5) <> 0 THEN
      PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(5).RevName;
      PRINT #UBRpt, TAB(20); USING "######.##"; UBBillRec(1).RevAmt(5);
    END IF
    PRINT #UBRpt, TAB(34); Message$

'    IF UBBillRec(1).RevAmt(6) <> 0 THEN
'      PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(6).RevName;
'      PRINT #UBRpt, TAB(20); USING "######.##"; UBBillRec(1).RevAmt(6)
'    ELSE
'      PRINT #UBRpt, ""
'    END IF

    IF Previous# <> 0 THEN
      PRINT #UBRpt, TAB(3); "Previous:"; TAB(20); USING "######.##"; Previous#
    ELSE
      PRINT #UBRpt, ""
    END IF

'    IF FinalFlag AND CDeposit# THEN
'      PRINT #UBRpt, TAB(4); "Deposit:"; TAB(20); USING "######.##"; -UBCustRec(1).DepositAmt
'      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
'    ELSE
      PRINT #UBRpt, TAB(50); USING "######.##"; TotalAmt#
'    END IF

    PRINT #UBRpt,
    PRINT #UBRpt, TAB(5); CustAcctNo&; TAB(20); USING "######.##"; TotalAmt#;
    PRINT #UBRpt, TAB(37); CustAcctNo&; TAB(50); USING "######.##"; Round#(TotalAmt# + TenPct#)
    PRINT #UBRpt, "~"

