
    FoundAMtr = 0
    FOR mChk = 1 TO 7
      IF UBBillRec(1).MtrTypes(mChk) > 0 THEN
        FoundAMtr = mChk
        EXIT FOR
      END IF
    NEXT

    IF FoundAMtr > 0 THEN
      UsageAmt& = UBBillRec(1).CurRead(FoundAMtr) - UBBillRec(1).PrevRead(FoundAMtr)
      IF UsageAmt& < 0 THEN
        MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(FoundAMtr))) - 1)
        UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(FoundAMtr)) + UBBillRec(1).CurRead(FoundAMtr)
      END IF
    ELSE
      UsageAmt& = 0
    END IF
    
    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    TotalAmt# = Round(UBBillRec(1).TransAmt + Previous#)
    
    PRINT #UBRpt, "~"; TAB(35); USING "########"; FBillNO& + PrintedCnt
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(41); USING "#####"; CustAcctNo&
    PRINT #UBRpt,

    IF UBBillRec(1).RevAmt(1) <> 0 THEN
      PRINT #UBRpt, LEFT$(UBSetUpRec(1).Revenues(1).RevName, 3);
      IF FoundAMtr THEN
        PRINT #UBRpt, TAB(4); USING "#######"; UBBillRec(1).CurRead(FoundAMtr);
        PRINT #UBRpt, TAB(12); USING "#######"; UBBillRec(1).PrevRead(FoundAMtr);
        PRINT #UBRpt, TAB(20); USING "#######"; UsageAmt&;
      END IF
      PRINT #UBRpt, TAB(28); USING "#####.##"; UBBillRec(1).RevAmt(1)
    ELSE
      PRINT #UBRpt,
    END IF
    IF UBBillRec(1).RevAmt(2) <> 0 THEN
      PRINT #UBRpt, LEFT$(UBSetUpRec(1).Revenues(2).RevName, 3);
      IF FoundAMtr THEN
        PRINT #UBRpt, TAB(4); USING "#######"; UBBillRec(1).CurRead(FoundAMtr);
        PRINT #UBRpt, TAB(12); USING "#######"; UBBillRec(1).PrevRead(FoundAMtr);
        PRINT #UBRpt, TAB(20); USING "#######"; UsageAmt&;
      END IF
      PRINT #UBRpt, TAB(28); USING "#####.##"; UBBillRec(1).RevAmt(2)
    ELSE
      PRINT #UBRpt,
    END IF
    IF UBBillRec(1).RevAmt(3) <> 0 THEN
      PRINT #UBRpt, LEFT$(UBSetUpRec(1).Revenues(3).RevName, 3);
      PRINT #UBRpt, TAB(28); USING "#####.##"; UBBillRec(1).RevAmt(3);
    END IF
    PRINT #UBRpt, TAB(39); UBCustRec(1).Book; TAB(42); UBCustRec(1).SeqNumb;
    PRINT #UBRpt, TAB(50); LEFT$(DateRead$, 2); TAB(53); MID$(DateRead$, 4, 2)
    IF UBBillRec(1).RevAmt(4) <> 0 THEN
      PRINT #UBRpt, LEFT$(UBSetUpRec(1).Revenues(4).RevName, 3);
      PRINT #UBRpt, TAB(28); USING "#####.##"; UBBillRec(1).RevAmt(4)
    ELSE
      PRINT #UBRpt,
    END IF

    IF FinalFlag AND CDeposit# THEN
      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
    END IF
    IF Previous# <> 0 THEN
      PRINT #UBRpt, "PRE"; TAB(28); USING "#####.##"; Previous#;
    END IF
    PRINT #UBRpt, TAB(47); USING "$$#####.##"; TotalAmt#
    IF FinalFlag AND CDeposit# THEN
      PRINT #UBRpt, "DEP"; TAB(28); USING "#####.##"; -UBCustRec(1).DepositAmt
    ELSE
      PRINT #UBRpt,
    END IF
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(22); LEFT$(UBCustRec(1).CustName, 25)
    PRINT #UBRpt, TAB(22); LEFT$(UBCustRec(1).Addr1, 25)
    PRINT #UBRpt, TAB(22); LEFT$(UBCustRec(1).Addr2, 25)
    PRINT #UBRpt, TAB(22); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; UBCustRec(1).ZipCode
    PRINT #UBRpt, Message$
    PRINT #UBRpt, Msg2$
    PRINT #UBRpt, Msg3$
    PRINT #UBRpt, Msg4$

