
    AcctNum = UBBillRec(1).CustAcctNo
    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)

    IF FinalFlag AND CDeposit# THEN
      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
    END IF

    IF TotalAmt# > 0 THEN
      TenPct# = Round#(UBBillRec(1).TransAmt * .1)
    ELSE
      TenPct# = 0
    END IF

    MtDy$ = LEFT$(DateRead$, 2) + " " + MID$(DateRead$, 4, 2)
    PDue$ = LEFT$(PastDueDate$, 2) + " " + MID$(PastDueDate$, 4, 2) + " " + RIGHT$(PastDueDate$, 2)

    WCurrRead& = 0
    WPrevRead& = 0
    Multi& = 0
    WFoundMtr = False
    FOR MChk = 1 TO 7
      IF UBBillRec(1).MtrTypes(MChk) > 0 THEN
        SELECT CASE UBBillRec(1).MtrTypes(MChk)
        CASE 1, 2, 3, 7
          WCurrRead& = UBBillRec(1).CurRead(MChk)
          WPrevRead& = UBBillRec(1).PrevRead(MChk)
          Multi& = UBCustRec(1).LocMeters(MChk).MTRMulti
          IF Multi& < 1 THEN Multi& = 1
          WUsageAmt& = WCurrRead& - WPrevRead&
          IF WUsageAmt& < 0 THEN
            MaxMeterAmt& = 10& ^ (LEN(STR$(WPrevRead&)) - 1)
            WUsageAmt& = (MaxMeterAmt& - WPrevRead&) + WCurrRead&
          END IF
          WUsageAmt& = WUsageAmt& * Multi&
          WFoundMtr = True
        END SELECT
      END IF
    NEXT

    PRINT #UBRpt, "~"; TAB(50); USING "########"; FBillNO& + PrintedCnt
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,

    IF UBBillRec(1).RevAmt(1) <> 0 THEN
      PRINT #UBRpt, TAB(2); LEFT$(UBSetUpRec(1).Revenues(1).RevName, 3);
      PRINT #UBRpt, TAB(5); USING "#######"; WCurrRead&;
      PRINT #UBRpt, TAB(13); USING "#######"; WPrevRead&;
      PRINT #UBRpt, TAB(21); USING "#######"; WUsageAmt&;
      PRINT #UBRpt, TAB(29); USING "#####.##"; UBBillRec(1).RevAmt(1);
    END IF
    PRINT #UBRpt, TAB(39); LEFT$(UBCustRec(1).CustName, 25)

    IF UBBillRec(1).RevAmt(2) <> 0 THEN
      PRINT #UBRpt, TAB(2); LEFT$(UBSetUpRec(1).Revenues(2).RevName, 3);
      PRINT #UBRpt, TAB(5); USING "#######"; WCurrRead&;
      PRINT #UBRpt, TAB(13); USING "#######"; WPrevRead&;
      PRINT #UBRpt, TAB(21); USING "#######"; WUsageAmt&;
      PRINT #UBRpt, TAB(29); USING "#####.##"; UBBillRec(1).RevAmt(2);
    END IF
    PRINT #UBRpt, TAB(39); LEFT$(UBCustRec(1).Addr1, 25)
    PRINT #UBRpt, TAB(39); LEFT$(UBCustRec(1).Addr2, 25)

    IF UBBillRec(1).RevAmt(3) <> 0 THEN
      PRINT #UBRpt, TAB(2); LEFT$(UBSetUpRec(1).Revenues(3).RevName, 3);
      PRINT #UBRpt, TAB(29); USING "#####.##"; UBBillRec(1).RevAmt(3);
    END IF
    PRINT #UBRpt, TAB(39); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State

    IF Previous# <> 0 THEN
      PRINT #UBRpt, TAB(5); "Previous:"; TAB(29); USING "#####.##"; Previous#;
    END IF
    PRINT #UBRpt, TAB(48); UBCustRec(1).ZipCode

    IF FinalFlag AND CDeposit# THEN
      PRINT #UBRpt, TAB(5); " Deposit:"; TAB(29); USING "#####.##"; -UBCustRec(1).DepositAmt
    ELSE
      PRINT #UBRpt,
    END IF
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(39); UBCustRec(1).Book; USING "######"; AcctNum;
    PRINT #UBRpt, TAB(51); PDue$
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(3); MtDy$; USING "#####.##"; TAB(12); TotalAmt#; TAB(20); TenPct#; TAB(29); Round#(TotalAmt# + TenPct#); TAB(41); TotalAmt#; TAB(52); Round#(TotalAmt# + TenPct#)
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, USING "#####.##"; TAB(12); TotalAmt#; TAB(29); Round#(TotalAmt# + TenPct#); TAB(41); TotalAmt#; TAB(52); Round#(TotalAmt# + TenPct#)
    PRINT #UBRpt, TAB(3); Message$

