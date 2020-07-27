'Allendale
'this is a test

    WFoundMtr = False
    WCurrRead& = 0
    WPrevRead& = 0
    FOR mChk = 1 TO 7
      IF UBBillRec(1).MtrTypes(mChk) > 0 THEN
        SELECT CASE UBBillRec(1).MtrTypes(mChk)
        CASE 1, 2, 3
          WCurrRead& = UBBillRec(1).CurRead(mChk)
          WPrevRead& = UBBillRec(1).PrevRead(mChk)
          WUsageAmt& = WCurrRead& - WPrevRead&
          IF WUsageAmt& < 0 THEN
            MaxMeterAmt& = 10& ^ (LEN(STR$(WPrevRead&)) - 1)
            WUsageAmt& = (MaxMeterAmt& - WPrevRead&) + WCurrRead&
          END IF
          WFoundMtr = True
        END SELECT
      END IF
    NEXT
    
    FoundAMtr = False
    FOR mChk = 1 TO 7
      IF UBBillRec(1).MtrTypes(mChk) > 0 THEN
        FoundAMtr = True
        EXIT FOR
      END IF
    NEXT
    
    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    TotalAmt# = Round(UBBillRec(1).TransAmt + Previous#)

    PRINT #UBRpt, "~"; USING "#####"; FBillNO& + PrintedCnt
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, USING "########"; WPrevRead&;
    PRINT #UBRpt, TAB(9); USING "########"; WCurrRead&;
    PRINT #UBRpt, TAB(17); USING "######"; WUsageAmt&
    PRINT #UBRpt,
    IF UBBillRec(1).RevAmt(1) <> 0 THEN
      PRINT #UBRpt, TAB(4); UBSetUpRec(1).Revenues(1).RevName; TAB(23); USING "######.##"; UBBillRec(1).RevAmt(1);
    END IF
    PRINT #UBRpt, TAB(37); LEFT$(UBCustRec(1).CustName, 25)

    IF UBBillRec(1).RevAmt(2) <> 0 THEN
      PRINT #UBRpt, TAB(4); UBSetUpRec(1).Revenues(2).RevName; TAB(23); USING "######.##"; UBBillRec(1).RevAmt(2);
    END IF
    PRINT #UBRpt, TAB(37); LEFT$(UBCustRec(1).Addr1, 25)
    IF UBBillRec(1).RevAmt(3) <> 0 THEN
      PRINT #UBRpt, TAB(4); UBSetUpRec(1).Revenues(3).RevName;
      PRINT #UBRpt, TAB(23); USING "######.##"; UBBillRec(1).RevAmt(3);
    END IF
    PRINT #UBRpt, TAB(37); LEFT$(UBCustRec(1).Addr2, 25)
    PRINT #UBRpt, TAB(37); QPTrim$(UBCustRec(1).City); " "; UBCustRec(1).State; " "; UBCustRec(1).ZipCode

    IF Previous# <> 0 THEN
      PRINT #UBRpt, TAB(4); "Previous:"; TAB(23); USING "######.##"; Previous#
    ELSE
      PRINT #UBRpt, ""
    END IF

    IF FinalFlag AND CDeposit# THEN
      PRINT #UBRpt, TAB(4); "Deposit:"; TAB(24); USING "######.##"; -UBCustRec(1).DepositAmt;
      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
      PRINT #UBRpt, TAB(37); CustAcctNo&
    ELSE
      PRINT #UBRpt, TAB(37); CustAcctNo&
    END IF
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(5); "DUE: "; PastDueDate$; TAB(23); USING "######.##"; TotalAmt#
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(38); LEFT$(PastDueDate$, 6); RIGHT$(PastDueDate$, 2); TAB(49); USING "######.##"; TotalAmt#
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(3); Message$
    PRINT #UBRpt, TAB(3); Msg2$
    PRINT #UBRpt, TAB(3); Msg3$
    PRINT #UBRpt, "~"
    'PRINT #UBRpt, TAB(5); CustAcctNo&; TAB(20); DateRead$

