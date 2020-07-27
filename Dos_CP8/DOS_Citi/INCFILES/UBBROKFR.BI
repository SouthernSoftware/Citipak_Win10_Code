    WCurrRead& = 0
    WPrevRead& = 0
    WUsageAmt& = 0
    WFoundMtr = False

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

    Other# = 0
    FOR zz = 3 TO 15
      Other# = Round#(Other# + UBBillRec(1).RevAmt(zz))
    NEXT

    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)
    IF FinalFlag AND CDeposit# THEN
      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
    END IF

    PRINT #UBRpt, "~"; TAB(2); USING "########"; FBillNO& + PrintedCnt
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,

    PRINT #UBRpt, USING "########"; WPrevRead&;
    PRINT #UBRpt, TAB(14); USING "####.##"; UBBillRec(1).RevAmt(1);
    PRINT #UBRpt, TAB(23); USING "####.##"; UBBillRec(1).RevAmt(2);
    PRINT #UBRpt, TAB(31); USING "####.##"; Other#;
    PRINT #UBRpt, TAB(44); USING "$$#####.##"; TotalAmt#
    PRINT #UBRpt, USING "########"; WCurrRead&;
    PRINT #UBRpt, TAB(9); USING "########"; WUsageAmt&
    PRINT #UBRpt, TAB(21); USING " PASTDUE AMT: ####.##"; Previous#
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(12); "Meter Read: "; DateRead$; TAB(37); "Billed: "; Num2Date$(UBBillRec(1).BillDate)
    PRINT #UBRpt, TAB(12); LEFT$(UBCustRec(1).CustName, 30); TAB(45); "Acct No:"
    PRINT #UBRpt, TAB(12); LEFT$(UBCustRec(1).Addr1, 30); TAB(45); CustAcctNo&
    PRINT #UBRpt, TAB(12); UBCustRec(1).Addr2
    PRINT #UBRpt, TAB(12); QPTrim$(UBCustRec(1).City); " "; UBCustRec(1).State; " "; LEFT$(UBCustRec(1).ZipCode, 5)
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, "~"

