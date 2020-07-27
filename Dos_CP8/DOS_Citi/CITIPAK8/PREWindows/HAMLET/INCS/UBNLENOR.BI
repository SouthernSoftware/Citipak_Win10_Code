'04-30-99 North Lenior

    BilDat$ = Num2Date$(UBBillRec(1).BillDate)
    MtrNum$ = ""

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
          FoundAMtr = True
          MtrNum$ = UBCustRec(1).LocMeters(mChk).MtrNum
        END SELECT
      END IF
    NEXT
    
    IF FoundAMtr = False AND BillInfoRec(1).PRDate <= 0 THEN
      'if no metered services then adjust read dates to billdate
      'and billdate - 30
      DateRead$ = Num2Date$(UBBillRec(1).BillDate)
      PrevDate$ = Num2Date$(UBBillRec(1).BillDate - 30)
    END IF

    AcctNum = UBBillRec(1).CustAcctNo
    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)

    IF FinalFlag AND CDeposit# THEN
      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
    END IF

    IF TotalAmt# > 0 THEN
      TenPct# = 15
    ELSE
      TenPct# = 0
    END IF

    AcctNum = UBBillRec(1).CustAcctNo
    Acct$ = QPTrim$(STR$(AcctNum))
    SELECT CASE AcctNum
    CASE IS < 10
      Acct$ = "00" + Acct$
    CASE IS < 100
      Acct$ = "0" + Acct$
    END SELECT
    AcctLen = LEN(Acct$)
    
    PRINT #UBRpt, CHR$(27); CHR$(48); "~"; TAB(38); USING "######"; FBillNO& + PrintedCnt
    PRINT #UBRpt, ""
    PRINT #UBRpt, ""
    PRINT #UBRpt, ""
    PRINT #UBRpt, ""
    PRINT #UBRpt, ""
    PRINT #UBRpt, TAB(2); LEFT$(UBCustRec(1).ServAddr, 28)
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(2); LEFT$(BilDat$, 6); RIGHT$(BilDat$, 2); TAB(11); LEFT$(DateRead$, 6); RIGHT$(DateRead$, 2); " "; LEFT$(MtrNum$, 6);
    PRINT #UBRpt, TAB(26); USING "######"; AcctNum;
    PRINT #UBRpt, TAB(36); LEFT$(PastDueDate$, 6); RIGHT$(PastDueDate$, 2); TAB(49); USING "######"; AcctNum

    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(2); USING "########"; WCurrRead&;
    PRINT #UBRpt, TAB(12); USING "########"; WPrevRead&;
    PRINT #UBRpt, TAB(22); USING "#######"; WUsageAmt&;
    PRINT #UBRpt, TAB(43); USING "#####.##"; Round#(TotalAmt# + TenPct#); TAB(51); TotalAmt#

    PRINT #UBRpt,
    PRINT #UBRpt,

    FOR WRevCnt = 1 TO 4
      IF UBBillRec(1).RevAmt(WRevCnt) <> 0 THEN
        PRINT #UBRpt, " "; QPTrim$(UBSetUpRec(1).Revenues(WRevCnt).RevName);
        PRINT #UBRpt, TAB(22); USING "#####.##"; UBBillRec(1).RevAmt(WRevCnt)
      ELSE
        PRINT #UBRpt,
      END IF
    NEXT
    IF Previous# <> 0 THEN
      PRINT #UBRpt, TAB(2); "Previous"; TAB(22); USING "#####.##"; Previous#
    ELSE
      PRINT #UBRpt,
    END IF
    
    IF FinalFlag AND CDeposit# THEN
      PRINT #UBRpt, TAB(2); "Deposit"; TAB(22); USING "#####.##"; -UBCustRec(1).DepositAmt;
    END IF

    PRINT #UBRpt, TAB(35); LEFT$(UBCustRec(1).CustName, 29)  'Line 23
    PRINT #UBRpt, TAB(35); UBCustRec(1).Addr1
    PRINT #UBRpt, TAB(35); UBCustRec(1).Addr2
    PRINT #UBRpt, TAB(35); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; LEFT$(UBCustRec(1).ZipCode, 5)

    PRINT #UBRpt, TAB(2); LEFT$(BilDat$, 6); RIGHT$(BilDat$, 2);
    PRINT #UBRpt, TAB(14); USING "#####.##"; Round#(TotalAmt# + TenPct#); TAB(23); TotalAmt#

    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, " "; Message$
    PRINT #UBRpt, " "; Msg2$
    PRINT #UBRpt, " "; Msg3$
    PRINT #UBRpt, "~"

