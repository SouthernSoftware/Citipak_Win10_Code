'New Utility Bill format 10-28-96 BAR CODE PRINTABLE
'MUST SHOW BOTH METERS OR, TOTAL CONSUMPTION ON THIS BILL

    CustMsg$ = QPTrim$(UBCustRec(1).BILLCMNT)
      
    IF NOT LPIFlag THEN
      LPIFlag = -2
      PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(58);
      'PRINT #UBRpt, CHR$(27); CHR$(48); 'CHR$(27); CHR$(77);
      ' put printer in     8 lpi             12 cpi  oki mode
    END IF
                                                      ':  M
    FoundAMtr = False

    FOR mChk = 1 TO 7
      IF UBBillRec(1).MtrTypes(mChk) > 0 THEN
        FoundAMtr = True
        EXIT FOR
      END IF
    NEXT

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
    
    IF FoundAMtr = False THEN
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
      'TenPct# = 2
      TenPct# = 0
      'TenPct# = 10
      'TenPct# = 5
      'TenPct# = Round#(TotalAmt# * .1)
      'TenPct# = Round#(UBBillRec(1).TransAmt * .1)
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
    
    PRINT #UBRpt, "!"; TAB(14); USING "#####"; FBillNO& + PrintedCnt
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    IF UBBillRec(1).RevAmt(1) <> 0 THEN
      PRINT #UBRpt, TAB(22); USING "#####.##"; UBBillRec(1).RevAmt(1)
    ELSE
      PRINT #UBRpt,
    END IF

    IF UBBillRec(1).RevAmt(2) <> 0 THEN
      PRINT #UBRpt, TAB(22); USING "#####.##"; UBBillRec(1).RevAmt(2);
    END IF
    PRINT #UBRpt, TAB(34); PastDueDate$; TAB(46); USING "##########"; UBBillRec(1).CustAcctNo;
    PRINT #UBRpt, TAB(60); USING "#####.##"; TotalAmt#
    IF UBBillRec(1).RevAmt(5) <> 0 THEN
      PRINT #UBRpt, TAB(22); USING "#####.##"; UBBillRec(1).RevAmt(5)
    ELSE
      PRINT #UBRpt,
    END IF
    IF UBBillRec(1).RevAmt(3) <> 0 THEN
      PRINT #UBRpt, TAB(22); USING "#####.##"; UBBillRec(1).RevAmt(3);
    END IF
    PRINT #UBRpt, TAB(45); LEFT$(UBCustRec(1).ServAddr, 26)

    IF UBBillRec(1).RevAmt(4) <> 0 THEN
      PRINT #UBRpt, TAB(22); USING "#####.##"; UBBillRec(1).RevAmt(4)
    ELSE
      PRINT #UBRpt,
    END IF
    IF Previous# <> 0 THEN
      PRINT #UBRpt, "     Previous"; TAB(22); USING "#####.##"; Previous#;
    END IF
    PRINT #UBRpt, TAB(45); PrevDate$; TAB(60); DateRead$

    IF FinalFlag AND CDeposit# THEN
      PRINT #UBRpt, " Less Deposit"; TAB(22); USING "#####.##"; -UBCustRec(1).DepositAmt
    ELSE
      PRINT #UBRpt,
    END IF
    
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(34); USING "########"; WCurrRead&;
    'PRINT #UBRpt, "00";
    PRINT #UBRpt, TAB(46); USING "########"; WPrevRead&;
    'PRINT #UBRpt, "00";
    PRINT #UBRpt, TAB(61); USING "#######"; WUsageAmt&
    'PRINT #UBRpt, "00"
    PRINT #UBRpt,
    PRINT #UBRpt, BillInfoRec(1).MsgLine1
    PRINT #UBRpt, BillInfoRec(1).MsgLine2
    PRINT #UBRpt, BillInfoRec(1).MsgLine3
    PRINT #UBRpt, BillInfoRec(1).MsgLine4
    
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(60); USING "##########"; UBBillRec(1).CustAcctNo
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(60); USING "######.##"; TotalAmt#
    PRINT #UBRpt,
    
    PRINT #UBRpt, TAB(10); LEFT$(UBCustRec(1).CustName, 29)
    PRINT #UBRpt, TAB(10); UBCustRec(1).Addr1
    PRINT #UBRpt, TAB(10); UBCustRec(1).Addr2
    Zip$ = QPTrim$(UBCustRec(1).ZipCode)
    IF LEN(Zip$) > 5 THEN
      Zip$ = LEFT$(Zip$, 5) + "-" + RIGHT$(Zip$, 4)
    END IF
    PRINT #UBRpt, TAB(10); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; Zip$
    PRINT #UBRpt,
    PRINT #UBRpt,
    
    IF DidADraftFlag THEN
      PRINT #UBRpt, TAB(47); "DRAFT NOTICE DO NOT PAY!!"
    ELSE
      PRINT #UBRpt, CustMsg$
    END IF

    IF DidADraftFlag THEN
      PRINT #UBRpt, TAB(47); "DRAFT DATE: "; Num2Date$(BillInfoRec(1).DrftDate)
    ELSE
      PRINT #UBRpt,
    END IF

    PRINT #UBRpt, "!"

