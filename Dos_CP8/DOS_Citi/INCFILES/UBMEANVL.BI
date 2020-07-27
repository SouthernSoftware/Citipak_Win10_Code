
'**********************************
'2800 REM alignment mask
'2810 FOR LOP = 1 TO 5: LPRINT "": NEXT LOP
'     LPRINT TAB(27); STRING$(8, "X");
'     LPRINT TAB(36); STRING$(8, "X");
'     LPRINT TAB(45); STRING$(8, "X");
'     LPRINT TAB(60); "Acct#  XXXXX"
'     LPRINT TAB(60); "Meter# XX-XXXXXXXXX"
'     LPRINT TAB(42); "$XXXX.XX";
'     LPRINT TAB(60); STRING$(20, "X")
'     LPRINT TAB(42); "$XXXX.XX";
'     LPRINT TAB(60); STRING$(20, "X")
'     LPRINT TAB(42); "$XXXX.XX";
'     LPRINT TAB(60); "XXXXXXXXX XX"
'     LPRINT TAB(42); "XX/XX/XX";
'     LPRINT TAB(60); "XXXXX-XXXX"
'     LPRINT TAB(30); "Meter : XX-XXXXXXX";
'     LPRINT TAB(60); "Net   : $XXXX.XX"
'     LPRINT TAB(42); "$XXXX.XX";
'     LPRINT TAB(60); "Gross : $XXXX.XX"
'     LPRINT TAB(42); "$XXXX.XX"
'     LPRINT
'     LPRINT
'     LPRINT
'     LPRINT
'     LPRINT
'     LPRINT
'     LPRINT
'********************************

    WCurrRead& = 0
    WPrevRead& = 0
    WUsageAmt& = 0
    WSerial$ = ""

    WFoundMtr = False
    FOR mChk = 1 TO 7
      IF UBBillRec(1).MtrTypes(mChk) > 0 THEN
        SELECT CASE UBBillRec(1).MtrTypes(mChk)
        CASE 1, 2, 3
          WCurrRead& = UBBillRec(1).CurRead(mChk)
          WPrevRead& = UBBillRec(1).PrevRead(mChk)
          WSerial$ = QPTrim$(UBCustRec(1).LocMeters(mChk).MTRNUM)
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
      TenPct# = Round#(UBBillRec(1).TransAmt * .1)
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
    
    PRINT #UBRpt, "~"; TAB(50); USING "########"; FBillNO& + PrintedCnt
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(26); USING "#########"; WPrevRead&;
    PRINT #UBRpt, TAB(35); USING "#########"; WCurrRead&;
    PRINT #UBRpt, TAB(44); USING "#########"; WUsageAmt&;
    PRINT #UBRpt, TAB(60); "Acct#"; TAB(66); AcctNum
    PRINT #UBRpt, TAB(60); "Meter#"; TAB(66); WSerial$

    IF Previous# <> 0 THEN
      PRINT #UBRpt, TAB(42); USING "$####.##"; Previous#;
    END IF
   
    PRINT #UBRpt, TAB(60); LEFT$(UBCustRec(1).CUSTNAME, 20)
    PRINT #UBRpt, TAB(42); USING "$####.##"; UBBillRec(1).TransAmt;
    PRINT #UBRpt, TAB(60); LEFT$(UBCustRec(1).Addr1, 20)
    PRINT #UBRpt, TAB(42); USING "$####.##"; TotalAmt#;
    PRINT #UBRpt, TAB(60); QPTrim$(UBCustRec(1).City); " "; UBCustRec(1).State
    PRINT #UBRpt, TAB(42); PastDueDate$;
    Zp$ = UBCustRec(1).ZipCode
    PRINT #UBRpt, TAB(60); LEFT$(Zp$, 5); "-"; RIGHT$(Zp$, 4)
    PRINT #UBRpt, TAB(30); "Meter : "; TAB(39); WSerial$;
    PRINT #UBRpt, TAB(60); "Net   :"; TAB(67); USING "$####.##"; TotalAmt#
    PRINT #UBRpt, TAB(42); USING "$####.##"; TenPct#;
    PRINT #UBRpt, TAB(60); "Gross :"; TAB(67); USING "$####.##"; Round#(TotalAmt# + TenPct#)
    PRINT #UBRpt, TAB(42); USING "$####.##"; Round#(TotalAmt# + TenPct#)
    'PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(30); Message$;
    PRINT #UBRpt, TAB(30); Msg2$
    PRINT #UBRpt, TAB(30); Msg3$
    'PRINT #UBRpt,
    'PRINT #UBRpt,
    PRINT #UBRpt, "~"

