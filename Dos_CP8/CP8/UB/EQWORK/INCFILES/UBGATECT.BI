'022698


  Loca$ = UBCustRec(1).Book + "-" + UBCustRec(1).SeqNumb
  FoundAMtr = False
  FOR mChk = 1 TO 7
    IF UBBillRec(1).MtrTypes(mChk) > 0 THEN
      FoundAMtr = True
      UsageAmt& = UBBillRec(1).CurRead(mChk) - UBBillRec(1).PrevRead(mChk)
      CRead# = UBBillRec(1).CurRead(mChk)
      PRead# = UBBillRec(1).PrevRead(mChk)
      IF UsageAmt& < 0 THEN
        MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(mChk))) - 1)
        UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(mChk)) + UBBillRec(1).CurRead(mChk)
      END IF
      EXIT FOR
    END IF
  NEXT

  IF FoundAMtr = False THEN
    'if no metered services then adjust read dates to billdate
    'and billdate - 30
    DateRead$ = Num2Date$(UBBillRec(1).BillDate)
    PrevDate$ = Num2Date$(UBBillRec(1).BillDate - 30)
  END IF

  AcctNum = CustAcctNo&
            
  Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
  TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)

  IF TotalAmt# > 0 THEN
    'TenPct# = 0
    TenPct# = Round#(UBBillRec(1).TransAmt * .1)
  ELSE
    TenPct# = 0
  END IF


  PMth$ = LEFT$(PrevDate$, 2)
  PDay$ = MID$(PrevDate$, 4, 2)
  TMth$ = LEFT$(DateRead$, 2)
  TDay$ = MID$(DateRead$, 4, 2)
  T$ = Num2Date$(UBBillRec(1).PastDueDate)
  DMth$ = LEFT$(T$, 2)
  DDay$ = MID$(T$, 4, 2)
  DYer$ = RIGHT$(T$, 2)

  PRINT #UBRpt, CHR$(27); CHR$(58);
  PRINT #UBRpt, "~"; TAB(25); USING "########"; FBillNO& + PrintedCnt
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(30); USING "#####"; AcctNum; TAB(44); AcctNum
  PRINT #UBRpt, TAB(27); Loca$
  PRINT #UBRpt,
  PRINT #UBRpt,

  PRINT #UBRpt, TAB(17); PMth$; TAB(20); PDay$;
  PRINT #UBRpt, TAB(24); TMth$; TAB(28); TDay$;
  PRINT #UBRpt, TAB(31); DMth$; TAB(35); DDay$; TAB(39); DYer$;
  PRINT #UBRpt, TAB(42); PastDueDate$
  PRINT #UBRpt,
  PRINT #UBRpt,

    IF UBBillRec(1).RevAmt(1) <> 0 THEN
      IF FoundAMtr THEN
        PRINT #UBRpt, TAB(3); USING "########"; PRead#;
        PRINT #UBRpt, TAB(13); USING "########"; CRead#;
        PRINT #UBRpt, TAB(22); USING "######"; UsageAmt&;
      END IF
      PRINT #UBRpt, TAB(29); USING "$###,#.##"; UBBillRec(1).RevAmt(1);
      PRINT #UBRpt, TAB(39); LEFT$(UBSetUpRec(1).Revenues(1).RevName, 2);
    END IF
    PRINT #UBRpt, TAB(42); LEFT$(UBCustRec(1).CustName, 30)
    
    IF UBBillRec(1).RevAmt(2) <> 0 THEN
      IF FoundAMtr THEN
        PRINT #UBRpt, TAB(3); USING "########"; PRead#;
        PRINT #UBRpt, TAB(13); USING "########"; CRead#;
        PRINT #UBRpt, TAB(22); USING "######"; UsageAmt&;
      END IF
      PRINT #UBRpt, TAB(29); USING "$###,#.##"; UBBillRec(1).RevAmt(2);
      PRINT #UBRpt, TAB(39); LEFT$(UBSetUpRec(1).Revenues(2).RevName, 2);
    END IF
    PRINT #UBRpt, TAB(42); LEFT$(UBCustRec(1).Addr1, 30)

    IF UBBillRec(1).RevAmt(3) <> 0 THEN
      PRINT #UBRpt, TAB(29); USING "$###,#.##"; UBBillRec(1).RevAmt(3);
      PRINT #UBRpt, TAB(39); LEFT$(UBSetUpRec(1).Revenues(3).RevName, 2);
    END IF
    PRINT #UBRpt, TAB(42); LEFT$(UBCustRec(1).Addr2, 30)

    IF UBBillRec(1).RevAmt(4) <> 0 THEN
      PRINT #UBRpt, TAB(29); USING "$###,#.##"; UBBillRec(1).RevAmt(4);
      PRINT #UBRpt, TAB(39); LEFT$(UBSetUpRec(1).Revenues(4).RevName, 2);
    END IF
    PRINT #UBRpt, TAB(42); LEFT$(UBCustRec(1).City, 16); " "; UBCustRec(1).State; " "; LEFT$(UBCustRec(1).ZipCode, 5)

    IF UBBillRec(1).RevAmt(5) <> 0 THEN
      PRINT #UBRpt, TAB(29); USING "$###,#.##"; UBBillRec(1).RevAmt(5);
      PRINT #UBRpt, TAB(39); LEFT$(UBSetUpRec(1).Revenues(5).RevName, 2)
    ELSE
      PRINT #UBRpt,
    END IF

    IF Previous# <> 0 THEN
      PRINT #UBRpt, TAB(13); "PAST DUE:"; TAB(29); USING "$###,#.##"; Previous#;
      PRINT #UBRpt, TAB(39); "PD"
    ELSE
      PRINT #UBRpt,
    END IF


    PRINT #UBRpt,
    PRINT #UBRpt, TAB(29); USING "$###,#.##"; TotalAmt#;
    PRINT #UBRpt, TAB(42); USING "$###,#.##"; TotalAmt#
    PRINT #UBRpt, TAB(29); USING "$###,#.##"; Round#(TotalAmt# + TenPct#);
    PRINT #UBRpt, TAB(42); USING "$###,#.##"; Round#(TotalAmt# + TenPct#)
    PRINT #UBRpt, TAB(2); Message$
    PRINT #UBRpt, "~"

