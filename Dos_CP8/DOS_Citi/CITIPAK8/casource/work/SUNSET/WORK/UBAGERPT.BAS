DEFINT A-Z

DEFSNG A-Z
DEFINT A-Z
SUB AgedBalanceListing
'  SHARED Choice$()
'
'  REDIM Choice$(0 TO 3, 0 TO 1)
'
'  REDIM MChoice$(1 TO 3)
'
'  REDIM Revenue#(15), BookTotal#(100), RevenueName$(15), BookRev#(100, 15)
'
'  UsingBook = False
'  UsingAcct = False
'  UsingName = False
'
'
'  '*********************
'  ' Get Date to Age On
'  LibName$ = "UB"
'  ScrnName$ = "UBAGEDAT"
'
'  DisplayUBScrn ScrnName$
'
'  ' Define Fields
'  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
'
'  ' Define Quick Screen Form Editing Arrays
'  REDIM Frm(1) AS FormInfo
'  REDIM Form$(NumFlds, 2)
'  REDIM Fld(NumFlds) AS FieldInfo
'
'  Frm(1).StayOnField = True
'  ' Get 1st & Last Fields
'  StartEl = 0
'  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
'
'  ' Clear Fields
'  FOR F = 1 TO NumFlds
'    LSET Form$(F, 0) = ""
'  NEXT F
'
'  Choice$(0, 0) = "5"
'  Choice$(1, 0) = "All Customers"
'  Choice$(2, 0) = "Active Customers"
'  Choice$(3, 0) = "InActive Customers"
'
'  Choice$(0, 1) = "3"
'  Choice$(1, 1) = "Customer Name Order"
'  Choice$(2, 1) = "Account Number Order"
'  Choice$(3, 1) = "Location Number Order"
'  FirstTime = True
'
'  Action = 1
'
'  DO
'    EditForm Form$(), Fld(), Frm(1), Cnf, Action
'
'    IF FirstTime THEN
'      FirstTime = False
'      LSET Form$(4, 0) = DATE$
'      LSET Form$(1, 0) = "0"
'      LSET Form$(2, 0) = "99"
'      Action = 1
'    END IF
'
'    SELECT CASE Frm(1).KeyCode
'    CASE F10Key
'      FromBook = Value#(Form$(1, 0), ECode)
'      ThruBook = Value#(Form$(2, 0), ECode)
'      IF FromBook > ThruBook THEN
'        Frm(1).FldNo = 1
'        Action = 1
'        GOTO NotDone
'      END IF
'      SELECT CASE LEFT$(Form$(3, 0), 1)
'      CASE "C"
'        IndexName$ = NameIndexFile
'        UsingName = True
'        OKFlag = True
'      CASE "A"
'        IndexName$ = ""
'        UsingAcct = True
'        OKFlag = True
'      CASE "L"
'        IndexName$ = BookIndexFile
'        UsingBook = True
'        OKFlag = True
'      END SELECT
'
'      AgeDate$ = Form$(4, 0)
'      IF VAL(LEFT$(AgeDate$, 2)) < 1 OR VAL(LEFT$(AgeDate$, 2)) > 12 THEN
'        Done = False
'        Action = 1
'        Frm(1).FldNo = 1
'      ELSEIF VAL(MID$(AgeDate$, 4, 2)) < 1 OR VAL(MID$(AgeDate$, 4, 2)) > 31 THEN
'        Done = False
'        Action = 1
'        Frm(1).FldNo = 1
'      ELSE
'        Done = True
'      END IF
'      Ctype = 1
'      IF LEFT$(Form$(5, 0), 6) = "Active" THEN Ctype = 2
'      IF LEFT$(Form$(5, 0), 8) = "InActive" THEN Ctype = 3
'
'    CASE ESC
'      EXIT SUB
'    CASE ELSE
'      Done = False
'    END SELECT
'NotDone:
'  LOOP UNTIL Done
'
'  '***************
'  AgeDate = Date2Num%(AgeDate$)
'  MaxLines = 55
'  PageNo = 0
'  Dash90$ = STRING$(99, "-")
'
'  REDIM UBCustRec(1) AS NewUBCustRecType
'  UBCustRecLen = LEN(UBCustRec(1))
'
'  REDIM UBTransRec(1) AS UBTransRecType
'  UBTransRecLen = LEN(UBTransRec(1))
'
'  REDIM UBSetup(1) AS UBSetupRecType
'  UBSetUpRecLen = LEN(UBSetup(1))
'
'  IF UsingName OR UsingBook THEN
'    IdxRecLen = 4               'we are using a long integer
'    IdxFileSize& = FileSize(IndexName$)
'    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
'    REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
'    FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
'    NumOfRecs = IdxNumOfRecs
'  ELSE
'    NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
'  END IF
'
'  UBCust = FREEFILE
'  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
'
'  UBTrans = FREEFILE
'  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTrans LEN = UBTransRecLen
'
'  UBRpt = FREEFILE
'  OPEN "UBABLIST.RPT" FOR OUTPUT AS UBRpt
'
'  UBSetup = FREEFILE
'  OPEN "UBSETUP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBSetup LEN = UBSetUpRecLen
'  IF LOF(UBSetup) / UBSetUpRecLen = 0 THEN
'    TownName$ = "Undefined"
'  ELSE
'    GET UBSetup, 1, UBSetup(1)
'    TownName$ = UBSetup(1).UTILNAME
'    TownLen = LEN(RTRIM$(TownName$))
'    TabStop = 40 - (TownLen / 2)
'    IF TabStop < 1 THEN TabStop = 1
'    FOR Rev = 1 TO 15
'      RevenueName$(Rev) = UBSetup(1).Revenues(Rev).REVNAME
'    NEXT Rev
'  END IF
'  CLOSE UBSetup
'
'  BlockClear
'  ShowProcessingScrn "Aged Balance Report."
'
'  GOSUB DoAgedBalRptHeader
'
'  FOR Cnt = 1 TO NumOfRecs
'    IF UsingName OR UsingBook THEN
'      GET UBCust, IdxBuff(Cnt).RecNum, UBCustRec(1)
'      acctnumber = IdxBuff(Cnt).RecNum
'    ELSE
'      GET UBCust, Cnt, UBCustRec(1)
'      acctnumber = Cnt
'    END IF
'
'    'IF acctnumber = 3684 THEN STOP
'
'    CustBook = Value#(UBCustRec(1).BOOK, ECode)
'    IF CustBook < FromBook OR CustBook > ThruBook THEN
'      GOTO SkipEm
'    END IF
'
'    IF Ctype = 3 THEN
'      IF NOT (UBCustRec(1).Status = "I" OR UBCustRec(1).Status = "B") THEN GOTO SkipEm
'    END IF
'    IF Ctype = 2 THEN
'      IF NOT (UBCustRec(1).Status = "A" OR UBCustRec(1).Status = "F") THEN GOTO SkipEm
'    END IF
'
'    'IF LEFT$(UBCustRec(1).CUSTNAME, 6) = "VACANT" OR LEN(QPTrim$(UBCustRec(1).CUSTNAME)) = 0 THEN GOTO SkipEm
'
'    IF UBCustRec(1).DelFlag <> 0 THEN
'      GOTO SkipEm
'    END IF
'
'    IF LineCnt > MaxLines THEN
'      PRINT #UBRpt, FF$
'      GOSUB DoAgedBalRptHeader
'    END IF
'
'    '*************************************
'    '   Main body of Printing goes here
'
'    AcctBal# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
'
'    'IF AcctBal# <= 0 THEN
'    '  GOTO SkipEm
'    'END IF
'
'    'GOTO SkipOverThis
'
'    FOR Rev = 1 TO 15
'      CurRevAmt# = UBCustRec(1).CurrRevAmts(Rev)
'      'PreRevAmt# = UBCustRec(1).PrevRevAmts(Rev)
'      TotRevAmt# = Round#(CurRevAmt# + PreRevAmt#)
'      Revenue#(Rev) = Round#(Revenue#(Rev) + TotRevAmt#)
'      BookRev#(BOOK, Rev) = Round#(BookRev#(BOOK, Rev) + TotRevAmt#)
'    NEXT Rev
'
'SkipOverThis:
'
'
'    'Trans& = UBCustRec(1).LastTrans
'    Process = 0
'
'    GOSUB CalcBal
'
'    LineTotal# = Round#(LineTotal#)
'
'    IF AcctBal# <> LineTotal# THEN
'      Dif# = Round#(AcctBal# - LineTotal#)
'      TDif# = Round#(TDif# + Dif#)
'    END IF
'
'    IF LineTotal# <> 0 THEN
'      '      FOR Rev = 1 TO 15
'      '        CurRevAmt# = UBCustRec(1).CurrRevAmts(Rev)
'      '        PreRevAmt# = UBCustRec(1).PrevRevAmts(Rev)
'      '        TotRevAmt# = Round#(CurRevAmt# + PreRevAmt#)
'      '        Revenue#(Rev) = Round#(Revenue#(Rev) + TotRevAmt#)
'      '        BookRev#(BOOK, Rev) = Round#(BookRev#(BOOK, Rev) + TotRevAmt#)
'      '      NEXT Rev
'
'      'Clean Up Customer Name
'      CUSTNAME$ = UBCustRec(1).CUSTNAME
'      IF ASC(LEFT$(CUSTNAME$, 1)) < 64 OR ASC(LEFT$(CUSTNAME$, 1)) > 90 THEN CUSTNAME$ = ""
'      CUSTNAME$ = QPTrim$(LEFT$(CUSTNAME$, 28))
'      'Clean Up Seq Number
'      SeqNumber$ = UBCustRec(1).SEQNUMB
'      IF ASC(LEFT$(SeqNumber$, 1)) < 45 OR ASC(LEFT$(SeqNumber$, 1)) > 57 THEN SeqNumber$ = ""
'      SeqNumber$ = QPTrim$(SeqNumber$)
'
'      PRINT #UBRpt, UBCustRec(1).BOOK; "-"; SeqNumber$;
'      PRINT #UBRpt, TAB(12); CUSTNAME$;
'      PRINT #UBRpt, TAB(41); USING "######.##"; Amt1#;
'      PRINT #UBRpt, TAB(51); USING "######.##"; Amt2#;
'      PRINT #UBRpt, TAB(61); USING "######.##"; Amt3#;
'      PRINT #UBRpt, TAB(71); USING "######.##"; Amt4#;
'      PRINT #UBRpt, TAB(81); USING "######.##"; LineTotal#;
'      PRINT #UBRpt, TAB(92); USING "#####"; acctnumber
'
'      LineCnt = LineCnt + 1
'      CustCnt = CustCnt + 1
'      TAmt1# = Round#(TAmt1# + Amt1#)
'      TAmt2# = Round#(TAmt2# + Amt2#)
'      TAmt3# = Round#(TAmt3# + Amt3#)
'      TAmt4# = Round#(TAmt4# + Amt4#)
'      TLineTotal# = Round#(TLineTotal# + LineTotal#)
'
'      '*************************************
'      IF AskAbandonPrint% THEN
'        AbortFlag = True
'        EXIT FOR
'      END IF
'
'      ShowPctComp Cnt, NumOfRecs
'    END IF
'SkipEm:
'    LineTotal# = 0: Amt1# = 0: Amt2# = 0: Amt3# = 0: Amt4# = 0
'    TotalPayments = 0
'  NEXT
'
'  GOSUB DoAgedBalRptFooter:
'
'  CLOSE
'
'  ERASE IdxBuff, UBCustRec, UBTransRec
'
'  IF NOT AbortFlag THEN
'    PrintRptFile "Aged Balance Report.", "UBABLIST.RPT", 1, RetCode, EntryPoint
'  END IF
'
'  'KillFile "UBABLIST.RPT"
'
'ExitAgedBalListing:
'
'  EXIT SUB
'
'DoAgedBalRptHeader:
'  PageNo = PageNo + 1
'  PRINT #UBRpt, TAB(TabStop + 4); TownName$
'  PRINT #UBRpt, TAB(26); "Aged Customer Balance Listing Report"; TAB(74); "Page #"; PageNo
'  PRINT #UBRpt, "Ageing Date: "; AgeDate$
'  PRINT #UBRpt, "Report Date: "; DATE$;
'  PRINT #UBRpt, TAB(44); "|------------- D A Y S -------------|"
'  PRINT #UBRpt, "Acct #"; TAB(10); "Customer Name"; TAB(43); "   1-30"; TAB(53); "  31-60"; TAB(63); "  61-90"; TAB(73); "Over 90"; TAB(84); "Balance"; TAB(93); "Acct #"
'  PRINT #UBRpt, Dash90$
'  LineCnt = 6
'  RETURN
'
'DoAgedBalRptFooter:
'  PRINT #UBRpt, Dash90$
'  PRINT #UBRpt, "Totals: Customers = "; CustCnt;
'  PRINT #UBRpt, TAB(41); USING "######.##"; TAmt1#;
'  PRINT #UBRpt, TAB(51); USING "######.##"; TAmt2#;
'  PRINT #UBRpt, TAB(61); USING "######.##"; TAmt3#;
'  PRINT #UBRpt, TAB(71); USING "######.##"; TAmt4#;
'  PRINT #UBRpt, TAB(81); USING "######.##"; TLineTotal#
'  PRINT #UBRpt, FF$;
'  PRINT #UBRpt, "Totals By Book"
'  PRINT #UBRpt, ""
'  PRINT #UBRpt, "Book"; TAB(10); "Book Total"
'  PRINT #UBRpt, STRING$(30, "=")
'  FOR BOOK = 0 TO 99
'    IF BookTotal#(BOOK) <> 0 THEN
'      GBTotal# = Round#(GBTotal# + BookTotal#(BOOK))
'      PRINT #UBRpt, USING "##"; BOOK;
'      PRINT #UBRpt, TAB(10); USING "$$######,#.##"; BookTotal#(BOOK)
'    END IF
'  NEXT
'  PRINT #UBRpt, " TOTAL:"; TAB(10); USING "$$######,#.##"; GBTotal#
'  PRINT #UBRpt, FF$;
'  PRINT #UBRpt, "Totals by Revenue"
'  PRINT #UBRpt, ""
'  PRINT #UBRpt, "Revenue"; TAB(30); "Amount"
'  PRINT #UBRpt, STRING$(50, "-")
'
'  FOR Rev = 1 TO 14
'    Revenue#(Rev) = 0
'  NEXT
'
'  FOR BOOK = 0 TO 99
'    FOR Rev = 1 TO 14
'       Revenue#(Rev) = Round(Revenue#(Rev) + BookRev#(BOOK, Rev))
'    NEXT
'  NEXT
'
'  'FOR Rev = 1 TO 15
'  'Dale
'  FOR Rev = 1 TO 14
'    IF Revenue#(Rev) <> 0 THEN
'      GRTotal# = Round#(GRTotal# + Revenue#(Rev))
'      PRINT #UBRpt, RevenueName$(Rev); TAB(30); USING "$$######,#.##"; Revenue#(Rev)
'    END IF
'  NEXT Rev
'
'
'  PRINT #UBRpt, " Total:"; TAB(30); USING "$$######,#.##"; GRTotal#
'  PRINT #UBRpt, FF$;
'  PRINT #UBRpt, "Totals by Book and Revenue"
'  PRINT #UBRpt, ""
'  PRINT #UBRpt, TAB(10); RevenueName$(1); TAB(30); RevenueName$(2); TAB(50); RevenueName$(3); TAB(70); RevenueName$(4)
'  PRINT #UBRpt, "Book"; TAB(10); RevenueName$(5); TAB(30); RevenueName$(6); TAB(50); RevenueName$(7); TAB(70); RevenueName$(8)
'  PRINT #UBRpt, Dash90$
'  FOR BOOK = 0 TO 99
'    IF BookTotal#(BOOK) <> 0 THEN
'      PRINT #UBRpt, USING "##"; BOOK;
'      PRINT #UBRpt, TAB(10); USING "$$######,#.##"; BookRev#(BOOK, 1);
'      PRINT #UBRpt, TAB(30); USING "$$######,#.##"; BookRev#(BOOK, 2);
'      PRINT #UBRpt, TAB(50); USING "$$######,#.##"; BookRev#(BOOK, 3);
'      PRINT #UBRpt, TAB(70); USING "$$######,#.##"; BookRev#(BOOK, 4)
'      PRINT #UBRpt, TAB(10); USING "$$######,#.##"; BookRev#(BOOK, 5);
'      PRINT #UBRpt, TAB(30); USING "$$######,#.##"; BookRev#(BOOK, 6);
'      PRINT #UBRpt, TAB(50); USING "$$######,#.##"; BookRev#(BOOK, 7);
'      PRINT #UBRpt, TAB(70); USING "$$######,#.##"; BookRev#(BOOK, 8)
'      PRINT #UBRpt, TAB(10); USING "$$######,#.##"; BookRev#(BOOK, 9);
'      PRINT #UBRpt, TAB(30); USING "$$######,#.##"; BookRev#(BOOK, 10);
'      PRINT #UBRpt, TAB(50); USING "$$######,#.##"; BookRev#(BOOK, 11);
'      PRINT #UBRpt, TAB(70); USING "$$######,#.##"; BookRev#(BOOK, 12)
'      PRINT #UBRpt, Dash90$
'    END IF
'  NEXT BOOK
'  RETURN
'
'CalcBal:
'  BOOK = VAL(UBCustRec(1).BOOK)
'
'  'Age1#=1-30 day  Age2#=31-60 day  Age3#=61-90 day   Age4#=Over 90
'  Age1# = 0
'  Age2# = 0
'  Age3# = 0
'  Age4# = 0
'  Amt1# = 0
'  Amt2# = 0
'  Amt3# = 0
'  Amt4# = 0
'
'  TotalPayments# = 0
'  Trans& = UBCustRec(1).LastTrans
'  Process = 0
'  'FirstTrans = True
'
'  DO WHILE Trans& > 0
'    GET UBTrans, Trans&, UBTransRec(1)
'
'    IF UBTransRec(1).RunBalance = 0 THEN
'      GOTO EndReview            'Get Out if Balance =0
'    END IF
'
'    'IF FirstTrans THEN
'    '  FirstTrans = False
'    '  TestDate = AgeDate - 90
'    '  IF UBTransRec(1).TransDate < TestDate THEN
'    '    EXIT DO
'    '  END IF
'    'END IF
'    IF NOT (UBTransRec(1).TransType = 107 OR UBTransRec(1).TransType = 7 OR UBTransRec(1).TransType = 9 OR UBTransRec(1).TransType = 109) THEN
'
'      IF UBTransRec(1).TransDate <= AgeDate THEN
'        'IF UBTransRec(1).RunBalance = 0 AND Process >= 2 THEN GOTO EndReview'Get Out if Balance =0
'        '    IF UBTransRec(1).RunBalance = 0 THEN
'        '      GOTO EndReview   'Get Out if Balance =0
'        '    END IF
'        GOSUB LineCalc
'        Process = Process + 1
'      END IF
'    ELSE
'
'    END IF
'    Trans& = UBTransRec(1).PrevTrans
'  LOOP
'
'EndReview:
'
'  GOSUB FinishCalc
'  RETURN
'
'FinishCalc:
'
'  IF TotalPayments# > Amt4# THEN
'    TotalPayments# = Round#(TotalPayments# - Amt4#)
'    Amt4# = 0
'  ELSE
'    Amt4# = Amt4# - TotalPayments#: TotalPayments# = 0
'  END IF
'  IF TotalPayments# > Amt3# THEN
'    TotalPayments# = Round#(TotalPayments# - Amt3#)
'    Amt3# = 0
'  ELSE
'    Amt3# = Round#(Amt3# - TotalPayments#)
'    TotalPayments# = 0
'  END IF
'  IF TotalPayments# > Amt2# THEN
'    TotalPayments# = Round#(TotalPayments# - Amt2#)
'    Amt2# = 0
'  ELSE
'    Amt2# = Round#(Amt2# - TotalPayments#)
'    TotalPayments# = 0
'  END IF
'
'  Amt1# = Round#(Amt1# - TotalPayments#)
'
'  LineTotal# = Round#(Amt1# + Amt2# + Amt3# + Amt4#)
'
'  'LineTotal# = Round#(LineTotal#)
'
'  BookTotal#(BOOK) = Round#(BookTotal#(BOOK) + LineTotal#)
'  RETURN
'
'LineCalc:
'  TransAmt# = UBTransRec(1).TransAmt
'  Revenue# = 0
'  FOR Rev = 1 TO 15
'    Revenue# = Round#(Revenue# + UBTransRec(1).RevAmt(Rev) + UBTransRec(1).TaxAmt(Rev))
'  NEXT Rev
'  TransAmt# = INT((TransAmt# * 100) + .5) / 100
'
'  'Revenue# = INT((Revenue# * 100) + .5) / 100
'  ' TranLateCharge = 2           '   2=late charge      'NOT USED
'  ' TranReconnectFee = 3         '   3=reconnect fee    'NOT USED
'  ' TranBillPayment = 4          '   4=Bill Payment
'  ' TranAppliedDeposit = 5       '   5=Applied Deposit
'  ' TranPenaltyCharge = 6        '   6=Penalty Charge
'  ' TranDepositPayment = 7       '   7=Deposit Payment
'  ' TranDraftPayment = 8         '   8=Draft Payment
'  ' TranRefundDeposit = 9        '   9=Refund Deposit
'  ' TranBeginBalance = 10        '  10=Beginning Balance
'  ' TranUpwardAdjustment = 11    '  11=Bill Adjustments
'  ' TranDownwardAdjustment = 12  '  12=Bill Adjustments
'  ' TranMiscPayment = 99         '  99=Misc Payment
'
'  '    FOR Rev = 1 TO 15
'  '      Revenue#(Rev) = Round#(Revenue#(Rev) + UBTransRec(1).RevAmt(Rev) + UBTransRec(1).TaxAmt(Rev))
'  '      BookRev#(BOOK, Rev) = Round#(BookRev#(BOOK, Rev) + UBTransRec(1).RevAmt(Rev) + UBTransRec(1).TaxAmt(Rev))
'  '    NEXT Rev
'
'  SELECT CASE UBTransRec(1).TransType
'    'CASE 4, 5, 8, 12, 104, 105, 109
'
'  CASE 4, 8, 12, 104, 109
'    'IF UBTransRec(1).TransType = 5 THEN
'    '  TotalPayments# = TotalPayments# + ABS(UBTransRec(1).TransAmt)
'    'ELSE
'    TotalPayments# = TotalPayments# + UBTransRec(1).TransAmt
'    'END IF
'    TotalPayments# = Round#(TotalPayments#)
'
'    'FOR Rev = 1 TO 15
'    '  Revenue#(Rev) = Round#(Revenue#(Rev) - UBTransRec(1).RevAmt(Rev) - UBTransRec(1).TaxAmt(Rev))
'    '  BookRev#(BOOK, Rev) = Round#(BookRev#(BOOK, Rev) - UBTransRec(1).RevAmt(Rev) - UBTransRec(1).TaxAmt(Rev))
'    'NEXT Rev
'
'    ' IF Round#(Revenue#) <> Round#(TransAmt#) THEN
'    '   Dif# = Round#(TransAmt#) - Round#(Revenue#)
'    '   Revenue#(4) = Revenue#(4) - Dif#
'    '   TransAmt# = TransAmt# - Dif#
'    ' END IF
'    RETURN
'  END SELECT
'
'  SELECT CASE UBTransRec(1).TransType
'  CASE 1, 2, 3, 6, 11, 10, 101, 102, 103, 112, 110
'    Days& = AgeDate - UBTransRec(1).TransDate
'    IF Days& <= 30 THEN
'      Amt1# = Amt1# + UBTransRec(1).TransAmt
'      Amt1# = Round#(Amt1#)
'      ' FOR Rev = 1 TO 15
'      '   Revenue#(Rev) = Round#(Revenue#(Rev) + (UBTransRec(1).RevAmt(Rev) + UBTransRec(1).TaxAmt(Rev)))
'      '   BookRev#(BOOK, Rev) = Round#(BookRev#(BOOK, Rev) + UBTransRec(1).RevAmt(Rev) + UBTransRec(1).TaxAmt(Rev))
'      ' NEXT Rev
'      RETURN
'    END IF
'    IF Days& > 30 AND Days& <= 60 THEN
'      Amt2# = Amt2# + UBTransRec(1).TransAmt
'      Amt2# = Round#(Amt2#)
'      'FOR Rev = 1 TO 15
'      '  Revenue#(Rev) = Revenue#(Rev) + (UBTransRec(1).RevAmt(Rev) + UBTransRec(1).TaxAmt(Rev))
'      '  BookRev#(BOOK, Rev) = Round#(BookRev#(BOOK, Rev) + UBTransRec(1).RevAmt(Rev) + UBTransRec(1).TaxAmt(Rev))
'      'NEXT Rev
'      RETURN
'    END IF
'    IF Days& > 60 AND Days& <= 90 THEN
'      Amt3# = Amt3# + UBTransRec(1).TransAmt
'      Amt3# = Round#(Amt3#)
'      'FOR Rev = 1 TO 15
'      '  Revenue#(Rev) = Round#(Revenue#(Rev) + UBTransRec(1).RevAmt(Rev) + UBTransRec(1).TaxAmt(Rev))
'      '  BookRev#(BOOK, Rev) = Round#(BookRev#(BOOK, Rev) + UBTransRec(1).RevAmt(Rev) + UBTransRec(1).TaxAmt(Rev))
'      'NEXT Rev
'
'      RETURN
'    END IF
'    IF Days& > 90 THEN
'      Amt4# = Amt4# + UBTransRec(1).TransAmt
'      Amt4# = Round#(Amt4#)
'      'FOR Rev = 1 TO 15
'      '  Revenue#(Rev) = Round#(Revenue#(Rev) + UBTransRec(1).RevAmt(Rev) + UBTransRec(1).TaxAmt(Rev))
'      '  BookRev#(BOOK, Rev) = Round#(BookRev#(BOOK, Rev) + UBTransRec(1).RevAmt(Rev) + UBTransRec(1).TaxAmt(Rev))
'      'NEXT Rev
'      RETURN
'    END IF
'  END SELECT
'
'  RETURN
'
END SUB

