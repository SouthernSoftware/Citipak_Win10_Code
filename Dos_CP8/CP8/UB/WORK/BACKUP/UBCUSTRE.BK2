  DEFINT A-Z
DECLARE FUNCTION CustHasMsg% (RecNo&)
DECLARE SUB CustStreetListing ()
DECLARE SUB MemberFeeReport ()
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION GetNumOfCust% ()
DECLARE FUNCTION GetNumRateRecs% ()
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION QPStrL$ (LongNumber&)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValI% (Numb$)
DECLARE FUNCTION Round# (N#)
DECLARE FUNCTION WEnvTest% ()
DECLARE SUB ARBalance ()
DECLARE SUB AgedBalanceListing ()
DECLARE SUB BlockClear ()
DECLARE SUB ClearScrn ()
DECLARE SUB CursorOff ()
DECLARE SUB CustCutOffListing ()
DECLARE SUB CustEditMeterEST (CustRec&)
DECLARE SUB CustMessageSystem (RecNo&)
DECLARE SUB CustTaxReport ()
DECLARE SUB CustomerFlatRates ()
DECLARE SUB CustomerInquiry (RecNo&)
DECLARE SUB DetailedTransJournal ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB HideCursor ()
DECLARE SUB KillFile (FileName$)
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetUpLen%)
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetUpLen%)
DECLARE SUB LookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, LocationFlag%)
DECLARE SUB MailingLabel ()
DECLARE SUB MasterBalanceListing ()
DECLARE SUB MasterCustomerListing ()
DECLARE SUB MasterDepositListing ()
DECLARE SUB PaymentSumReport ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
DECLARE SUB RestScrn (array())
DECLARE SUB SaveScrn (array())
DECLARE SUB ShowCursor ()
DECLARE SUB ShowCustConsHist (CustRec&)
DECLARE SUB ShowCustHistory (CustRec&)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB ShowWrkOrdHistory (RecNo&)
DECLARE SUB SortServAddress (IndexText$)
DECLARE SUB SortServiceAddrs (IndexText$)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB WaitForAction ()
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE SUB WestCaroConsumpList ()
DECLARE SUB MakeSequenceIndex (Text$)
  
  
  CONST False = 0, True = NOT False
  
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'ubrate.BI'
  '$INCLUDE: 'newcust.bi'
  
  FF$ = CHR$(12)
  
  REDIM UBSetup(1) AS UBSetupRecType
  UBSetUpRecLen = LEN(UBSetup(1))
  
  UBSetup = FREEFILE
  OPEN "UBSETUP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBSetup LEN = UBSetUpRecLen
  IF LOF(UBSetup) / UBSetUpRecLen = 0 THEN
    TownName$ = "Undefined"
  ELSE
    GET UBSetup, 1, UBSetup(1)
    TownName$ = UBSetup(1).UTILNAME
    TownLen = LEN(RTRIM$(TownName$))
    TabStop = 40 - (TownLen / 2)
    IF TabStop < 1 THEN TabStop = 1
  END IF
  CLOSE UBSetup
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 16)
  
  MChoice$(1) = "Customer Inquiry"
  MChoice$(2) = "Master Customer Listing"
  MChoice$(3) = "Master Balance Listing"
  MChoice$(4) = "Master Deposit Listing"
  MChoice$(5) = "Membership Fees Report"
  MChoice$(6) = "Customers Using Flat Rates"
  MChoice$(7) = "Customer Cut-Off Listing"
  MChoice$(8) = "Customer Street Listing"
  MChoice$(9) = "Detailed Transaction Journal"
  MChoice$(10) = "Payment Summary Report"
  MChoice$(11) = "Mailing Labels"
  MChoice$(12) = "Aged Balance Listing"
  'MChoice$(13) = "A/R Balance Report"
  MChoice$(13) = "Billing/Payment Tax Report"
  
  IF INSTR(TownName$, "TUCKA") THEN
    SelectStop = 15
    MChoice$(14) = "WCU Consumption Report"
    MChoice$(15) = "Exit to OS"
  ELSE
    SelectStop = 14
    MChoice$(14) = "Exit to OS"
  END IF
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 18   'limit the box length to go no lower than line 20
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  '--Find max menu width
  FOR Cnt = 1 TO UBOUND(MChoice$)
    TLen = LEN(MChoice$(Cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT
  
  '--Center Menu within Screen
  Row = ((25 - (UBOUND(MChoice$))) \ 2) + 2
  Col = ((80 - MaxLen) \ 2) - 1
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row + 1, Col, 0
    BlockClear
    
    TitleBox 2, Col, MaxLen + 3, "Customer Reports Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      CustomerInquiry 0&
    CASE 2
      MasterCustomerListing
    CASE 3
      MasterBalanceListing
    CASE 4
      MasterDepositListing
    CASE 5
      MemberFeeReport
    CASE 6
      CustomerFlatRates
    CASE 7
      CustCutOffListing
    CASE 8
      CustStreetListing
    CASE 9
      DetailedTransJournal
    CASE 10
      PaymentSumReport
    CASE 11
      RUN "UBMAILLB"
    CASE 12
      AgedBalanceListing
    CASE 13
      CustTaxReport
    CASE 14
      IF SelectStop = 15 THEN
        WestCaroConsumpList
      ELSE
        HideCursor
        CLS
        END
      END IF
    CASE 14
      HideCursor
      CLS
      END
    END SELECT
  LOOP
  
  IF INSTR(COMMAND$, "TEST") <= 0 THEN
    RUN "ubmenu"
  END IF
  
  HideCursor
  ClearScrn
  END

SUB AgedBalanceListing
  SHARED Choice$()
  
  REDIM Choice$(0 TO 3, 0 TO 1)
  
  REDIM MChoice$(1 TO 3)
  
  REDIM Revenue#(15), BookTotal#(100), RevenueName$(15), BookRev#(100, 15)
  
  UsingBook = False
  UsingAcct = False
  UsingName = False
  
  
  '*********************
  ' Get Date to Age On
  LibName$ = "UB"
  ScrnName$ = "UBAGEDAT"
  
  DisplayUBScrn ScrnName$
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  Frm(1).StayOnField = True
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  Choice$(0, 0) = "5"
  Choice$(1, 0) = "All Customers"
  Choice$(2, 0) = "Active Customers"
  Choice$(3, 0) = "InActive Customers"
  
  Choice$(0, 1) = "3"
  Choice$(1, 1) = "Customer Name Order"
  Choice$(2, 1) = "Account Number Order"
  Choice$(3, 1) = "Location Number Order"
  FirstTime = True
  
  Action = 1
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      FirstTime = False
      LSET Form$(4, 0) = DATE$
      LSET Form$(1, 0) = "0"
      LSET Form$(2, 0) = "99"
      Action = 1
    END IF
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      FromBook = Value#(Form$(1, 0), ECode)
      ThruBook = Value#(Form$(2, 0), ECode)
      IF FromBook > ThruBook THEN
        Frm(1).FldNo = 1
        Action = 1
        GOTO NotDone
      END IF
      SELECT CASE LEFT$(Form$(3, 0), 1)
      CASE "C"
        IndexName$ = NameIndexFile
        UsingName = True
        OKFlag = True
      CASE "A"
        IndexName$ = ""
        UsingAcct = True
        OKFlag = True
      CASE "L"
        IndexName$ = BookIndexFile
        UsingBook = True
        OKFlag = True
      END SELECT
      
      AgeDate$ = Form$(4, 0)
      IF VAL(LEFT$(AgeDate$, 2)) < 1 OR VAL(LEFT$(AgeDate$, 2)) > 12 THEN
        Done = False
        Action = 1
        Frm(1).FldNo = 1
      ELSEIF VAL(MID$(AgeDate$, 4, 2)) < 1 OR VAL(MID$(AgeDate$, 4, 2)) > 31 THEN
        Done = False
        Action = 1
        Frm(1).FldNo = 1
      ELSE
        Done = True
      END IF
      Ctype = 1
      IF LEFT$(Form$(5, 0), 6) = "Active" THEN Ctype = 2
      IF LEFT$(Form$(5, 0), 8) = "InActive" THEN Ctype = 3
      
    CASE ESC
      EXIT SUB
    CASE ELSE
      Done = False
    END SELECT
NotDone:
  LOOP UNTIL Done
  
  '***************
  AgeDate = Date2Num%(AgeDate$)
  MaxLines = 55
  PageNo = 0
  Dash90$ = STRING$(99, "-")
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBTransRec(1) AS UBTransRecType
  UBTransRecLen = LEN(UBTransRec(1))
  
  REDIM UBSetup(1) AS UBSetupRecType
  UBSetUpRecLen = LEN(UBSetup(1))
  
  IF UsingName OR UsingBook THEN
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize(IndexName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
    FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
    NumOfRecs = IdxNumOfRecs
  ELSE
    NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
  END IF
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  UBTrans = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTrans LEN = UBTransRecLen
  
  UBRpt = FREEFILE
  OPEN "UBABLIST.RPT" FOR OUTPUT AS UBRpt
  
  UBSetup = FREEFILE
  OPEN "UBSETUP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBSetup LEN = UBSetUpRecLen
  IF LOF(UBSetup) / UBSetUpRecLen = 0 THEN
    TownName$ = "Undefined"
  ELSE
    GET UBSetup, 1, UBSetup(1)
    TownName$ = UBSetup(1).UTILNAME
    TownLen = LEN(RTRIM$(TownName$))
    TabStop = 40 - (TownLen / 2)
    IF TabStop < 1 THEN TabStop = 1
    FOR Rev = 1 TO 15
      RevenueName$(Rev) = UBSetup(1).Revenues(Rev).REVNAME
    NEXT Rev
  END IF
  CLOSE UBSetup
  
  BlockClear
  ShowProcessingScrn "Aged Balance Report."
  
  GOSUB DoAgedBalRptHeader
  
  FOR Cnt = 1 TO NumOfRecs
    IF UsingName OR UsingBook THEN
      GET UBCust, IdxBuff(Cnt).RecNum, UBCustRec(1)
      acctnumber = IdxBuff(Cnt).RecNum
    ELSE
      GET UBCust, Cnt, UBCustRec(1)
      acctnumber = Cnt
    END IF
    
    'IF acctnumber = 3684 THEN STOP
    
    CustBook = Value#(UBCustRec(1).BOOK, ECode)
    IF CustBook < FromBook OR CustBook > ThruBook THEN
      GOTO SkipEm
    END IF
    
    IF Ctype = 3 THEN
      IF NOT (UBCustRec(1).Status = "I" OR UBCustRec(1).Status = "B") THEN GOTO SkipEm
    END IF
    IF Ctype = 2 THEN
      IF NOT (UBCustRec(1).Status = "A" OR UBCustRec(1).Status = "F") THEN GOTO SkipEm
    END IF
    
    'IF LEFT$(UBCustRec(1).CUSTNAME, 6) = "VACANT" OR LEN(QPTrim$(UBCustRec(1).CUSTNAME)) = 0 THEN GOTO SkipEm
    
    IF UBCustRec(1).DelFlag <> 0 THEN
      GOTO SkipEm
    END IF
    
    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, FF$
      GOSUB DoAgedBalRptHeader
    END IF
    
    '*************************************
    '   Main body of Printing goes here
    
    AcctBal# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
    
    'IF AcctBal# <= 0 THEN
    '  GOTO SkipEm
    'END IF
    
    'GOTO SkipOverThis
    
    FOR Rev = 1 TO 15
      CurRevAmt# = UBCustRec(1).CurrRevAmts(Rev)
      'PreRevAmt# = UBCustRec(1).PrevRevAmts(Rev)
      TotRevAmt# = Round#(CurRevAmt# + PreRevAmt#)
      Revenue#(Rev) = Round#(Revenue#(Rev) + TotRevAmt#)
      BookRev#(BOOK, Rev) = Round#(BookRev#(BOOK, Rev) + TotRevAmt#)
    NEXT Rev
    
SkipOverThis:
    
    
    'Trans& = UBCustRec(1).LastTrans
    Process = 0
    
    GOSUB CalcBal
    
    LineTotal# = Round#(LineTotal#)
    
    IF AcctBal# <> LineTotal# THEN
      Dif# = Round#(AcctBal# - LineTotal#)
      TDif# = Round#(TDif# + Dif#)
    END IF
    
    IF LineTotal# <> 0 THEN
      '      FOR Rev = 1 TO 15
      '        CurRevAmt# = UBCustRec(1).CurrRevAmts(Rev)
      '        PreRevAmt# = UBCustRec(1).PrevRevAmts(Rev)
      '        TotRevAmt# = Round#(CurRevAmt# + PreRevAmt#)
      '        Revenue#(Rev) = Round#(Revenue#(Rev) + TotRevAmt#)
      '        BookRev#(BOOK, Rev) = Round#(BookRev#(BOOK, Rev) + TotRevAmt#)
      '      NEXT Rev
      
      'Clean Up Customer Name
      CUSTNAME$ = UBCustRec(1).CUSTNAME
      IF ASC(LEFT$(CUSTNAME$, 1)) < 64 OR ASC(LEFT$(CUSTNAME$, 1)) > 90 THEN CUSTNAME$ = ""
      CUSTNAME$ = QPTrim$(LEFT$(CUSTNAME$, 28))
      'Clean Up Seq Number
      SeqNumber$ = UBCustRec(1).SEQNUMB
      IF ASC(LEFT$(SeqNumber$, 1)) < 45 OR ASC(LEFT$(SeqNumber$, 1)) > 57 THEN SeqNumber$ = ""
      SeqNumber$ = QPTrim$(SeqNumber$)
      
      PRINT #UBRpt, UBCustRec(1).BOOK; "-"; SeqNumber$;
      PRINT #UBRpt, TAB(12); CUSTNAME$;
      PRINT #UBRpt, TAB(41); USING "######.##"; Amt1#;
      PRINT #UBRpt, TAB(51); USING "######.##"; Amt2#;
      PRINT #UBRpt, TAB(61); USING "######.##"; Amt3#;
      PRINT #UBRpt, TAB(71); USING "######.##"; Amt4#;
      PRINT #UBRpt, TAB(81); USING "######.##"; LineTotal#;
      PRINT #UBRpt, TAB(92); USING "#####"; acctnumber
      
      LineCnt = LineCnt + 1
      CustCnt = CustCnt + 1
      TAmt1# = Round#(TAmt1# + Amt1#)
      TAmt2# = Round#(TAmt2# + Amt2#)
      TAmt3# = Round#(TAmt3# + Amt3#)
      TAmt4# = Round#(TAmt4# + Amt4#)
      TLineTotal# = Round#(TLineTotal# + LineTotal#)
      
      '*************************************
      IF AskAbandonPrint% THEN
        AbortFlag = True
        EXIT FOR
      END IF
      
      ShowPctComp Cnt, NumOfRecs
    END IF
SkipEm:
    LineTotal# = 0: Amt1# = 0: Amt2# = 0: Amt3# = 0: Amt4# = 0
    TotalPayments = 0
  NEXT
  
  GOSUB DoAgedBalRptFooter:
  
  CLOSE
  
  ERASE IdxBuff, UBCustRec, UBTransRec
  
  IF NOT AbortFlag THEN
    PrintRptFile "Aged Balance Report.", "UBABLIST.RPT", 1, RetCode, EntryPoint
  END IF
  
  'KillFile "UBABLIST.RPT"
  
ExitAgedBalListing:
  
  EXIT SUB
  
DoAgedBalRptHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, TAB(TabStop + 4); TownName$
  PRINT #UBRpt, TAB(26); "Aged Customer Balance Listing Report"; TAB(74); "Page #"; PageNo
  PRINT #UBRpt, "Ageing Date: "; AgeDate$
  PRINT #UBRpt, "Report Date: "; DATE$;
  PRINT #UBRpt, TAB(44); "|------------- D A Y S -------------|"
  PRINT #UBRpt, "Acct #"; TAB(10); "Customer Name"; TAB(43); "   1-30"; TAB(53); "  31-60"; TAB(63); "  61-90"; TAB(73); "Over 90"; TAB(84); "Balance"; TAB(93); "Acct #"
  PRINT #UBRpt, Dash90$
  LineCnt = 6
  RETURN
  
DoAgedBalRptFooter:
  PRINT #UBRpt, Dash90$
  PRINT #UBRpt, "Totals: Customers = "; CustCnt;
  PRINT #UBRpt, TAB(41); USING "######.##"; TAmt1#;
  PRINT #UBRpt, TAB(51); USING "######.##"; TAmt2#;
  PRINT #UBRpt, TAB(61); USING "######.##"; TAmt3#;
  PRINT #UBRpt, TAB(71); USING "######.##"; TAmt4#;
  PRINT #UBRpt, TAB(81); USING "######.##"; TLineTotal#
  PRINT #UBRpt, FF$;
  PRINT #UBRpt, "Totals By Book"
  PRINT #UBRpt, ""
  PRINT #UBRpt, "Book"; TAB(10); "Book Total"
  PRINT #UBRpt, STRING$(30, "=")
  FOR BOOK = 0 TO 99
    IF BookTotal#(BOOK) <> 0 THEN
      GBTotal# = Round#(GBTotal# + BookTotal#(BOOK))
      PRINT #UBRpt, USING "##"; BOOK;
      PRINT #UBRpt, TAB(10); USING "$$######,#.##"; BookTotal#(BOOK)
    END IF
  NEXT
  PRINT #UBRpt, " TOTAL:"; TAB(10); USING "$$######,#.##"; GBTotal#
  PRINT #UBRpt, FF$;
  PRINT #UBRpt, "Totals by Revenue"
  PRINT #UBRpt, ""
  PRINT #UBRpt, "Revenue"; TAB(30); "Amount"
  PRINT #UBRpt, STRING$(50, "-")
  
  FOR Rev = 1 TO 14
    Revenue#(Rev) = 0
  NEXT

  FOR BOOK = 0 TO 99
    FOR Rev = 1 TO 14
       Revenue#(Rev) = Round(Revenue#(Rev) + BookRev#(BOOK, Rev))
    NEXT
  NEXT
    
  'FOR Rev = 1 TO 15
  'Dale
  FOR Rev = 1 TO 14
    IF Revenue#(Rev) <> 0 THEN
      GRTotal# = Round#(GRTotal# + Revenue#(Rev))
      PRINT #UBRpt, RevenueName$(Rev); TAB(30); USING "$$######,#.##"; Revenue#(Rev)
    END IF
  NEXT Rev


  PRINT #UBRpt, " Total:"; TAB(30); USING "$$######,#.##"; GRTotal#
  PRINT #UBRpt, FF$;
  PRINT #UBRpt, "Totals by Book and Revenue"
  PRINT #UBRpt, ""
  PRINT #UBRpt, TAB(10); RevenueName$(1); TAB(30); RevenueName$(2); TAB(50); RevenueName$(3); TAB(70); RevenueName$(4)
  PRINT #UBRpt, "Book"; TAB(10); RevenueName$(5); TAB(30); RevenueName$(6); TAB(50); RevenueName$(7); TAB(70); RevenueName$(8)
  PRINT #UBRpt, Dash90$
  FOR BOOK = 0 TO 99
    IF BookTotal#(BOOK) <> 0 THEN
      PRINT #UBRpt, USING "##"; BOOK;
      PRINT #UBRpt, TAB(10); USING "$$######,#.##"; BookRev#(BOOK, 1);
      PRINT #UBRpt, TAB(30); USING "$$######,#.##"; BookRev#(BOOK, 2);
      PRINT #UBRpt, TAB(50); USING "$$######,#.##"; BookRev#(BOOK, 3);
      PRINT #UBRpt, TAB(70); USING "$$######,#.##"; BookRev#(BOOK, 4)
      PRINT #UBRpt, TAB(10); USING "$$######,#.##"; BookRev#(BOOK, 5);
      PRINT #UBRpt, TAB(30); USING "$$######,#.##"; BookRev#(BOOK, 6);
      PRINT #UBRpt, TAB(50); USING "$$######,#.##"; BookRev#(BOOK, 7);
      PRINT #UBRpt, TAB(70); USING "$$######,#.##"; BookRev#(BOOK, 8)
      PRINT #UBRpt, TAB(10); USING "$$######,#.##"; BookRev#(BOOK, 9);
      PRINT #UBRpt, TAB(30); USING "$$######,#.##"; BookRev#(BOOK, 10);
      PRINT #UBRpt, TAB(50); USING "$$######,#.##"; BookRev#(BOOK, 11);
      PRINT #UBRpt, TAB(70); USING "$$######,#.##"; BookRev#(BOOK, 12)
      PRINT #UBRpt, Dash90$
    END IF
  NEXT BOOK
  RETURN
  
CalcBal:
  BOOK = VAL(UBCustRec(1).BOOK)
  
  'Age1#=1-30 day  Age2#=31-60 day  Age3#=61-90 day   Age4#=Over 90
  Age1# = 0
  Age2# = 0
  Age3# = 0
  Age4# = 0
  Amt1# = 0
  Amt2# = 0
  Amt3# = 0
  Amt4# = 0
  
  TotalPayments# = 0
  Trans& = UBCustRec(1).LastTrans
  Process = 0
  'FirstTrans = True
  
  DO WHILE Trans& > 0
    GET UBTrans, Trans&, UBTransRec(1)
    
    IF UBTransRec(1).RunBalance = 0 THEN
      GOTO EndReview            'Get Out if Balance =0
    END IF
    
    'IF FirstTrans THEN
    '  FirstTrans = False
    '  TestDate = AgeDate - 90
    '  IF UBTransRec(1).TransDate < TestDate THEN
    '    EXIT DO
    '  END IF
    'END IF
    IF NOT (UBTransRec(1).TransType = 107 OR UBTransRec(1).TransType = 7 OR UBTransRec(1).TransType = 9 OR UBTransRec(1).TransType = 109) THEN
      
      IF UBTransRec(1).TransDate <= AgeDate THEN
        'IF UBTransRec(1).RunBalance = 0 AND Process >= 2 THEN GOTO EndReview'Get Out if Balance =0
        '    IF UBTransRec(1).RunBalance = 0 THEN
        '      GOTO EndReview   'Get Out if Balance =0
        '    END IF
        GOSUB LineCalc
        Process = Process + 1
      END IF
    ELSE
      
    END IF
    Trans& = UBTransRec(1).PrevTrans
  LOOP
  
EndReview:
  
  GOSUB FinishCalc
  RETURN
  
FinishCalc:
  
  IF TotalPayments# > Amt4# THEN
    TotalPayments# = Round#(TotalPayments# - Amt4#)
    Amt4# = 0
  ELSE
    Amt4# = Amt4# - TotalPayments#: TotalPayments# = 0
  END IF
  IF TotalPayments# > Amt3# THEN
    TotalPayments# = Round#(TotalPayments# - Amt3#)
    Amt3# = 0
  ELSE
    Amt3# = Round#(Amt3# - TotalPayments#)
    TotalPayments# = 0
  END IF
  IF TotalPayments# > Amt2# THEN
    TotalPayments# = Round#(TotalPayments# - Amt2#)
    Amt2# = 0
  ELSE
    Amt2# = Round#(Amt2# - TotalPayments#)
    TotalPayments# = 0
  END IF
  
  Amt1# = Round#(Amt1# - TotalPayments#)
  
  LineTotal# = Round#(Amt1# + Amt2# + Amt3# + Amt4#)
  
  'LineTotal# = Round#(LineTotal#)
  
  BookTotal#(BOOK) = Round#(BookTotal#(BOOK) + LineTotal#)
  RETURN
  
LineCalc:
  TransAmt# = UBTransRec(1).TransAmt
  Revenue# = 0
  FOR Rev = 1 TO 15
    Revenue# = Round#(Revenue# + UBTransRec(1).RevAmt(Rev) + UBTransRec(1).TaxAmt(Rev))
  NEXT Rev
  TransAmt# = INT((TransAmt# * 100) + .5) / 100
  
  'Revenue# = INT((Revenue# * 100) + .5) / 100
  ' TranLateCharge = 2           '   2=late charge      'NOT USED
  ' TranReconnectFee = 3         '   3=reconnect fee    'NOT USED
  ' TranBillPayment = 4          '   4=Bill Payment
  ' TranAppliedDeposit = 5       '   5=Applied Deposit
  ' TranPenaltyCharge = 6        '   6=Penalty Charge
  ' TranDepositPayment = 7       '   7=Deposit Payment
  ' TranDraftPayment = 8         '   8=Draft Payment
  ' TranRefundDeposit = 9        '   9=Refund Deposit
  ' TranBeginBalance = 10        '  10=Beginning Balance
  ' TranUpwardAdjustment = 11    '  11=Bill Adjustments
  ' TranDownwardAdjustment = 12  '  12=Bill Adjustments
  ' TranMiscPayment = 99         '  99=Misc Payment
  
  '    FOR Rev = 1 TO 15
  '      Revenue#(Rev) = Round#(Revenue#(Rev) + UBTransRec(1).RevAmt(Rev) + UBTransRec(1).TaxAmt(Rev))
  '      BookRev#(BOOK, Rev) = Round#(BookRev#(BOOK, Rev) + UBTransRec(1).RevAmt(Rev) + UBTransRec(1).TaxAmt(Rev))
  '    NEXT Rev
  
  SELECT CASE UBTransRec(1).TransType
    'CASE 4, 5, 8, 12, 104, 105, 109
    
  CASE 4, 8, 12, 104, 109
    'IF UBTransRec(1).TransType = 5 THEN
    '  TotalPayments# = TotalPayments# + ABS(UBTransRec(1).TransAmt)
    'ELSE
    TotalPayments# = TotalPayments# + UBTransRec(1).TransAmt
    'END IF
    TotalPayments# = Round#(TotalPayments#)
    
    'FOR Rev = 1 TO 15
    '  Revenue#(Rev) = Round#(Revenue#(Rev) - UBTransRec(1).RevAmt(Rev) - UBTransRec(1).TaxAmt(Rev))
    '  BookRev#(BOOK, Rev) = Round#(BookRev#(BOOK, Rev) - UBTransRec(1).RevAmt(Rev) - UBTransRec(1).TaxAmt(Rev))
    'NEXT Rev
    
    ' IF Round#(Revenue#) <> Round#(TransAmt#) THEN
    '   Dif# = Round#(TransAmt#) - Round#(Revenue#)
    '   Revenue#(4) = Revenue#(4) - Dif#
    '   TransAmt# = TransAmt# - Dif#
    ' END IF
    RETURN
  END SELECT
  
  SELECT CASE UBTransRec(1).TransType
  CASE 1, 2, 3, 6, 11, 10, 101, 102, 103, 112, 110
    Days& = AgeDate - UBTransRec(1).TransDate
    IF Days& <= 30 THEN
      Amt1# = Amt1# + UBTransRec(1).TransAmt
      Amt1# = Round#(Amt1#)
      ' FOR Rev = 1 TO 15
      '   Revenue#(Rev) = Round#(Revenue#(Rev) + (UBTransRec(1).RevAmt(Rev) + UBTransRec(1).TaxAmt(Rev)))
      '   BookRev#(BOOK, Rev) = Round#(BookRev#(BOOK, Rev) + UBTransRec(1).RevAmt(Rev) + UBTransRec(1).TaxAmt(Rev))
      ' NEXT Rev
      RETURN
    END IF
    IF Days& > 30 AND Days& <= 60 THEN
      Amt2# = Amt2# + UBTransRec(1).TransAmt
      Amt2# = Round#(Amt2#)
      'FOR Rev = 1 TO 15
      '  Revenue#(Rev) = Revenue#(Rev) + (UBTransRec(1).RevAmt(Rev) + UBTransRec(1).TaxAmt(Rev))
      '  BookRev#(BOOK, Rev) = Round#(BookRev#(BOOK, Rev) + UBTransRec(1).RevAmt(Rev) + UBTransRec(1).TaxAmt(Rev))
      'NEXT Rev
      RETURN
    END IF
    IF Days& > 60 AND Days& <= 90 THEN
      Amt3# = Amt3# + UBTransRec(1).TransAmt
      Amt3# = Round#(Amt3#)
      'FOR Rev = 1 TO 15
      '  Revenue#(Rev) = Round#(Revenue#(Rev) + UBTransRec(1).RevAmt(Rev) + UBTransRec(1).TaxAmt(Rev))
      '  BookRev#(BOOK, Rev) = Round#(BookRev#(BOOK, Rev) + UBTransRec(1).RevAmt(Rev) + UBTransRec(1).TaxAmt(Rev))
      'NEXT Rev
      
      RETURN
    END IF
    IF Days& > 90 THEN
      Amt4# = Amt4# + UBTransRec(1).TransAmt
      Amt4# = Round#(Amt4#)
      'FOR Rev = 1 TO 15
      '  Revenue#(Rev) = Round#(Revenue#(Rev) + UBTransRec(1).RevAmt(Rev) + UBTransRec(1).TaxAmt(Rev))
      '  BookRev#(BOOK, Rev) = Round#(BookRev#(BOOK, Rev) + UBTransRec(1).RevAmt(Rev) + UBTransRec(1).TaxAmt(Rev))
      'NEXT Rev
      RETURN
    END IF
  END SELECT
  
  RETURN
  
END SUB

SUB CustCutOffListing
  
  SHARED Choice$()
  REDIM Choice$(0 TO 4, 0 TO 1)
  Choice$(0, 0) = "4"
  Choice$(1, 0) = "Past Due Balance"
  Choice$(2, 0) = "Current Balance "
  Choice$(3, 0) = "Total Balance   "
  
  Choice$(0, 1) = "5"
  Choice$(1, 1) = "Customer Name Order   "
  Choice$(2, 1) = "Account Number Order  "
  Choice$(3, 1) = "Location Number Order "
  Choice$(4, 1) = "Sequence Number Order "
  
  AbortFlag = False
  PageNo = 0
  
  REDIM UBSetup(1) AS UBSetupRecType
  UBSetUpRecLen = LEN(UBSetup(1))
  LoadUBSetUpFile UBSetup(), UBSetUpRecLen
  
  TownName$ = QPTrim$(UBSetup(1).UTILNAME)
  TownLen = LEN(TownName$)
  IF TownLen = 0 THEN
    TownName$ = "Undefined"
    TownLen = LEN(TownName$)
  END IF
  TabStop = 40 - (TownLen / 2)
  IF TabStop < 1 THEN TabStop = 1
  
  IF UBSetup(1).BILLCYCL = "Y" THEN
    UseCycle = True
  END IF
  
  LibName$ = "UB"
  ScrnName$ = "UBCOLIST"
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  Frm(1).StayOnField = True
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Set Defaults as Anticpated Response
  Action = 1
  
  FirstTime = True
  
  BlockClear
  DisplayUBScrn ScrnName$
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      FirstTime = False
      Action = 1
      LSET Form$(3, 0) = "0"
      GOSUB ShowBookorCycle
      LSET Form$(1, 0) = "0"
      LSET Form$(2, 0) = "99"
    END IF
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      First = Value#(Form$(1, 0), ECode)
      Last = Value#(Form$(2, 0), ECode)
      IF Last < First THEN
        GOSUB ShowErrScrn
        GOTO LoopJump
      END IF
      MINAMT# = Value#(Form$(3, 0), ECode)
      BalanceType$ = LEFT$(Form$(4, 0), 1)
      Order$ = LEFT$(QPTrim$(Form$(5, 0)), 1)
      
      SELECT CASE Order$
      CASE "C"
        IndexName$ = NameIndexFile
        UsingName = True
        OKFlag = True
      CASE "A"
        IndexName$ = ""
        UsingAcct = True
        OKFlag = True
      CASE "L"
        IndexName$ = BookIndexFile
        UsingBook = True
        OKFlag = True
      CASE "S"
        IndexName$ = "UBTEMP.IDX"
        UsingSeq = True
        OKFlag = True
      CASE ELSE
        GOSUB ShowErrScrn
      END SELECT
    CASE ESC
      GOTO ExitCutOffListing
    CASE ELSE
      Done = False
    END SELECT
    
LoopJump:
  LOOP UNTIL OKFlag
  
  '***************
  MaxLines = 55
  
  PageNo = 0
  Dash80$ = STRING$(80, "-")
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  REDIM UBTransRec(1) AS UBTransRecType
  UBTransRecLen = LEN(UBTransRec(1))
  REDIM UBSetup(1) AS UBSetupRecType
  UBSetUpRecLen = LEN(UBSetup(1))
  
  'AgeDate = Date2Num%(DATE$)
  
  IF UsingName OR UsingBook THEN
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize(IndexName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
    FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
    NumOfRecs = IdxNumOfRecs
  ELSEIF UsingSeq THEN
    MakeSequenceIndex "Sequence Number"
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize(IndexName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
    FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
    NumOfRecs = IdxNumOfRecs
  ELSE
    NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
  END IF
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  UBTrans = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTrans LEN = UBTransRecLen
  UBRpt = FREEFILE
  OPEN "UBCOLIST.RPT" FOR OUTPUT AS UBRpt
  
  IF INSTR(TownName$, "LONGVIEW") THEN
    LVFlag = True
  END IF
  
  BlockClear
  ShowProcessingScrn "Customer Cut-Off Report."
  
  GOSUB DoCutOffRptHeader
  
  FOR Cnt = 1 TO NumOfRecs
    IF UsingName OR UsingBook OR UsingSeq THEN
      AcctNo& = IdxBuff(Cnt).RecNum
    ELSE
      AcctNo& = Cnt
    END IF
    GET UBCust, AcctNo&, UBCustRec(1)
    IF UseCycle THEN
      CustCheck = UBCustRec(1).BILLCYCL
    ELSE
      CustCheck = Value#(UBCustRec(1).BOOK, ECode)
    END IF
    IF CustCheck >= First AND CustCheck <= Last THEN
      IF UBCustRec(1).CUTOFFYN = "Y" THEN
        IF UBCustRec(1).Status = "A" THEN
          SELECT CASE BalanceType$
          CASE "P"
            CutOffBalance# = UBCustRec(1).PrevBalance
          CASE "C"
            CutOffBalance# = UBCustRec(1).CurrBalance
          CASE "T"
            CutOffBalance# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
          END SELECT
          RealBalance# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
          IF CutOffBalance# >= MINAMT# THEN
            IF RealBalance# > 0 AND CutOffBalance# > 0 THEN
              GOSUB printline
            END IF
          END IF
        END IF
      END IF
    END IF
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, FF$
      GOSUB DoCutOffRptHeader
    END IF
    ShowPctComp Cnt, NumOfRecs
  NEXT
  
  GOSUB DoCutOffRptFooter:
  
  CLOSE
  
  ERASE IdxBuff, UBCustRec
  
  IF NOT AbortFlag THEN
    PrintRptFile "Customer Cut-Off Report.", "UBCOLIST.RPT", 1, RetCode, EntryPoint
  END IF
  
  KillFile "UBCOLIST.RPT"
  
ExitCutOffListing:
  
  EXIT SUB
  
DoCutOffRptHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, TAB(TabStop); TownName$
  PRINT #UBRpt, TAB(26); "Customer Cut Off Listing Report"; TAB(70); "Page #"; PageNo
  PRINT #UBRpt, "Report Date: "; DATE$
  PRINT #UBRpt, "Location  Acct#  Customer Name"; TAB(43); "Service Address"; TAB(74); "Balance"
  IF NOT LVFlag > 0 THEN
    PRINT #UBRpt, "Meter Number"; TAB(17); "Last Read"; TAB(30); "Reading at Cut-Off"
  END IF
  PRINT #UBRpt, Dash80$
  LineCnt = 6
  RETURN
  
  'DoCutOffRptHeader:
  '  PageNo = PageNo + 1
  '  PRINT #UBRpt, TAB(TabStop); TownName$
  '  PRINT #UBRpt, TAB(26); "Customer Cut Off Listing Report"; TAB(70); "Page #"; PageNo
  '  PRINT #UBRpt, "Report Date: "; DATE$
  '  PRINT #UBRpt, "Location  Acct#  Customer Name"; TAB(43); "Service Address"; TAB(64); "Previous"; TAB(74); "Current"
  '  IF NOT LVFlag > 0 THEN
  '    PRINT #UBRpt, "Meter Number"; TAB(17); "Last Read"; TAB(30); "Reading at Cut-Off"; TAB(64); "Balance"; TAB(74); "Balance"
  '  END IF
  '  PRINT #UBRpt, Dash80$
  '  LineCnt = 6
  '  RETURN
  
DoCutOffRptFooter:
  PRINT #UBRpt, ""
  PRINT #UBRpt, "Total Customers to Cut Off: "; USING "#####,#"; CustomerCnt
  PRINT #UBRpt, "     Total Cut Off Balance: "; USING "#####.##"; TBalance#
  PRINT #UBRpt,
  PRINT #UBRpt, "         Report Parameters"
  PRINT #UBRpt, "                      From:"; First; "  Thru:"; Last
  PRINT #UBRpt, "              Balance Type: "; Form$(4, 0)
  
  RETURN
  
ShowErrScrn:
  
  REDIM TempScrn(0)
  SaveScrn TempScrn()
  CursorOff
  WazzWind 13, 27, 19, 63, 78, 1, True
  IF Last < First THEN
    IF UseCycle THEN
      QPrintRC "Invalid Cycle Selection!", 15, 34, 79
    ELSE
      QPrintRC "Invalid Book Selection!", 15, 34, 79
    END IF
    Frm(1).FldNo = 1
  ELSE
    QPrintRC "Invalid Printing Order!", 15, 34, 79
    Frm(1).FldNo = 5
  END IF
  QPrintRC "Press Any Key To Continue.", 17, 33, 79
  WaitForAction
  RestScrn TempScrn()
  
  Action = 1
  RETURN
  
  
printline:
  TBalance# = Round#(TBalance# + RealBalance#)
  '*************************************
  '   Main body of Printing goes here
  PRINT #UBRpt, UBCustRec(1).BOOK; "-"; UBCustRec(1).SEQNUMB; USING "######"; AcctNo&;
  PRINT #UBRpt, TAB(18); LEFT$(UBCustRec(1).CUSTNAME, 25); LEFT$(UBCustRec(1).SERVADDR, 22);
  PRINT #UBRpt, TAB(70); USING "$$####,#.##"; RealBalance#
  'PRINT #UBRpt, TAB(65); USING "($$####,#.##)"; CutOffBalance#
  LineCnt = LineCnt + 1
  
  IF LVFlag THEN
    GOTO Skip1
  END IF
  
  FOR MrtCnt = 1 TO 7
    IF LEN(QPTrim$(UBCustRec(1).LocMeters(MrtCnt).MTRType)) > 0 THEN
      PRINT #UBRpt, UBCustRec(1).LocMeters(MrtCnt).MTRNUM;
      PRINT #UBRpt, TAB(17); UBCustRec(1).LocMeters(MrtCnt).CurRead;
      PRINT #UBRpt, TAB(30); "_________________"
      LineCnt = LineCnt + 1
    END IF
  NEXT
  PRINT #UBRpt, STRING$(79, "-")
  LineCnt = LineCnt + 1
  
Skip1:
  IF Landis THEN
    PRINT #UBRpt,
    LineCnt = LineCnt + 2
  END IF
  CustomerCnt = CustomerCnt + 1
  '*************************************
  RETURN
  
ShowBookorCycle:
  IF UseCycle THEN
    QPrintRC "From Cycle", 7, 27, -1
    QPrintRC "Thru Cycle", 9, 27, -1
  ELSE
    QPrintRC "From Book", 7, 28, -1
    QPrintRC "Thru Book", 9, 28, -1
  END IF
  
  RETURN
END SUB

SUB CustomerFlatRates
  
  SHARED Choice$()
  REDIM Choice$(0 TO 15, 0 TO 2)
  
  REDIM UBSetup(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetup(), UBSetUpRecLen
  
  FOR Cnt = 1 TO 15
    Rev$ = QPTrim$(UBSetup(1).Revenues(Cnt).REVNAME)
    IF LEN(Rev$) > 0 THEN
      Choice$(Cnt, 1) = QPTrim$(STR$(Cnt)) + ")  " + Rev$
    ELSE
      EXIT FOR
    END IF
  NEXT
  
  Choice$(0, 1) = "1"
  
  
  Choice$(0, 2) = "2"
  Choice$(1, 2) = " Customer Name Order   "
  Choice$(2, 2) = " Account Number Order  "
  Choice$(3, 2) = " Location Number Order "
  
  UsingBook = False
  UsingAcct = False
  UsingName = False
  
  AbortFlag = False
  PageNo = 0
  
  LibName$ = "UB"
  ScrnName$ = "UBFLCUST"
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  Frm(1).StayOnField = True
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Set Defaults as Anticpated Response
  
  Action = 1
  BlockClear
  DisplayUBScrn ScrnName$
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      RevSource = Value#(Form$(1, 0), ECode)
      Order$ = LEFT$(QPTrim$(Form$(2, 0)), 1)
      SELECT CASE Order$
      CASE "C"
        IndexName$ = NameIndexFile
        UsingName = True
        OKFlag = True
      CASE "A"
        IndexName$ = ""
        UsingAcct = True
        OKFlag = True
      CASE "L"
        IndexName$ = BookIndexFile
        UsingBook = True
        OKFlag = True
      CASE ELSE
        'GOSUB ShowErrScrn
      END SELECT
    CASE ESC
      GOTO ExitCustomerFlat
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL OKFlag
  
  '***************
  MaxLines = 55
  PageNo = 0
  Dash80$ = STRING$(80, "-")
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  IF UsingName OR UsingBook THEN
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize(IndexName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
    FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
    NumOfRecs = IdxNumOfRecs
  ELSE
    NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
  END IF
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  UBRpt = FREEFILE
  OPEN "UBFRLIST.RPT" FOR OUTPUT AS UBRpt
  
  
  TownName$ = QPTrim$(UBSetup(1).UTILNAME)
  TabStop = 40 - (LEN(TownName$) / 2)
  IF TabStop < 1 THEN TabStop = 1
  TownName$ = SPACE$(TabStop) + TownName$
  
  BlockClear
  ShowProcessingScrn "Flat Rate Listing."
  
  GOSUB DoMastFlatHeader
  
  FOR Cnt = 1 TO NumOfRecs
    IF UsingName OR UsingBook THEN
      AcctNo = IdxBuff(Cnt).RecNum
    ELSE
      AcctNo = Cnt
    END IF
    
    GET UBCust, AcctNo, UBCustRec(1)
    
    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, FF$
      GOSUB DoMastFlatHeader
    END IF
    
    IF UBCustRec(1).DelFlag <> -1 AND UBCustRec(1).Status = "A" THEN
      DidOne = False
      FOR FRCnt = 1 TO 4
        IF UBCustRec(1).FlatRates(FRCnt).REVSRC = RevSource THEN
          DidOne = True
          PRINT #UBRpt, USING "Acct:#####"; AcctNo;
          PRINT #UBRpt, TAB(15); LEFT$(UBCustRec(1).CUSTNAME, 30);
          IF UBCustRec(1).FlatRates(FRCnt).NumMin < 1 THEN
            UBCustRec(1).FlatRates(FRCnt).NumMin = 1
          END IF
          CustFlat# = Round#(UBCustRec(1).FlatRates(FRCnt).FRAMT * UBCustRec(1).FlatRates(FRCnt).NumMin)
          PRINT #UBRpt, TAB(45); QPTrim$(UBCustRec(1).FlatRates(FRCnt).FRDESC); TAB(65); USING "#####.##"; CustFlat#
          TotalFlat# = Round#(TotalFlat# + CustFlat#)
          EXIT FOR
        END IF
      NEXT
      IF DidOne THEN
        PRINT #UBRpt, Dash80$
        LineCnt = LineCnt + 1
        CustCnt = CustCnt + 1
      END IF
      '*************************************
      IF AskAbandonPrint% THEN
        AbortFlag = True
        EXIT FOR
      END IF
    END IF
    ShowPctComp Cnt, NumOfRecs
  NEXT
  
  GOSUB DoMastFlatFooter:
  
  CLOSE UBCust, UBRpt
  
  ERASE IdxBuff, UBCustRec
  
  IF NOT AbortFlag THEN
    PrintRptFile "Master Customer Report.", "UBFRLIST.RPT", 1, RetCode, EntryPoint
  END IF
  
  KillFile "UBFRLIST.RPT"
  
ExitCustomerFlat:
  
  EXIT SUB
  
DoMastFlatHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, TAB(TabStop); TownName$
  PRINT #UBRpt, TAB(28); "Customer Flat Rate Listing"; TAB(70); "Page #"; PageNo
  PRINT #UBRpt, "Report Date: "; DATE$
  PRINT #UBRpt, Dash80$
  LineCnt = 4
  RETURN
  
DoMastFlatFooter:
  PRINT #UBRpt, "Totals:"; TAB(10); "Customers: "; USING "#####,#"; CustCnt
  PRINT #UBRpt, TAB(10); "   Amount:"; USING "######,#.##"; TotalFlat#
  RETURN
  
  
  
END SUB

SUB DetailedTransJournal
  
  SHARED Choice$()
  
  REDIM TempScrn(0)
  REDIM Choice$(0 TO 14, 0 TO 1)
  REDIM RevenueName$(15)
  
  LibName$ = "UB"
  ScrnName$ = "UBTRJOU"
  
  BlockClear
  DisplayUBScrn ScrnName$
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  Frm(1).StayOnField = True
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  Choice$(0, 0) = "5"
  Choice$(1, 0) = " 1 - Utility Bill"
  Choice$(2, 0) = " 4 - Payment"
  Choice$(3, 0) = " 5 - Applied Deposit"
  Choice$(4, 0) = " 6 - Penalty Charge"
  Choice$(5, 0) = " 7 - Deposit Payment"
  Choice$(6, 0) = " 9 - Refunded Deposit"
  Choice$(7, 0) = "11 - Up Adjustment"
  Choice$(8, 0) = "12 - Down Adjustment"
  
  Choice$(0, 1) = "9"
  Choice$(1, 1) = "Customer Name Order   "
  Choice$(2, 1) = "Account Number Order  "
  Choice$(3, 1) = "Location Number Order "
  
  
  ' Set Defaults as Anticpated Response
  Form$(1, 0) = DATE$
  Form$(2, 0) = DATE$
  Form$(6, 0) = "0"
  Form$(7, 0) = "N"
  FirstTime = True
  Action = 1
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      FirstTime = False
      LSET Form$(1, 0) = DATE$
      LSET Form$(2, 0) = DATE$
      LSET Form$(3, 0) = "0"
      LSET Form$(4, 0) = "99"
      LSET Form$(6, 0) = "0"
      LSET Form$(7, 0) = "N"
      Action = 1
    END IF
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key, -103
      GOSUB CheckDetailParms
    CASE ESC
      EXIT SUB
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  
  '***************
  'Check Date Range for Valid Dates
  ' Set Up Specifications from Input Screen
  
  '***************
  MaxLines = 55
  PageNo = 0
  Dash120$ = STRING$(121, "-")
  
  REDIM RevTotals(1 TO 15)  AS DOUBLE
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBSetup(1) AS UBSetupRecType
  UBSetUpRecLen = LEN(UBSetup(1))
  
  REDIM UBTrans(1) AS UBTransRecType
  UBTransRecLen = LEN(UBTrans(1))
  
  IF UsingName OR UsingBook THEN
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize(IndexName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
    FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
    NumOfRecs = IdxNumOfRecs
  ELSE
    NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
  END IF
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  UBTrans = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTrans LEN = UBTransRecLen
  UBRpt = FREEFILE
  OPEN "UBDJLIST.RPT" FOR OUTPUT AS UBRpt
  UBSetup = FREEFILE
  
  OPEN "UBSETUP.DAT" FOR RANDOM SHARED AS UBSetup LEN = UBSetUpRecLen
  IF LOF(UBSetup) / UBSetUpRecLen = 0 THEN
    TownName$ = "Undefined"
    ' Set Revenue Names to Nothing
    FOR RCnt = 1 TO 15
      RevenueName$(RCnt) = "Not Set"
    NEXT RCnt
  ELSE
    GET UBSetup, 1, UBSetup(1)
    FOR RCnt = 1 TO 15
      RevenueName$(RCnt) = QPTrim$(UBSetup(1).Revenues(RCnt).REVNAME)
    NEXT RCnt
    RCnt = 1
    DO WHILE RCnt <= 15
      IF RevenueName$(RCnt) = "" THEN
        MaxRevenue = RCnt - 1
        EXIT DO
      END IF
      RCnt = RCnt + 1
    LOOP
    TownName$ = UBSetup(1).UTILNAME
    TownLen = LEN(RTRIM$(TownName$))
    TabStop = 40 - (TownLen / 2)
    IF TabStop < 1 THEN TabStop = 1
  END IF
  CLOSE UBSetup
  
  'Special Code just for ellenboro!!
  IF INSTR(TownName$, "ELLENBO") > 0 THEN
    EllenFlag = True
  END IF
  
  BlockClear
  ShowProcessingScrn "Detailed Journal Report."
  
  GOSUB DoDetailedRptHeader
  
  FOR Cnt = 1 TO NumOfRecs
    IF UsingName OR UsingBook THEN
      GET UBCust, IdxBuff(Cnt).RecNum, UBCustRec(1)
    ELSE
      GET UBCust, Cnt, UBCustRec(1)
    END IF
    
    IF UBCustRec(1).DelFlag <> 0 THEN
      GOTO SkipThisOne
    END IF
    
    IF UseType THEN
      ThisType$ = QPTrim$(UBCustRec(1).CUSTTYPE)
      IF ThisType$ <> CUSTTYPE$ THEN
        GOTO SkipThisOne
      END IF
    END IF
    
    CustBook = VAL(UBCustRec(1).BOOK)
    IF CustBook < FromBook OR CustBook > ThruBook THEN
      GOTO SkipThisOne
    END IF
    
    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, FF$
      GOSUB DoDetailedRptHeader
    END IF
    '*************************************
    '   Main Body of Printing goes here
    BadCount = 0
    Trans& = UBCustRec(1).LastTrans
    DO WHILE Trans& <> 0
      GET UBTrans, Trans&, UBTrans(1)
      IF NOT EllenFlag THEN
        IF UBTrans(1).TransDate < BegDate THEN
          BadCount = BadCount + 1
          IF BadCount > 3 THEN
            EXIT DO
          END IF
        END IF
      END IF
      'Check Date, Operator and Trans Type
      IF (UBTrans(1).TransDate >= BegDate AND UBTrans(1).TransDate <= EndDate) THEN
        IF (UBTrans(1).OperatorNumber >= BegOperator AND UBTrans(1).OperatorNumber <= EndOperator) THEN
          IF (UBTrans(1).TransType >= BegTrans AND UBTrans(1).TransType <= EndTrans OR (UBTrans(1).TransType >= BegTrans + 100 AND UBTrans(1).TransType <= EndTrans + 100)) THEN
            GOSUB DefineType
            PRINT #UBRpt, Num2Date$(UBTrans(1).TransDate); TAB(11); USING "#####"; UBTrans(1).CustAcctNo;
            'PRINT #UBRpt, Num2Date$(UBTrans(1).TransDate); TAB(11); ASC(UBTrans(1).Posted2GL); 'USING "#####"; UBTrans(1).CustAcctNo;
            PRINT #UBRpt, TAB(20); LEFT$(UBCustRec(1).CUSTNAME, 33);
            PRINT #UBRpt, TAB(55); TransDesc$;
            'PRINT #UBRpt, TAB(55); Trans&;
            PRINT #UBRpt, TAB(80); UBTrans(1).OperatorNumber;
            PRINT #UBRpt, TAB(90); LEFT$(UBTrans(1).TransDesc, 20);
            PRINT #UBRpt, TAB(110); USING "$$#####,#.##"; Amount#
            'PRINT #UBRpt, "  "; "!"; UBTrans(1).Posted2GL; "!"
            LineCnt = LineCnt + 1
            TotalTrans# = Round#(TotalTrans# + Amount#)
            TransCnt& = TransCnt& + 1
            IF Detail$ = "Y" THEN
              PRINT #UBRpt, "Revenue Source Breakdown ................................."
              LineCnt = LineCnt + 1
              FOR RCnt = 1 TO MaxRevenue STEP 3
                PRINT #UBRpt, RevenueName$(RCnt); TAB(16); USING "#####.##"; UBTrans(1).RevAmt(RCnt);
                PRINT #UBRpt, TAB(30); RevenueName$(RCnt + 1); TAB(46); USING "#####.##"; UBTrans(1).RevAmt(RCnt + 1);
                PRINT #UBRpt, TAB(60); RevenueName$(RCnt + 2); TAB(76); USING "#####.##"; UBTrans(1).RevAmt(RCnt + 2)
                LineCnt = LineCnt + 1
              NEXT RCnt
              'IF UBTrans(1).TransType = TranUpwardAdjustment OR UBTrans(1).TransType = TranDownwardAdjustment THEN
              '  FOR RCnt = 1 TO 7
              '    PRINT #UBRpt, RevenueName$(RCnt); TAB(16); USING "#####.##"; UBTrans(1).RevAmt(RCnt);
              '  PRINT #UBRpt, TAB(30); RevenueName$(RCnt + 1); TAB(46); USING "#####.##"; UBTrans(1).RevAmt(RCnt + 1);
              '  PRINT #UBRpt, TAB(60); RevenueName$(RCnt + 2); TAB(76); USING "#####.##"; UBTrans(1).RevAmt(RCnt + 2)
              '  LineCnt = LineCnt + 1
              'NEXT RCnt
              PRINT #UBRpt, Dash120$
              LineCnt = LineCnt + 1
            END IF
            FOR RCnt = 1 TO MaxRevenue
              RevTotals(RCnt) = Round#(RevTotals(RCnt) + UBTrans(1).RevAmt(RCnt))
            NEXT
            IF LineCnt > MaxLines THEN
              PRINT #UBRpt, FF$
              GOSUB DoDetailedRptHeader
            END IF
          END IF
        END IF
      END IF
      
      IF AskAbandonPrint% THEN
        AbortFlag = True
        EXIT FOR
      END IF
      Trans& = UBTrans(1).PrevTrans
    LOOP
SkipThisOne:
    ShowPctComp Cnt, NumOfRecs
  NEXT
  
  GOSUB DoDetailedRptFooter
  PRINT #UBRpt, FF$;
  
  CLOSE
  
  ERASE IdxBuff, UBCustRec
  
  IF NOT AbortFlag THEN
    PrintRptFile "Detailed Journal Report.", "UBDJLIST.RPT", 1, RetCode, EntryPoint
  END IF
  
  'KillFile "UBDJLIST.RPT"
  
ExitDetailedListing:
  
  EXIT SUB
  
DoDetailedRptHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, TAB(TabStop); TownName$
  PRINT #UBRpt, TAB(28); "Detailed Transaction Report"; TAB(90); "Page #"; PageNo
  PRINT #UBRpt, "Report Date: "; DATE$
  PRINT #UBRpt, "Beginning Transaction Date: "; BegDate$;
  IF VAL(Operator$) = 0 THEN
    PRINT #UBRpt, TAB(90); " Operator #: ALL"
  ELSE
    PRINT #UBRpt, TAB(90); " Operator #: "; Operator$
  END IF
  PRINT #UBRpt, "   Ending Transaction Date: "; EndDate$;
  PRINT #UBRpt, TAB(90); "Show Detail: "; Detail$
  PRINT #UBRpt, "             Customer Type: ";
  
  IF UseType THEN
    PRINT #UBRpt, CUSTTYPE$
  ELSE
    PRINT #UBRpt, "N/A"
  END IF
  
  PRINT #UBRpt, "  Date"; TAB(11); "Acct #"; TAB(20); "Customer Name"; TAB(55); "Description"; TAB(80); "Oper#"; TAB(90); "Trans Type"; TAB(113); "   Amount"
  PRINT #UBRpt, Dash120$
  LineCnt = 8
  RETURN
  
DoDetailedRptFooter:
  PRINT #UBRpt, Dash120$
  PRINT #UBRpt, "Transactions: "; TransCnt&; "                                                       Total of Transactions: "; USING "$$#######,#.##"; TotalTrans#
  PRINT #UBRpt, FF$
  PageNo = PageNo + 1
  PRINT #UBRpt, TAB(TabStop); TownName$
  PRINT #UBRpt, TAB(28); "Detailed Transaction Report"; TAB(90); "Page #"; PageNo
  PRINT #UBRpt, "Report Date: "; DATE$
  PRINT #UBRpt, "Beginning Transaction Date: "; BegDate$;
  IF VAL(Operator$) = 0 THEN
    PRINT #UBRpt, TAB(90); " Operator #: ALL"
  ELSE
    PRINT #UBRpt, TAB(90); " Operator #: "; Operator$
  END IF
  PRINT #UBRpt, "   Ending Transaction Date: "; EndDate$;
  PRINT #UBRpt, TAB(90); "Show Detail: "; Detail$
  PRINT #UBRpt, ""
  PRINT #UBRpt, "Revenue Summary"; TAB(38); "Amount"
  PRINT #UBRpt, Dash120$
  TotalRevsAmt# = 0
  FOR RCnt = 1 TO MaxRevenue
    TotalRevsAmt# = Round#(TotalRevsAmt# + RevTotals(RCnt))
    PRINT #UBRpt, RevenueName$(RCnt), TAB(35); USING "########.##"; RevTotals(RCnt)
  NEXT
  PRINT #UBRpt,
  PRINT #UBRpt, "Total Amount"; TAB(35); USING "########.##"; TotalRevsAmt#
  RETURN
  
DefineType:
  SELECT CASE UBTrans(1).TransType
  CASE 1, 101
    TransDesc$ = "Util Bill"
    Amount# = UBTrans(1).TransAmt
  CASE 2, 102
    TransDesc$ = "Late Charge"
    Amount# = UBTrans(1).TransAmt
  CASE 3
    TransDesc$ = "Reconnect"
    Amount# = UBTrans(1).TransAmt
  CASE 4, 104
    TransDesc$ = "Reg Payment"
    Amount# = UBTrans(1).TransAmt
  CASE 5, 105
    TransDesc$ = "Applied Dep"
    'Amount# = -UBTrans(1).TransAmt
    Amount# = ABS(UBTrans(1).TransAmt)
  CASE 6
    TransDesc$ = "Penalty Chg"
    Amount# = UBTrans(1).TransAmt
  CASE 7, 107
    TransDesc$ = "Dep. Payment"
    Amount# = UBTrans(1).TransAmt
  CASE 8
    TransDesc$ = "Draft Paymt"
    Amount# = UBTrans(1).TransAmt * -1
  CASE 9, 109
    TransDesc$ = "Refunded Dep"
    Amount# = ABS(UBTrans(1).TransAmt)
  CASE 10, 110
    TransDesc$ = "Beg Balance"
    Amount# = UBTrans(1).TransAmt
  CASE 11, 111
    TransDesc$ = UBTrans(1).BillMsg
    Amount# = UBTrans(1).TransAmt
  CASE 12, 112
    TransDesc$ = UBTrans(1).BillMsg
    Amount# = UBTrans(1).TransAmt
  CASE 99
    TransDesc$ = "Misc Payment"
    Amount# = UBTrans(1).TransAmt
  CASE ELSE
    TransDesc$ = "UNKNOWN"
    Amount# = UBTrans(1).TransAmt
  END SELECT
  RETURN
  
CheckDetailParms:
  
  BegDate$ = Form$(1, 0)
  EndDate$ = Form$(2, 0)
  
  BegDate = Date2Num%(BegDate$)
  EndDate = Date2Num%(EndDate$)
  
  IF BegDate = -32768 OR EndDate = -32768 THEN
    Done = False
    Action = 1
    Frm(1).FldNo = 1
    ECode = 1
    GOSUB ShowDetailError
    GOTO DetailReturn
  ELSEIF EndDate < BegDate THEN
    Done = False
    Action = 1
    Frm(1).FldNo = 1
    ECode = 1
    GOSUB ShowDetailError
    GOTO DetailReturn
  END IF
  
  FromBook = VAL(Form$(3, 0))
  ThruBook = VAL(Form$(4, 0))
  IF FromBook > ThruBook THEN
    Done = False
    Action = 1
    Frm(1).FldNo = 3
    ECode = 2
    GOSUB ShowDetailError
    GOTO DetailReturn
  END IF
  
  TrType$ = QPTrim$(LEFT$(Form$(5, 0), 2))
  TrType = VAL(TrType$)
  
  IF TrType = 0 AND Frm(1).KeyCode = F10Key THEN
    Done = False
    Action = 1
    Frm(1).FldNo = 5
    ECode = 3
    GOSUB ShowDetailError
    GOTO DetailReturn
  ELSEIF TrType = 0 AND Frm(1).KeyCode = -103 THEN
    BegTrans = 1
    EndTrans = 999
  ELSE
    BegTrans = TrType
    EndTrans = TrType
  END IF
  
  Operator$ = Form$(6, 0)
  Operator = VAL(Operator$)
  IF Operator = 0 THEN
    BegOperator = 0
    EndOperator = 99
  ELSE
    BegOperator = Operator
    EndOperator = Operator
  END IF
  
  Detail$ = Form$(7, 0)
  
  CUSTTYPE$ = QPTrim$(Form$(8, 0))
  IF LEN(CUSTTYPE$) > 0 THEN
    UseType = True
  END IF
  
  SELECT CASE LEFT$(Form$(9, 0), 1)
  CASE "C"
    IndexName$ = NameIndexFile
    UsingName = True
    OKFlag = True
  CASE "A"
    IndexName$ = ""
    UsingAcct = True
    OKFlag = True
  CASE "L"
    IndexName$ = BookIndexFile
    UsingBook = True
    OKFlag = True
  CASE ELSE
    Done = False
    Action = 1
    Frm(1).FldNo = 9
    ECode = 4
    GOSUB ShowDetailError
    GOTO DetailReturn
  END SELECT
  Done = True
  
DetailReturn:
  RETURN
  
ShowDetailError:
  SaveScrn TempScrn()
  DisplayUBScrn "ERRSCRN1"
  SELECT CASE ECode
  CASE 1
    QPrintRC "Invalid Transaction Date!", 10, 28, -1
  CASE 2
    QPrintRC "Invalid Book Selection!", 10, 29, -1
  CASE 3
    QPrintRC "Invalid Transaction Type!", 10, 28, -1
  CASE 4
    QPrintRC "Invalid Printing Order!", 10, 29, -1
  CASE ELSE
  END SELECT
  QPrintRC "Please Correct and try again.", 12, 26, -1
  
  WaitForAction
  RestScrn TempScrn()
  
  RETURN
END SUB

SUB MasterBalanceListing
  
  SHARED Choice$()
  REDIM Choice$(0 TO 15, 0 TO 3)
  
  Choice$(0, 3) = "8"
  REDIM UBSetUpRec(1) AS UBSetupRecType
  UBSetUpLen = LEN(UBSetUpRec(1))
  FGetAH "UBSETUP.DAT", UBSetUpRec(1), UBSetUpLen, 1            'load it
  TownName$ = UBSetUpRec(1).UTILNAME
  IF INSTR(TownName$, "MOWAS") > 0 THEN
    MowasaFlag = True
  END IF
  
  RCnt = RCnt + 1
  Choice$(RCnt, 3) = "  ALL Revenues"
  
  REDIM Fmt$(1 TO 3)
  Fmt$(1) = "####,#.##"
  Fmt$(2) = "#####"
  Fmt$(3) = "######,#.##"
  
  FOR Cnt = 1 TO 15
    TempRev$ = QPTrim$(UBSetUpRec(1).Revenues(Cnt).REVNAME)
    IF LEN(TempRev$) = 0 THEN
      EXIT FOR
    ELSE
      RCnt = RCnt + 1
      Choice$(RCnt, 3) = QPTrim$(STR$(RCnt - 1)) + ") " + TempRev$
    END IF
  NEXT
  ERASE UBSetUpRec
  
  Choice$(0, 0) = "3"
  Choice$(1, 0) = "Past Due Balance"
  Choice$(2, 0) = "Current Balance"
  Choice$(3, 0) = "Total Balance"
  Choice$(4, 0) = "Credit Balance"
  
  Choice$(0, 1) = "5"
  Choice$(1, 1) = "  ALL"
  Choice$(2, 1) = "Active"
  Choice$(3, 1) = "Inactive"
  Choice$(4, 1) = "Balance"
  
  Choice$(0, 2) = "6"
  Choice$(1, 2) = "Customer Name Order"
  Choice$(2, 2) = "Account Number Order"
  Choice$(3, 2) = "Location Number Order"
  
  LibName$ = "UB"
  ScrnName$ = "UBMBLIST"
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  Fld(1).LowRange = 0
  Fld(2).LowRange = 0
  
  Action = 1
  FirstTime = True
  BlockClear
  
  DisplayUBScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      LSET Form$(1, 0) = "0"
      LSET Form$(2, 0) = "99"
      LSET Form$(4, 0) = "0"
      LSET Form$(7, 0) = "Y"
      IF MowasaFlag = False THEN
        Fld(8).Protected = True
        QPrintRC SPACE$(16), 18, 19, -1
      ELSE
        LSET Form$(8, 0) = Choice$(1, 3)
      END IF
      FirstTime = False
      Action = 1
    END IF
    
    '--Check for Key presses
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      GOSUB CheckBalInfo
      
    CASE EscKey
      ExitFlag = True
      EXIT SUB
    END SELECT
    
  LOOP UNTIL OKFlag OR ExitFlag
  IF ExitFlag THEN GOTO ExitBalRpt
  
  'Main Body Start
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  REDIM UBSetup(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetup(), UBSetUpRecLen
  
  TownName$ = QPTrim$(UBSetup(1).UTILNAME)
  TabStop = 40 - (LEN(TownName$) / 2)
  IF TabStop < 1 THEN TabStop = 1
  TownName$ = SPACE$(TabStop) + TownName$
  
  REDIM RevenueName(1 TO 15) AS STRING * 10
  FOR RCnt = 1 TO 15
    TRevName$ = QPTrim$(UBSetup(1).Revenues(RCnt).REVNAME)
    IF LEN(TRevName$) > 0 THEN
      AndPos = INSTR(TRevName$, "&")
      IF AndPos THEN
        MID$(TRevName$, AndPos) = " "
      END IF
      RevenueName(RCnt) = TRevName$
    ELSE
      MaxRevenue = RCnt - 1
      EXIT FOR
    END IF
  NEXT
  
  'MaxRevenue = 15
  
  REDIM RevTotals(1 TO MaxRevenue) AS DOUBLE
  
  IF UsingName OR UsingBook THEN
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize(IndexName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
    FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
    NumOfRecs = IdxNumOfRecs
  ELSE
    NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
  END IF
  
  MaxLines = 55
  Dash80$ = STRING$(80, "-")
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  UBRpt = FREEFILE
  OPEN "UBBALIST.RPT" FOR OUTPUT AS UBRpt
  
  'RecFile = FREEFILE
  'OPEN "RECLIST.IDX" FOR RANDOM AS RecFile LEN = 4
  
  BlockClear
  ShowProcessingScrn "Master Balance Listing."
  
  GOSUB DoCustRptHeader
  
  FOR Cnt = 1 TO NumOfRecs
    IF UsingName OR UsingBook THEN
      AcctNo& = IdxBuff(Cnt).RecNum
    ELSE
      AcctNo& = Cnt
    END IF
    
    GET UBCust, AcctNo&, UBCustRec(1)
    
    IF UBCustRec(1).DelFlag <> 0 THEN
      GOTO BSkipEm
    END IF
    
    RealBalance# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
    
    '110598 Old code could skew totals, A credit in one revenue and a debit
    '       in another revenue, that totaled to zero wouldn't show!!!
    'IF RealBalance# = 0 THEN
    '  GOTO BSkipEM
    'END IF
    
    '110598 Code to fix the above bug
    'If all there rev sources are "0" then skipem
    OKToSkip = True
    FOR RevChk = 1 TO 15        'MaxRevenue '15
      IF Round#(UBCustRec(1).CurrRevAmts(RevChk)) <> 0 THEN
        OKToSkip = False
        EXIT FOR
      END IF
    NEXT
    
    IF OKToSkip THEN
      GOTO BSkipEm
    END IF
    
    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, FF$
      GOSUB DoCustRptHeader
    END IF
    
    IF UseStatus THEN           'if they care about the cust status, or want all.
      CStatus$ = LEFT$(QPTrim$(UBCustRec(1).Status), 1)
      IF CStatus$ <> RStatus$ THEN
        GOTO BSkipEm
      END IF
    END IF
    
    BOOK$ = QPTrim$(UBCustRec(1).BOOK)
    SEQNUMB$ = QPTrim$(UBCustRec(1).SEQNUMB)
    IF LEN(BOOK$) = 0 THEN
      BOOK$ = "  "
    END IF
    BOOK = Value#(BOOK$, ECode)
    IF BOOK < BegRoute OR BOOK > EndRoute THEN
      GOTO BSkipEm
    END IF
    
    'IF (RealBalance# = 0) AND (NOT Ok2Skip) THEN STOP
    
    SELECT CASE BalType$
    CASE "PA"
      ChkBalance# = UBCustRec(1).PrevBalance
      IF (ChkBalance# >= MinBal#) AND (ChkBalance# > 0) THEN
        IF RealBalance# <= 0 THEN
          GOTO BSkipEm
        END IF
      ELSE
        GOTO BSkipEm
      END IF
    CASE "CU"
      ChkBalance# = UBCustRec(1).CurrBalance
      IF MinBal# > 0 THEN
        IF (ChkBalance# < MinBal#) OR (ChkBalance# <= 0) THEN
          GOTO BSkipEm
        END IF
      ELSE
        IF ChkBalance# = 0 THEN
          GOTO BSkipEm
        END IF
      END IF
    CASE "TO"
      ChkBalance# = RealBalance#
      IF MinBal# > 0 THEN
        IF ChkBalance# < MinBal# THEN
          GOTO BSkipEm
        END IF
      END IF
    CASE "CR"
      IF RealBalance# >= 0 THEN
        GOTO BSkipEm
      END IF
    END SELECT
    
    IF RevSource > 0 THEN
      IF UBCustRec(1).CurrRevAmts(RevSource) <> 0 THEN
        PRINT #UBRpt, USING Fmt$(2); AcctNo&;
        PRINT #UBRpt, TAB(8); BOOK$; "-"; SEQNUMB$;
        PRINT #UBRpt, TAB(18); LEFT$(UBCustRec(1).CUSTNAME, 31);
        PRINT #UBRpt, TAB(60); USING Fmt$(1); Round#(UBCustRec(1).CurrRevAmts(RevSource))
        LineCnt = LineCnt + 1
        TCurrBalance# = Round#(TCurrBalance# + UBCustRec(1).CurrRevAmts(RevSource))
        CustCnt = CustCnt + 1
      END IF
    ELSE
      PRINT #UBRpt, USING Fmt$(2); AcctNo&;
      PRINT #UBRpt, TAB(8); BOOK$; "-"; SEQNUMB$;
      PRINT #UBRpt, TAB(18); LEFT$(UBCustRec(1).CUSTNAME, 31);
      PRINT #UBRpt, TAB(50); USING Fmt$(1); Round#(UBCustRec(1).CurrBalance);
      PRINT #UBRpt, TAB(61); USING Fmt$(1); Round#(UBCustRec(1).PrevBalance);
      PRINT #UBRpt, TAB(72); USING Fmt$(1); RealBalance#
      LineCnt = LineCnt + 1
      TCurrBalance# = Round#(TCurrBalance# + UBCustRec(1).CurrBalance)
      TPrevBalance# = Round#(TPrevBalance# + UBCustRec(1).PrevBalance)
      CustCnt = CustCnt + 1
    END IF
    
    GOSUB PrintDetail
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
    
BSkipEm:
    ShowPctComp Cnt, NumOfRecs
  NEXT
  
  GOSUB DoCustRptFooter
  
  CLOSE
  ERASE IdxBuff, UBCustRec
  
  IF NOT AbortFlag THEN
    PrintRptFile "Balance Listing Report.", "UBBALIST.RPT", 1, RetCode, EntryPoint
  END IF
  
  'Main Body Exit
ExitBalRpt:
  
  EXIT SUB
  
DoCustRptHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, TownName$
  PRINT #UBRpt, TAB(26); "Customer Balance Listing Report"; TAB(70); "Page #"; PageNo
  PRINT #UBRpt, "Report Date: "; DATE$
  PRINT #UBRpt, "Acct #"; TAB(9); "Location"; TAB(18); "Customer Name";
  IF RevSource > 0 THEN
    PRINT #UBRpt, TAB(63); "Amount"
  ELSE
    PRINT #UBRpt, TAB(52); "Cur Bal"; TAB(62); "Past Due"; TAB(73); "Acct Bal"
  END IF
  PRINT #UBRpt, Dash80$
  LineCnt = 5
  RETURN
  
DoCustRptFooter:
  PRINT #UBRpt, Dash80$
  PRINT #UBRpt, "Totals:"; TAB(10); "Customers: "; USING "#####,#"; CustCnt;
  IF RevSource = 0 THEN
    PRINT #UBRpt, TAB(47); USING Fmt$(3); TCurrBalance#;
    PRINT #UBRpt, TAB(58); USING Fmt$(3); TPrevBalance#;
    PRINT #UBRpt, TAB(69); USING Fmt$(3); (TCurrBalance# + TPrevBalance#)
    'IF Detail THEN
    FOR Cnt = 1 TO MaxRevenue
      Detail$ = SPACE$(28)
      LSET Detail$ = RevenueName(Cnt)
      MID$(Detail$, 15) = Fmt$(3)
      GTotal# = Round#(GTotal# + RevTotals(Cnt))
      PRINT #UBRpt, USING Detail$; RevTotals(Cnt)
    NEXT
    LSET Detail$ = "Grand Total:"
    MID$(Detail$, 15) = Fmt$(3)
    PRINT #UBRpt, USING Detail$; GTotal#
    'ELSE
    '  PRINT #UBRpt,
    'END IF
  ELSE
    PRINT #UBRpt, TAB(58); USING Fmt$(3); TCurrBalance#
  END IF
  
  PRINT #UBRpt,
  PRINT #UBRpt, "Report Parameters"
  PRINT #UBRpt, "                Routes:"; STR$(BegRoute); " to"; STR$(EndRoute)
  IF CoFlag THEN
    PRINT #UBRpt, "       Minimum Balance: N/A       ";
  ELSE
    PRINT #UBRpt, USING "       Minimum Balance:$$######.##"; MinBal#;
  END IF
  PRINT #UBRpt, "            Customer Status:"; Stat$
  PRINT #UBRpt, "        Printing Order:"; POrder$;
  PRINT #UBRpt, "           Balance Type:"; Bal$;
  IF RevSource > 0 THEN
    PRINT #UBRpt, TAB(39); "Revenue Source: "; Choice$(RevSource + 1, 3);
  END IF
  PRINT #UBRpt,
  PRINT #UBRpt, Dash80$
  LineCnt = LineCnt + 4
  
  PRINT #UBRpt, FF$
  RETURN
  
PrintDetail:
  DLineCnt = LineCnt
  TCnt = 0
  Detail$ = SPACE$(18)
  First = 1
  Last = MaxRevenue
  IF MaxRevenue < Last THEN
    Last = MaxRevenue
  END IF
  FOR RCnt = First TO Last
    TCnt = TCnt + 1
    TabStop = (TCnt * 21) - 20
    IF TabStop > 81 THEN
      LineCnt = LineCnt + 1
      TCnt = 1
      TabStop = (TCnt * 21) - 20
    END IF
    LSET Detail$ = RevenueName(RCnt)
    MID$(Detail$, 10) = "#####.##"
    RevTotals(RCnt) = Round#(RevTotals(RCnt) + UBCustRec(1).CurrRevAmts(RCnt))
    'IF RCnt = 15 THEN
    '  IF UBCustRec(1).CurrRevAmts(RCnt) <> 0 THEN STOP
    'END IF
    IF Detail THEN
      PRINT #UBRpt, TAB(TabStop); USING Detail$; UBCustRec(1).CurrRevAmts(RCnt);
    END IF
  NEXT
  
  IF Detail THEN
    PRINT #UBRpt,
    PRINT #UBRpt, Dash80$
    LineCnt = LineCnt + 2
  ELSE
    LineCnt = DLineCnt
  END IF
  RETURN
  
  
CheckBalInfo:
  
  REDIM TempScrn(0)
  SaveScrn TempScrn()
  
  BegRoute = QPValI(Form$(1, 0))
  EndRoute = QPValI(Form$(2, 0))
  BalType$ = UCASE$(LEFT$(Form$(3, 0), 2))
  Order$ = LEFT$(QPTrim$(Form$(6, 0)), 1)
  MinBal# = Value#(Form$(4, 0), ECode)
  RStatus$ = LEFT$(Form$(5, 0), 1)
  Detail = LEFT$(Form$(7, 0), 1) = "Y"          'Evaluates to TRUE if the field = "Y"
  Rev$ = QPTrim$(LEFT$(Form$(8, 0), 1))
  
  RevSource = Value#(Rev$, ECode)
  
  IF RevSource > 0 THEN
    Detail = False
  END IF
  
  IF BegRoute > EndRoute THEN
    DisplayUBScrn "ERRSCRN1"
    QPrintRC "Invalid Route Order", 10, 30, -1
    QPrintRC "Please Correct and Try Again.", 12, 26, -1
    WaitForAction
    RestScrn TempScrn()
    Frm(1).FldNo = 1
    Action = 2
    GOTO ParmErrorRet
  END IF
  SELECT CASE BalType$
  CASE "PA"
    Bal$ = " PAST DUE"
  CASE "CU"
    Bal$ = " CURRENT"
  CASE "TO"
    Bal$ = " TOTAL BALANCE"
  CASE "CR"
    Bal$ = " CREDIT BALANCE"
    CoFlag = True
  CASE ELSE
    DisplayUBScrn "ERRSCRN1"
    QPrintRC "Invalid Balance Type", 10, 29, -1
    QPrintRC "Please Correct and Try Again.", 12, 26, -1
    WaitForAction
    RestScrn TempScrn()
    Frm(1).FldNo = 3
    Action = 2
    GOTO ParmErrorRet
  END SELECT
  
  SELECT CASE Order$
  CASE "C"
    IndexName$ = NameIndexFile
    UsingName = True
    POrder$ = " CUSTOMER NAME"
  CASE "A"
    POrder$ = " ACCOUNT NUMBER"
    IndexName$ = ""
    UsingAcct = True
  CASE "L"
    POrder$ = " LOCATION NUMBER"
    IndexName$ = BookIndexFile
    UsingBook = True
  CASE ELSE
    DisplayUBScrn "ERRSCRN1"
    QPrintRC "Invalid Printing Order", 10, 28, -1
    QPrintRC "Please Correct and Try Again.", 12, 26, -1
    WaitForAction
    RestScrn TempScrn()
    Frm(1).FldNo = 6
    Action = 2
    GOTO ParmErrorRet
  END SELECT
  SELECT CASE RStatus$
  CASE "A"
    UseStatus = True
    Stat$ = " ACTIVE"
  CASE "I"
    UseStatus = True
    Stat$ = " INACTIVE"
  CASE "B"
    UseStatus = True
    Stat$ = " BALANCE DUE"
  CASE ELSE
    Stat$ = " ALL"
    UseStatus = False
  END SELECT
  
  OKFlag = True
  
ParmErrorRet:
  
  RETURN
  
END SUB

SUB MasterCustomerListing
  
  REDIM MChoice$(1 TO 4)
  
  MChoice$(1) = " Customer Name Order"
  MChoice$(2) = " Account Number Order"
  MChoice$(3) = " Location Number Order"
  MChoice$(4) = " Service Address Order  "
  '^^ Don't change this!
  MaxLen = 0    'Set menu width to zero
  BoxBot = 17   'limit the box length to go no lower than line 20
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  '--Find max menu width
  FOR Cnt = 1 TO UBOUND(MChoice$)
    TLen = LEN(MChoice$(Cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT
  
  
  '--Center Menu within Screen
  Row = ((25 - (UBOUND(MChoice$))) \ 2) - 1
  Col = ((80 - MaxLen) \ 2) - 1
  
  UsingBook = False
  UsingAcct = False
  UsingName = False
  
  AbortFlag = False
  PageNo = 0
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    TitleBox 2, Col, MaxLen + 3, "Master Customer Report ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN
      Choice = 0
      ExitFlag = True
    END IF
    
    SELECT CASE Choice
    CASE 0
      ExitFlag = True
    CASE 1
      IndexName$ = NameIndexFile
      UsingName = True
      OKFlag = True
    CASE 2
      IndexName$ = ""
      UsingAcct = True
      OKFlag = True
    CASE 3
      IndexName$ = BookIndexFile
      UsingBook = True
      OKFlag = True
    CASE 4
      IndexName$ = TempIndexName
      UsingAddr = True
      OKFlag = True
    END SELECT
    
  LOOP UNTIL OKFlag OR ExitFlag
  
  IF ExitFlag THEN GOTO ExitMastCustListing
  
  '***************
  MaxLines = 52
  PageNo = 0
  Dash80$ = STRING$(80, "-")
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBSetup(1) AS UBSetupRecType
  UBSetUpRecLen = LEN(UBSetup(1))
  
  IF UsingName OR UsingBook THEN
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize(IndexName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
    FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
    NumOfRecs = IdxNumOfRecs
  ELSEIF UsingAddr THEN
    SortServiceAddrs "Service Address"
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize&(IndexName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
    FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
    NumOfRecs = IdxNumOfRecs
  ELSE
    NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
  END IF
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  '  UBLoca = FREEFILE
  '  OPEN "UBLOCA.DAT" FOR RANDOM SHARED AS UBLoca LEN = UBLocaRecLen
  
  UBRpt = FREEFILE
  OPEN "UBMCLIST.RPT" FOR OUTPUT AS UBRpt
  
  UBSetup = FREEFILE
  OPEN "UBSETUP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBSetup LEN = UBSetUpRecLen
  IF LOF(UBSetup) / UBSetUpRecLen = 0 THEN
    TownName$ = "Undefined"
  ELSE
    GET UBSetup, 1, UBSetup(1)
    TownName$ = QPTrim$(UBSetup(1).UTILNAME)
    TownLen = LEN(RTRIM$(TownName$))
    TabStop = 40 - (TownLen / 2)
    IF TabStop < 1 THEN TabStop = 1
    TownName$ = SPACE$(TabStop) + TownName$
  END IF
  CLOSE UBSetup
  
  
  BlockClear
  ShowProcessingScrn "Master Customer Report."
  
  GOSUB DoMastCustRptHeader
  
  FOR Cnt = 1 TO NumOfRecs
    IF UsingName OR UsingBook OR UsingAddr THEN
      GET UBCust, IdxBuff(Cnt).RecNum, UBCustRec(1)
      acctnumber = IdxBuff(Cnt).RecNum
    ELSE
      GET UBCust, Cnt, UBCustRec(1)
      acctnumber = Cnt
    END IF
    
    LocationNumber$ = UBCustRec(1).BOOK + "-" + UBCustRec(1).SEQNUMB
    
    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, FF$
      GOSUB DoMastCustRptHeader
    END IF
    
    IF UBCustRec(1).DelFlag <> -1 THEN
      PRINT #UBRpt, "Acct# :"; USING "#####"; acctnumber;
      PRINT #UBRpt, TAB(15); RTRIM$(UBCustRec(1).CUSTNAME);
      PRINT #UBRpt, TAB(48); "Location: "; LocationNumber$
      PRINT #UBRpt, "Status: "; UBCustRec(1).Status; TAB(15); QPTrim$(UBCustRec(1).ADDR1);
      PRINT #UBRpt, TAB(48); "H Phone#: "; QPTrim$(UBCustRec(1).HPHONE)
      PRINT #UBRpt, "Type  : "; QPTrim$(UBCustRec(1).CUSTTYPE); TAB(15); QPTrim$(UBCustRec(1).ADDR2);
      PRINT #UBRpt, TAB(48); "W Phone#: "; UBCustRec(1).WPHONE
      PRINT #UBRpt, "Draft : "; QPTrim$(UBCustRec(1).USEDRAFT);
      Zip$ = QPTrim$(UBCustRec(1).ZIPCODE)
      IF LEN(Zip$) > 5 THEN
        Zip$ = LEFT$(Zip$, 5) + "-" + MID$(Zip$, 6)
      END IF
      PRINT #UBRpt, TAB(15); QPTrim$(UBCustRec(1).CITY); ", "; QPTrim$(UBCustRec(1).STATE); " "; Zip$;
      PRINT #UBRpt, TAB(48); "Ser Addr: "; QPTrim$(UBCustRec(1).SERVADDR)
      PRINT #UBRpt, Dash80$
      LineCnt = LineCnt + 5
      CustCnt = CustCnt + 1
      IF AskAbandonPrint% THEN
        AbortFlag = True
        EXIT FOR
      END IF
    END IF
    ShowPctComp Cnt, NumOfRecs
  NEXT
  
  GOSUB DoMastCustRptFooter:
  
  CLOSE
  
  ERASE IdxBuff, UBCustRec
  
  IF NOT AbortFlag THEN
    PrintRptFile "Master Customer Report.", "UBMCLIST.RPT", 1, RetCode, EntryPoint
  END IF
  
ExitMastCustListing:
  
  EXIT SUB
  
DoMastCustRptHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, TownName$
  PRINT #UBRpt, TAB(24); "Master Customer Listing Report"; TAB(70); "Page #"; PageNo
  PRINT #UBRpt, "Report Date: "; DATE$
  PRINT #UBRpt, Dash80$
  LineCnt = 4
  RETURN
  
DoMastCustRptFooter:
  PRINT #UBRpt, "Totals:"; TAB(10); "Customers: "; USING "#####,#"; CustCnt
  RETURN
  
  
END SUB

SUB MasterDepositListing
  
  REDIM MChoice$(1 TO 3)
  
  MChoice$(1) = " Customer Name Order    "
  MChoice$(2) = " Location Number Order  "
  MChoice$(3) = " Account Number Order   "
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 17   'limit the box length to go no lower than line 20
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  '--Find max menu width
  FOR Cnt = 1 TO UBOUND(MChoice$)
    TLen = LEN(MChoice$(Cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT
  
  'ReStart:
  '--Center Menu within Screen
  Row = ((25 - (UBOUND(MChoice$))) \ 2) - 1
  Col = ((80 - MaxLen) \ 2) - 1
  
  UsingBook = False
  UsingAcct = False
  UsingName = False
  
  AbortFlag = False
  PageNo = 0
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    TitleBox 2, Col, MaxLen + 3, "Deposit Listing Report ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN
      Choice = 0
      ExitFlag = True
    END IF
    
    SELECT CASE Choice
    CASE 0
      ExitFlag = True
    CASE 1
      IndexName$ = NameIndexFile
      UsingName = True
      OKFlag = True
      
    CASE 2
      IndexName$ = BookIndexFile
      UsingBook = True
      OKFlag = True
      
    CASE 3
      IndexName$ = ""
      UsingAcct = True
      OKFlag = True
    END SELECT
    
  LOOP UNTIL OKFlag OR ExitFlag
  
  IF ExitFlag THEN GOTO ExitDepositListing
  
  '***************
  MaxLines = 52
  PageNo = 0
  Dash80$ = STRING$(80, "-")
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBSetup(1) AS UBSetupRecType
  UBSetUpRecLen = LEN(UBSetup(1))
  
  IF UsingName OR UsingBook THEN
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize(IndexName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
    FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
    NumOfRecs = IdxNumOfRecs
  ELSE
    NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
  END IF
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  UBRpt = FREEFILE
  OPEN "UBDPLIST.RPT" FOR OUTPUT AS UBRpt
  
  UBSetup = FREEFILE
  OPEN "UBSETUP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBSetup LEN = UBSetUpRecLen
  IF LOF(UBSetup) / UBSetUpRecLen = 0 THEN
    TownName$ = "Undefined"
  ELSE
    GET UBSetup, 1, UBSetup(1)
    TownName$ = QPTrim$(UBSetup(1).UTILNAME)
    TownLen = LEN(RTRIM$(TownName$))
    TabStop = 40 - (TownLen / 2)
    IF TabStop < 1 THEN TabStop = 1
    TownName$ = SPACE$(TabStop) + TownName$
  END IF
  CLOSE UBSetup
  
  
  BlockClear
  ShowProcessingScrn "Deposit Listing Report."
  
  GOSUB DoDepositRptHeader
  
  FOR Cnt = 1 TO NumOfRecs
    IF UsingName OR UsingBook THEN
      acctnumber = IdxBuff(Cnt).RecNum
    ELSE
      acctnumber = Cnt
    END IF
    
    GET UBCust, acctnumber, UBCustRec(1)
    IF UBCustRec(1).DelFlag = 0 THEN
      IF Round#(UBCustRec(1).DepositAmt) <> 0 THEN
        PRINT #UBRpt, USING "#####"; acctnumber;
        PRINT #UBRpt, TAB(8); QPTrim$(UBCustRec(1).BOOK); "-"; QPTrim$(UBCustRec(1).SEQNUMB);
        PRINT #UBRpt, TAB(20); LEFT$(UBCustRec(1).CUSTNAME, 38);
        PRINT #UBRpt, TAB(61); USING "####,#.##"; UBCustRec(1).DepositAmt;
        PRINT #UBRpt, TAB(78); UBCustRec(1).Status
        LineCnt = LineCnt + 1
        TDeposit# = Round#(TDeposit# + UBCustRec(1).DepositAmt)
        CustCnt = CustCnt + 1
      END IF
    END IF
    
    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, FF$
      GOSUB DoDepositRptHeader
    END IF
    
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
    
    ShowPctComp Cnt, NumOfRecs
  NEXT
  
  GOSUB DoDepositRptFooter:
  
  CLOSE UBCust, UBRpt
  
  ERASE IdxBuff, UBCustRec
  
  IF NOT AbortFlag THEN
    PrintRptFile "Deposit Listing Report.", "UBDPLIST.RPT", 1, RetCode, EntryPoint
  END IF
  
  'KillFile "UBDPLIST.RPT"
  
ExitDepositListing:
  
  EXIT SUB
  
DoDepositRptHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, TownName$
  PRINT #UBRpt, TAB(24); "Customer Deposit Listing Report"; TAB(70); "Page #"; PageNo
  PRINT #UBRpt, "Report Date: "; DATE$
  PRINT #UBRpt, "Acct #";
  PRINT #UBRpt, TAB(9); "Location"; TAB(20); "Customer Name"; TAB(60); "Deposit Amt"; TAB(74); "Status"
  PRINT #UBRpt, Dash80$
  LineCnt = 4
  
  RETURN
  
DoDepositRptFooter:
  PRINT #UBRpt, Dash80$
  PRINT #UBRpt, "Totals:"; TAB(10); "Customers: "; USING "#####,#"; CustCnt;
  PRINT #UBRpt, TAB(60); USING "#####,#.##"; TDeposit#
  
  PRINT #UBRpt, FF$
  
  RETURN
  
END SUB

