DEFINT A-Z
DECLARE SUB CycleCount ()
DECLARE SUB MeterInstallDate ()
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
DECLARE SUB MasterBalanceListing (AutoFlag)
DECLARE SUB MasterCustomerListing ()
DECLARE SUB MasterDepositListing ()
DECLARE SUB PaymentSumReport ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LptPort%, RetCode%, EntryPoint%)
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


TYPE MtrDateSortType
  MtrDate AS INTEGER
  RecNum   AS INTEGER
END TYPE

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'ubrate.BI'
  '$INCLUDE: 'newcust.bi'
AgedBalanceListing

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

    CustBook = Value#(UBCustRec(1).Book, ECode)
    IF CustBook < FromBook OR CustBook > ThruBook THEN
      GOTO SkipEm
    END IF
    
    IF Ctype = 3 THEN
      IF NOT (UBCustRec(1).Status = "I" OR UBCustRec(1).Status = "B") THEN GOTO SkipEm
    END IF
    IF Ctype = 2 THEN
      IF NOT (UBCustRec(1).Status = "A" OR UBCustRec(1).Status = "F") THEN GOTO SkipEm
    END IF
    
    IF LEFT$(UBCustRec(1).CustName, 6) = "VACANT" OR LEN(QPTrim$(UBCustRec(1).CustName)) = 0 THEN GOTO SkipEm
    
    IF UBCustRec(1).DelFlag THEN
      GOTO SkipEm
    END IF
    
    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, FF$
      GOSUB DoAgedBalRptHeader
    END IF
    
    '*************************************
    '   Main body of Printing goes here
    
    AcctBal# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
    Trans& = UBCustRec(1).LastTrans
    Process = 0
    IF Trans& > 0 THEN
      GET UBTrans, Trans&, UBTransRec(1)
      TestDate = AgeDate - 90
      IF UBTransRec(1).TransDate < TestDate THEN
        GOTO SkipEm
      END IF
    END IF

    'IF AcctBal# = 0 THEN
    '  GOTO SkipEm
    'END IF

    GOSUB CalcBal
    LineTotal# = Round#(LineTotal#)
    
    IF AcctBal# <> LineTotal# THEN
      Dif# = AcctBal# - LineTotal#
    END IF
    
    IF LineTotal# <> 0 THEN
      'Clean Up Customer Name
      CustName$ = UBCustRec(1).CustName
      IF ASC(LEFT$(CustName$, 1)) < 64 OR ASC(LEFT$(CustName$, 1)) > 90 THEN CustName$ = ""
      CustName$ = QPTrim$(LEFT$(CustName$, 28))
      'Clean Up Seq Number
      SeqNumber$ = UBCustRec(1).SEQNUMB
      IF ASC(LEFT$(SeqNumber$, 1)) < 45 OR ASC(LEFT$(SeqNumber$, 1)) > 57 THEN SeqNumber$ = ""
      SeqNumber$ = QPTrim$(SeqNumber$)

      PRINT #UBRpt, UBCustRec(1).Book; "-"; SeqNumber$;
      PRINT #UBRpt, TAB(12); CustName$;
      PRINT #UBRpt, TAB(41); USING "######.##"; Amt1#;
      PRINT #UBRpt, TAB(51); USING "######.##"; Amt2#;
      PRINT #UBRpt, TAB(61); USING "######.##"; Amt3#;
      PRINT #UBRpt, TAB(71); USING "######.##"; Amt4#;
      PRINT #UBRpt, TAB(81); USING "######.##"; LineTotal#;
      PRINT #UBRpt, TAB(92); USING "#####"; acctnumber
      
      LineCnt = LineCnt + 1
      CustCnt = CustCnt + 1
      TAmt1# = TAmt1# + Amt1#
      TAmt2# = TAmt2# + Amt2#
      TAmt3# = TAmt3# + Amt3#
      TAmt4# = TAmt4# + Amt4#
      TLineTotal# = TLineTotal# + LineTotal#
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
  
  KillFile "UBABLIST.RPT"
  
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
  FOR Book = 0 TO 99
    IF BookTotal#(Book) <> 0 THEN
      PRINT #UBRpt, USING "##"; Book;
      PRINT #UBRpt, TAB(10); USING "$$######,#.##"; BookTotal#(Book)
    END IF
  NEXT Book
  PRINT #UBRpt, FF$;
  PRINT #UBRpt, "Totals by Revenue"
  PRINT #UBRpt, ""
  PRINT #UBRpt, "Revenue"; TAB(30); "Amount"
  PRINT #UBRpt, STRING$(50, "-")
  FOR Rev = 1 TO 15
    IF Revenue#(Rev) <> 0 THEN
      PRINT #UBRpt, RevenueName$(Rev); TAB(30); USING "$$######,#.##"; Revenue#(Rev)
    END IF
  NEXT Rev
  PRINT #UBRpt, FF$;
  PRINT #UBRpt, "Totals by Book and Revenue"
  PRINT #UBRpt, ""
  PRINT #UBRpt, TAB(10); RevenueName$(1); TAB(30); RevenueName$(2); TAB(50); RevenueName$(3); TAB(70); RevenueName$(4)
  PRINT #UBRpt, "Book"; TAB(10); RevenueName$(5); TAB(30); RevenueName$(6); TAB(50); RevenueName$(7); TAB(70); RevenueName$(8)
  PRINT #UBRpt, Dash90$
  FOR Book = 0 TO 99
    IF BookTotal#(Book) <> 0 THEN
      PRINT #UBRpt, USING "##"; Book;
      PRINT #UBRpt, TAB(10); USING "$$######,#.##"; BookRev#(Book, 1);
      PRINT #UBRpt, TAB(30); USING "$$######,#.##"; BookRev#(Book, 2);
      PRINT #UBRpt, TAB(50); USING "$$######,#.##"; BookRev#(Book, 3);
      PRINT #UBRpt, TAB(70); USING "$$######,#.##"; BookRev#(Book, 4)
      PRINT #UBRpt, TAB(10); USING "$$######,#.##"; BookRev#(Book, 5);
      PRINT #UBRpt, TAB(30); USING "$$######,#.##"; BookRev#(Book, 6);
      PRINT #UBRpt, TAB(50); USING "$$######,#.##"; BookRev#(Book, 7);
      PRINT #UBRpt, TAB(70); USING "$$######,#.##"; BookRev#(Book, 8)
      PRINT #UBRpt, TAB(10); USING "$$######,#.##"; BookRev#(Book, 9);
      PRINT #UBRpt, TAB(30); USING "$$######,#.##"; BookRev#(Book, 10);
      PRINT #UBRpt, TAB(50); USING "$$######,#.##"; BookRev#(Book, 11);
      PRINT #UBRpt, TAB(70); USING "$$######,#.##"; BookRev#(Book, 12)
      PRINT #UBRpt, Dash90$
    END IF
  NEXT Book
  RETURN
  
CalcBal:
  Book = VAL(UBCustRec(1).Book)
  
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
    'IF FirstTrans THEN
    '  FirstTrans = False
    '  TestDate = AgeDate - 90
    '  IF UBTransRec(1).TransDate < TestDate THEN
    '    EXIT DO
    '  END IF
    'END IF
    IF NOT (UBTransRec(1).TransType = 107 OR UBTransRec(1).TransType = 7 OR UBTransRec(1).TransType = 9) THEN
      IF UBTransRec(1).TransDate <= AgeDate THEN
        'IF UBTransRec(1).RunBalance = 0 AND Process >= 2 THEN GOTO EndReview'Get Out if Balance =0
        GOSUB LineCalc
        Process = Process + 1
      END IF
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
    Amt3# = Amt3# - TotalPayments#
    TotalPayments# = 0
  END IF
  IF TotalPayments# > Amt2# THEN
    TotalPayments# = Round#(TotalPayments# - Amt2#)
    Amt2# = 0
  ELSE
    Amt2# = Amt2# - TotalPayments#: TotalPayments# = 0
  END IF
  Amt1# = Amt1# - TotalPayments#
  
  LineTotal# = Round#(Amt1# + Amt2# + Amt3# + Amt4#)
  
Gotit:
  
  LineTotal# = Round#(LineTotal#)
  
  BookTotal#(Book) = BookTotal#(Book) + LineTotal#
  RETURN
  
  
LineCalc:
  TransAmt# = UBTransRec(1).TransAmt
  Revenue# = 0
  FOR Rev = 1 TO 15
    Revenue# = Round#(Revenue# + UBTransRec(1).RevAmt(Rev) + UBTransRec(1).TaxAmt(Rev))
  NEXT Rev
  TransAmt# = INT((TransAmt# * 100) + .5) / 100
  
  'Revenue# = INT((Revenue# * 100) + .5) / 100
  
  IF UBTransRec(1).TransType = 4 OR UBTransRec(1).TransType = 5 OR UBTransRec(1).TransType = 8 OR UBTransRec(1).TransType = 12 OR UBTransRec(1).TransType = 104 OR UBTransRec(1).TransType = 105 OR UBTransRec(1).TransType = 109 THEN
    IF UBTransRec(1).TransType = 5 THEN
      TotalPayments# = TotalPayments# + ABS(UBTransRec(1).TransAmt)
    ELSE
      TotalPayments# = TotalPayments# + UBTransRec(1).TransAmt
    END IF
    TotalPayments# = Round#(TotalPayments#)
    
    FOR Rev = 1 TO 15
      Revenue#(Rev) = Round#(Revenue#(Rev) - UBTransRec(1).RevAmt(Rev) - UBTransRec(1).TaxAmt(Rev))
      BookRev#(Book, Rev) = Round#(BookRev#(Book, Rev) - UBTransRec(1).RevAmt(Rev) - UBTransRec(1).TaxAmt(Rev))
    NEXT Rev
    
    IF Round#(Revenue#) <> Round#(TransAmt#) THEN
      Dif# = Round#(TransAmt#) - Round#(Revenue#)
      Revenue#(4) = Revenue#(4) - Dif#
      TransAmt# = TransAmt# - Dif#
    END IF
    
    
    RETURN
  END IF
  
  SELECT CASE UBTransRec(1).TransType
  CASE 1, 2, 3, 6, 11, 10, 101, 102, 103, 112, 110
    
    Days& = AgeDate - UBTransRec(1).TransDate
    IF Days& <= 30 THEN
      Amt1# = Amt1# + UBTransRec(1).TransAmt
      Amt1# = Round#(Amt1#)
      FOR Rev = 1 TO 15
        Revenue#(Rev) = Round#(Revenue#(Rev) + (UBTransRec(1).RevAmt(Rev) + UBTransRec(1).TaxAmt(Rev)))
        BookRev#(Book, Rev) = Round#(BookRev#(Book, Rev) + UBTransRec(1).RevAmt(Rev) + UBTransRec(1).TaxAmt(Rev))
      NEXT Rev
      RETURN
    END IF
    IF Days& > 30 AND Days& <= 60 THEN
      Amt2# = Amt2# + UBTransRec(1).TransAmt
      Amt2# = Round#(Amt2#)
      FOR Rev = 1 TO 15: Revenue#(Rev) = Revenue#(Rev) + (UBTransRec(1).RevAmt(Rev) + UBTransRec(1).TaxAmt(Rev))
        BookRev#(Book, Rev) = Round#(BookRev#(Book, Rev) + UBTransRec(1).RevAmt(Rev) + UBTransRec(1).TaxAmt(Rev))
      NEXT Rev
      RETURN
    END IF
    IF Days& > 60 AND Days& <= 90 THEN
      Amt3# = Amt3# + UBTransRec(1).TransAmt
      Amt3# = Round#(Amt3#)
      FOR Rev = 1 TO 15
        Revenue#(Rev) = Round#(Revenue#(Rev) + UBTransRec(1).RevAmt(Rev) + UBTransRec(1).TaxAmt(Rev))
        BookRev#(Book, Rev) = Round#(BookRev#(Book, Rev) + UBTransRec(1).RevAmt(Rev) + UBTransRec(1).TaxAmt(Rev))
      NEXT Rev
      RETURN
    END IF
    IF Days& > 90 THEN
      Amt4# = Amt4# + UBTransRec(1).TransAmt
      Amt4# = Round#(Amt4#)
      FOR Rev = 1 TO 15
        Revenue#(Rev) = Round#(Revenue#(Rev) + UBTransRec(1).RevAmt(Rev) + UBTransRec(1).TaxAmt(Rev))
        BookRev#(Book, Rev) = Round#(BookRev#(Book, Rev) + UBTransRec(1).RevAmt(Rev) + UBTransRec(1).TaxAmt(Rev))
      NEXT Rev
      RETURN
    END IF
  END SELECT
  
  RETURN
  
END SUB

