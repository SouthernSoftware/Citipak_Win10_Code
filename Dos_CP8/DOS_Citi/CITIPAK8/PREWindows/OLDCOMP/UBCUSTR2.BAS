DECLARE SUB SortMtrDateIndex (FromDate%, ThruDate%)
DEFINT A-Z
DECLARE SUB SortServiceAddrs (IndexText$)
DECLARE SUB ShowWrkOrdHistory (RecNo&)
DECLARE SUB ShowCustConsHist (CustRec&)
DECLARE SUB ShowCustHistory (CustRec&)
DECLARE SUB FCreate (FileName$)
DECLARE SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Round# (N#)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB BlockClear ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetUpLen%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
DECLARE SUB RestScrn (array())
DECLARE SUB SaveScrn (array())
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE SUB WaitForAction ()
DECLARE SUB HideCursor ()
DECLARE SUB ShowCursor ()
DECLARE SUB CursorOff ()
DECLARE FUNCTION FileSize& (FileName$)
DECLARE SUB KillFile (FileName$)
DECLARE SUB LookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, LocationFlag%)
DECLARE SUB CustMessageSystem (RecNo&)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE FUNCTION GetNumOfCust% ()

  CONST False = 0, True = NOT False

  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'ubtrans.BI'

   FF$ = CHR$(12)

TYPE TaxRpt2Type
  ERevAmt    AS DOUBLE
  Rev1Amt    AS DOUBLE
  Rev1Tax    AS DOUBLE
  Rev2Amt    AS DOUBLE
  Rev2Tax    AS DOUBLE
END TYPE

TYPE MtrDateSortType
  MtrDate AS INTEGER
  RecNum   AS INTEGER
END TYPE

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

    CustBook = Value#(UBCustRec(1).BOOK, ECode)
    IF CustBook < FromBook OR CustBook > ThruBook THEN
      GOTO AgedSkipEm
    END IF
    
    IF Ctype = 3 THEN
      IF NOT (UBCustRec(1).Status = "I" OR UBCustRec(1).Status = "B") THEN GOTO AgedSkipEm
    END IF
    IF Ctype = 2 THEN
      IF NOT (UBCustRec(1).Status = "A" OR UBCustRec(1).Status = "F") THEN GOTO AgedSkipEm
    END IF
    
    'IF LEFT$(UBCustRec(1).CUSTNAME, 6) = "VACANT" OR LEN(QPTrim$(UBCustRec(1).CUSTNAME)) = 0 THEN GOTO AgedSkipEm
    
    IF UBCustRec(1).DelFlag <> 0 THEN
      GOTO AgedSkipEm
    END IF
    
    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, FF$
      GOSUB DoAgedBalRptHeader
    END IF
    
    '*************************************
    '   Main body of Printing goes here
    
    AcctBal# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
    FOR Rev = 1 TO 15
      CurRevAmt# = UBCustRec(1).CurrRevAmts(Rev)
      PreRevAmt# = UBCustRec(1).PrevRevAmts(Rev)
      TotRevAmt# = Round#(CurRevAmt# + PreRevAmt#)
      Revenue#(Rev) = Round#(Revenue#(Rev) + TotRevAmt#)
      BookRev#(BOOK, Rev) = Round#(BookRev#(BOOK, Rev) + TotRevAmt#)
    NEXT Rev

    'IF AcctBal# = 0 THEN
    '  GOTO AgedSkipEm
    'END IF


    'Trans& = UBCustRec(1).LastTrans
    Process = 0
    'IF Trans& > 0 THEN
    '  GET UBTrans, Trans&, UBTransRec(1)
    '  TestDate = AgeDate - 90
    '  IF UBTransRec(1).TransDate < TestDate THEN
    '    GOTO AgedSkipEm
    '  END IF
    'END IF
'dale

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
AgedSkipEm:
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
  FOR Rev = 1 TO 15
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
  CASE 4, 5, 8, 12, 104, 105, 109
    IF UBTransRec(1).TransType = 5 THEN
      TotalPayments# = TotalPayments# + ABS(UBTransRec(1).TransAmt)
    ELSE
      TotalPayments# = TotalPayments# + UBTransRec(1).TransAmt
    END IF
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

SUB CustStreetListing

  SHARED Choice$()
  REDIM Choice$(0 TO 3, 0 TO 1)
  
  Choice$(0, 1) = "3"
  Choice$(1, 1) = "Customer Name Order   "
  Choice$(2, 1) = "Account Number Order  "
  Choice$(3, 1) = "Location Number Order "

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

  IF INSTR(TownName$, "GROVER") > 0 THEN
    GroverFlag = True
  END IF

  TabStop = 40 - (TownLen / 2)
  IF TabStop < 1 THEN TabStop = 1
  
  IF UBSetup(1).BILLCYCL = "Y" THEN
    UseCycle = True
  END IF

  LibName$ = "UB"
  ScrnName$ = "UBSTRADD"

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
      GOSUB ShowBookorCycle
      LSET Form$(1, 0) = "0"
      LSET Form$(2, 0) = "99"
      LSET Form$(6, 0) = "Y"
    END IF

    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      First = Value#(Form$(1, 0), ECode)
      Last = Value#(Form$(2, 0), ECode)
      IF Last < First THEN
        GOSUB ShowErrScrn
        GOTO LoopJump
      END IF
      Order$ = LEFT$(QPTrim$(Form$(3, 0)), 1)
      Ctype$ = QPTrim$(Form$(4, 0))
      IF LEN(Ctype$) > 0 THEN
        ChkCType = True
      END IF
      StrSrch$ = QPTrim$(Form$(5, 0))
      InactiveFlag = LEFT$(Form$(6, 0), 1) = "Y"

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
  ELSE
    NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
  END IF

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  UBTrans = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTrans LEN = UBTransRecLen
  UBRpt = FREEFILE
  OPEN "UBSTLIST.RPT" FOR OUTPUT AS UBRpt

  IF INSTR(TownName$, "LONGVIEW") THEN
    LVFlag = True
  END IF

  BlockClear
  ShowProcessingScrn "Customer Street Listing."

  GOSUB DoCutOffRptHeader

  ThisBook = 0
  FirstBook = True

  FOR Cnt = 1 TO NumOfRecs
    IF UsingName OR UsingBook THEN
      AcctNo& = IdxBuff(Cnt).RecNum
    ELSE
      AcctNo& = Cnt
    END IF
    GET UBCust, AcctNo&, UBCustRec(1)
    IF UBCustRec(1).DelFlag = 0 THEN
      IF UseCycle THEN
        CustCheck = UBCustRec(1).BILLCYCL
      ELSE
        CustCheck = Value#(UBCustRec(1).BOOK, ECode)
      END IF

      IF CustCheck >= First AND CustCheck <= Last THEN
        IF NOT InactiveFlag THEN
          IF UBCustRec(1).Status <> "A" THEN
            GOTO SkipEMStreet
          END IF
        END IF
        IF ChkCType THEN
          CustType$ = QPTrim$(UBCustRec(1).CustType)
          IF CustType$ <> Ctype$ THEN
            GOTO SkipEMStreet
          END IF
        END IF
        IF INSTR(UBCustRec(1).ServAddr, StrSrch$) > 0 THEN
          IF FirstBook THEN
            FirstBook = False
            ThisBook = Value#(UBCustRec(1).BOOK, ECode)
          END IF
          GOSUB printline
        END IF
      END IF
SkipEMStreet:
      IF AskAbandonPrint% THEN
        AbortFlag = True
        EXIT FOR
      END IF
'    ELSE
'      Deleted = Deleted + 1
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
    PrintRptFile "Customer Street Listing.", "UBSTLIST.RPT", 1, RetCode, EntryPoint
  END IF

  KillFile "UBSTLIST.RPT"

ExitCutOffListing:

  EXIT SUB

DoCutOffRptHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, TAB(TabStop); TownName$
  PRINT #UBRpt, TAB(26); "Customer Street Listing Report"; TAB(70); "Page #"; PageNo
  PRINT #UBRpt, "Report Date: "; DATE$
  PRINT #UBRpt, "Location  Acct#  Customer Name"; TAB(53); "Service Address"
  PRINT #UBRpt, Dash80$
  LineCnt = 6
RETURN

DoCutOffRptFooter:
  PRINT #UBRpt, ""
  PRINT #UBRpt, "Total Customers: "; USING "#####,#"; CustomerCnt
  PRINT #UBRpt,
  PRINT #UBRpt, "         Report Parameters"
  PRINT #UBRpt, "                      From:"; First; "  Thru:"; Last
  PRINT #UBRpt, "               Search Text: ";
  IF LEN(StrSrch$) = 0 THEN
    PRINT #UBRpt, "NONE."
  ELSE
    PRINT #UBRpt, StrSrch$
  END IF
  PRINT #UBRpt, "             Customer Type: ";
  IF ChkCType THEN
    PRINT #UBRpt, Ctype$
  ELSE
    PRINT #UBRpt, "ALL"
  END IF
  PRINT #UBRpt, "          Include Inactive: ";
  IF InactiveFlag THEN
    PRINT #UBRpt, "Y"
  ELSE
    PRINT #UBRpt, "N"
  END IF
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

  IF GroverFlag THEN
    IF UsingBook THEN
      CustBook = Value#(UBCustRec(1).BOOK, ECode)
      IF ThisBook <> CustBook THEN
        ThisBook = CustBook
        PRINT #UBRpt, FF$
        GOSUB DoCutOffRptHeader
      END IF
    END IF
  END IF

  '*************************************
  '   Main body of Printing goes here
  PRINT #UBRpt, UBCustRec(1).BOOK; "-"; UBCustRec(1).SEQNUMB; USING "######"; AcctNo&;
  PRINT #UBRpt, TAB(18); LEFT$(UBCustRec(1).CUSTNAME, 30); TAB(52); UBCustRec(1).ServAddr
  LineCnt = LineCnt + 1
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
    QPrintRC "From Cycle", 6, 27, -1
    QPrintRC "Thru Cycle", 8, 27, -1
  ELSE
    QPrintRC "From Book", 6, 28, -1
    QPrintRC "Thru Book", 8, 28, -1
  END IF

RETURN

END SUB

SUB CustTaxReport
  
  SHARED Choice$()
  
  TheDate$ = DATE$
  Dash80$ = STRING$(80, "-")
  MaxLines = 60
  
  REDIM Choice$(0 TO 3, 0 TO 1)
  REDIM RevText$(1 TO MaxRevsCnt)
  REDIM TaxRates(1 TO 15) AS SINGLE
  REDIM RevTotals(1 TO 15) AS DOUBLE
  REDIM TaxAmt(1 TO 15) AS DOUBLE
  REDIM TempScrn(0)
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBTrans(1) AS UBTransRecType
  UBTransRecLen = LEN(UBTrans(1))
  
  REDIM UBSetup(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetup(), UBSetUpRecLen
  TownName$ = UBSetup(1).UTILNAME
  TownLen = LEN(RTRIM$(TownName$))
  TabStop = 40 - (TownLen / 2)
  IF TabStop < 1 THEN TabStop = 1
  
  FOR Cnt = 1 TO MaxRevsCnt
    TempRev$ = QPTrim$(UBSetup(1).Revenues(Cnt).REVNAME)
    IF LEN(TempRev$) = 0 THEN
      LastRev = Cnt - 1
      EXIT FOR
    ELSE
      RevText$(Cnt) = TempRev$
      TaxRates(Cnt) = UBSetup(1).Revenues(Cnt).TaxRate
      IF TaxRates(Cnt) > 0 THEN
        GotTaxFlag = True
      END IF
    END IF
  NEXT

  IF NOT GotTaxFlag THEN
    CursorOff
    HideCursor
    BlockClear
    DisplayUBScrn "ERRSCRN1"
    QPrintRC "You do not have any taxes to report on.", 10, 21, -1
    QPrintRC "This report is ONLY for taxed revenues.", 11, 21, -1
    QPrintRC "Press any key to continue.", 13, 27, -1
    ShowCursor
    WaitForAction
    GOTO ExitTaxReport
  END IF
  
  LibName$ = "UB"
  ScrnName$ = "UBTAXRPT"
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  Frm(1).StayOnField = True
  
  'Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  'Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  Choice$(0, 0) = "3"
  Choice$(1, 0) = "Payments"
  Choice$(2, 0) = "Billings"
  
  ' Set Defaults as Anticpated Response
  FirstTime = True
  
  Action = 1
  BlockClear
  DisplayUBScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      FirstTime = False
      LSET Form$(1, 0) = TheDate$
      LSET Form$(2, 0) = TheDate$
      LSET Form$(4, 0) = "Y"
      Action = 1
    END IF
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      GOSUB CheckInfo
      IF ParmsOK THEN
        ExitFlag = True
      END IF
    CASE ESC
      ParmsOK = False
      ExitFlag = True
    END SELECT
  LOOP UNTIL ExitFlag
  
  IF NOT ParmsOK THEN
    GOTO ExitTaxReport
  END IF
  
  BlockClear
  ShowProcessingScrn "Transaction Tax Report."
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  UBTrans = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTrans LEN = UBTransRecLen
  
  UBRpt = FREEFILE
  OPEN "UBTAXRPT.RPT" FOR OUTPUT AS UBRpt
  
  GOSUB DoTaxRptHeader
  
  NumOfRecs& = LOF(UBCust) / UBCustRecLen
  FOR Cnt& = 1 TO NumOfRecs&
    GET UBCust, Cnt&, UBCustRec(1)
    'IF Cnt& = 2016 THEN STOP

    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, FF$
      GOSUB DoTaxRptHeader
    END IF
    Trans& = UBCustRec(1).LastTrans
    DO WHILE Trans& <> 0
      GET UBTrans, Trans&, UBTrans(1)
      IF UBTrans(1).TransType = TransType THEN
        IF (UBTrans(1).TransDate >= BegDate AND UBTrans(1).TransDate <= EndDate) THEN
          IF UBTrans(1).TaxExempt <> "Y" THEN
            CustTax# = 0
            FOR RCnt = 1 TO LastRev
              IF TaxRates(RCnt) > 0 THEN
                Diff# = Round#(UBTrans(1).RevAmt(RCnt) / (1 + TaxRates(RCnt)))
                Tax# = Round#(UBTrans(1).RevAmt(RCnt) - Diff#)
                CustTax# = Round#(CustTax# + Tax#)
                TaxAmt(RCnt) = Round#(TaxAmt(RCnt) + Tax#)
              END IF
            NEXT
            TTax# = Round#(TTax# + CustTax#)
            '*********************
            PRINT #UBRpt, Num2Date$(UBTrans(1).TransDate); TAB(11); USING "#####"; UBTrans(1).CustAcctNo;
            IF UBCustRec(1).DelFlag THEN
              PRINT #UBRpt, TAB(24); "*";
            END IF
            PRINT #UBRpt, TAB(25); LEFT$(UBCustRec(1).CUSTNAME, 33);
            PRINT #UBRpt, TAB(65); USING "$$#####,#.##"; CustTax#
            LineCnt = LineCnt + 1
            IF DetailFlag THEN
              FOR RCnt = 1 TO LastRev
                IF TaxRates(RCnt) > 0 THEN
                  Diff# = Round#(UBTrans(1).RevAmt(RCnt) / (1 + TaxRates(RCnt)))
                  Tax# = Round#(UBTrans(1).RevAmt(RCnt) - Diff#)
                  PRINT #UBRpt, RevText$(RCnt); TAB(16); USING "#####.##"; Tax#
                  LineCnt = LineCnt + 1
                END IF
              NEXT RCnt
              PRINT #UBRpt, Dash80$
              LineCnt = LineCnt + 1
            END IF
            IF LineCnt > MaxLines THEN
              PRINT #UBRpt, FF$
              GOSUB DoTaxRptHeader:
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
    ShowPctComp Cnt&, NumOfRecs&
SkipThisOne:
  NEXT

  PRINT #UBRpt, Dash80$
  PRINT #UBRpt, "Total Tax:"; TAB(65); USING "$$######,.##"; TTax#
  PRINT #UBRpt,
  PRINT #UBRpt, "Tax Breakdown:"
  FOR RCnt = 1 TO LastRev
    IF TaxRates(RCnt) > 0 THEN
      PRINT #UBRpt, TAB(5); RevText$(RCnt); TAB(20); USING "$$######,.##"; TaxAmt(RCnt)
    END IF
  NEXT
  PRINT #UBRpt,
  PRINT #UBRpt, "Report Parameters"
  PRINT #UBRpt, "     Report Type: "; Form$(3, 0)
  PRINT #UBRpt, "      Start Date: "; Form$(1, 0)
  PRINT #UBRpt, "     Ending Date: "; Form$(2, 0)
  PRINT #UBRpt, "          Detail: "; Form$(4, 0)

  CLOSE
  
  IF NOT AbortFlag THEN
    PrintRptFile "Transaction Tax Report.", "UBTAXRPT.RPT", 1, RetCode, EntryPoint
  END IF

ExitTaxReport:
  EXIT SUB

CheckInfo:
  BegDate = Date2Num(Form$(1, 0))
  EndDate = Date2Num(Form$(2, 0))
  Rpt$ = LEFT$(Form$(3, 0), 1)
  IF Form$(4, 0) = "Y" THEN
    DetailFlag = True
  ELSE
    DetailFlag = False
  END IF
  IF (BegDate < 0 OR EndDate < 0) OR (BegDate > EndDate) THEN
    Frm(1).FldNo = 1
    ErrCode = 1
    GOSUB ShowTaxParmErr
  ELSEIF LEN(QPTrim$(Form$(3, 0))) = 0 THEN
    Frm(1).FldNo = 3
    ErrCode = 2
    GOSUB ShowTaxParmErr
  ELSE
    ParmsOK = True
    SELECT CASE Rpt$
    CASE "P"
      TransType = TranBillPayment
    CASE "B"
      TransType = TranUtilityBill
    END SELECT
  END IF
  RETURN
  
ShowTaxParmErr:
  SaveScrn TempScrn()
  DisplayUBScrn "ERRSCRN1"
  SELECT CASE ErrCode
  CASE 1
    QPrintRC "Invalid Start/Ending Dates!", 10, 26, -1
  CASE 2
    QPrintRC "Invalid Report Type!", 10, 29, -1
  END SELECT
  QPrintRC "Correct and try again.", 13, 29, -1
  WaitForAction
  RestScrn TempScrn()
  Action = 1
  RETURN
  
DoTaxRptHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, TAB(30); "Transaction Tax Report"
  PRINT #UBRpt, TownName$; TAB(70); "Page #"; PageNo
  PRINT #UBRpt, "Report Date: "; TheDate$
  PRINT #UBRpt, "  Date      Acct #             Customer Name                      Tax Amount"
  PRINT #UBRpt, Dash80$
  LineCnt = 5
RETURN
  
  
END SUB

SUB MemberFeeReport

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
    TitleBox 2, Col, MaxLen + 3, "Membership Fees Report ", Cnf
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

  IF ExitFlag THEN GOTO ExitMemberShipRpt

  MaxLines = 55
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
  OPEN "UBMFEES.RPT" FOR OUTPUT AS UBRpt

  UBSetup = FREEFILE
  OPEN "UBSETUP.DAT" FOR RANDOM SHARED AS UBSetup LEN = UBSetUpRecLen
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
  ShowProcessingScrn "Membership Fees Report."

  GOSUB DoMemberShipRptHeader

  FOR Cnt = 1 TO NumOfRecs
    IF UsingName OR UsingBook THEN
      acctnumber = IdxBuff(Cnt).RecNum
    ELSE
      acctnumber = Cnt
    END IF

    GET UBCust, acctnumber, UBCustRec(1)
    IF UBCustRec(1).DelFlag = 0 THEN
      IF Round#(UBCustRec(1).MFee1) > 0 OR Round#(UBCustRec(1).MFee2) > 0 THEN
        PRINT #UBRpt, USING "#####"; acctnumber;
        PRINT #UBRpt, TAB(8); QPTrim$(UBCustRec(1).BOOK); "-"; QPTrim$(UBCustRec(1).SEQNUMB);
        PRINT #UBRpt, TAB(20); LEFT$(UBCustRec(1).CUSTNAME, 30);
        PRINT #UBRpt, TAB(54); USING "#####.##"; UBCustRec(1).MFee1; TAB(67); UBCustRec(1).MFee2
        LineCnt = LineCnt + 1
        IF UBCustRec(1).MFee1 >= 0 THEN
          RTotal# = Round#(RTotal# + UBCustRec(1).MFee1)
        END IF
        IF UBCustRec(1).MFee2 >= 0 THEN
          NTotal# = Round#(NTotal# + UBCustRec(1).MFee2)
        END IF
        CustCnt = CustCnt + 1
      END IF
    END IF

    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, FF$
      GOSUB DoMemberShipRptHeader
    END IF

    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF

    ShowPctComp Cnt, NumOfRecs
  NEXT

  GOSUB DoMemberShipRptFooter:

  CLOSE UBCust, UBRpt

  ERASE IdxBuff, UBCustRec

  IF NOT AbortFlag THEN
    PrintRptFile "Membership Fees Report.", "UBMFEES.RPT", 1, RetCode, EntryPoint
  END IF

  'KillFile "UBMFEES.RPT"

ExitMemberShipRpt:

  EXIT SUB

DoMemberShipRptHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, TownName$
  PRINT #UBRpt, TAB(24); "Customer Membership Fees Report"; TAB(70); "Page #"; PageNo
  PRINT #UBRpt, "Report Date: "; DATE$; TAB(71); "NON"
  PRINT #UBRpt, "Acct #  Location   Customer Name                      Refundable   Refundable"
  PRINT #UBRpt, Dash80$
  LineCnt = 5
RETURN

DoMemberShipRptFooter:
  PRINT #UBRpt, Dash80$
  PRINT #UBRpt, "Totals:"; TAB(10); "Customers: "; USING "#####,#"; CustCnt;
  PRINT #UBRpt, TAB(52); USING "#######.##"; RTotal#; TAB(65); NTotal#
  PRINT #UBRpt, FF$
RETURN

END SUB

SUB MeterInstallDate

  LibName$ = "UBHH"
  ScrnName$ = "MINSDATE"
  RptName$ = "UBMTINDT.RPT"

  MaxLines = 52
  PageNo = 0
  Dash80$ = STRING$(80, "-")

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
  FirstTime = True
  BlockClear
  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%

  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      'LSET Form$(1, 0) = "01-01-2000"
      LSET Form$(2, 0) = DATE$
      Action = 1
    END IF

    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      GOSUB CheckDates
    CASE ESC
      ExitFlag = True
      EXIT DO
    CASE ELSE
      Done = False
    END SELECT

  LOOP UNTIL OKFlag

  IF ExitFlag THEN
    GOTO ExitMtrInsRpt
  END IF

  MaxLines = 52
  PageNo = 0
  Dash80$ = STRING$(80, "-")
  IndexName$ = TempIndexName

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBSetup(1) AS UBSetupRecType
  UBSetUpRecLen = LEN(UBSetup(1))

  SortMtrDateIndex FromDate, ThruDate

  'SortServiceAddrs "Service Address"

  IdxRecLen = 4               'we are using a long integer
  IdxFileSize& = FileSize&(IndexName$)
  IdxNumRecs = IdxFileSize& \ IdxRecLen
  REDIM IdxBuff(1 TO IdxNumRecs) AS UBCustIndexRecType
  FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumRecs      'load it
  NumOfRecs& = IdxNumRecs
  
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

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen

  UBRpt = FREEFILE
  OPEN RptName$ FOR OUTPUT AS UBRpt

  BlockClear
  ShowProcessingScrn "Meter Installed Report."

  GOSUB DoMtrInsDateRptHeader

  FOR Cnt = 1 TO NumOfRecs&
    GET UBCust, IdxBuff(Cnt).RecNum, UBCustRec(1)
    AcctNumb& = IdxBuff(Cnt).RecNum

    Status$ = UBCustRec(1).Status
    'LocationNumber$ = UBCustRec(1).BOOK + "-" + UBCustRec(1).SEQNUMB

    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, FF$
      GOSUB DoMtrInsDateRptHeader
    END IF
    CustPrinted = False
    FOR MTCnt = 1 TO 7
      MtrDate = UBCustRec(1).LocMeters(MTCnt).InsDate
      IF MtrDate >= FromDate AND MtrDate <= ThruDate THEN
        IF NOT CustPrinted THEN
          CustPrinted = True
          'PRINT #UBRpt, LocationNumber$;
          PRINT #UBRpt, USING "#####"; AcctNumb&;
          PRINT #UBRpt, "  "; Status$;
          PRINT #UBRpt, TAB(17); QPTrim$(UBCustRec(1).CUSTNAME);
          PRINT #UBRpt, TAB(48); LEFT$(UBCustRec(1).ServAddr, 30)
          LineCnt = LineCnt + 1
        END IF
        PRINT #UBRpt, Num2Date(UBCustRec(1).LocMeters(MTCnt).InsDate);
        PRINT #UBRpt, TAB(17); QPTrim$(UBCustRec(1).LocMeters(MTCnt).MtrNum);
        PRINT #UBRpt, USING "##########"; TAB(35); UBCustRec(1).LocMeters(MTCnt).CurRead
        LineCnt = LineCnt + 1
      END IF
    NEXT
    IF CustPrinted THEN
      PRINT #UBRpt, Dash80$
      LineCnt = LineCnt + 1
    END IF
'MtrDateSkip:
    ShowPctComp Cnt, NumOfRecs&
  NEXT
  CLOSE

  IF NOT AbortFlag THEN
    PrintRptFile "Meter Installed Date Report.", RptName$, 1, RetCode, EntryPoint
  END IF

ExitMtrInsRpt:

EXIT SUB


CheckDates:
  FromDate = Date2Num(Form$(1, 0))
  ThruDate = Date2Num(Form$(2, 0))

  IF FromDate < -29219 THEN
    FromDate = -29219        '01-01-1900
  END IF
  IF FromDate <= ThruDate THEN
    OKFlag = True
  END IF
RETURN

DoMtrInsDateRptHeader:

  PageNo = PageNo + 1
  PRINT #UBRpt, TownName$; TAB(70); "Page #"; PageNo
  PRINT #UBRpt, TAB(28); "Meter Installed Date Report"; TAB(61); " Date: "; DATE$
  PRINT #UBRpt, "Acct #  Status   Cust Name                    Service Addr"
  PRINT #UBRpt, "Ins Date         Mtr Numb              Current Reading"
  PRINT #UBRpt, Dash80$
  LineCnt = 5
  RETURN

RETURN
END SUB

SUB PaymentSumReport
  
  REDIM TempScrn(0)
  REDIM RevenueName$(15)
  REDIM Revenues(1 TO 15) AS DOUBLE
  REDIM TaxRates(1 TO 15) AS SINGLE
  REDIM TaxAmt(1 TO 15) AS DOUBLE
  
  REDIM UBTrans(1) AS UBTransRecType
  UBTransRecLen = LEN(UBTrans(1))
  
  REDIM UBSetup(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetup(), UBSetUpRecLen
  
  FOR RCnt = 1 TO 15
    TRevName$ = QPTrim$(UBSetup(1).Revenues(RCnt).REVNAME)
    IF LEN(TRevName$) > 0 THEN
      RevenueName$(RCnt) = TRevName$
      TaxRates(RCnt) = UBSetup(1).Revenues(RCnt).TaxRate
    ELSE
      MaxRevenue = RCnt - 1
      EXIT FOR
    END IF
  NEXT
  
  TownName$ = UBSetup(1).UTILNAME
  TownLen = LEN(RTRIM$(TownName$))
  TabStop = 40 - (TownLen / 2)
  IF TabStop < 1 THEN TabStop = 1
  
  '*******************************
  
SelectFromList3:
  LibName$ = "UB"
  ScrnName$ = "UBPAYSUM"
  
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
  Form$(1, 0) = DATE$
  Form$(2, 0) = DATE$
  Action = 1
  
  BlockClear
  DisplayUBScrn ScrnName$
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      BegDate = Date2Num(Form$(1, 0))
      
      EndDate = Date2Num(Form$(2, 0))
      IF EndDate < BegDate THEN
        Done = False
        BEEP
        SaveScrn TempScrn()
        CursorOff
        DisplayUBScrn "UBBADDAT"
        WaitForAction
        RestScrn TempScrn()
        Frm(1).FldNo = 1
        Action = 1
      ELSE
        Done = True
      END IF
      
    CASE ESC
      GOTO ExitDetailedListing3
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  
  '***************
  ' Set Up Specifications from Input Screen
  Operator$ = Form$(3, 0)
  Operator = VAL(Operator$)
  FromDate$ = Form$(1, 0)
  ToDate$ = Form$(2, 0)
  
  IF Operator = 0 THEN
    BegOperator = 0
    EndOperator = 99
  ELSE
    BegOperator = Operator
    EndOperator = Operator
  END IF
  
  '***************
  MaxLines = 55
  PageNo = 0
  Dash80$ = STRING$(80, "-")
  
  UBRpt = FREEFILE
  OPEN "UBPAYSUM.RPT" FOR OUTPUT AS UBRpt
  
  UBTrans = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTrans LEN = UBTransRecLen
  
  NumOfRecs& = LOF(UBTrans) \ UBTransRecLen
  
  BlockClear
  ShowProcessingScrn "Payment Summary Report."
  
  GOSUB DoDetailedRptHeader3
  
  FOR Cnt& = 1 TO NumOfRecs&
    GET UBTrans, Cnt&, UBTrans(1)
    TransOK = False
    IF (UBTrans(1).TransDate >= BegDate AND UBTrans(1).TransDate <= EndDate) AND (UBTrans(1).OperatorNumber >= BegOperator AND UBTrans(1).OperatorNumber <= EndOperator) THEN
      SELECT CASE UBTrans(1).TransType
      CASE TranBillPayment, TranBillPayment + 100
        TransOK = True
      CASE TranDraftPayment, TranDraftPayment + 100
        TransOK = True
      'CASE TranDepositPayment, TranDepositPayment + 100
      '  TransOK = True
      END SELECT
      
      IF TransOK THEN
        IF UBTrans(1).TaxExempt = "Y" THEN
          TaxExempt = True
          GOTO SkipEm
          'LPRINT UBTrans(1).CustAcctNo
        ELSE
          TaxExempt = False
        END IF
        
        FOR RevCnt = 1 TO 15
          IF NOT TaxExempt THEN
            IF TaxRates(RevCnt) > 0 THEN
              Diff# = Round#(UBTrans(1).RevAmt(RevCnt) / (1 + TaxRates(RevCnt)))
              Tax# = Round#(UBTrans(1).RevAmt(RevCnt) - Diff#)
              TaxAmt(RevCnt) = Round#(TaxAmt(RevCnt) + Tax#)
              Revenues(RevCnt) = Round#(Revenues(RevCnt) + (UBTrans(1).RevAmt(RevCnt) - Tax#))
            ELSE
              Revenues(RevCnt) = Round#(Revenues(RevCnt) + UBTrans(1).RevAmt(RevCnt))
            END IF
          ELSE
            Revenues(RevCnt) = Round#(Revenues(RevCnt) + UBTrans(1).RevAmt(RevCnt))
          END IF
        NEXT
        TransCnt& = TransCnt& + 1
      END IF
    END IF
    
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
SkipEm:
    ShowPctCompL Cnt&, NumOfRecs&
  NEXT
  
  GOSUB DoDetailedRptFooter3
  
  CLOSE
  
  IF NOT AbortFlag THEN
    PrintRptFile "Payment Summary Report.", "UBPAYSUM.RPT", 1, RetCode, EntryPoint
  END IF
  
  KillFile "UBPAYSUM.RPT"
  
ExitDetailedListing3:
  
  EXIT SUB
  
DoDetailedRptHeader3:
  PRINT #UBRpt, TAB(TabStop); TownName$
  PRINT #UBRpt, TAB(29); "Payment Summary Report"
  PRINT #UBRpt, "Beginning Date: "; FromDate$;
  IF VAL(Operator$) = 0 THEN
    PRINT #UBRpt, TAB(65); " Operator #: ALL"
  ELSE
    PRINT #UBRpt, TAB(65); " Operator #: "; Operator$
  END IF
  PRINT #UBRpt, "   Ending Date: "; ToDate$
  PRINT #UBRpt,
  PRINT #UBRpt, "    Source                           Revenue Amt                 Tax"
RETURN
  
DoDetailedRptFooter3:
  PRINT #UBRpt, Dash80$
  FOR Cnt = 1 TO MaxRevenue
    PRINT #UBRpt, TAB(5); RevenueName$(Cnt); TAB(35); USING "$$#######,#.##"; Revenues(Cnt); TAB(55); TaxAmt(Cnt)
    TotalTrans# = Round#(TotalTrans# + Revenues(Cnt))
    TaxTotal# = Round#(TaxTotal# + TaxAmt(Cnt))
  NEXT
  PRINT #UBRpt, Dash80$
  PRINT #UBRpt, "Total Payments: "; TAB(20); USING "######"; TransCnt&
  PRINT #UBRpt, "Revenue Totals: "; TAB(35); USING "$$#######,#.##"; TotalTrans#; TAB(55); TaxTotal#
  PRINT #UBRpt, FF$
RETURN
  
END SUB

SUB SortMtrDateIndex (FromDate%, ThruDate%)

  ShowProcessingScrn "Creating Meter Date Index"
  QPrintRC "    Reading Customer Records     ", 11, 25, -1

  REDIM UBCustRec(1) AS NewUBCustRecType
  CustRecLen = LEN(UBCustRec(1))

  NumCustRecs = GetNumOfCust%

  REDIM MDateIdx(1 TO 1)  AS MtrDateSortType
  IndexRecLen = LEN(MDateIdx(1))

  CHandle = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CHandle LEN = CustRecLen
  FOR Cnt = 1 TO NumCustRecs
    GET CHandle, Cnt, UBCustRec(1)
    IF UBCustRec(1).DelFlag <> -1 THEN
      FOR MTCnt = 1 TO 7
        MtrDate = UBCustRec(1).LocMeters(MTCnt).InsDate
        IF MtrDate >= FromDate AND MtrDate <= ThruDate THEN
          FoundCnt = FoundCnt + 1
          REDIM PRESERVE MDateIdx(1 TO FoundCnt)  AS MtrDateSortType
          MDateIdx(FoundCnt).MtrDate = UBCustRec(1).LocMeters(MTCnt).InsDate
          MDateIdx(FoundCnt).RecNum = Cnt
          EXIT FOR
        END IF
      NEXT
    END IF
    ShowPctComp Cnt, NumCustRecs                'show user percentage complete
  NEXT

  CLOSE CHandle

  QPrintRC "         Sorting Index.        ", 11, 25, -1

  SortT MDateIdx(1), FoundCnt, 0, 4, 0, -1
  'Elemen, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
  QPrintRC "      Writing Index Records      ", 11, 25, -1
  IHandle = FREEFILE
  IndexName$ = TempIndexName
  FCreate IndexName$
  OPEN IndexName$ FOR RANDOM SHARED AS IHandle LEN = 4
  FOR Cnt = 1 TO FoundCnt
    CRec& = MDateIdx(Cnt).RecNum
    PUT IHandle, Cnt, CRec&
    ShowPctComp Cnt, NumCustRecs                'show user percentage complete
  NEXT
  CLOSE IHandle

  ERASE UBCustRec, MDateIdx

END SUB

SUB SortServiceAddrs (IndexText$)
  
  ShowProcessingScrn "Creating " + IndexText$ + " Index"
  QPrintRC "    Reading Customer Records     ", 11, 25, -1
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  CustRecLen = LEN(UBCustRec(1))
  
  NumCustRecs = GetNumOfCust%
  
  REDIM ServIndex(1 TO NumCustRecs)  AS UBServiceAddressIndexType
  IndexRecLen = LEN(ServIndex(1))
  
  CHandle = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CHandle LEN = CustRecLen
  FOR Cnt = 1 TO NumCustRecs
    GET CHandle, Cnt, UBCustRec(1)
    ServIndex(Cnt).ServiceAddress = UBCustRec(1).ServAddr
    ServIndex(Cnt).RecNum = Cnt
    ShowPctComp Cnt, NumCustRecs                'show user percentage complete
  NEXT
  
  CLOSE CHandle
  
  QPrintRC "         Sorting Index.        ", 11, 25, -1
  
  SortT ServIndex(1), NumCustRecs, 0, 16, 0, 14
  ' SortT (Elemen, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
  QPrintRC "      Writing Index Records      ", 11, 25, -1
  IHandle = FREEFILE
  IndexName$ = TempIndexName
  FCreate IndexName$
  OPEN IndexName$ FOR RANDOM SHARED AS IHandle LEN = 4
  FOR Cnt = 1 TO NumCustRecs
    CRec& = ServIndex(Cnt).RecNum
    PUT IHandle, Cnt, CRec&
    ShowPctComp Cnt, NumCustRecs                'show user percentage complete
  NEXT
  CLOSE IHandle
  
  ERASE UBCustRec, ServIndex
  
  
END SUB

SUB WestCaroConsumpList
  
  BlockClear
  ShowProcessingScrn "WCU Consumption Report."
  
  RptTitle$ = "W.C.U. Consumption Report"
  
  REDIM UBSetup(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetup(), UBSetUpRecLen
  TownName$ = UBSetup(1).UTILNAME
  TownLen = LEN(RTRIM$(TownName$))
  TabStop = 40 - (TownLen / 2)
  IF TabStop < 1 THEN TabStop = 1
  
  IndexName$ = BookIndexFile
  
  MaxLines = 55
  PageNo = 0
  Dash80$ = STRING$(80, "-")
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  IdxRecLen = 4 'we are using a long integer
  IdxFileSize& = FileSize(IndexName$)
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs        'load it
  NumOfRecs = IdxNumOfRecs
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  UBRpt = FREEFILE
  OPEN "UBWCUCON.RPT" FOR OUTPUT AS UBRpt
  
  GOSUB DoWestRptHeader
  
  FOR CCnt = 1 TO NumOfRecs
    GET UBCust, IdxBuff(CCnt).RecNum, UBCustRec(1)
    PCode$ = UCASE$(QPTrim$(UBCustRec(1).PUMPCODE))
    IF PCode$ = "W" THEN
      GOSUB PrintThisOne
    END IF
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, FF$
      GOSUB DoWestRptHeader
    END IF
    ShowPctComp CCnt, NumOfRecs
  NEXT
  
  GOSUB DoWestRptFooter
  
  CLOSE
  
  EntryPoint = 0
  LPTPort = 1
  ERASE UBSetup, UBCustRec, IdxBuff
  
  IF NOT AbortFlag THEN
    PrintRptFile RptTitle$, "UBWCUCON.RPT", LPTPort, RetCode, EntryPoint
  END IF
  
  EXIT SUB
  
PrintThisOne:
  CustCnt = CustCnt + 1
  CustConsump& = UBCustRec(1).LocMeters(1).CurRead - UBCustRec(1).LocMeters(1).PrevRead
  TotalConsump& = TotalConsump& + CustConsump&
  PRINT #UBRpt, UBCustRec(1).BOOK; "-"; UBCustRec(1).SEQNUMB; USING "######"; IdxBuff(CCnt).RecNum;
  PRINT #UBRpt, TAB(18); LEFT$(UBCustRec(1).CUSTNAME, 25);
  PRINT #UBRpt, TAB(44); USING "############"; UBCustRec(1).LocMeters(1).CurRead; UBCustRec(1).LocMeters(1).PrevRead; CustConsump&
  LineCnt = LineCnt + 1
  RETURN
  
DoWestRptHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, TAB(TabStop); TownName$
  PRINT #UBRpt, TAB(28); RptTitle$
  PRINT #UBRpt, "Report Date: "; DATE$; TAB(70); "Page #"; PageNo
  PRINT #UBRpt, TAB(49); "Current    Previous"
  PRINT #UBRpt, "Location  ActNo. Customer Name"; TAB(49); "Reading     Reading  Consumption"
  PRINT #UBRpt, Dash80$
  LineCnt = 6
  RETURN
  
DoWestRptFooter:
  PRINT #UBRpt, Dash80$
  PRINT #UBRpt, "Totals"
  PRINT #UBRpt, "      Customers:"; USING "############"; CustCnt
  PRINT #UBRpt, "    Consumption:"; USING "############"; TotalConsump&
  PRINT #UBRpt, FF$
  RETURN
  
END SUB

