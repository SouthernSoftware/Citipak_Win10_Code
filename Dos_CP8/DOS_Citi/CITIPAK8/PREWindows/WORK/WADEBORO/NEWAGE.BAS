DEFINT A-Z

DEFSNG A-Z
DEFINT A-Z
SUB AgedBalanceListing
  REDIM MChoice$(1 TO 3)
  REDIM Revenue#(15), BookTotal#(100), RevenueName$(15), BookRev#(100, 15)
  
  MChoice$(1) = " Customer Name Order   "
  MChoice$(2) = " Account Number Order  "
  MChoice$(3) = " Location Number Order  "
  
  
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
  Col = ((80 - MaxLen) \ 2) - 2
  
  UsingBook = False
  UsingAcct = False
  UsingName = False
  
  AbortFlag = False
  PageNo = 0
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    TitleBox 3, Col, MaxLen + 3, "Aged Balance Report ", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
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
    END SELECT
    
  LOOP UNTIL OKFlag OR ExitFlag
  
  IF ExitFlag THEN GOTO ExitAgedBalListing
  
  
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
  Form$(1, 0) = DATE$
  Action = 1
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      AgeDate$ = Form$(1, 0)
      IF VAL(LEFT$(AgeDate$, 2)) < 1 OR VAL(LEFT$(AgeDate$, 2)) > 12 THEN
        Done = False: Action = 1: Frm(1).FldNo = 1
      ELSEIF VAL(MID$(AgeDate$, 4, 2)) < 1 OR VAL(MID$(AgeDate$, 4, 2)) > 31 THEN
        Done = False: Action = 1: Frm(1).FldNo = 1
      ELSE
        Done = True
      END IF
      
    CASE ESC
      EXIT SUB
    CASE ELSE
      Done = False
    END SELECT
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
      AcctNumber = IdxBuff(Cnt).RecNum
    ELSE
      GET UBCust, Cnt, UBCustRec(1)
      AcctNumber = Cnt
    END IF
    
    
    IF UBCustRec(1).DelFlag THEN
      GOTO skipem
    END IF
    
    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, CHR$(12)
      GOSUB DoAgedBalRptHeader
    END IF
    
    '*************************************
    '   Main body of Printing goes here
    
    AcctBalanceHere# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
    'AcctBalanceHere# = Round#(AcctBalanceHere#)
    
    GOSUB CalcBal
    LineTotal# = Round#(LineTotal#)
    
    IF LineTotal# <> 0 THEN
      
      'Clean Up Customer Name
      CustName$ = UBCustRec(1).CustName
      IF ASC(LEFT$(CustName$, 1)) < 64 OR ASC(LEFT$(CustName$, 1)) > 90 THEN CustName$ = ""
      CustName$ = QPTrim$(LEFT$(CustName$, 28))
      
      'Clean Up Seq Number
      SeqNumber$ = UBCustRec(1).SeqNumb
      IF ASC(LEFT$(SeqNumber$, 1)) < 45 OR ASC(LEFT$(SeqNumber$, 1)) > 57 THEN SeqNumber$ = ""
      SeqNumber$ = QPTrim$(SeqNumber$)
      
      PRINT #UBRpt, UBCustRec(1).Book; "-"; SeqNumber$;
      PRINT #UBRpt, TAB(12); CustName$;
      PRINT #UBRpt, TAB(41); USING "######.##"; Amt1#;
      PRINT #UBRpt, TAB(51); USING "######.##"; Amt2#;
      PRINT #UBRpt, TAB(61); USING "######.##"; Amt3#;
      PRINT #UBRpt, TAB(71); USING "######.##"; Amt4#;
      PRINT #UBRpt, TAB(81); USING "######.##"; LineTotal#;
      PRINT #UBRpt, TAB(92); USING "#####"; AcctNumber
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
    LineTotal# = 0: Amt1# = 0: Amt2# = 0: Amt3# = 0: Amt4# = 0
    TotalPayments = 0
skipem:
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
  PRINT #UBRpt, CHR$(12);
  
  
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
  PRINT #UBRpt, CHR$(12);
  PRINT #UBRpt, "Totals by Revenue"
  PRINT #UBRpt, ""
  PRINT #UBRpt, "Revenue"; TAB(30); "Amount"
  PRINT #UBRpt, STRING$(50, "-")
  FOR Rev = 1 TO 15
    IF Revenue#(Rev) <> 0 THEN
      PRINT #UBRpt, RevenueName$(Rev); TAB(30); USING "$$######,#.##"; Revenue#(Rev)
    END IF
  NEXT Rev
  PRINT #UBRpt, CHR$(12);
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
  
  DO WHILE Trans& > 0
    GET UBTrans, Trans&, UBTransRec(1)
    IF UBTransRec(1).TransType <> 107 OR UBTransRec(1).TransType <> 7 THEN
      IF UBTransRec(1).TransDate <= AgeDate THEN
        GOSUB LineCalc
      END IF
    END IF
    Trans& = UBTransRec(1).PrevTrans
  LOOP
  
  GOSUB FinishCalc
  RETURN
  
FinishCalc:
  
  'SELECT CASE acctnumber
  'CASE 28
  '  STOP
  'END SELECT
  
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
  
  IF Amt1# < 0 AND AcctBalanceHere# > 0 THEN
    Amt1# = AcctBalanceHere#
  END IF
  
  'IF Amt1# <> AcctBalanceHere# THEN
  '  AmtDif# = AcctBalanceHere#
  'END IF
  
  LineTotal# = Round#(Amt1# + Amt2# + Amt3# + Amt4#)
  
  IF LineTotal# < AcctBalanceHere# THEN
    IF AcctBalanceHere# > 0 THEN
      'STOP
      AmtDif# = Round#(AcctBalanceHere# - LineTotal#)
      Amt4# = Round#(Amt4# + AmtDif#)
    END IF
    
  ELSEIF LineTotal# > AcctBalanceHere# THEN
    AmtDif# = Round#(AcctBalanceHere# - LineTotal#)
  END IF
  
  IF AmtDif# <> 0 THEN
    IF Amt4# <> 0 THEN
      Amt4# = Round#(Amt4# - AmtDif#)
      AmtDif# = 0
      GOTO Gotit
    END IF
    IF Amt3# <> 0 THEN
      Amt3# = Round#(Amt3# - AmtDif#)
      AmtDif# = 0
      GOTO Gotit
    END IF
    IF Amt2# <> 0 THEN
      Amt2# = Round#(Amt2# - AmtDif#)
      AmtDif# = 0
      GOTO Gotit
    END IF
    IF Amt1# <> 0 THEN
      Amt1# = Round#(Amt1# - AmtDif#)
      AmtDif# = 0
    END IF
    
  END IF
  
Gotit:
  
  LineTotal# = AcctBalanceHere#
  
  BookTotal#(Book) = BookTotal#(Book) + LineTotal#
  RETURN
  
  
LineCalc:
  TransAmt# = UBTransRec(1).TransAmt
  Revenue# = 0
  FOR Rev = 1 TO 15
    Revenue# = Revenue# + UBTransRec(1).RevAmt(Rev)
  NEXT Rev
  TransAmt# = INT((TransAmt# * 100) + .5) / 100
  Revenue# = INT((Revenue# * 100) + .5) / 100
  IF TransAmt# <> Revenue# AND UBTransRec(1).TransType <> 9 THEN
    RETURN
  END IF
  
  IF UBTransRec(1).TransType = 4 OR UBTransRec(1).TransType = 5 OR UBTransRec(1).TransType = 8 OR UBTransRec(1).TransType = 12 OR UBTransRec(1).TransType = 104 OR UBTransRec(1).TransType = 105 OR UBTransRec(1).TransType = 109 THEN
    TotalPayments# = TotalPayments# + UBTransRec(1).TransAmt
    TotalPayments# = Round#(TotalPayments#)
    FOR Rev = 1 TO 15
      Revenue#(Rev) = Revenue#(Rev) - UBTransRec(1).RevAmt(Rev)
      BookRev#(Book, Rev) = BookRev#(Book, Rev) - UBTransRec(1).RevAmt(Rev)
    NEXT Rev
    
    RETURN
  END IF
  
  SELECT CASE UBTransRec(1).TransType
  CASE 1, 2, 3, 6, 11, 10, 101, 102, 103, 112, 110
    
    Days& = AgeDate - UBTransRec(1).TransDate
    IF Days& <= 30 THEN
      Amt1# = Amt1# + UBTransRec(1).TransAmt
      Amt1# = Round#(Amt1#)
      FOR Rev = 1 TO 15
        Revenue#(Rev) = Round#(Revenue#(Rev) + (UBTransRec(1).RevAmt(Rev)))
        BookRev#(Book, Rev) = Round#(BookRev#(Book, Rev) + UBTransRec(1).RevAmt(Rev))
      NEXT Rev
      RETURN
    END IF
    IF Days& > 30 AND Days& <= 60 THEN
      Amt2# = Amt2# + UBTransRec(1).TransAmt
      Amt2# = Round#(Amt2#)
      FOR Rev = 1 TO 15: Revenue#(Rev) = Revenue#(Rev) + (UBTransRec(1).RevAmt(Rev))
        BookRev#(Book, Rev) = BookRev#(Book, Rev) + UBTransRec(1).RevAmt(Rev)
      NEXT Rev
      RETURN
    END IF
    IF Days& > 60 AND Days& <= 90 THEN
      Amt3# = Amt3# + UBTransRec(1).TransAmt
      Amt3# = Round#(Amt3#)
      FOR Rev = 1 TO 15: Revenue#(Rev) = Revenue#(Rev) + (UBTransRec(1).RevAmt(Rev))
        BookRev#(Book, Rev) = BookRev#(Book, Rev) + UBTransRec(1).RevAmt(Rev)
      NEXT Rev
      RETURN
    END IF
    IF Days& > 90 THEN
      Amt4# = Amt4# + UBTransRec(1).TransAmt
      Amt4# = Round#(Amt4#)
      FOR Rev = 1 TO 15
        Revenue#(Rev) = Revenue#(Rev) + (UBTransRec(1).RevAmt(Rev))
        BookRev#(Book, Rev) = BookRev#(Book, Rev) + UBTransRec(1).RevAmt(Rev)
      NEXT Rev
      RETURN
    END IF
  END SELECT
  
  RETURN
  
  
  
  
END SUB

