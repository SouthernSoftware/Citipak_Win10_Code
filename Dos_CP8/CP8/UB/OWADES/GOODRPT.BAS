
DEFINT A-Z
DECLARE SUB CustomerFlatRates ()
DECLARE SUB WestCaroConsumpList ()
DECLARE FUNCTION Monitor% ()
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetupLen%)

DECLARE SUB SortServAddress (IndexText$)
DECLARE SUB CustEditMeterEST (CustRec&)
DECLARE SUB CustByRate ()
DECLARE SUB CustMessageSystem (RecNo&)
DECLARE SUB ShowWrkOrdHistory (RecNo&)
DECLARE SUB CustWorkOrders (CustRec&)
DECLARE SUB ShowCustConsHist (CustRec&)
DECLARE SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE SUB PaymentSumReport ()
DECLARE SUB MailingLabel ()
DECLARE SUB AgedBalanceListing ()
DECLARE SUB CustomerCutOffListing ()
DECLARE SUB CustomerInquiry ()
DECLARE SUB DetailedTransJournal ()
DECLARE SUB MasterBalanceListing ()
DECLARE SUB MasterCustomerListing ()
DECLARE SUB MasterDepositListing ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE FUNCTION GetNumRateRecs% ()
DECLARE FUNCTION FLof& (Handle%)
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE FUNCTION FmtBook$ (Book$)
DECLARE FUNCTION FmtSeqN$ (SeqN$)
DECLARE FUNCTION FindRateTbl% (RateCode$, NumOfRates%, UBRateTbls() AS ANY)
DECLARE FUNCTION GetNumOfRevs% ()
DECLARE SUB AddEditLocation (RecNo&, FromFlag%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB Search4Cust (SEARCH$, RecNo&, CLSFlag%, LocationFlag%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB ShowCustHistory (CustRec&)
DECLARE FUNCTION GetNumOfCust% ()
DECLARE FUNCTION ConvDateStr$ (D$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FudgeFactor& ()
DECLARE FUNCTION NovellThere% ()
DECLARE FUNCTION QPStrI$ (Num%)
DECLARE FUNCTION QPStrL$ (Num&)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Round# (N#)
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE SUB BlockClear ()
DECLARE SUB ButtonPress (ButNo%, Down%, Presses%, x%, Y%)
DECLARE SUB ClearScrn ()
DECLARE SUB CursorOff ()
DECLARE SUB DisplayScrn (BTmp%(), Element%, MonoCode%, WipeType%)
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB ExplodBox (UlRow%, UlCol%, BrRow%, BrCol%, Scr(), El%)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB GetCursor (x, Y, Button)
DECLARE SUB HideCursor ()
DECLARE SUB BlockClear ()
DECLARE SUB MPaintBox (UlRow%, UlCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB MScrnRest (UlRow, UlCol, LRRow, LRCol, SEG Address)
DECLARE SUB MScrnSave (UlRow, UlCol, LRRow, LRCol, SEG Address)
DECLARE SUB Pause3 (MillaSecs%, ProcessorSpeed&)
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
DECLARE SUB ShowCursor ()
DECLARE SUB SaveScrn (array())
DECLARE SUB RestScrn (array())
DECLARE SUB StuffBuf (Ky$)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE SUB FOpenS (FileName$, Handle)
DECLARE SUB FClose (Handle%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE FUNCTION QPValL& (Number$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice, MaxLen%, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB WaitForAction ()
DECLARE SUB Get.Moose.OR.Key (Ky$, MooseButton%, MRow%, MCol%)
DECLARE FUNCTION Chk4DupeBookSeqNum (Book$, SeqNum$)
DECLARE FUNCTION PromptSaveData% ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB KillFile (FileName$)
DECLARE SUB LookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, LocationFlag%)
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetupLen%)
DECLARE SUB MakeSequenceIndex (IndexText$)
DECLARE SUB MakePostalIndex (IdxTypeText$)
DECLARE FUNCTION QPStrL$ (LongNumber&)
DECLARE FUNCTION WEnvTest% ()
  
  CONST False = 0, True = NOT False
  
  TYPE Struct
    TrDate AS INTEGER
    RecNum AS LONG
  END TYPE
  
  DIM SHARED PctC(1) AS STRING * 3
  
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'PageInfo.BI'                  'Form Page info
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'ubrate.BI'
  '$INCLUDE: 'newcust.bi'

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
  REDIM MChoice$(1 TO 11)
  
  MChoice$(1) = "Customer Inquiry"
  MChoice$(2) = "Master Customer Listing"
  MChoice$(3) = "Master Balance Listing"
  MChoice$(4) = "Aged Balance Listing"
  MChoice$(5) = "Master Deposit Listing"
  MChoice$(6) = "Customer Flat-Rates Listing"
  MChoice$(7) = "Customer Cut-Off Listing"
  MChoice$(8) = "Detailed Transaction Journal"
  MChoice$(9) = "Payment Summary Report"
  MChoice$(10) = "Mailing Labels"
  MChoice$(11) = "Exit to OS"

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
  Row = ((25 - (UBOUND(MChoice$))) \ 2)
  Col = ((80 - MaxLen) \ 2) - 2
  help$ = "Customer Reports Menu"
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    
    TitleBox 2, Col, MaxLen + 3, "Customer Reports Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    'PrintTitle user$
    'PrintHelp Help$
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    
    SELECT CASE Choice
    CASE 1
      CustomerInquiry
    CASE 2
      MasterCustomerListing
    CASE 3
      MasterBalanceListing
    CASE 4
      AgedBalanceListing
    CASE 5
      MasterDepositListing
    CASE 6
      CustomerFlatRates
    CASE 7
      CustomerCutOffListing
    CASE 8
      DetailedTransJournal
    CASE 9
      PaymentSumReport
    CASE 10
      MailingLabel
    CASE 11
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
      Acctnumber = IdxBuff(Cnt).RecNum
    ELSE
      GET UBCust, Cnt, UBCustRec(1)
      Acctnumber = Cnt
    END IF

   AcctBalanceHere# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
   AcctBalanceHere# = Round#(AcctBalanceHere#)


    IF UBCustRec(1).DelFlag THEN
      GOTO SkipEm
    END IF

    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, CHR$(12)
      GOSUB DoAgedBalRptHeader
    END IF

    '*************************************
    '   Main body of Printing goes here
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
      PRINT #UBRpt, TAB(92); USING "#####"; Acctnumber
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
'   IF Round#(LineTotal#) <> Round#(AcctBalanceHere#) THEN
'   LPRINT Acctnumber; "AB= "; AcctBalanceHere#; "  LB= "; LineTotal#; "  Diff= "; AcctBalanceHere# - LineTotal#
'  END IF
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
    IF UBTransRec(1).TransType <> 107 OR UBTransRec(1).TransType <> 7 OR UBTransRec(1).TransType <> 9 THEN
      IF UBTransRec(1).TransDate <= AgeDate THEN
        GOSUB LineCalc
      END IF
    END IF
    Trans& = UBTransRec(1).PrevTrans
  LOOP

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

  IF Amt1# < 0 AND AcctBalanceHere# > 0 THEN
    Amt1# = AcctBalanceHere#
  END IF


  LineTotal# = Round#(Amt1# + Amt2# + Amt3# + Amt4#)


Gotit:

  LineTotal# = Round#(LineTotal#)

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
'  IF TransAmt# <> Revenue# AND UBTransRec(1).TransType <> 9 THEN
'    RETURN
'  END IF

  IF UBTransRec(1).TransType = 4 OR UBTransRec(1).TransType = 5 OR UBTransRec(1).TransType = 8 OR UBTransRec(1).TransType = 12 OR UBTransRec(1).TransType = 104 OR UBTransRec(1).TransType = 105 OR UBTransRec(1).TransType = 109 THEN
    IF UBTransRec(1).TransType = 5 THEN
    TotalPayments# = TotalPayments# + ABS(UBTransRec(1).TransAmt)
    TotalPayments# = Round#(TotalPayments#)
    ELSE
    TotalPayments# = TotalPayments# + UBTransRec(1).TransAmt
    TotalPayments# = Round#(TotalPayments#)
    END IF
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

SUB CustByRate
  'here
  
  NumOfRateRecs = GetNumRateRecs
  
  REDIM UBRateTblRec(1 TO NumOfRateRecs) AS UBRateTblRecType
  UBRateTblRecLen = LEN(UBRateTblRec(1))
  
  UBFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS UBFile LEN = UBRateTblRecLen
  FOR Cnt = 1 TO NumOfRateRecs
    GET UBFile, Cnt, UBRateTblRec(Cnt)
  NEXT
  CLOSE
  
  REDIM CustCnt(1 TO NumOfRateRecs) AS INTEGER
  
  '***************
  MaxLines = 55
  PageNo = 0
  Dash80$ = STRING$(80, "-")
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBSetup(1) AS UBSetupRecType
  UBSetUpRecLen = LEN(UBSetup(1))
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  NumOfRecs = LOF(UBCust) \ UBCustRecLen
  
  UBRpt = FREEFILE
  OPEN "UBCBRATE.RPT" FOR OUTPUT AS UBRpt
  
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
  
  
  BlockClear
  ShowProcessingScrn "Customer By Rate Code Report."
  
  GOSUB DOCustRateHeader
  
  FOR Cnt = 1 TO NumOfRecs
    GET UBCust, Cnt, UBCustRec(1)
    IF UBCustRec(1).Status = "A" THEN
      'IF LineCnt > MaxLines THEN
      '  PRINT #UBRpt, CHR$(12)
      '  GOSUB DOCustRateHeader
      'END IF
      FOR RCnt = 1 TO 7
        '*************************************
        '   Main body of Printing goes here
        ThisCode$ = QPTrim$(UBCustRec(1).Serv(RCnt).RateCode)
        IF LEN(ThisCode$) > 0 THEN
          FOR ZCnt = 1 TO NumOfRateRecs
            IF INSTR(UBRateTblRec(ZCnt).RateCode, ThisCode$) > 0 THEN
              CustCnt(ZCnt) = CustCnt(ZCnt) + 1
              EXIT FOR
            END IF
          NEXT
        END IF
      NEXT
      'LineCnt = LineCnt + 1
      CustomerCnt = CustomerCnt + 1
      '*************************************
      IF AskAbandonPrint% THEN
        AbortFlag = True
        EXIT FOR
      END IF
    END IF
    ShowPctComp Cnt, NumOfRecs
  NEXT
  FOR Cnt = 1 TO NumOfRateRecs
    PRINT #UBRpt, UBRateTblRec(Cnt).RateCode; TAB(12); UBRateTblRec(Cnt).RATEDESC; TAB(61); USING "####"; CustCnt(Cnt)
  NEXT
  
  GOSUB DOCustRateFooter
  
  CLOSE
  
  ERASE UBCustRec
  
  IF NOT AbortFlag THEN
    PrintRptFile "Customer By Rate Code Report.", "UBCBRATE.RPT", 1, RetCode, EntryPoint
  END IF
  
  'KillFile "UBCOLIST.RPT"
  
ExitCustRateReport:
  
  EXIT SUB
  
DOCustRateHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, TAB(TabStop); TownName$
  PRINT #UBRpt, TAB(26); "Customers By Rate Code Report"; TAB(70); "Page #"; PageNo
  PRINT #UBRpt, "Report Date: "; DATE$;
  PRINT #UBRpt, ""
  PRINT #UBRpt, "Rate"; TAB(12); "Description"; TAB(60); "Count"                '; TAB(43); "Service Address"; TAB(72); "Balance"
  PRINT #UBRpt, Dash80$
  LineCnt = 6
  RETURN
  
DOCustRateFooter:
  PRINT #UBRpt, ""
  PRINT #UBRpt, "Total Customers Count: "; USING "####,#"; CustomerCnt
  RETURN
  
END SUB

SUB CustomerCutOffListing
  
  SHARED Choice$()
  REDIM Choice$(0 TO 3, 0 TO 1)
  Choice$(0, 0) = "2"
  Choice$(1, 0) = "Thirty Days or Older"
  Choice$(2, 0) = "Sixty Days or Older"
  Choice$(3, 0) = "Ninety Days or Older"
  
  Choice$(0, 1) = "3"
  Choice$(1, 1) = " Customer Name Order   "
  Choice$(2, 1) = " Account Number Order  "
  Choice$(3, 1) = " Location Number Order "
  
  UsingBook = False
  UsingAcct = False
  UsingName = False
  
  AbortFlag = False
  PageNo = 0
  
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
      LSET Form$(1, 0) = "0"
    END IF
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      MinAmt# = Value#(Form$(1, 0), ECode)
      Days$ = Form$(2, 0)
      Order$ = LEFT$(QPTrim$(Form$(3, 0)), 1)
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
  
  AgeDate = Date2Num%(DATE$)
  
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
  OPEN "UBCOLIST.RPT" FOR OUTPUT AS UBRpt
  
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
  
  
  BlockClear
  ShowProcessingScrn "Customer Cut-Off Report."
  
  GOSUB DoCutOffRptHeader
  
  LongView = INSTR(TownName$, "LONGVIEW")
  Landis = INSTR(TownName$, "LANDIS")
  Tucka = INSTR(TownName$, "TUCK")
  Ellen = INSTR(TownName$, "ELLEN")
  Wades = INSTR(TownName$, "WADES")

  FOR Cnt = 1 TO NumOfRecs
    IF UsingName OR UsingBook THEN
      AcctNo& = IdxBuff(Cnt).RecNum
    ELSE
      AcctNo& = Cnt
    END IF

    'IF AcctNo& = 841 THEN STOP

    GET UBCust, AcctNo&, UBCustRec(1)

'special mod for ellenboro
    IF Wades THEN
      TotBalance# = UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance
      AcctBalance# = UBCustRec(1).PrevBalance
      IF AcctBalance# > 0 THEN
        IF AcctBalance# >= MinAmt# THEN
          IF TotBalance# > 0 THEN
            IF UBCustRec(1).Status = "A" THEN
              GOSUB printline
            END IF
          END IF
        END IF
      END IF
      GOTO skip2here
    END IF
    'goto
'end of ellenboro mod

'special mod for ellenboro
    IF Ellen THEN
      AcctBalance# = UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance
      IF AcctBalance# > 0 THEN
        IF AcctBalance# >= MinAmt# THEN
          GOSUB printline
        END IF
      END IF
      GOTO skip2here
    END IF
'end of ellenboro mod

    'IF AcctNo& = 403 THEN STOP

    IF UBCustRec(1).Status = "A" THEN
      ' Calc Balance
      GOSUB CutOffCalcBal

'not sure about this
'      IF Landis OR Tucka THEN
        AcctBalance# = UBCustRec(1).PrevBalance
'      ELSE
'        AcctBalance# = UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance
'      END IF

      TotBalance# = UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance

      'Check if Town of LongView if so remove taps rev's 5 & 6
      
      IF LongView > 0 THEN
        AcctBalance# = AcctBalance# - (UBCustRec(1).CurrRevAmts(5) + UBCustRec(1).CurrRevAmts(6))
      END IF
      Days$ = LEFT$(Days$, 1)
      
      IF Days$ = "T" AND (Amt1# + Amt2# + Amt3# + Amt4#) > MinAmt# AND AcctBalance# > 0 THEN
        GOSUB printline
      END IF
      IF Days$ = "S" AND (Amt3# + Amt4#) > MinAmt# AND AcctBalance# > 0 THEN
        GOSUB printline
      END IF
      IF Days$ = "N" AND Amt4# > MinAmt# AND AcctBalance# > 0 THEN
        GOSUB printline
      END IF
      
skip2here:
      IF AskAbandonPrint% THEN
        AbortFlag = True
        EXIT FOR
      END IF
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
  PRINT #UBRpt, "Report Date: "; DATE$;
  PRINT #UBRpt, ""
  PRINT #UBRpt, "Location  ActNo. Customer Name"; TAB(43); "Service Address"; TAB(74); "Balance"
  PRINT #UBRpt, "Meter Number"; TAB(17); "Last Read"; TAB(30); "Reading at Cut-Off"
  PRINT #UBRpt, Dash80$
  LineCnt = 6
  RETURN
  
DoCutOffRptFooter:
  PRINT #UBRpt, ""
  PRINT #UBRpt, "Total Customers to Cut Off: "; USING "####,#"; CustomerCnt
  RETURN
  
ShowErrScrn:
  
  REDIM TempScrn(0)
  SaveScrn TempScrn()
  CursorOff
  WazzWind 13, 27, 19, 63, 78, 1, True
  QPrintRC "Invalid Printing Order!", 15, 34, 79
  QPrintRC "Press Any Key To Continue.", 17, 33, 79
  WaitForAction
  RestScrn TempScrn()
  Frm(1).FldNo = 2
  Action = 1
  RETURN
  
CutOffCalcBal:
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
    IF UBTransRec(1).TransDate <= AgeDate THEN
      IF AgeDate - UBTransRec(1).TransDate > 120 THEN
        IF UBTransRec(1).RunBalance = 0 THEN EXIT DO
      END IF
      GOSUB CutOffLineCalc
    END IF
    Trans& = UBTransRec(1).PrevTrans
  LOOP
  
  GOSUB CutOffFinishCalc
  RETURN
  
CutOffFinishCalc:
  IF TotalPayments# > Amt4# THEN
    TotalPayments# = TotalPayments# - Amt4#
    Amt4# = 0
  ELSE
    Amt4# = Amt4# - TotalPayments#
    TotalPayments# = 0
  END IF
  IF TotalPayments# > Amt3# THEN
    TotalPayments# = TotalPayments# - Amt3#
    Amt3# = 0
  ELSE
    Amt3# = Amt3# - TotalPayments#
    TotalPayments# = 0
  END IF
  IF TotalPayments# > Amt2# THEN
    TotalPayments# = TotalPayments# - Amt2#
    Amt2# = 0
  ELSE
    Amt2# = Amt2# - TotalPayments#
    TotalPayments# = 0
  END IF
  Amt1# = Amt1# - TotalPayments#
  LineTotal# = Amt1# + Amt2# + Amt3# + Amt4#
  RETURN
  
  
CutOffLineCalc:
  
  IF UBTransRec(1).TransType = 107 OR UBTransRec(1).TransType = 7 THEN
    RETURN
  END IF
  IF UBTransRec(1).TransType = 4 OR UBTransRec(1).TransType = 5 OR UBTransRec(1).TransType = 8 OR UBTransRec(1).TransType = 12 OR UBTransRec(1).TransType = 104 OR UBTransRec(1).TransType = 105 OR UBTransRec(1).TransType = 109 THEN
    TotalPayments# = TotalPayments# + UBTransRec(1).TransAmt
    TotalPayments# = Round#(TotalPayments#)
    RETURN
  END IF
  
  SELECT CASE UBTransRec(1).TransType
  CASE 1, 2, 3, 6, 11, 10, 101, 102, 103, 112
    Days& = AgeDate - UBTransRec(1).TransDate
    IF Days& <= 30 THEN
      Amt1# = Amt1# + UBTransRec(1).TransAmt
      Amt1# = Round#(Amt1#)
      'RETURN
    END IF
    IF Days& > 30 AND Days& <= 60 THEN
      Amt2# = Amt2# + UBTransRec(1).TransAmt
      Amt2# = Round#(Amt2#)
      'RETURN
    END IF
    IF Days& > 60 AND Days& <= 90 THEN
      Amt3# = Amt3# + UBTransRec(1).TransAmt
      Amt3# = Round#(Amt3#)
      'RETURN
    END IF
    IF Days& > 90 THEN
      Amt4# = Amt4# + UBTransRec(1).TransAmt
      Amt4# = Round#(Amt4#)
      'RETURN
    END IF
  END SELECT
  'STOP
  
  RETURN

printline:
  IF LineCnt > MaxLines THEN
    PRINT #UBRpt, CHR$(12)
    GOSUB DoCutOffRptHeader
  END IF
  '*************************************
  '   Main body of Printing goes here
  PRINT #UBRpt, UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb; USING "######"; AcctNo&;
  PRINT #UBRpt, TAB(18); LEFT$(UBCustRec(1).CustName, 25); LEFT$(UBCustRec(1).SERVADDR, 22);
  PRINT #UBRpt, TAB(70); USING "$$####,#.##"; TotBalance#
  LineCnt = LineCnt + 1
  FOR MrtCnt = 1 TO 7
    IF QPTrim$(UBCustRec(1).LocMeters(MrtCnt).MtrType) <> "" THEN
      PRINT #UBRpt, UBCustRec(1).LocMeters(MrtCnt).MtrNum;
      PRINT #UBRpt, TAB(17); UBCustRec(1).LocMeters(MrtCnt).PrevRead;
      PRINT #UBRpt, TAB(30); "_________________"
      LineCnt = LineCnt + 1
    END IF
  NEXT
  PRINT #UBRpt, STRING$(79, "-"): LineCnt = LineCnt + 1
  IF Landis THEN
    PRINT #UBRpt,
    LineCnt = LineCnt + 2
  END IF
  CustomerCnt = CustomerCnt + 1
  '*************************************
  RETURN
  
END SUB

SUB CustomerFlatRates

  SHARED Choice$()
  REDIM Choice$(0 TO 3, 0 TO 1)

  Choice$(0, 1) = "2"
  Choice$(1, 1) = " Customer Name Order   "
  Choice$(2, 1) = " Account Number Order  "
  Choice$(3, 1) = " Location Number Order "

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

  FirstTime = True

  BlockClear
  DisplayUBScrn ScrnName$

  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      FirstTime = False
      Action = 1
      LSET Form$(1, 0) = "1"
    END IF

    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      RevSource = Value#(Form$(1, 0), ECode)
      Days$ = Form$(2, 0)
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
  
  'IF ExitFlag THEN GOTO ExitCustomerFlat

  '***************
  MaxLines = 55
  PageNo = 0
  Dash80$ = STRING$(80, "-")

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  '  REDIM UBLocaRec(1) AS UBLocationRecType
  '  UBLocaRecLen = LEN(UBLocaRec(1))

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

  '  UBLoca = FREEFILE
  '  OPEN "UBLOCA.DAT" FOR RANDOM SHARED AS UBLoca LEN = UBLocaRecLen

  UBRpt = FREEFILE
  OPEN "UBFRLIST.RPT" FOR OUTPUT AS UBRpt

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


  BlockClear
  ShowProcessingScrn "Flat Rate Listing."

  GOSUB DoMastFlatHeader

  FOR Cnt = 1 TO NumOfRecs
    IF UsingName OR UsingBook THEN
      GET UBCust, IdxBuff(Cnt).RecNum, UBCustRec(1)
      Acctnumber = IdxBuff(Cnt).RecNum
    ELSE
      GET UBCust, Cnt, UBCustRec(1)
      Acctnumber = Cnt
    END IF

    'LocationNumber$ = UBCustRec(1).Book + "-" + UBCustRec(1).SeqNumb
    'LocationNumber$ = QPTrim$(LocationNumber$)

    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, CHR$(12)
      GOSUB DoMastFlatHeader
    END IF

    '*************************************
    '   Main body of Printing goes here
    IF UBCustRec(1).DelFlag <> -1 AND UBCustRec(1).Status = "A" THEN
      PrnCnt = 0
      DidOne = False
      FOR FRCnt = 1 TO 4
        IF UBCustRec(1).FlatRates(FRCnt).RevSrc = RevSource THEN
          DidOne = True
          PRINT #UBRpt, USING "Acct:#####"; Acctnumber;
          PRINT #UBRpt, TAB(15); LEFT$(UBCustRec(1).CustName, 30);
          CustFlat# = Round#(UBCustRec(1).FlatRates(FRCnt).FRAmt * UBCustRec(1).FlatRates(FRCnt).NumMin)
          PRINT #UBRpt, TAB(45); QPTrim$(UBCustRec(1).FlatRates(FRCnt).FRDesc); TAB(65); USING "#####.##"; CustFlat#
          TotalFlat# = Round#(TotalFlat# + CustFlat#)
          PrnCnt = PrnCnt + 1
         EXIT FOR
        END IF
      NEXT

      IF DidOne THEN
        PrnCnt = PrnCnt + 1
        PRINT #UBRpt, Dash80$
        LineCnt = LineCnt + PrnCnt
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

SUB CustomerInquiry
  
  BlockClear
  LookUp RecNo&, "Customer Inquiry", 2, True, False
  IF RecNo& = 0 THEN
    GOTO ExitInquiry
  END IF
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
MainBody:
  CustRec& = RecNo&
  
  LibName$ = "UB"
  ScrnName$ = "UBCUSINQ"
  
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
  
  GOSUB LoadCustInfo
  
  Action = 1
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F2Key
      
    CASE F4KEY  'Customer History
      IF CustRec& > 0 THEN
        ShowCustHistory CustRec&
        Action = 2
      END IF
    CASE F6KEY
      ShowCustConsHist CustRec&
      Action = 2
    CASE F3Key
      ShowWrkOrdHistory RecNo&
      Action = 1
    CASE F7KEY
      CustMessageSystem CustRec&
      Action = 1
      GOTO MainBody
    CASE F8KEY
      'CustEditMeterEST CustRec&
      Action = 1
      GOTO MainBody
    CASE F10Key
      Done = True
    CASE ESC
      EXIT SUB
    CASE ELSE
      Done = False
    END SELECT
    
  LOOP UNTIL Done
  
ExitInquiry:
  EXIT SUB
  '***************
  
LoadCustInfo:
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  GET UBCust, CustRec&, UBCustRec(1)
  CLOSE
  
  Form$(1, 0) = STR$(CustRec&)
  Form$(2, 0) = UBCustRec(1).Book
  Form$(3, 0) = UBCustRec(1).SeqNumb
  Form$(4, 0) = UBCustRec(1).Status
  Form$(5, 0) = UBCustRec(1).CustName
  Form$(6, 0) = UBCustRec(1).ADDR1
  Form$(7, 0) = UBCustRec(1).ADDR2
  Form$(8, 0) = UBCustRec(1).city
  Form$(9, 0) = UBCustRec(1).State
  Form$(10, 0) = UBCustRec(1).ZipCode
  Form$(11, 0) = UBCustRec(1).HPHONE
  Form$(12, 0) = Num2Date$(UBCustRec(1).OPENDATE)
  Form$(13, 0) = UBCustRec(1).CUTOFFYN
  Soc$ = UBCustRec(1).SOSEC
  Form$(14, 0) = LEFT$(Soc$, 3) + "-" + MID$(Soc$, 4, 2) + "-" + MID$(Soc$, 6)
  Form$(15, 0) = UBCustRec(1).DRVLIC
  Form$(16, 0) = UBCustRec(1).WPHONE
  Form$(17, 0) = UBCustRec(1).TAXEXPT
  Form$(18, 0) = UBCustRec(1).Custtype
  Form$(19, 0) = STR$(UBCustRec(1).BILLCOPY)
  Form$(20, 0) = UBCustRec(1).USEDRAFT
  Form$(21, 0) = UBCustRec(1).BANKNAME
  Form$(22, 0) = UBCustRec(1).BANKLOC
  Form$(23, 0) = UBCustRec(1).TRANSIT
  Form$(24, 0) = UBCustRec(1).BANKACCT
  Form$(25, 0) = STR$((Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)))
  Form$(26, 0) = STR$(Round#(UBCustRec(1).CurrBalance))
  Form$(27, 0) = STR$(Round#(UBCustRec(1).PrevBalance))
  Form$(28, 0) = STR$(Round#(UBCustRec(1).DepositAmt))
  RETURN
END SUB

SUB CustWorkOrders (CustRec&)
  
END SUB

SUB DetailedTransJournal
  
  SHARED Choice$()
  
  REDIM TempScrn(0)
  REDIM Choice$(0 TO 14, 0 TO 1)
  REDIM RevenueName$(15)
  REDIM MChoice$(3)
  
  MChoice$(1) = " Customer Name Order   "
  MChoice$(2) = " Account Number Order  "
  MChoice$(3) = " Location Number Order "
  
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
  Col = ((80 - MaxLen) \ 2) - 2
  'Help$ = "Utility Billing Main Menu"
  
  UsingBook = False
  UsingAcct = False
  UsingName = False
  
  AbortFlag = False
  PageNo = 0
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    TitleBox 3, Col, MaxLen + 3, "Detailed Journal Report ", Cnf
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
  
  IF ExitFlag THEN GOTO ExitDetailedListing
  '*******************************
  
  'Select What They Want From the List
SelectFromList:
  LibName$ = "UB"
  ScrnName$ = "UBTRJOU"
  
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
  
  Choice$(0, 0) = "3"
  Choice$(1, 0) = " 0 - All Transactions"
  Choice$(2, 0) = " 1 - Utility Bill"
  Choice$(3, 0) = " 2 - Late Charge"
  Choice$(4, 0) = " 3 - Reconnect Charge"
  Choice$(5, 0) = " 4 - Payment"
  Choice$(6, 0) = " 5 - Applied Deposit"
  Choice$(7, 0) = " 6 - Penalty Charge"
  Choice$(8, 0) = " 7 - Deposit Payment"
  Choice$(9, 0) = " 8 - Draft Payment"
  Choice$(10, 0) = " 9 - Refunded Deposit"
  Choice$(11, 0) = "10 - Beginning Balance"
  Choice$(12, 0) = "11 - Up Adjustment"
  Choice$(13, 0) = "12 - Down Adjustment"
  Choice$(14, 0) = "99 - Misc Payment"
  
  ' Set Defaults as Anticpated Response
  Form$(1, 0) = DATE$
  Form$(2, 0) = DATE$
  Form$(4, 0) = "0"
  Form$(5, 0) = "N"
  Action = 1
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      BegDate$ = Form$(1, 0)
      EndDate$ = Form$(2, 0)
      IF VAL(LEFT$(BegDate$, 2)) < 1 OR VAL(LEFT$(BegDate$, 2)) > 12 THEN
        Done = False: Action = 1: Frm(1).FldNo = 1
      ELSEIF VAL(MID$(BegDate$, 4, 2)) < 1 OR VAL(MID$(BegDate$, 4, 2)) > 31 THEN
        Done = False: Action = 1: Frm(1).FldNo = 1
      ELSEIF VAL(LEFT$(EndDate$, 2)) < 1 OR VAL(LEFT$(EndDate$, 2)) > 12 THEN
        Done = False: Action = 1: Frm(1).FldNo = 1
      ELSEIF VAL(MID$(EndDate$, 4, 2)) < 1 OR VAL(MID$(EndDate$, 4, 2)) > 31 THEN
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
  'Check Date Range for Valid Dates
  ' Set Up Specifications from Input Screen
  BegDate = Date2Num%(BegDate$)
  EndDate = Date2Num%(EndDate$)
  
  IF EndDate < BegDate THEN
    PRINT CHR$(7);
    SaveScrn TempScrn()
    LOCATE , , 0
    DisplayUBScrn "UBBADDAT"
    WaitForAction
    RestScrn TempScrn()
    LOCATE , , 1
    Frm(1).FldNo = 1
    Action = 1
    GOTO SelectFromList
  END IF
  
  '  BegDate = 0
  '  EndDate = BegDate
  
  TrType$ = QPTrim$(LEFT$(Form$(3, 0), 2))
  
  TrType = VAL(TrType$)
  
  IF TrType = 0 THEN
    BegTrans = 1: EndTrans = 999
  ELSE
    BegTrans = TrType: EndTrans = TrType
  END IF
  
  Operator$ = Form$(4, 0)
  Operator = VAL(Operator$)
  IF Operator = 0 THEN
    BegOperator = 0: EndOperator = 99
  ELSE
    BegOperator = Operator: EndOperator = Operator
  END IF
  
  Detail$ = Form$(5, 0)
  
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
  OPEN "UBTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBTrans LEN = UBTransRecLen
  
  UBRpt = FREEFILE
  OPEN "UBDJLIST.RPT" FOR OUTPUT AS UBRpt
  
  UBSetup = FREEFILE
  OPEN "UBSETUP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBSetup LEN = UBSetUpRecLen
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
  
  BlockClear
  ShowProcessingScrn "Detailed Journal Report."
  
  GOSUB DoDetailedRptHeader
  
  FOR Cnt = 1 TO NumOfRecs
    IF UsingName OR UsingBook THEN
      GET UBCust, IdxBuff(Cnt).RecNum, UBCustRec(1)
    ELSE
      GET UBCust, Cnt, UBCustRec(1)
    END IF
    
    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, CHR$(12)
      GOSUB DoDetailedRptHeader
    END IF
    
    '*************************************
    '   Main Body of Printing goes here
    BadCount = 0
    Trans& = UBCustRec(1).LastTrans
    DO WHILE Trans& <> 0
      GET UBTrans, Trans&, UBTrans(1)
      'Check Date, Operator and Trans Type
      IF UBTrans(1).TransDate < BegDate THEN
        BadCount = BadCount + 1
        IF BadCount > 125 THEN
          EXIT DO
        END IF
      END IF
      IF (UBTrans(1).TransDate >= BegDate AND UBTrans(1).TransDate <= EndDate) THEN
        IF (UBTrans(1).OperatorNumber >= BegOperator AND UBTrans(1).OperatorNumber <= EndOperator) THEN
          IF (UBTrans(1).TransType >= BegTrans AND UBTrans(1).TransType <= EndTrans OR (UBTrans(1).TransType >= BegTrans + 100 AND UBTrans(1).TransType <= EndTrans + 100)) THEN
            GOSUB DefineType
            PRINT #UBRpt, Num2Date$(UBTrans(1).TransDate); TAB(11); USING "#####"; UBTrans(1).CustAcctNo;
            PRINT #UBRpt, TAB(20); LEFT$(UBCustRec(1).CustName, 33);
            PRINT #UBRpt, TAB(55); TransDesc$;
            PRINT #UBRpt, TAB(80); Trans&;   'UBTrans(1).OperatorNumber;
            PRINT #UBRpt, TAB(90); LEFT$(UBTrans(1).TransDesc, 20);
            PRINT #UBRpt, TAB(110); USING "$$#####,#.##"; amount#
            'LPRINT UBTrans(1).CustAcctNo, Trans&
            LineCnt = LineCnt + 1
            TotalTrans# = TotalTrans# + amount#
            TransCnt! = TransCnt! + 1
            
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
              PRINT #UBRpt, CHR$(12)
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
    ShowPctComp Cnt, NumOfRecs
  NEXT
  
  GOSUB DoDetailedRptFooter
  PRINT #UBRpt, CHR$(12);
  
  CLOSE UBCust, UBRpt, UBTrans
  
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
  PRINT #UBRpt, ""
  PRINT #UBRpt, "  Date"; TAB(11); "Acct #"; TAB(20); "Customer Name"; TAB(55); "Description"; TAB(80); "Oper#"; TAB(90); "Trans Type"; TAB(113); "   Amount"
  PRINT #UBRpt, Dash120$
  LineCnt = 8
  RETURN
  
DoDetailedRptFooter:
  PRINT #UBRpt, Dash120$
  PRINT #UBRpt, "Transactions: "; TransCnt!; "                                                       Total of Transactions: "; USING "$$#######,#.##"; TotalTrans#
  PRINT #UBRpt, CHR$(12)
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

  IF UBTrans(1).TransType = 1 OR UBTrans(1).TransType = 101 THEN
    TransDesc$ = "Util Bill"
    amount# = UBTrans(1).TransAmt
    RETURN
  END IF
  IF UBTrans(1).TransType = 2 OR UBTrans(1).TransType = 102 THEN
    TransDesc$ = "Late Charge"
    amount# = UBTrans(1).TransAmt
    RETURN
  END IF
  IF UBTrans(1).TransType = 3 THEN
    TransDesc$ = "Reconnect"
    amount# = UBTrans(1).TransAmt
    RETURN
  END IF

  IF UBTrans(1).TransType = 4 OR UBTrans(1).TransType = 104 THEN
    TransDesc$ = "Reg Payment"
    amount# = UBTrans(1).TransAmt * -1
    RETURN
  END IF
  
  IF UBTrans(1).TransType = 5 OR UBTrans(1).TransType = 105 THEN
    TransDesc$ = "Applied Dep"
    amount# = UBTrans(1).TransAmt * -1
    RETURN
  END IF

  IF UBTrans(1).TransType = 9 OR UBTrans(1).TransType = 109 THEN
    TransDesc$ = "Refunded Dep"
    amount# = ABS(UBTrans(1).TransAmt) * -1
    RETURN
  END IF

  IF UBTrans(1).TransType = 6 THEN
    TransDesc$ = "Penalty Chg"
    amount# = UBTrans(1).TransAmt * -1
    RETURN
  END IF
  
  IF UBTrans(1).TransType = 7 OR UBTrans(1).TransType = 107 THEN
    TransDesc$ = "Dep. Payment"
    amount# = UBTrans(1).TransAmt * -1
    RETURN
  END IF

  IF UBTrans(1).TransType = 8 THEN
    TransDesc$ = "Draft Paymt"
    amount# = UBTrans(1).TransAmt * -1
    RETURN
  END IF

  IF UBTrans(1).TransType = 10 OR UBTrans(1).TransType = 110 THEN
    TransDesc$ = "Beg Balance"
    amount# = UBTrans(1).TransAmt
    RETURN
  END IF

  IF UBTrans(1).TransType = 11 OR UBTrans(1).TransType = 111 THEN
    TransDesc$ = UBTrans(1).BillMsg
    amount# = UBTrans(1).TransAmt
    RETURN
  END IF

  IF UBTrans(1).TransType = 12 OR UBTrans(1).TransType = 112 THEN
    TransDesc$ = UBTrans(1).BillMsg
    amount# = UBTrans(1).TransAmt * -1
    RETURN
  END IF

  IF UBTrans(1).TransType = 99 THEN
    TransDesc$ = "Misc Payment"
    amount# = UBTrans(1).TransAmt * -1
    RETURN
  END IF

  TransDesc$ = "UNKNOWN"
  amount# = UBTrans(1).TransAmt
  
  
  RETURN
  
END SUB

SUB MailingLabel
  
  SHARED Choice$()
  REDIM Choice$(0 TO 4, 0 TO 1)
  
  Choice$(0, 0) = "3"
  Choice$(1, 0) = " Customer Name Order   "
  Choice$(2, 0) = " Account Number Order  "
  Choice$(3, 0) = " Location Number Order "
  Choice$(4, 0) = " Service Address Order "
  
  UsingBook = False
  UsingAcct = False
  UsingName = False
  UsingAddr = False
  
  AbortFlag = False
  PageNo = 0
  
  LibName$ = "UBWRKORD"
  ScrnName$ = "UBPRNLAB"
  
  ' Define Fields
  
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  '--Set screen number to one and display screen
  Scr = 1
  
  BlockClear
  
  LibFile2Scrn LibName$, ScrnName$, MonoCode, Attribute%, ErrCode
  
  ShowCursor
  
  FirstTime = True
  
  
  Action = 1
  
  FirstTime = True
  
  'BlockClear
  
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      FirstTime = False
      LSET Form$(1, 0) = "1"
      LSET Form$(2, 0) = "99"
      LSET Form$(5, 0) = "N"
      Action = 1
    END IF
    
    
    SELECT CASE Frm(1).KeyCode
    CASE F5Key
      GOSUB DoAlignLabelMask
    CASE F10Key
      Order$ = LEFT$(QPTrim$(Form$(3, 0)), 1)
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
        IndexName$ = TempIndexName
        UsingAddr = True
        OKFlag = True
      CASE ELSE
        GOSUB ShowErrScrn1
      END SELECT
    CASE ESC
      GOTO ExitMailLabListing
    CASE ELSE
      Done = False
    END SELECT
    
  LOOP UNTIL OKFlag
  
  '***************
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBSetup(1) AS UBSetupRecType
  UBSetUpRecLen = LEN(UBSetup(1))
  
  CustomerType$ = QPTrim$(Form$(4, 0))
  
  FirstBook = Value#(Form$(1, 0), ECode)
  LastBook = Value#(Form$(2, 0), ECode)
  
  IF Form$(5, 0) = "Y" THEN
    IncInactive = True
  END IF
  
  ERASE Frm, Form$, Fld, Choice$
  
  IF UsingAddr THEN
    SortServAddress "Service Address"
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize&(IndexName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
    FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
    NumOfRecs = IdxNumOfRecs
  ELSEIF UsingName OR UsingBook OR UsingAddr THEN
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
  OPEN "UBLABEL.RPT" FOR OUTPUT AS UBRpt
  
  BlockClear
  ShowProcessingScrn "Mailing Labels"
  
  FOR Cnt = 1 TO NumOfRecs
    IF UsingName OR UsingBook OR UsingAddr THEN
      Acctnumber& = IdxBuff(Cnt).RecNum
    ELSE
      Acctnumber& = Cnt
    END IF
    
    GET UBCust, Acctnumber&, UBCustRec(1)
    
    IF UBCustRec(1).DelFlag = True THEN
      GOTO NextLabel
    END IF
    
    CustBook = Value#(UBCustRec(1).Book, ECode)
    
    IF CustBook < FirstBook OR CustBook > LastBook THEN
      GOTO NextLabel
    END IF
    
    IF (UBCustRec(1).Status = "I") AND (IncInactive = 0) THEN
      GOTO NextLabel
    END IF
    
    IF LEN(CustomerType$) > 0 THEN
      IF UCASE$(QPTrim$(UBCustRec(1).Custtype)) <> UCASE$(CustomerType$) THEN
        GOTO NextLabel
      END IF
    END IF
    
    IF RTRIM$(UBCustRec(1).CustName) = "VACANT" THEN
      GOTO NextLabel
    END IF
    
    PRINT #UBRpt, "Cust #"; Acctnumber&
    PRINT #UBRpt, UBCustRec(1).CustName
    PRINT #UBRpt, UBCustRec(1).ADDR1
    
    IF LEN(QPTrim$(UBCustRec(1).ADDR2)) > 0 THEN
      PRINT #UBRpt, UBCustRec(1).ADDR2
      PRINT #UBRpt, RTRIM$(UBCustRec(1).city) + ", " + UBCustRec(1).State + "  " + UBCustRec(1).ZipCode
    ELSE
      PRINT #UBRpt, RTRIM$(UBCustRec(1).city) + ", " + UBCustRec(1).State + "  " + UBCustRec(1).ZipCode
      PRINT #UBRpt, ""
    END IF
    PRINT #UBRpt,
    
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
    
NextLabel:
    ShowPctComp Cnt, NumOfRecs
    
  NEXT
  
  PRINT #UBRpt, CHR$(12);
  CLOSE
  
  ERASE IdxBuff, UBCustRec
  
  IF NOT AbortFlag THEN
    PrintRptFile "Mailing Labels", "UBLABEL.RPT", 1, RetCode, EntryPoint
  END IF
  
  KillFile "UBLABEL.RPT"
  
ExitMailLabListing:
  
  EXIT SUB
  
DoAlignLabelMask:
  LPRINT STRING$(35, "X")
  LPRINT STRING$(35, "X")
  LPRINT STRING$(35, "X")
  LPRINT STRING$(35, "X")
  LPRINT STRING$(35, "X")
  LPRINT
  RETURN
  
ShowErrScrn1:
  
  REDIM TempScrn(0)
  SaveScrn TempScrn()
  CursorOff
  WazzWind 13, 27, 19, 63, 78, 1, True
  QPrintRC "Invalid Printing Order!", 15, 34, 79
  QPrintRC "Press Any Key To Continue.", 17, 33, 79
  WaitForAction
  RestScrn TempScrn()
  ERASE TempScrn
  Frm(1).FldNo = 3
  Action = 1
  RETURN
  
END SUB

SUB MasterBalanceListing
  
  ShowDetail = True
  
  REDIM MChoice$(1 TO 3)
  
  MChoice$(1) = " Customer Name Order   "
  MChoice$(2) = " Location Number Order "
  MChoice$(3) = " Account Number Order  "
  
  
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
  
ReStart:
  '--Center Menu within Screen
  Row = ((25 - (UBOUND(MChoice$))) \ 2) - 1
  Col = ((80 - MaxLen) \ 2) - 2
  'Help$ = "Utility Billing Main Menu"
  
  UsingBook = False
  UsingAcct = False
  UsingName = False
  
  AbortFlag = False
  PageNo = 0
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    TitleBox 3, Col, MaxLen + 3, "Balance Listing Report  ", Cnf
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
      IndexName$ = BookIndexFile
      UsingBook = True
      OKFlag = True
    CASE 3
      IndexName$ = ""
      UsingAcct = True
      OKFlag = True
    END SELECT
    
  LOOP UNTIL OKFlag OR ExitFlag
  
  IF ExitFlag THEN GOTO ExitBalanceListing
  
  '***************
  MaxLines = 55
  PageNo = 0
  Dash80$ = STRING$(80, "-")
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBSetup(1) AS UBSetupRecType
  UBSetUpRecLen = LEN(UBSetup(1))
  
  LoadUBSetUpFile UBSetup(), UBSetUpRecLen
  
  REDIM RevenueName(1 TO 15) AS STRING * 10
  FOR RCnt = 1 TO 15
    TRevName$ = QPTrim$(UBSetup(1).Revenues(RCnt).REVNAME)
    IF LEN(TRevName$) > 0 THEN
      RevenueName(RCnt) = TRevName$
    ELSE
      MaxRevenue = RCnt - 1
      EXIT FOR
    END IF
  NEXT
  
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
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBCust LEN = UBCustRecLen
  
  UBRpt = FREEFILE
  OPEN "UBBALIST.RPT" FOR OUTPUT AS UBRpt
  
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
  
  BlockClear
  ShowProcessingScrn "Balance Listing Report."
  
  GOSUB DoCustRptHeader
  
  FOR Cnt = 1 TO NumOfRecs
    IF UsingName OR UsingBook THEN
      Acctnumber = IdxBuff(Cnt).RecNum
    ELSE
      Acctnumber = Cnt
    END IF
    
    GET UBCust, Acctnumber, UBCustRec(1)
    
    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, CHR$(12)
      GOSUB DoCustRptHeader
    END IF
    
    Book$ = QPTrim$(UBCustRec(1).Book)
    SeqNumb$ = QPTrim$(UBCustRec(1).SeqNumb)
    IF LEN(Book$) = 0 THEN
      Book$ = "  "
    END IF
    IF LEN(SeqNumb$) = 0 THEN
      SeqNumb$ = "      "
    END IF
    'IF (Round#(UBCustRec(1).CurrBalance) = 0 AND Round#(UBCustRec(1).PrevBalance) = 0) OR UBCustRec(1).DelFlag < 0 THEN
'110598 Code to fix the above bug
    'If all there rev sources are "0" then skipem
    OKToSkip = True
    FOR RevChk = 1 TO 15
      IF Round#(UBCustRec(1).CurrRevAmts(RevChk)) <> 0 THEN
        OKToSkip = False
        EXIT FOR
      END IF
    NEXT
    IF OKToSkip THEN
      GOTO BSkipEm
    END IF

      REM Do Not Print


      PRINT #UBRpt, USING "#####"; Acctnumber;
      PRINT #UBRpt, TAB(8); Book$; "-"; SeqNumb$;
      PRINT #UBRpt, TAB(18); LEFT$(UBCustRec(1).CustName, 31);
      PRINT #UBRpt, TAB(50); USING "####,#.##"; Round#(UBCustRec(1).CurrBalance);
      PRINT #UBRpt, TAB(61); USING "####,#.##"; Round#(UBCustRec(1).PrevBalance);
      PRINT #UBRpt, TAB(72); USING "####,#.##"; Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
      'FOR zzz = 1 TO 15
      '  IF UBCustRec(1).CurrRevAmts(zzz) < -10000 OR UBCustRec(1).CurrRevAmts(zzz) > 10000 THEN
      '    LPRINT USING "#####"; AcctNumber;
      '    LPRINT , zzz, UBCustRec(1).CurrRevAmts(zzz)
      '    EXIT FOR
      '  END IF
      'NEXT
      'TBal# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
      'IF TBal# < -10000 OR TBal# > 10000 THEN
      '   PRINT USING "#####"; AcctNumber
      'END IF

      'IF UBCustRec(1).PrevBalance > 0 THEN
      '  PrevCnt = PrevCnt + 1
      'END IF

      IF ShowDetail THEN
        GOSUB PrintDetail
      END IF
      TCurrBalance# = Round#(TCurrBalance# + UBCustRec(1).CurrBalance)
      TPrevBalance# = Round#(TPrevBalance# + UBCustRec(1).PrevBalance)
      CustCnt = CustCnt + 1
      LineCnt = LineCnt + 1
BSkipEm:
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
    'BSkipEm:
    ShowPctComp Cnt, NumOfRecs
  NEXT
  
  GOSUB DoCustRptFooter
  
  CLOSE UBCust, UBRpt
  ERASE IdxBuff, UBCustRec
  
  IF NOT AbortFlag THEN
    PrintRptFile "Balance Listing Report.", "UBBALIST.RPT", 1, RetCode, EntryPoint
  END IF
  
  '    KillFile "UBBALIST.RPT"
  
ExitBalanceListing:
  
  EXIT SUB
  
DoCustRptHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, TAB(TabStop); TownName$
  PRINT #UBRpt, TAB(24); "Customer Balance Listing Report"; TAB(70); "Page #"; PageNo
  PRINT #UBRpt, "Report Date: "; DATE$
  PRINT #UBRpt, "Acct #";
  PRINT #UBRpt, TAB(9); "Location"; TAB(18); "Customer Name"; TAB(52); "Cur Bal"; TAB(62); "Past Due"; TAB(73); "Acct Bal"
  PRINT #UBRpt, Dash80$
  LineCnt = 4
  RETURN
  
DoCustRptFooter:
  PRINT #UBRpt, Dash80$
  PRINT #UBRpt, "Totals:"; TAB(10); "Customers: "; USING "#####,#"; CustCnt;
  PRINT #UBRpt, TAB(49); USING "######,#.##"; TCurrBalance#;
  PRINT #UBRpt, TAB(60); USING "######,#.##"; TPrevBalance#;
  PRINT #UBRpt, TAB(71); USING "######,#.##"; (TCurrBalance# + TPrevBalance#)
  PRINT #UBRpt,
  FOR Cnt = 1 TO MaxRevenue
    Detail$ = SPACE$(28)
    LSET Detail$ = RevenueName(Cnt)
    MID$(Detail$, 15) = "#######.##"
    PRINT #UBRpt, USING Detail$; RevTotals(Cnt)
  NEXT
  PRINT #UBRpt, CHR$(12)
  RETURN
  
PrintDetail:
  TCnt = 0
  Detail$ = SPACE$(18)
  First = 1: Last = MaxRevenue
  IF MaxRevenue < Last THEN
    Last = MaxRevenue
  END IF
  FOR RCnt = First TO Last
    TCnt = TCnt + 1
    TabStop = (TCnt * 21) - 20
    IF TabStop > 81 THEN
      TCnt = 1
      TabStop = (TCnt * 21) - 20
    END IF
    LSET Detail$ = RevenueName(RCnt)
    MID$(Detail$, 10) = "#####.##"
    RevTotals(RCnt) = Round#(RevTotals(RCnt) + UBCustRec(1).CurrRevAmts(RCnt))
    PRINT #UBRpt, TAB(TabStop); USING Detail$; UBCustRec(1).CurrRevAmts(RCnt);
  NEXT
  PRINT #UBRpt,
  LineCnt = LineCnt + 1
  
  
DoneDetail:
  PRINT #UBRpt, Dash80$
  LineCnt = LineCnt + 1
  RETURN
  
END SUB

SUB MasterCustomerListing
  
  REDIM MChoice$(1 TO 3)
  
  MChoice$(1) = " Customer Name Order    "
  MChoice$(2) = " Account Number Order   "
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
  
  
  '--Center Menu within Screen
  Row = ((25 - (UBOUND(MChoice$))) \ 2) - 1
  Col = ((80 - MaxLen) \ 2) - 2
  'Help$ = "Utility Billing Main Menu"
  
  UsingBook = False
  UsingAcct = False
  UsingName = False
  
  AbortFlag = False
  PageNo = 0
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    TitleBox 3, Col, MaxLen + 3, "Master Customer Report  ", Cnf
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
  
  IF ExitFlag THEN GOTO ExitMastCustListing
  
  '***************
  MaxLines = 55
  PageNo = 0
  Dash80$ = STRING$(80, "-")
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  '  REDIM UBLocaRec(1) AS UBLocationRecType
  '  UBLocaRecLen = LEN(UBLocaRec(1))
  
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
    TownName$ = UBSetup(1).UTILNAME
    TownLen = LEN(RTRIM$(TownName$))
    TabStop = 40 - (TownLen / 2)
    IF TabStop < 1 THEN TabStop = 1
  END IF
  CLOSE UBSetup
  
  
  BlockClear
  ShowProcessingScrn "Master Customer Report."
  
  GOSUB DoMastCustRptHeader
  
  FOR Cnt = 1 TO NumOfRecs
    IF UsingName OR UsingBook THEN
      GET UBCust, IdxBuff(Cnt).RecNum, UBCustRec(1)
      Acctnumber = IdxBuff(Cnt).RecNum
    ELSE
      GET UBCust, Cnt, UBCustRec(1)
      Acctnumber = Cnt
    END IF
    
    LocationNumber$ = UBCustRec(1).Book + "-" + UBCustRec(1).SeqNumb
    LocationNumber$ = QPTrim$(LocationNumber$)
    
    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, CHR$(12)
      GOSUB DoMastCustRptHeader
    END IF
    
    '*************************************
    '   Main body of Printing goes here
    IF UBCustRec(1).DelFlag <> -1 THEN
      PRINT #UBRpt, "Acct# :"; USING "#####"; Acctnumber;
      PRINT #UBRpt, TAB(15); RTRIM$(UBCustRec(1).CustName);
      PRINT #UBRpt, TAB(52); "Location: "; LocationNumber$
      PRINT #UBRpt, "Status: "; UBCustRec(1).Status; TAB(15); QPTrim$(UBCustRec(1).ADDR1);
      PRINT #UBRpt, TAB(52); "H Phone#: "; QPTrim$(UBCustRec(1).HPHONE)
      PRINT #UBRpt, "Type  : "; QPTrim$(UBCustRec(1).Custtype); TAB(15); QPTrim$(UBCustRec(1).ADDR2);
      PRINT #UBRpt, TAB(52); "W Phone#: "; UBCustRec(1).WPHONE
      PRINT #UBRpt, "Draft : "; QPTrim$(UBCustRec(1).USEDRAFT);
      Zip$ = QPTrim$(UBCustRec(1).ZipCode)
      IF LEN(Zip$) > 5 THEN
        Zip$ = LEFT$(Zip$, 5) + "-" + MID$(Zip$, 6)
      END IF
      PRINT #UBRpt, TAB(15); QPTrim$(UBCustRec(1).city); ", "; QPTrim$(UBCustRec(1).State); " "; Zip$;
      PRINT #UBRpt, TAB(52); "DrvLic #: "; QPTrim$(UBCustRec(1).DRVLIC)
      PRINT #UBRpt, Dash80$
      LineCnt = LineCnt + 5: CustCnt = CustCnt + 1
      '*************************************
      
      IF AskAbandonPrint% THEN
        AbortFlag = True
        EXIT FOR
      END IF
    END IF
    ShowPctComp Cnt, NumOfRecs
  NEXT
  
  GOSUB DoMastCustRptFooter:
  
  CLOSE UBCust, UBRpt
  
  ERASE IdxBuff, UBCustRec
  
  IF NOT AbortFlag THEN
    PrintRptFile "Master Customer Report.", "UBMCLIST.RPT", 1, RetCode, EntryPoint
  END IF
  
  KillFile "UBMCLIST.RPT"
  
ExitMastCustListing:
  
  EXIT SUB
  
DoMastCustRptHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, TAB(TabStop); TownName$
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
  
  'REDIM MChoice$(1 TO 2)
  '
  'MChoice$(1) = " Customer Name Order   "
  'MChoice$(2) = " Account Number Order  "
  '
  REDIM MChoice$(1 TO 3)
  
  MChoice$(1) = " Customer Name Order   "
  MChoice$(2) = " Location Number Order "
  MChoice$(3) = " Account Number Order  "
  
  
  
  
  
  
  
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
  'Help$ = "Utility Billing Main Menu"
  
  UsingBook = False
  UsingAcct = False
  UsingName = False
  
  AbortFlag = False
  PageNo = 0
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    TitleBox 3, Col, MaxLen + 3, "Deposit Listing Report  ", Cnf
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
  OPEN "UBDPLIST.RPT" FOR OUTPUT AS UBRpt
  
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
  
  
  BlockClear
  ShowProcessingScrn "Deposit Listing Report."
  
  GOSUB DoDepositRptHeader
  
  FOR Cnt = 1 TO NumOfRecs
    IF UsingName OR UsingBook THEN
      Acctnumber = IdxBuff(Cnt).RecNum
    ELSE
      Acctnumber = Cnt
    END IF
    
    GET UBCust, Acctnumber, UBCustRec(1)
    
    IF UBCustRec(1).DepositAmt <> 0 THEN
      PRINT #UBRpt, USING "#####"; Acctnumber;
      PRINT #UBRpt, TAB(8); QPTrim$(UBCustRec(1).Book); "-"; QPTrim$(UBCustRec(1).SeqNumb);
      PRINT #UBRpt, TAB(20); LEFT$(UBCustRec(1).CustName, 38);
      PRINT #UBRpt, TAB(61); USING "####,#.##"; UBCustRec(1).DepositAmt;
      PRINT #UBRpt, TAB(78); UBCustRec(1).Status
      LineCnt = LineCnt + 1
      TDeposit# = Round#(TDeposit# + UBCustRec(1).DepositAmt)
      CustCnt = CustCnt + 1
    END IF
    
    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, CHR$(12)
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
  
  KillFile "UBDPLIST.RPT"
  
ExitDepositListing:
  
  EXIT SUB
  
DoDepositRptHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, TAB(TabStop); TownName$
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
  
  PRINT #UBRpt, CHR$(12)
  
  RETURN
  
END SUB

SUB oShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&) STATIC
  RSET PctC(1) = QPStrL$(INT((RecNo& / NumOfRecs&) * 100))
  '  HideCursor
  QPrintRC PctC(1), 13, 40, Cnf.HiLite
  '  ShowCursor
  '  QPrintRC STR$(FRE("")), 25, 1, Cnf.HiLite
END SUB

SUB PaymentSumReport
  
  REDIM TempScrn(0)
  REDIM RevenueName$(15)
  REDIM Revenues(1 TO 15) AS DOUBLE
  
  REDIM UBTrans(1) AS UBTransRecType
  UBTransRecLen = LEN(UBTrans(1))
  
  REDIM UBSetup(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetup(), UBSetUpRecLen
  
  FOR RCnt = 1 TO 15
    TRevName$ = QPTrim$(UBSetup(1).Revenues(RCnt).REVNAME)
    IF LEN(TRevName$) > 0 THEN
      RevenueName$(RCnt) = TRevName$
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
    '*************************************
    '   Main Body of Printing goes here
    TransOK = False
    IF (UBTrans(1).TransDate >= BegDate AND UBTrans(1).TransDate <= EndDate) AND (UBTrans(1).OperatorNumber >= BegOperator AND UBTrans(1).OperatorNumber <= EndOperator) THEN
      SELECT CASE UBTrans(1).TransType
      CASE TranBillPayment, TranBillPayment + 100
        TransOK = True
      CASE TranPenaltyPayment, TranPenaltyPayment + 100
        TransOK = True
      CASE TranDepositPayment, TranDepositPayment + 100
        TransOK = True
        '            CASE TranDraftPayment, TranDraftPayment + 100
        '              TransOK = True
      END SELECT
      
      IF TransOK THEN
        FOR RevCnt = 1 TO 15
          Revenues(RevCnt) = Round#(Revenues(RevCnt) + UBTrans(1).RevAmt(RevCnt))
        NEXT
        TransCnt& = TransCnt& + 1
      END IF
    END IF
    
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
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
  PRINT #UBRpt, TAB(28); "Payment Summary Report"               '; TAB(70); "Page #1"
  PRINT #UBRpt, "   Report Date: "; DATE$
  PRINT #UBRpt, "Beginning Date: "; BegDate$;
  IF VAL(Operator$) = 0 THEN
    PRINT #UBRpt, TAB(65); " Operator #: ALL"
  ELSE
    PRINT #UBRpt, TAB(65); " Operator #: "; Operator$
  END IF
  PRINT #UBRpt, "   Ending Date: "; EndDate$;
  PRINT #UBRpt, ""
  RETURN
  
DoDetailedRptFooter3:
  PRINT #UBRpt, Dash80$
  FOR Cnt = 1 TO MaxRevenue
    PRINT #UBRpt, TAB(5); RevenueName$(Cnt); TAB(35); USING "$$#######,#.##"; Revenues(Cnt)
    TotalTrans# = Round#(TotalTrans# + Revenues(Cnt))
  NEXT
  PRINT #UBRpt, Dash80$
  PRINT #UBRpt, "Total Number Transactions: "; USING "##############"; TransCnt&
  PRINT #UBRpt, "    Total of Transactions: "; USING "$$#######,#.##"; TotalTrans#
  PRINT #UBRpt, CHR$(12)
  RETURN
  
END SUB

SUB SortServAddress (IndexText$)
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
    ServIndex(Cnt).ServiceAddress = UBCustRec(1).SERVADDR
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

  IdxRecLen = 4               'we are using a long integer
  IdxFileSize& = FileSize(IndexName$)
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
  NumOfRecs = IdxNumOfRecs

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen

  UBRpt = FREEFILE
  OPEN "UBWCUCON.RPT" FOR OUTPUT AS UBRpt

  GOSUB DoWestRptHeader

  FOR CCnt = 1 TO NumOfRecs
    GET UBCust, IdxBuff(CCnt).RecNum, UBCustRec(1)
    PCode$ = UCASE$(QPTrim$(UBCustRec(1).PumpCode))
    IF PCode$ = "W" THEN
      GOSUB PrintThisOne
    END IF
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, CHR$(12)
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
  PRINT #UBRpt, UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb; USING "######"; IdxBuff(CCnt).RecNum;
  PRINT #UBRpt, TAB(18); LEFT$(UBCustRec(1).CustName, 25);
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
  PRINT #UBRpt, CHR$(12)
RETURN

END SUB

