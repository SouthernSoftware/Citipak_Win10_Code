DEFINT A-Z

DEFSNG A-Z
DEFINT A-Z
SUB MasterBalanceListing
  
  REDIM MChoice$(1 TO 3)
  
  MChoice$(1) = " Customer Name Order   "
  MChoice$(2) = " Location Number Order "
  MChoice$(3) = " Account Number Order  "
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 17   'limit the box length to go no lower than line 17
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
  
  'UsingBook = False
  'UsingAcct = False
  'UsingName = False
  'AbortFlag = False
  'PageNo = 0
  
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
      GOSUB GetBooks
    CASE 3
      IndexName$ = ""
      UsingAcct = True
      OKFlag = True
    END SELECT
    
  LOOP UNTIL OKFlag OR ExitFlag
  
  IF ExitFlag THEN GOTO ExitBalanceListing
  
'***************
  MaxLines = 55
'  PageNo = 0

  Dash80$ = STRING$(80, "-")
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBSetup(1) AS UBSetupRecType
  UBSetUpRecLen = LEN(UBSetup(1))
  
  LoadUBSetUpFile UBSetup(), UBSetUpRecLen

  TownName$ = QPTrim$(UBSetup(1).UTILNAME)
  TabStop = 40 - (LEN(TownName$) / 2)
  IF TabStop < 1 THEN TabStop = 1
  TownName$ = SPACE$(TabStop) + TownName$
  
  REDIM RevenueName(1 TO 15) AS STRING * 10
  FOR RCnt = 1 TO 15
    TRevName$ = QPTrim$(UBSetup(1).Revenues(RCnt).RevName)
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
  
  BlockClear
  ShowProcessingScrn "Balance Listing Report."
  
  GOSUB DoCustRptHeader
  
  FOR Cnt = 1 TO NumOfRecs
    IF UsingName OR UsingBook THEN
      AcctNumber = IdxBuff(Cnt).RecNum
    ELSE
      AcctNumber = Cnt
    END IF
    
    GET UBCust, AcctNumber, UBCustRec(1)
    
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
    IF UsingBook THEN
      Book% = Value#(Book$, ECode)
      IF Book% < BegRoute OR Book% > EndRoute THEN
        GOTO BSkipEm
      END IF
    END IF

    IF (Round#(UBCustRec(1).CurrBalance) = 0 AND Round#(UBCustRec(1).PrevBalance) = 0) OR UBCustRec(1).DelFlag < 0 THEN
      REM Do Not Print
    ELSE
      PRINT #UBRpt, USING "#####"; AcctNumber;
      PRINT #UBRpt, TAB(8); Book$; "-"; SeqNumb$;
      PRINT #UBRpt, TAB(18); LEFT$(UBCustRec(1).CustName, 31);
      PRINT #UBRpt, TAB(50); USING "####,#.##"; Round#(UBCustRec(1).CurrBalance);
      PRINT #UBRpt, TAB(61); USING "####,#.##"; Round#(UBCustRec(1).PrevBalance);
      PRINT #UBRpt, TAB(72); USING "####,#.##"; Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
      LineCnt = LineCnt + 1

      GOSUB PrintDetail

      TCurrBalance# = Round#(TCurrBalance# + UBCustRec(1).CurrBalance)
      TPrevBalance# = Round#(TPrevBalance# + UBCustRec(1).PrevBalance)
      CustCnt = CustCnt + 1
    END IF
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF

BSkipEm:
    ShowPctComp Cnt, NumOfRecs
  NEXT
  
  GOSUB DoCustRptFooter
  
  CLOSE UBCust, UBRpt
  ERASE IdxBuff, UBCustRec
  
  IF NOT AbortFlag THEN
    PrintRptFile "Balance Listing Report.", "UBBALIST.RPT", 1, RetCode, EntryPoint
  END IF
  
ExitBalanceListing:
  
  EXIT SUB
  
DoCustRptHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, TownName$
  PRINT #UBRpt, TAB(26); "Customer Balance Listing Report"; TAB(70); "Page #"; PageNo
  PRINT #UBRpt, "Report Date: "; DATE$
  PRINT #UBRpt, "Acct #"; TAB(9); "Location"; TAB(18); "Customer Name"; TAB(52); "Cur Bal"; TAB(62); "Past Due"; TAB(73); "Acct Bal"
  PRINT #UBRpt, Dash80$
  LineCnt = 5
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
  First = 1
  Last = MaxRevenue
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
  PRINT #UBRpt, Dash80$
  LineCnt = LineCnt + 2
RETURN

GetBooks:

  LibName$ = "UB"
  ScrnName$ = "UBMTRRPT"

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
  'LineCnt& = 0

  Action = 1
  FirstTime = True
  BlockClear

  DisplayUBScrn ScrnName$
  QPrintRC "Master Balance Report", 8, 24, -1
  QPrintRC " ]", 8, 45, -1

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF FirstTime THEN
      LSET Form$(1, 0) = "0"
      LSET Form$(2, 0) = "99"
      FirstTime = False
      Action = 1
    END IF

    '--Check for Key presses

    SELECT CASE Frm(1).KeyCode

    CASE F10Key
      BegRoute = QPValI(Form$(1, 0))
      EndRoute = QPValI(Form$(2, 0))

      IF BegRoute > EndRoute THEN
        COLOR 12
        LOCATE 13, 25
        PRINT "Invalid Route Order : Please Redo"
        BEEP
        WaitForAction
        LOCATE 13, 25
        PRINT STRING$(35, " ")
        Form$(1, 0) = ""
        Form$(2, 0) = ""
        Frm(1).FldNo = 1
        Action = 1
      ELSE
        inputting = True
      END IF

    CASE EscKey
      EXIT SUB
    END SELECT

  LOOP WHILE inputting = False

RETURN

END SUB

