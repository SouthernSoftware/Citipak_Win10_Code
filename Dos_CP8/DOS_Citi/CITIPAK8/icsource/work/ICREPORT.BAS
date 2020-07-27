DEFINT A-Z
DECLARE SUB DisplayICScrn (ScrnName$)
DECLARE SUB ClearBack ()
DECLARE SUB MasterItemListing ()
DECLARE SUB CustomerListing ()
DECLARE SUB TransactionJournal ()
DECLARE SUB OpenARCustIdxFile (NumOfARIdxRecs%, ARIdxFile%)
DECLARE SUB ARFixMess ()
DECLARE SUB ShowNoCode ()
DECLARE SUB OpenARCustFile (NumOfArRecs, ARFile)
DECLARE SUB printhelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB ShowCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB HideCursor ()
DECLARE SUB QPrint (x$, Colr%, page%)
DECLARE SUB QPrintRC (t$, r%, C%, clr%)
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
  
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'IC.bi'                        'A/R FILE LAYOUTS
  '$INCLUDE: 'GL.bi'
  
  TYPE Struct
    who AS STRING * 14
    RecNum AS INTEGER
  END TYPE
  
  DIM SHARED ItemRec(1) AS ItemRecType
  
  CONST False = 0, True = NOT False
  
  STACK 8000
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 3)
  
  MChoice$(1) = " Item Transaction Journal "
  MChoice$(2) = " Master Item Listing"
  MChoice$(3) = " Exit to OS"
  
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
  Col = ((80 - MaxLen) \ 2)
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    ClearBack
    
    TitleBox 3, Col, MaxLen + 3, " Reports Menu ", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      '          TransactionJournal
    CASE 2
      MasterItemListing
    CASE 3
      END
    END SELECT
  LOOP
  RUN "icmenu"

SUB MasterItemListing
  
  SHARED Choice$()
  ReportFile$ = "ICITEM.PRN"    'Report File Name
  CommaFmt$ = "########,.##"    'format takes 13 chars
  TotalFmt$ = "#########,.##"   'format takes 14 chars
  SumLine$ = STRING$(13, "-")   'column summary line
  DivLine$ = STRING$(77, "-")   'dashed line
  DivLine2$ = STRING$(77, "=")  'Double Line
  FF$ = CHR$(12)
  MaxLines = 53
  LineCnt = 0
  TotDr# = 0
  TotCr# = 0
  size = 2500
  Start = 1     'start at array element 1
  Dir = 0       'sort direction - use anything else for descending
  SSize = 16    'total size of each TYPE element
  MOff = 0      'offset into the TYPE for the key element
  MSize = 16    'size of the key element - coded as follows:
  '   -1 = integer
  '   -2 = long integer
  '   -3 = single precision
  '   -4 = double precision
  '   +N = TYPE array/fixed-length string of length N
  
  REDIM array(1 TO size) AS Struct
  
  GOSUB SelectOutput
  
  
  ItemRecLen = LEN(ItemRec(1))
  ItemFile = FREEFILE
  OPEN "ICITEM.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ItemFile LEN = ItemRecLen
  NumOfItemRecs = LOF(ItemFile) \ ItemRecLen
  
  
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  GOSUB PrintRptHeader
  
  FOR Cnt& = 1 TO NumOfItemRecs
    GET ItemFile, Cnt&, ItemRec(1)
    IF ItemRec(1).Deleted <> "Y" THEN
      IF LineCnt >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintRptHeader
      END IF
      PRINT #RptHandle, USING "#####"; Cnt&;
      PRINT #RptHandle, TAB(10); ItemRec(1).Desc1; TAB(60); "QOH: "; USING "######,#.###"; ItemRec(1).QtyOnHand
      PRINT #RptHandle, TAB(10); ItemRec(1).Desc2; TAB(59); "Dept: "; ItemRec(1).DeptNumb
      PRINT #RptHandle, STRING$(79, "-")
      TotalItems = TotalItems + 1
      LineCnt = LineCnt + 3
    END IF
  NEXT Cnt&
  GOSUB PrintRptEnding
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi
  CLOSE         'Close all open files now
  
  IF DevSpec$ = "S" THEN
    EntryPoint = 2
  ELSE
    EntryPoint = 5
  END IF
  
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  KILL ReportFile$
  
  EXIT SUB
  
  
PrintRptHeader:
  page = page + 1
  PRINT #RptHandle, TAB(25); "Item Inventory Listing"
  PRINT #RptHandle, STRING$(80, "=")
  LineCnt = 5
  RETURN
  
PrintRptEnding:
  PRINT #RptHandle, STRING$(80, "-")
  PRINT #RptHandle, "Number of Items .. "; USING "####,#"; TotalItems
  PRINT #RptHandle, FF$
  RETURN
  
  
  
  
SelectOutput:
  LibName$ = "IC"
  ScrnName$ = "WHERPRNT"
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  REDIM Choice$(2, 0)
  
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "SCREEN"
  Choice$(2, 0) = "PRINTER"
  
  Action = 1
  ClearBack
  DisplayICScrn ScrnName$
  ShowCursor
  
  Action = 1
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      DevSpec$ = LEFT$(Form$(1, 0), 1)
      RETURN
    CASE EscKey
      Canceled$ = "Y"
      RETURN
    END SELECT
  LOOP
  RETURN
  
END SUB

SUB TransactionJournal
  
  SHARED Choice$()
  ReportFile$ = "ICTRANS.PRN"   'Report File Name
  CommaFmt$ = "########,.##"    'format takes 13 chars
  TotalFmt$ = "#########,.##"   'format takes 14 chars
  SumLine$ = STRING$(13, "-")   'column summary line
  DivLine$ = STRING$(77, "-")   'dashed line
  DivLine2$ = STRING$(77, "=")  'Double Line
  LPTPort% = 1
  FF$ = CHR$(12)
  MaxLines = 53
  LineCnt = 0
  TotDr# = 0
  TotCr# = 0
  size = 4000
  Start = 1     'start at array element 1
  Dir = 0       'sort direction - use anything else for descending
  SSize = 16    'total size of each TYPE element
  MOff = 0      'offset into the TYPE for the key element
  MSize = 16    'size of the key element - coded as follows:
  '   -1 = integer
  '   -2 = long integer
  '   -3 = single precision
  '   -4 = double precision
  '   +N = TYPE array/fixed-length string of length N
  
  REDIM array(1 TO size) AS Struct
  REDIM ICTransRec(1) AS ICTransRecType
  
  GOSUB SelectOutput2
  IF Canceled$ = "Y" THEN EXIT SUB
  
  GOSUB GetReportInformation2
  
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  
  'PRINT #RptHandle, CHR$(27); CHR$(58); ' oki 320 12 cpi
  
  GOSUB PrintRptHeader2
  
  FOR Cnt = 1 TO Count
    ' GET ARTRFile, array(Cnt).RecNum, ARTransRec(1)
    
    IF LineCnt >= MaxLines THEN
      PRINT #RptHandle, FF$
      GOSUB PrintRptHeader2
    END IF
    
    'help$ = ARCustRec(1).CustName
    'printhelp help$
    'BillCat$ = QPTrim$(ARCustRec(1).BillCat)
    
    '  PRINT #RptHandle, Num2Date$(ARTransRec(1).TransDate);
    '   PRINT #RptHandle, TAB(13); LEFT$(ARCustRec(1).CustName, 25);
    '    PRINT #RptHandle, TAB(40); "";
    '     IF ARTransRec(1).TransType = 1 THEN PRINT #RptHandle, "Charge";
    '     IF ARTransRec(1).TransType = 2 THEN PRINT #RptHandle, "Payment";
    '     IF ARTransRec(1).TransType = 9 THEN PRINT #RptHandle, "Beg Bal";
    '     PRINT #RptHandle, TAB(50); LEFT$(ARTransRec(1).TransDesc, 18);
    '     PRINT #RptHandle, TAB(69); USING "$$#####,#.##"; ARTransRec(1).TransAmount
    '
    '  CLOSE ARFile
    '  TotalTrans = TotalTrans + 1
    '  LineCnt = LineCnt + 1
    'TotalAmt# = TotalAmt# + ARTransRec(1).TransAmount
    REM total by catagory
    '  IF ARTransRec(1).TransType = 1 THEN
    '   IF ChrgeCat = 0 THEN
    '      ChrgeCat = 1: ChrgeCat$(1) = BillCat$: ChrgeAmt#(1) = ARTransRec(1).TransAmount
    '      GOTO TotalUp
    '   END IF
    '   FOR Snt! = 1 TO ChrgeCat
    '    IF BillCat$ = ChrgeCat$(Snt!) THEN
    '     ChrgeAmt#(Snt!) = ChrgeAmt#(Snt!) + ARTransRec(1).TransAmount
    '     GOTO TotalUp
    '    END IF
    '   NEXT Snt!
    '   ChrgeCat = ChrgeCat + 1
    '   ChrgeCat$(ChrgeCat) = BillCat$: ChrgeAmt#(ChrgeCat) = ARTransRec(1).TransAmount
    '   GOTO TotalUp
    '  END IF
    '  IF ARTransRec(1).TransType = 2 THEN
    '   IF PayCat = 0 THEN
    '      PayCat = 1: PayCat$(1) = BillCat$: PayAmt#(1) = ARTransRec(1).TransAmount
    '      GOTO TotalUp
    '   END IF
    '   FOR Snt! = 1 TO PayCat
    '    IF BillCat$ = PayCat$(Snt!) THEN
    '     PayAmt#(Snt!) = PayAmt#(Snt!) + ARTransRec(1).TransAmount
    '     GOTO TotalUp
    '    END IF
    '   NEXT Snt!
    '   PayCat = PayCat + 1
    '   PayCat$(PayCat) = BillCat$: PayAmt#(PayCat) = ARTransRec(1).TransAmount
    '   GOTO TotalUp
    '
    '  END IF
    '
    '
    '
    
    
TotalUp:
    'TotalAmt#(ARTransRec(1).TransType) = TotalAmt#(ARTransRec(1).TransType) + ARTransRec(1).TransAmount
    
  NEXT Cnt
  
  GOSUB PrintRptEnding2
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi
  CLOSE         'Close all open files now
  
  IF DevSpec$ = "S" THEN
    EntryPoint = 2
  ELSE
    EntryPoint = 5
  END IF
  Header$ = "Transaction Journal"
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  KILL ReportFile$
  
  EXIT SUB
  
  
PrintRptHeader2:
  page = page + 1
  PRINT #RptHandle, TAB(22); "Business License : Transactions Journal"
  PRINT #RptHandle, ""
  PRINT #RptHandle, "               Report Date: "; DATE$; TAB(65); "Page #"; page
  PRINT #RptHandle, "Beginning Transaction Date: "; BegDate$
  PRINT #RptHandle, "   Ending Transaction Date: "; EndDate$
  PRINT #RptHandle, ""
  PRINT #RptHandle, "  Date"; TAB(13); "Customer Name"; TAB(40); "Type"; TAB(50); "Description"; TAB(70); "Amount"
  PRINT #RptHandle, STRING$(80, "=")
  LineCnt = 5
  RETURN
  
PrintRptEnding2:
  PRINT #RptHandle, STRING$(80, "-")
  PRINT #RptHandle, TAB(69); USING "$$#####,#.##"; TotalAmt#
  '    FOR Cnt! = 1 TO 99
  '    IF TotalAmt#(Cnt!) <> 0 THEN
  '     PRINT #RptHandle, "Trans Type : ";
  '     IF Cnt! = 1 THEN PRINT #RptHandle, "Charges  ";
  '     IF Cnt! = 2 THEN PRINT #RptHandle, "Payments ";
  '     IF Cnt! = 9 THEN PRINT #RptHandle, "Beg Bal  ";
  '     PRINT #RptHandle, "     Total Amount: "; USING "$$######,#.##"; TotalAmt#(Cnt!)
  '    END IF
  '    NEXT Cnt!
  '    PRINT #RptHandle, FF$
  '    IF ChrgeCat > 0 THEN
  '     page = page + 1
  '     PRINT #RptHandle, TAB(20); "Business License : Transactions Journal "
  '     PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; page
  '     PRINT #RptHandle, "Total Charges by Catagory"
  '     LCnt = 1
  ' FOR ll = 1 TO ChrgeCat
  '
  '   IF LCnt > 55 THEN
  '    page = page + 1
  '    PRINT #RptHandle, TAB(20); "Business License : Transactions Journal "
  '    PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; page
  '    PRINT #RptHandle, "Total Charges by Catagory"
  '    LCnt = 3
  '   END IF
  
  
  
  
  'PRINT #RptHandle, ChrgeCat$(ll); TAB(10); CatagoryDesc$; TAB(50); USING "$$#######,#.##"; ChrgeAmt#(ll)
  ' LCnt = LCnt + 1
  '  NEXT ll
  '  PRINT #RptHandle, FF$
  '  END IF
  IF PayCat > 0 THEN
    page = page + 1
    PRINT #RptHandle, TAB(20); "Business License : Transactions Journal "
    PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; page
    PRINT #RptHandle, "Total Payments by Catagory"
    LCnt = 1
    FOR ll = 1 TO PayCat
      
      IF LCnt > 55 THEN
        page = page + 1
        PRINT #RptHandle, TAB(20); "Business License : Transactions Journal "
        PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; page
        PRINT #RptHandle, "Total Payments by Catagory"
        LCnt = 3
      END IF
      
      
      
      'Get Catagory Desc First
      CatagoryDesc$ = ""
      
      'PRINT #RptHandle, PayCat$(ll); TAB(10); CatagoryDesc$; TAB(50); USING "$$#######,#.##"; PayAmt#(ll)
      LCnt = LCnt + 1
    NEXT ll
    PRINT #RptHandle, FF$
  END IF
  
  RETURN
  
GetReportInformation2:
  'ARTRFile = FREEFILE
  'OPEN "ARTrans.Dat" FOR RANDOM ACCESS READ WRITE SHARED AS #ARTRFile LEN = LEN(ARTransRec(1))
  'TrNumRecs = LOF(ARTRFile) / LEN(ARTransRec(1))
  
  FOR Cnt! = 1 TO TrNumRecs
    'GET ARTRFile, Cnt!, ARTransRec(1)
    'IF LEFT$(Type$, 1) = "A" THEN
    'IF ARTransRec(1).TransDate >= BegDate AND ARTransRec(1).TransDate <= EndDate THEN
    ' Count = Count + 1
    ' array(Count).who = LTRIM$(STR$(ARTransRec(1).TransDate)) + ARTransRec(1).CustomerNumber + STRING$(4, " ")
    ' array(Count).RecNum = Cnt!
    'END IF
    'END IF
    'IF LEFT$(Type$, 1) = "B" THEN
    'IF ARTransRec(1).TransType = 9 THEN
    'IF ARTransRec(1).TransDate >= BegDate AND ARTransRec(1).TransDate <= EndDate THEN
    'Count = Count + 1
    'array(Count).who = LTRIM$(STR$(ARTransRec(1).TransDate)) + ARTransRec(1).CustomerNumber + STRING$(4, " ")
    'array(Count).RecNum = Cnt!
    'END IF
    'END IF
    'END IF
    'IF LEFT$(Type$, 1) = "C" THEN
    'IF ARTransRec(1).TransType = 1 THEN
    ' IF ARTransRec(1).TransDate >= BegDate AND ARTransRec(1).TransDate <= EndDate THEN
    '  Count = Count + 1
    '  array(Count).who = LTRIM$(STR$(ARTransRec(1).TransDate)) + ARTransRec(1).CustomerNumber + STRING$(4, " ")
    '  array(Count).RecNum = Cnt!
    ' END IF
    'END IF
    'END IF
    'IF LEFT$(Type$, 1) = "P" THEN
    'IF ARTransRec(1).TransType = 2 THEN
    'IF ARTransRec(1).TransDate >= BegDate AND ARTransRec(1).TransDate <= EndDate THEN
    'Count = Count + 1
    'array(Count).who = LTRIM$(STR$(ARTransRec(1).TransDate)) + ARTransRec(1).CustomerNumber + STRING$(4, " ")
    'array(Count).RecNum = Cnt!
    'END IF
    'END IF
    'END IF
  NEXT Cnt!
  
  
  
  SortT array(Start), Count, Dir, SSize, MOff, MSize
  RETURN
  
  
  
  
  
  
SelectOutput2:
  LibName$ = "AR"
  ScrnName$ = "ARTRREP"
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  REDIM Choice$(4, 1)
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "All"
  Choice$(2, 0) = "Charges"
  Choice$(3, 0) = "Payments"
  Choice$(4, 0) = "Beg Bal"
  Choice$(0, 1) = "4"
  Choice$(1, 1) = "SCREEN"
  Choice$(2, 1) = "PRINTER"
  
  
  
  
  Action = 1
  ShowCursor
  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%
  'printhelp help$
  Form$(2, 0) = DATE$
  Form$(3, 0) = DATE$
  Action = 1
  
  DO
    
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      DevSpec$ = LEFT$(Form$(4, 0), 1)
      Type$ = Form$(1, 0)
      BegDate = Date2Num%(Form$(2, 0))
      BegDate$ = Form$(2, 0)
      EndDate = Date2Num%(Form$(3, 0))
      EndDate$ = Form$(3, 0)
      IF EndDate < BegDate THEN CLOSE : EXIT SUB
      RETURN
    CASE EscKey
      Canceled$ = "Y"
      RETURN
    END SELECT
  LOOP
  RETURN
  
  
END SUB

