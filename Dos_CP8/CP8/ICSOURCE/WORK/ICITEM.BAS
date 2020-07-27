DEFINT A-Z
DECLARE SUB ICursorOff ()
DECLARE SUB DisplayICScrn (ScrnName$)
DECLARE SUB OpenICItemFile (NumOfICRecs%, ICFile%)
DECLARE SUB ClearBack ()
DECLARE SUB ILookUp (RecNo&, Text$, ChkBalFlag%, CLSFlag%, SSNFlag%)
DECLARE SUB SortARNameIndex ()
DECLARE SUB AddItem ()
DECLARE SUB EditItem ()
DECLARE SUB PrintItem ()
DECLARE SUB ShowNoItems ()
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB ShowCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB HideCursor ()
DECLARE SUB QPrint (x$, Colr%, page%)
DECLARE SUB QPrintRC (T$, r%, c%, clr%)
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
  
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'IC.bi'                        'I/C FILE LAYOUTS

  TYPE Struct
    who AS STRING * 14
    RecNum AS INTEGER
  END TYPE

  STACK 8000
  
  DIM SHARED ICITEM(1) AS ItemRecType
  DIM SHARED ItemRec(1) AS ItemRecType
  
  CONST False = 0, True = NOT False
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 4)
  MChoice$(1) = " Add New Item"
  MChoice$(2) = " Edit Existing Item"
  MChoice$(3) = " Print Quick Item List "
  MChoice$(4) = " Quit to O/S"
  
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
  help$ = "Add/Edit/Print Items"
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    ClearBack
    'LibFile2Scrn "IC.QSL", "MENUBAK", MonoCode, -1, ErrorCode
    
    TitleBox 3, Col, MaxLen + 3, "Item Maintenance", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      AddItem
    CASE 2
      EditItem
    CASE 3
      PrintItem
    CASE 4
      END
    END SELECT
  LOOP
  RUN "icmenu"

SUB AddItem
  
mainbody:
  
  LibName$ = "IC"
  ScrnName$ = "ITEM"
  'help$ = "NEW Item Entry"

  ICursorOff
  ShowCursor
  
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
  
  Form$(8, 0) = "EACH"
  Fld(7).Protected = True
  Fld(9).Protected = True
  Fld(10).Protected = True
  Fld(11).Protected = True
  Fld(12).Protected = True
  
  Action = 1
  
  REM check for code file
  ItemRecLen = LEN(ItemRec(1))
  
  Frm(1).FldNo = 2

  ClearBack
  DisplayICScrn ScrnName$
  QPrintRC "        PENDING", 4, 28, 15

  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      GOSUB SaveRecord
      DONE = True
      GOTO mainbody
    CASE EscKey
      NeedtoSort = True         ' set to true for testing
      EXIT SUB
    END SELECT
  LOOP
  
  
SaveRecord:
  ItemFile = FREEFILE
  OPEN "ICITEM.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ItemFile LEN = ItemRecLen
  NumOfItemRecs = LOF(ItemFile) \ ItemRecLen
  NextRecord = NumOfItemRecs + 1
  
  ItemRec(1).ITEMNUMBER = NextRecord
  ItemRec(1).DESC1 = Form$(2, 0)
  ItemRec(1).DESC2 = Form$(3, 0)
  ItemRec(1).Vendor = Form$(4, 0)
  ItemRec(1).VItem = Form$(5, 0)
  ItemRec(1).DeptNumb = Form$(6, 0)
  ItemRec(1).COST = 0
  ItemRec(1).Unit = Form$(8, 0)
  ItemRec(1).QtyOnHand = 0
  ItemRec(1).QtyOnOrd = 0
  ItemRec(1).UsedYTD = 0
  ItemRec(1).ReOrder = Value#(Form$(12, 0), ecode)
  ItemRec(1).QTD = Value#(Form$(13, 0), ecode)
  ItemRec(1).Note1 = Form$(14, 0)
  ItemRec(1).Note2 = Form$(15, 0)
  ItemRec(1).Deleted = "N"
  ItemRec(1).Future = ""
  PUT ItemFile, NextRecord, ItemRec(1)
  CLOSE ItemFile
  QPrintRC "  ASSIGNED         ", 4, 28, 15
  BEEP
  SLEEP 2
  NeedtoSort = True

RETURN
  
  
END SUB

SUB EditItem '(RecNo&)
  
  ItemRecLen = LEN(ItemRec(1))
  
BeginAgain:
  ClearBack
  ILookUp RecNo&, "Item", False, True, False
  IF RecNo& = 0 THEN
    EXIT SUB
  END IF
  
  AccountRec& = RecNo&
  CLOSE ICFile
  OpenICItemFile NumOfICRecs, ICFile
  
EditMainBody:
  
  LibName$ = "IC"
  ScrnName$ = "ITEM"
  help$ = "Edit Item"
  
  ICursorOff
  
  ShowCursor
  DisplayICScrn ScrnName$
  'PrintHelp help$
  
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
  
  Action = 1
  FirstTime = True
  
  
  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF FirstTime THEN
      FirstTime = False
      GOSUB GetItem1
      MPaintBox 11, 28, 11, 37, 14
      MPaintBox 11, 53, 11, 65, 14
      MPaintBox 13, 28, 15, 39, 14
    END IF

    SELECT CASE Frm(1).KeyCode
    CASE F3Key
      IF AccountRec& > 0 THEN
        GOSUB DeleteRecord
        CLOSE ICFile
        ExitFlag = True
      END IF
    CASE F10Key
      GOSUB EditSaveRecord
      CLOSE ICFile
      ExitFlag = True
    CASE EscKey
      CLOSE ICFile
      ExitFlag = True
    END SELECT
    
  LOOP UNTIL ExitFlag

EXIT SUB
  
EditSaveRecord:
  IF AccountRec& = 0 THEN RETURN
  
  ItemRec(1).DESC1 = Form$(2, 0)
  ItemRec(1).DESC2 = Form$(3, 0)
  ItemRec(1).Vendor = Form$(4, 0)
  ItemRec(1).VItem = Form$(5, 0)
  ItemRec(1).DeptNumb = Form$(6, 0)
  ItemRec(1).Unit = Form$(9, 0)

  ItemRec(1).ReOrder = Value#(Form$(13, 0), ecode)
  ItemRec(1).Note1 = Form$(15, 0)
  ItemRec(1).Note2 = Form$(16, 0)
  PUT ICFile, AccountRec&, ItemRec(1)
  CLOSE ICFile

RETURN
  
  
GetItem1:

  LOCK #ICFile, AccountRec&
  GET ICFile, AccountRec&, ItemRec(1)
  IF ItemRec(1).Deleted = "Y" THEN
    GOTO ItemDeleted
  END IF
  LSET Form$(1, 0) = LTRIM$(STR$(AccountRec&))
  LSET Form$(2, 0) = ItemRec(1).DESC1
  LSET Form$(3, 0) = ItemRec(1).DESC2
  LSET Form$(4, 0) = ItemRec(1).Vendor
  LSET Form$(5, 0) = ItemRec(1).VItem
  LSET Form$(6, 0) = ItemRec(1).DeptNumb
  LSET Form$(7, 0) = QPTrim$(STR$(ItemRec(1).COST))
  LSET Form$(8, 0) = QPTrim$(STR$(ItemRec(1).YTDCOST))
  LSET Form$(9, 0) = ItemRec(1).Unit
  LSET Form$(10, 0) = QPTrim$(STR$(ItemRec(1).QtyOnHand))
  LSET Form$(11, 0) = QPTrim$(STR$(ItemRec(1).QtyOnOrd))
  LSET Form$(12, 0) = QPTrim$(STR$(ItemRec(1).UsedYTD))
  LSET Form$(13, 0) = QPTrim$(STR$(ItemRec(1).ReOrder))
  LSET Form$(14, 0) = QPTrim$(STR$(ItemRec(1).QTD))
  LSET Form$(15, 0) = ItemRec(1).Note1
  LSET Form$(16, 0) = ItemRec(1).Note2

  Fld(1).Protected = True
  Fld(7).Protected = True
  Fld(8).Protected = True
  Fld(10).Protected = True
  Fld(11).Protected = True
  Fld(12).Protected = True

  Action = 1

RETURN
    
ItemDeleted:

  LibName$ = "IC"
  ScrnName$ = "ICDELITEM"
  help$ = "Edit Item"
  ICursorOff

  ok = MsgBox%(LibName$, ScrnName$)
  Action = 1
  GOTO BeginAgain

DeleteRecord:
  LibName$ = "IC"
  ScrnName$ = "ICITEMDEL"
  help$ = "Delete Item"
  ICursorOff
  
  ShowCursor
  LibFile2Scrn "IC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  'PrintHelp help$

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
  
  Action = 1
  Form$(1, 0) = "Y"
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
      
      
    CASE F10Key
      IF Form$(1, 0) = "Y" THEN
        help$ = "Item Deleted!!!"
        'PrintHelp help$
        PRINT CHR$(7);
        ItemRec(1).Deleted = "Y"
        PUT ICFile, AccountRec&, ItemRec(1)
        CLOSE ICFile
        NeedtoSort = True
        RETURN
      END IF
    CASE EscKey
      RETURN
    END SELECT
    
  LOOP
END SUB

SUB OpenICItemFile (NumOfICRecs, ICFile)
  ItemRecLen = LEN(ItemRec(1))
  ICFile = FREEFILE
  OPEN "ICITEM.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ICFile LEN = ItemRecLen
  NumOfICRecs = LOF(ICFile) \ ItemRecLen
  
END SUB

SUB PrintItem
  
  SHARED Choice$()

  ReportFile$ = "ICITEM.PRN"    'Report File Name
  FF$ = CHR$(12)

  MaxLines = 53
  LineCnt = 0
  
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
      EXIT DO
    CASE EscKey
      ExitFlag = True
      EXIT DO
    END SELECT
  LOOP

  IF ExitFlag THEN
    GOTO QuickRptExit
  END IF
  
  'REDIM ItemRec(1) AS ItemRecType
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
      PRINT #RptHandle, TAB(10); ItemRec(1).DESC1; TAB(60); "QOH: "; USING "######,#.###"; ItemRec(1).QtyOnHand
      PRINT #RptHandle, TAB(10); ItemRec(1).DESC2; TAB(59); "Dept: "; ItemRec(1).DeptNumb
      PRINT #RptHandle, STRING$(79, "-")
      TotalItems = TotalItems + 1
      LineCnt = LineCnt + 3
    END IF
  NEXT

  GOSUB PrintRptEnding
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi

  CLOSE         'Close all open files now
  
  IF DevSpec$ = "S" THEN
    EntryPoint = 2
  ELSE
    EntryPoint = 4
  END IF
  
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  KILL ReportFile$
  
QuickRptExit:

  EXIT SUB
  
  
PrintRptHeader:
  page = page + 1
  PRINT #RptHandle, TAB(25); "Item Inventory 'Quick' Listing"
  PRINT #RptHandle, STRING$(80, "=")
  LineCnt = 5
RETURN
  
PrintRptEnding:
  PRINT #RptHandle, STRING$(80, "-")
  PRINT #RptHandle, "Number of Items .. "; USING "####,#"; TotalItems
  PRINT #RptHandle, FF$
RETURN
  
  
  
  
  
  
  
END SUB

SUB ShowNoItems
  LibName$ = "IC"
  ScrnName$ = "ICNOCUST"
  help$ = "New Item Entry"
  ICursorOff
  
  
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
  
  
  PRINT CHR$(7);
  ShowCursor
  LibFile2Scrn "IC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  Action = 1
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE EscKey
      EXIT SUB
    END SELECT
    
  LOOP
  
END SUB

