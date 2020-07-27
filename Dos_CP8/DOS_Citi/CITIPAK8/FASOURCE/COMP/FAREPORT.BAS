DEFINT A-Z
DECLARE SUB ItemSummaryReport ()
DECLARE SUB ItemDisposedReport ()
DECLARE SUB ItemNewDel ()
DECLARE SUB MasterList ()
DECLARE SUB ItemChkList ()
DECLARE SUB ItemValuation ()
DECLARE SUB ItemDepreciation ()
DECLARE SUB FASearch4SerNo (SerialNum$, RecNo&, CLSFlag%, ActiveOnly%)
DECLARE SUB FASearch4Desc (TText$, RecNo&, CLSFlag%, ActiveOnly%)
DECLARE SUB FAShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB FAShowProcessingScrn (RptTitle$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE SUB FASearch4Tag (TNumb$, RecNo&, CLSFlag%, ActiveOnly%)
DECLARE SUB DisplayFAScrn (ScrnName$)
DECLARE SUB ClearScrn ()
DECLARE SUB PrintHelp (H$)
DECLARE SUB BlockClear ()
DECLARE SUB PrintTitle (Title$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION WEnvTest% ()
DECLARE FUNCTION Round# (DblNum#)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB HideCursor ()
DECLARE SUB CursorOff ()
DECLARE SUB TextCursor (MouseFg%, MouseBg%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE SUB SaveScrn (Array())
DECLARE SUB RestScrn (Array())
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
DECLARE SUB FOpenS (FileName$, Handle)
DECLARE SUB FClose (Handle%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FGetA (Handle%, SEG Dest AS ANY, NumBytes%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB WaitForAction ()
DECLARE SUB MPaintBox (UlRow%, UlCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice, MaxLen%, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION Num2Date$ (TheDate%)
DECLARE FUNCTION QPValL& (Number$)
DECLARE FUNCTION QPStrI$ (Num%)
DECLARE FUNCTION QPStrL$ (Num&)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
  
  CONST False = 0, True = NOT False
  
  TYPE FLen2
    V AS STRING * 64
  END TYPE
  
  TYPE Struct
    Who AS STRING * 14
    RecNum AS INTEGER
  END TYPE

TYPE FundTotals
  Fund AS INTEGER
  OrgVal   AS DOUBLE
  CurVal   AS DOUBLE
  YtdDep   AS DOUBLE
END TYPE
  
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'newfa.BI'
  '$INCLUDE: 'Qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  
  
  DIM SHARED PctC(1) AS STRING * 3
  DIM SHARED FaItemRec(1) AS FAItemRecType

  STACK 5000
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 7)
  
  MChoice$(1) = " Master Item Listing - (8.5 Format)"
  MChoice$(2) = " Item Check List Report"
  MChoice$(3) = " New/Deleted Item Report"
  MChoice$(4) = " Item Depreciation Report "
  MChoice$(5) = " Item Disposed Of Report"
  MChoice$(6) = " Fund/Group Depreciation Report"
  MChoice$(7) = " Exit to OS"
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 18   'limit the box length to go no lower than line 18
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
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    
    TitleBox 2, Col, MaxLen + 3, "Fixed Assets Report Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      MasterList
    CASE 2
      ItemChkList
    CASE 3
      ItemNewDel
    CASE 4
      ItemDepreciation
    CASE 5
      ItemDisposedReport
    CASE 6
      ItemSummaryReport
    CASE 7
      HideCursor
      CLS
      END
    END SELECT
  LOOP
  
  IF WEnvTest THEN
    Ext$ = ".bas"
  ELSE
    Ext$ = ".exe"
  END IF
  IF Exist("famenu" + Ext$) THEN
    RUN "famenu"
  ELSE
    HideCursor
    ClearScrn
  END IF
  
  END

SUB ItemChkList
  
  SHARED Choice$()
  
  REDIM Array(1 TO 1) AS Struct 'Template for the sort array
  ReportFile$ = "FACHK.PRN"     'Report File Name
  Dash80$ = STRING$(80, "=")
  FF$ = CHR$(12)
  
  MaxLines = 50
  LineCnt& = 0
  ItemCnt& = 0
  
  LibName$ = "FA"
  ScrnName$ = "MASTRPT"
  
  FAItemRecLen = LEN(FaItemRec(1))
  
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
  
  REDIM Choice$(0 TO 2, 0 TO 1)
  
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "TAG NUMBER"
  Choice$(2, 0) = "DEPT"
  Choice$(0, 1) = "5"
  Choice$(1, 1) = "Screen"
  Choice$(2, 1) = "Printer"
  Form$(1, 0) = "DEPT"
  Form$(2, 0) = "N"
  Form$(3, 0) = "ALL"
  CYear = VAL(RIGHT$(DATE$, 4))
  CYear$ = LTRIM$(STR$(CYear))
  Form$(4, 0) = "N/A"           'CYear$
  Fld(1).Protected = 1
  Fld(2).Protected = 1
  Fld(4).Protected = 1
  
  Action = 1
  
  BlockClear
  ShowCursor
  
  DisplayFAScrn ScrnName$
  
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      Index$ = Form$(1, 0)
      Dispose$ = Form$(2, 0)
      Dept$ = RTRIM$(Form$(3, 0))
      DevSpec$ = LEFT$(Form$(5, 0), 1)
      ExitFlag = True
    CASE EscKey
      AbortFlag = True
      ExitFlag = True           'EXIT DO
    END SELECT
  LOOP UNTIL ExitFlag
  
  IF AbortFlag THEN EXIT SUB
  
  RptHandle = FREEFILE
  
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  
  FAFile = FREEFILE
  OPEN FAItemFile FOR RANDOM AS FAFile LEN = FAItemRecLen
  NumOfFaRecs = LOF(FAFile) / FAItemRecLen
  
  GOSUB GetIndex2
  
  
  FOR Cnt& = 1 TO NumOfFaRecs
    
    
    
    ItemRecNo = Array(Cnt&).RecNum
    GET FAFile, ItemRecNo, FaItemRec(1)
    
    IF Cnt& = 1 AND Dept$ = "ALL" THEN DeptNumber = VAL(FaItemRec(1).IDEPT)
    
    PValue# = 0
    
    IF LineCnt& >= MaxLines THEN
      PRINT #RptHandle, FF$
      GOSUB PrintMasterHeader2
    END IF
    
    'Check For Disposed Of
    
    IF Dispose$ = "N" THEN
      IF FaItemRec(1).DISPDATE > 0 THEN GOTO SkipEm2
    END IF
    
    IF Dept$ = "ALL" THEN
    ELSE
      IF VAL(Dept$) <> VAL(FaItemRec(1).IDEPT) THEN GOTO SkipEm2
    END IF
    
    IF RTRIM$(Index$) = "DEPT" AND RTRIM$(Dept$) = "ALL" THEN
      IF DeptNumber <> VAL(FaItemRec(1).IDEPT) THEN
        PRINT #RptHandle, FF$
        DeptNumber = VAL(FaItemRec(1).IDEPT)
        GOSUB PrintMasterHeader2
      END IF
    END IF
    
    
    
    'Figure Values
    
    IF FaItemRec(1).DISPDATE > 0 THEN Disp$ = "Y" ELSE Disp$ = "N"
    
    
    IF FirstTime = 0 THEN
      GOSUB PrintMasterHeader2
      FirstTime = 1
    END IF
    
    PRINT #RptHandle, FaItemRec(1).ITEMTAG;
    PRINT #RptHandle, TAB(22); RTRIM$(FaItemRec(1).IDESC1);
    PRINT #RptHandle, TAB(53); LEFT$(QPTrim$(FaItemRec(1).ITEMLOC), 24);
    PRINT #RptHandle, TAB(78); "___"
    PRINT #RptHandle, STRING$(80, "-")
    
    'SubTotal Here
    LineCnt& = LineCnt& + 2
    ItemCnt& = ItemCnt& + 1
    
    
    
SkipEm2:
  NEXT Cnt&
  
  GOSUB PrintChkEnding
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi
  CLOSE         'Close all open files now
  
  IF DevSpec$ = "P" THEN
    EntryPoint = 4
  ELSEIF DevSpec$ = "S" THEN
    EntryPoint = 2
  ELSE
    EntryPoint = 1
  END IF
  
  ERASE Array, Frm, Form$, Fld, FaItemRec
  
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  
  KILL ReportFile$
  
  EXIT SUB
  
PrintMasterHeader2:
  Page = Page + 1
  PRINT #RptHandle, TAB(29); "Asset Check List Report"
  PRINT #RptHandle, "Dept # "; DeptNumber
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
  PRINT #RptHandle, "Asset Tag Number"; TAB(22); "Description"; TAB(52); "Location"; TAB(77); "CHK"
  PRINT #RptHandle, Dash80$
  LineCnt& = 6
  RETURN
  
PrintChkEnding:
  PRINT #RptHandle, FF$
  RETURN
  
GetIndex2:
  REDIM Array(1 TO NumOfFaRecs) AS Struct
  FOR Cnt& = 1 TO NumOfFaRecs
    GET FAFile, Cnt&, FaItemRec(1)
    IF LEFT$(Index$, 3) = "TAG" THEN
      Array(Cnt&).Who = LTRIM$(LEFT$(UCASE$(FaItemRec(1).ITEMTAG), 14))
    ELSE
      RSET Array(Cnt&).Who = LTRIM$(RTRIM$(UCASE$(FaItemRec(1).IDEPT)) + RTRIM$(FaItemRec(1).AssetCode))
    END IF
    Array(Cnt&).RecNum = Cnt&
  NEXT
  
  'Sort Them Here
  SortT Array(1), NumOfFaRecs, 0, LEN(Array(1)), 0, 14
  RETURN
  
END SUB

SUB ItemDepreciation
  
  SHARED Choice$()
  DIM DOrigCost#(2), DBookTotal#(2), DCDep#(2), DYDep#(2), OrigCost#(2), BookTotal#(2), CDep#(2), YDep#(2)
  
  REDIM Array(1 TO 1) AS Struct 'Template for the sort array
  ReportFile$ = "FADEPREP.PRN"  'Report File Name
  Dash80$ = STRING$(80, "=")
  FF$ = CHR$(12)
  
  MaxLines = 50
  LineCnt& = 0
  ItemCnt& = 0
  
  LibName$ = "FA"
  ScrnName$ = "DEPREP"
  
  FAItemRecLen = LEN(FaItemRec(1))
  
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
  
  REDIM Choice$(0 TO 2, 0 TO 1)
  
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "TAG NUMBER"
  Choice$(2, 0) = "DEPT"
  Choice$(0, 1) = "5"
  Choice$(1, 1) = "Screen"
  Choice$(2, 1) = "Printer"
  Form$(1, 0) = "DEPT"
  Form$(2, 0) = "ALL"
  Form$(3, 0) = "01-01-1900"
  Form$(1, 0) = "DEPT":  Fld(1).Protected = 1
  Action = 1
  
  BlockClear
  ShowCursor
  
  DisplayFAScrn ScrnName$
  LOCATE 6, 22: COLOR 14: PRINT "Depreciation"
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      Index$ = Form$(1, 0)
      Dept$ = RTRIM$(Form$(2, 0))
      BDate$ = Form$(3, 0): BDate = Date2Num(BDate$)
      EDate$ = Form$(4, 0): EDate = Date2Num(EDate$)
      DevSpec$ = LEFT$(Form$(5, 0), 1)
      ExitFlag = True
      
    CASE EscKey
      AbortFlag = True
      ExitFlag = True           'EXIT DO
    END SELECT
  LOOP UNTIL ExitFlag
  
  IF AbortFlag THEN EXIT SUB
  
  RptHandle = FREEFILE
  
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  GOSUB PrintMasterHeader3
  
  FAFile = FREEFILE
  OPEN FAItemFile FOR RANDOM AS FAFile LEN = FAItemRecLen
  NumOfFaRecs = LOF(FAFile) / FAItemRecLen
  
  GOSUB GetIndex3
  
  
  FOR Cnt& = 1 TO NumOfFaRecs
    
    ItemRecNo = Array(Cnt&).RecNum
    GET FAFile, ItemRecNo, FaItemRec(1)
    IF Cnt& = 1 THEN DeptNumber = VAL(FaItemRec(1).IDEPT)
    
    PValue# = 0
    
    IF LineCnt& >= MaxLines THEN
      PRINT #RptHandle, FF$
      GOSUB PrintMasterHeader3
    END IF
    
    
    IF Dept$ = "ALL" THEN
    ELSE
      IF VAL(Dept$) <> VAL(FaItemRec(1).IDEPT) THEN GOTO SkipEm3
    END IF
    
    'Check For Disposed Date
    DisposeDate = FaItemRec(1).DISPDATE
    'Check for Acquired Date
    AcquireDate = FaItemRec(1).AQURDATE
    
    
    IF (DisposeDate >= BDate AND DisposeDate <= EDate) OR (AcquireDate >= BDate AND AcquireDate <= EDate) THEN
    ELSE
      GOTO SkipEm3
    END IF
    
    
    IF RTRIM$(Index$) = "DEPT" AND RTRIM$(Dept$) = "ALL" THEN
      IF DeptNumber <> VAL(FaItemRec(1).IDEPT) THEN
        'Print Subtotals and Clear
        PRINT #RptHandle, STRING$(105, "-")
        PRINT #RptHandle, "Totals for Dept Number: "; DeptNumber;
        PRINT #RptHandle, TAB(62); USING "########,#.##"; DOrigCost#;
        PRINT #RptHandle, TAB(78); USING "########,#.##"; DBookTotal#;
        PRINT #RptHandle, TAB(92); USING "########,#.##"; DYDep#
        'PRINT #RptHandle, "  "; FaItemRec(1).GLACCT
        LineCnt& = LineCnt& + 1
        
        PRINT #RptHandle, "": LineCnt& = LineCnt& + 1
        PRINT #RptHandle, "": LineCnt& = LineCnt& + 1
        
        
        'GOSUB PrintMasterHeader1
        DeptNumber = VAL(FaItemRec(1).IDEPT)
        DOrigCost# = 0
        DBookTotal# = 0
        DYDep# = 0
      END IF
    END IF
    
    'Figure Values
    'Calc Depreciation for This Period
    YtdDep# = FaItemRec(1).DEP2DATE
    BookVal# = FaItemRec(1).ORGCOST - FaItemRec(1).DEP2DATE
    IF BookVal# < 0 THEN BookVal# = 0

    'IF INSTR(FaItemRec(1).ITEMTAG, "ADM5") > 0 THEN STOP

    PRINT #RptHandle, FaItemRec(1).ITEMTAG; TAB(22); LEFT$(FaItemRec(1).IDESC1, 28);
    PRINT #RptHandle, TAB(51); FaItemRec(1).IDEPT;
    PRINT #RptHandle, TAB(58); USING "###"; FaItemRec(1).ILIFE;
    PRINT #RptHandle, TAB(62); USING "########,#.##"; FaItemRec(1).ORGCOST;
    PRINT #RptHandle, TAB(78); USING "########,#.##"; BookVal#;
    PRINT #RptHandle, TAB(92); USING "########,#.##"; YtdDep#
    
    
    'SubTotal Here
    LineCnt& = LineCnt& + 1
    ItemCnt& = ItemCnt& + 1
    'Grand Totals Here
    OrigCost# = OrigCost# + FaItemRec(1).ORGCOST
    BookTotal# = BookTotal# + BookVal#
    YDep# = YDep# + YtdDep#
    'Dept Totals Here
    DOrigCost# = DOrigCost# + FaItemRec(1).ORGCOST
    DBookTotal# = DBookTotal# + BookVal#
    DYDep# = DYDep# + YtdDep#
    
    
SkipEm3:
  NEXT Cnt&
  'First Print Subtotals
  PRINT #RptHandle, STRING$(105, "-")
  PRINT #RptHandle, "Totals for Dept Number: "; DeptNumber;
  PRINT #RptHandle, TAB(62); USING "########,#.##"; DOrigCost#;
  PRINT #RptHandle, TAB(78); USING "########,#.##"; DBookTotal#;
  PRINT #RptHandle, TAB(92); USING "########,#.##"; DYDep#
  
  LineCnt& = LineCnt& + 1
  
  PRINT #RptHandle, "": LineCnt& = LineCnt& + 1
  PRINT #RptHandle, "": LineCnt& = LineCnt& + 1
  
  GOSUB PrintDepRepEnding1
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi
  CLOSE         'Close all open files now
  
  IF DevSpec$ = "P" THEN
    EntryPoint = 4
  ELSEIF DevSpec$ = "S" THEN
    EntryPoint = 2
  ELSE
    EntryPoint = 1
  END IF
  
  ERASE Array, Frm, Form$, Fld, FaItemRec
  
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  
  KILL ReportFile$
  
  EXIT SUB
  
PrintMasterHeader3:
  Page = Page + 1
  PRINT #RptHandle, TAB(20); "Master Asset Listing : Depreciation Status Report"
  PRINT #RptHandle, "Depreciated Thru Year: "; CYear$
  PRINT #RptHandle, "Dept # "; Dept$
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
  PRINT #RptHandle, "Asset Number"; TAB(22); "Description"; TAB(51); "Dept"; TAB(58); "Life"; TAB(64); "Original Cost"; TAB(81); "Book Value"; TAB(94); "YTD Deprec"
  PRINT #RptHandle, STRING$(105, "=")
  LineCnt& = 6
  RETURN
  
PrintDepRepEnding1:
  PRINT #RptHandle, STRING$(105, "-")
  PRINT #RptHandle, "Grand Totals ";
  PRINT #RptHandle, TAB(62); USING "########,#.##"; OrigCost#;
  PRINT #RptHandle, TAB(78); USING "########,#.##"; BookTotal#;
  PRINT #RptHandle, TAB(92); USING "########,#.##"; YDep#
  
  PRINT #RptHandle, FF$
  RETURN
  
GetIndex3:


  REDIM Array(1 TO NumOfFaRecs) AS Struct

  FOR Cnt& = 1 TO NumOfFaRecs
    GET FAFile, Cnt&, FaItemRec(1)

'    FundCode$ = QPTrim$(LEFT$(FaItemRec(1).GLACCT, 2))
'    AssetCode$ = QPTrim$(FaItemRec(1).AssetCode)
'    Array(Cnt&).who = FundCode$ + AssetCode$

    IF LEFT$(Index$, 3) = "TAG" THEN
      Array(Cnt&).Who = LTRIM$(LEFT$(UCASE$(FaItemRec(1).ITEMTAG), 14))
    ELSE
      RSET Array(Cnt&).Who = LTRIM$(RTRIM$(UCASE$(FaItemRec(1).IDEPT)) + RTRIM$(FaItemRec(1).AssetCode))
    END IF

    Array(Cnt&).RecNum = Cnt&

  NEXT
  CLOSE CodeFile

  'Sort Them Here
  SortT Array(1), NumOfFaRecs, 0, LEN(Array(1)), 0, 14
  RETURN
  
  
END SUB

SUB ItemDisposedReport

  SHARED Choice$()
  DIM DOrigCost#(2), DBookTotal#(2), DCDep#(2), DYDep#(2), OrigCost#(2), BookTotal#(2), CDep#(2), YDep#(2)

  REDIM Array(1 TO 1) AS Struct 'Template for the sort array
  ReportFile$ = "FADISREP.PRN"  'Report File Name
  Dash80$ = STRING$(80, "=")
  FF$ = CHR$(12)

  MaxLines = 50
  LineCnt& = 0
  ItemCnt& = 0

  LibName$ = "FA"
  ScrnName$ = "DISREP"

  FAItemRecLen = LEN(FaItemRec(1))

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

  REDIM Choice$(0 TO 2, 0 TO 1)

  Choice$(0, 0) = "1"
  Choice$(1, 0) = "TAG NUMBER"
  Choice$(2, 0) = "DEPT"
  Choice$(0, 1) = "4"
  Choice$(1, 1) = "Screen"
  Choice$(2, 1) = "Printer"
  Form$(1, 0) = "DEPT"
  Form$(2, 0) = "ALL"
  Form$(3, 0) = RIGHT$(DATE$, 4)
  'Form$(1, 0) = "DEPT":  Fld(1).Protected = 1
  Action = 1

  BlockClear
  ShowCursor

  DisplayFAScrn ScrnName$
  LOCATE 6, 22: COLOR 14: PRINT "Depreciation"

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      Index$ = Form$(1, 0)
      Dept$ = RTRIM$(Form$(2, 0))
      DisYearDate$ = Form$(3, 0)
      DevSpec$ = LEFT$(Form$(4, 0), 1)
      ExitFlag = True

    CASE EscKey
      AbortFlag = True
      ExitFlag = True           'EXIT DO
    END SELECT
  LOOP UNTIL ExitFlag

  IF AbortFlag THEN EXIT SUB

  RptHandle = FREEFILE

  OPEN ReportFile$ FOR OUTPUT AS #RptHandle

  GOSUB PrintMasterHeader3ID

  FAFile = FREEFILE
  OPEN FAItemFile FOR RANDOM AS FAFile LEN = FAItemRecLen
  NumOfFaRecs = LOF(FAFile) / FAItemRecLen

  GOSUB GetIndex3ID


  FOR Cnt& = 1 TO NumOfFaRecs



    ItemRecNo = Array(Cnt&).RecNum
    GET FAFile, ItemRecNo, FaItemRec(1)
    IF Cnt& = 1 THEN DeptNumber = VAL(FaItemRec(1).IDEPT)

    PValue# = 0

    IF LineCnt& >= MaxLines THEN
      PRINT #RptHandle, FF$
      GOSUB PrintMasterHeader3ID
    END IF


    IF Dept$ = "ALL" THEN
    ELSE
      IF VAL(Dept$) <> VAL(FaItemRec(1).IDEPT) THEN GOTO SkipEm3ID
    END IF

    'Check For Disposed Date
    DisposeDate = FaItemRec(1).DISPDATE
    DisposeYear$ = RIGHT$(Num2Date$(DisposeDate), 4)
    'Check for Acquired Date
    AcquireDate = FaItemRec(1).AQURDATE


    IF VAL(DisposeYear$) <> VAL(DisYearDate$) THEN
    GOTO SkipEm3ID
    END IF


    IF RTRIM$(Index$) = "DEPT" AND RTRIM$(Dept$) = "ALL" THEN
      IF DeptNumber <> VAL(FaItemRec(1).IDEPT) THEN
        'Print Subtotals and Clear
        PRINT #RptHandle, STRING$(105, "-")
        PRINT #RptHandle, "Totals for Dept Number: "; DeptNumber;
        PRINT #RptHandle, TAB(62); USING "########,#.##"; DOrigCost#;
        PRINT #RptHandle, TAB(78); USING "########,#.##"; DBookTotal#;
        PRINT #RptHandle, TAB(92); USING "########,#.##"; DYDep#
        LineCnt& = LineCnt& + 1

        PRINT #RptHandle, "": LineCnt& = LineCnt& + 1
        PRINT #RptHandle, "": LineCnt& = LineCnt& + 1


        'GOSUB PrintMasterHeader1
        DeptNumber = VAL(FaItemRec(1).IDEPT)
        DOrigCost# = 0
        DBookTotal# = 0
        DYDep# = 0
      END IF
    END IF

    'Figure Values
    'Calc Depreciation for This Period
    YtdDep# = FaItemRec(1).DEP2DATE
    BookVal# = FaItemRec(1).ORGCOST - FaItemRec(1).DEP2DATE
    IF BookVal# < 0 THEN BookVal# = 0

    PRINT #RptHandle, FaItemRec(1).ITEMTAG; TAB(22); LEFT$(FaItemRec(1).IDESC1, 28);
    PRINT #RptHandle, TAB(51); FaItemRec(1).IDEPT;
    PRINT #RptHandle, TAB(58); USING "###"; FaItemRec(1).ILIFE;
    PRINT #RptHandle, TAB(62); USING "########,#.##"; FaItemRec(1).ORGCOST;
    PRINT #RptHandle, TAB(78); USING "########,#.##"; BookVal#;
    PRINT #RptHandle, TAB(92); USING "########,#.##"; YtdDep#;
    PRINT #RptHandle, TAB(109); Num2Date$(DisposeDate)


    'SubTotal Here
    LineCnt& = LineCnt& + 1
    ItemCnt& = ItemCnt& + 1
    'Grand Totals Here
    OrigCost# = OrigCost# + FaItemRec(1).ORGCOST
    BookTotal# = BookTotal# + BookVal#
    YDep# = YDep# + YtdDep#
    'Dept Totals Here
    DOrigCost# = DOrigCost# + FaItemRec(1).ORGCOST
    DBookTotal# = DBookTotal# + BookVal#
    DYDep# = DYDep# + YtdDep#


SkipEm3ID:
  NEXT Cnt&
  'First Print Subtotals
  PRINT #RptHandle, STRING$(120, "-")
  PRINT #RptHandle, "Totals for Dept Number: "; DeptNumber;
  PRINT #RptHandle, TAB(62); USING "########,#.##"; DOrigCost#;
  PRINT #RptHandle, TAB(78); USING "########,#.##"; DBookTotal#;
  PRINT #RptHandle, TAB(92); USING "########,#.##"; DYDep#

  LineCnt& = LineCnt& + 1

  PRINT #RptHandle, "": LineCnt& = LineCnt& + 1
  PRINT #RptHandle, "": LineCnt& = LineCnt& + 1

  GOSUB PrintDepRepEnding1id
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi
  CLOSE         'Close all open files now

  IF DevSpec$ = "P" THEN
    EntryPoint = 4
  ELSEIF DevSpec$ = "S" THEN
    EntryPoint = 2
  ELSE
    EntryPoint = 1
  END IF

  ERASE Array, Frm, Form$, Fld, FaItemRec

  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint


  KILL ReportFile$

  EXIT SUB

PrintMasterHeader3ID:
  Page = Page + 1
  PRINT #RptHandle, TAB(30); "Disposed of Asset Report"
  PRINT #RptHandle, "Year of Disposal: "; DisYearDate$
  PRINT #RptHandle, "Dept # "; Dept$
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
  PRINT #RptHandle, "Asset Number"; TAB(22); "Description"; TAB(51); "Dept"; TAB(58); "Life"; TAB(64); "Original Cost"; TAB(81); "Book Value"; TAB(94); "YTD Deprec"; TAB(108); "Date Disposed"
  PRINT #RptHandle, STRING$(120, "=")
  LineCnt& = 6
  RETURN

PrintDepRepEnding1id:
  PRINT #RptHandle, STRING$(120, "-")
  PRINT #RptHandle, "Grand Totals ";
  PRINT #RptHandle, TAB(62); USING "########,#.##"; OrigCost#;
  PRINT #RptHandle, TAB(78); USING "########,#.##"; BookTotal#;
  PRINT #RptHandle, TAB(92); USING "########,#.##"; YDep#

  PRINT #RptHandle, FF$
  RETURN

GetIndex3ID:
  REDIM Array(1 TO NumOfFaRecs) AS Struct
  FOR Cnt& = 1 TO NumOfFaRecs
    GET FAFile, Cnt&, FaItemRec(1)
    IF LEFT$(Index$, 3) = "TAG" THEN
      Array(Cnt&).Who = LTRIM$(LEFT$(UCASE$(FaItemRec(1).ITEMTAG), 14))
    ELSE
      RSET Array(Cnt&).Who = LTRIM$(RTRIM$(UCASE$(FaItemRec(1).IDEPT)) + RTRIM$(FaItemRec(1).AssetCode))
    END IF
    Array(Cnt&).RecNum = Cnt&
  NEXT

  'Sort Them Here
  SortT Array(1), NumOfFaRecs, 0, LEN(Array(1)), 0, 14
  RETURN

END SUB

SUB ItemNewDel
  SHARED Choice$()
  DIM DOrigCost#(2), DBookTotal#(2), DCDep#(2), DYDep#(2), OrigCost#(2), BookTotal#(2), CDep#(2), YDep#(2)
  
  REDIM Array(1 TO 1) AS Struct 'Template for the sort array
  ReportFile$ = "FAMaster.PRN"  'Report File Name
  Dash80$ = STRING$(80, "=")
  FF$ = CHR$(12)
  
  MaxLines = 50
  LineCnt& = 0
  ItemCnt& = 0
  
  LibName$ = "FA"
  ScrnName$ = "MASTNEW"
  
  FAItemRecLen = LEN(FaItemRec(1))
  
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
  
  REDIM Choice$(0 TO 2, 0 TO 1)
  
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "TAG NUMBER"
  Choice$(2, 0) = "DEPT"
  Choice$(0, 1) = "5"
  Choice$(1, 1) = "Screen"
  Choice$(2, 1) = "Printer"
  Form$(1, 0) = "DEPT":  Fld(1).Protected = 1
  Form$(2, 0) = "ALL"
  
  Action = 1
  
  BlockClear
  ShowCursor
  
  DisplayFAScrn ScrnName$
  
  
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      Index$ = Form$(1, 0)
      Dept$ = RTRIM$(Form$(2, 0))
      BDate$ = Form$(3, 0): BDate = Date2Num(BDate$)
      EDate$ = Form$(4, 0): EDate = Date2Num(EDate$)
      DevSpec$ = LEFT$(Form$(5, 0), 1)
      ExitFlag = True
    CASE EscKey
      AbortFlag = True
      ExitFlag = True           'EXIT DO
    END SELECT
  LOOP UNTIL ExitFlag
  
  IF AbortFlag THEN EXIT SUB
  
  RptHandle = FREEFILE
  
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  GOSUB PrintMasterHeader1
  
  FAFile = FREEFILE
  OPEN FAItemFile FOR RANDOM AS FAFile LEN = FAItemRecLen
  NumOfFaRecs = LOF(FAFile) / FAItemRecLen
  
  GOSUB GetIndex1
  
  
  FOR Cnt& = 1 TO NumOfFaRecs
    
    
    
    ItemRecNo = Array(Cnt&).RecNum
    GET FAFile, ItemRecNo, FaItemRec(1)
    IF Cnt& = 1 THEN DeptNumber = VAL(FaItemRec(1).IDEPT)
    PValue# = 0
    
    IF LineCnt& >= MaxLines THEN
      PRINT #RptHandle, FF$
      GOSUB PrintMasterHeader1
    END IF
    
    'Check For Disposed Date
    DisposeDate = FaItemRec(1).DISPDATE
    'Check for Acquired Date
    AcquireDate = FaItemRec(1).AQURDATE
    
    
    IF DisposeDate >= BDate AND DisposeDate <= EDate OR AcquireDate >= BDate AND AcquireDate <= EDate THEN
      IF DisposeDate >= BDate AND DisposeDate <= EDate THEN
        DFlag = 1
      ELSE
        DFlag = 0
      END IF
      IF AcquireDate >= BDate AND AcquireDate <= EDate THEN
        AFlag = 1
      ELSE
        AFlag = 0
      END IF
    ELSE
      GOTO SkipEm1
    END IF
    
    IF Dept$ = "ALL" THEN
    ELSE
      IF VAL(Dept$) <> VAL(FaItemRec(1).IDEPT) THEN GOTO SkipEm1
    END IF
    
    IF RTRIM$(Index$) = "DEPT" AND RTRIM$(Dept$) = "ALL" THEN
      IF DeptNumber <> VAL(FaItemRec(1).IDEPT) THEN
        'Print Subtotals and Clear
        PRINT #RptHandle, STRING$(120, "-")
        PRINT #RptHandle, "Current for Dept Number: "; DeptNumber;
        PRINT #RptHandle, TAB(63); USING "#######,#.##"; DOrigCost#;
        PRINT #RptHandle, TAB(79); USING "#######,#.##"; DBookTotal#;
        PRINT #RptHandle, TAB(93); USING "#######,#.##"; DCDep#;
        PRINT #RptHandle, TAB(105); USING "#######,#.##"; DYDep#
        
        PRINT #RptHandle, "Additions for Dept Number: "; DeptNumber;
        PRINT #RptHandle, TAB(63); USING "#######,#.##"; DOrigCost#(2);
        PRINT #RptHandle, TAB(79); USING "#######,#.##"; DBookTotal#(2);
        PRINT #RptHandle, TAB(93); USING "#######,#.##"; DCDep#(2);
        PRINT #RptHandle, TAB(107); USING "#######,#.##"; DYDep#(2)
        
        PRINT #RptHandle, "Deletions for Dept Number: "; DeptNumber;
        PRINT #RptHandle, TAB(63); USING "#######,#.##"; DOrigCost#(1);
        PRINT #RptHandle, TAB(79); USING "#######,#.##"; DBookTotal#(1);
        PRINT #RptHandle, TAB(93); USING "#######,#.##"; DCDep#(1);
        PRINT #RptHandle, TAB(107); USING "#######,#.##"; DYDep#(1)
        
        PRINT #RptHandle, "Totals for Dept Number: "; DeptNumber;
        PRINT #RptHandle, TAB(63); USING "#######,#.##"; DOrigCost# + DOrigCost#(1) + DOrigCost#(2);
        PRINT #RptHandle, TAB(79); USING "#######,#.##"; DBookTotal# + DBookTotal#(1) + DBookTotal#(2);
        PRINT #RptHandle, TAB(93); USING "#######,#.##"; DCDep# + DCDep#(1) + DCDep#(2);
        PRINT #RptHandle, TAB(107); USING "#######,#.##"; DYDep# + DYDep#(1) + DYDep#(2)
        LineCnt& = LineCnt& + 4
        
        PRINT #RptHandle, "": LineCnt& = LineCnt& + 1
        PRINT #RptHandle, "": LineCnt& = LineCnt& + 1
        
        
        'GOSUB PrintMasterHeader1
        DeptNumber = VAL(FaItemRec(1).IDEPT)
        DOrigCost# = 0
        DBookTotal# = 0
        DCDep# = 0
        DYDep# = 0
        DOrigCost#(1) = 0
        DBookTotal#(1) = 0
        DCDep#(1) = 0
        DYDep#(1) = 0
        DOrigCost#(2) = 0
        DBookTotal#(2) = 0
        DCDep#(2) = 0
        DYDep#(2) = 0
      END IF
    END IF
    
    'Figure Values
    IF DFlag = 1 THEN PRINT #RptHandle, "D";
    IF AFlag = 1 THEN PRINT #RptHandle, "A";
    'Calc Depreciation for This Period
    IF FaItemRec(1).ILIFE > 0 THEN
      CurDep# = Round#(FaItemRec(1).ORGCOST / FaItemRec(1).ILIFE)
      MaxDep# = Round#(FaItemRec(1).ORGCOST - FaItemRec(1).DEP2DATE)
      IF CurDep# > MaxDep# THEN CurDep# = MaxDep#
      YtdDep# = Round#(CurDep# + FaItemRec(1).DEP2DATE)
    ELSE
      CurDep# = 0
      YtdDep# = 0
    END IF
    
    PRINT #RptHandle, FaItemRec(1).ITEMTAG; TAB(22); LEFT$(FaItemRec(1).IDESC1, 28);
    PRINT #RptHandle, TAB(51); FaItemRec(1).IDEPT;
    PRINT #RptHandle, TAB(58); USING "###"; FaItemRec(1).ILIFE;
    PRINT #RptHandle, TAB(63); USING "#######,#.##"; FaItemRec(1).ORGCOST;
    PRINT #RptHandle, TAB(79); USING "#######,#.##"; FaItemRec(1).ORGCOST - FaItemRec(1).DEP2DATE;
    PRINT #RptHandle, TAB(93); USING "#######,#.##"; CurDep#;
    PRINT #RptHandle, TAB(107); USING "#######,#.##"; YtdDep#
    
    
    'SubTotal Here
    LineCnt& = LineCnt& + 1
    ItemCnt& = ItemCnt& + 1
    'Grand Totals Here
    IF DFlag = 0 AND AFlag = 0 THEN
      OrigCost# = OrigCost# + FaItemRec(1).ORGCOST
      BookTotal# = BookTotal# + (FaItemRec(1).ORGCOST - FaItemRec(1).DEP2DATE)
      CDep# = CDep# + CurDep#
      YDep# = YDep# + YtdDep#
    END IF
    IF DFlag = 1 THEN
      OrigCost#(1) = OrigCost#(1) + FaItemRec(1).ORGCOST
      BookTotal#(1) = BookTotal#(1) + (FaItemRec(1).ORGCOST - FaItemRec(1).DEP2DATE)
      CDep#(1) = CDep#(1) + CurDep#
      YDep#(1) = YDep#(1) + YtdDep#
    END IF
    IF AFlag = 1 THEN
      OrigCost#(2) = OrigCost#(2) + FaItemRec(1).ORGCOST
      BookTotal#(2) = BookTotal#(2) + (FaItemRec(1).ORGCOST - FaItemRec(1).DEP2DATE)
      CDep#(2) = CDep#(2) + CurDep#
      YDep#(2) = YDep#(2) + YtdDep#
    END IF
    
    'Dept Totals Here
    IF DFlag = 0 AND AFlag = 0 THEN
      DOrigCost# = DOrigCost# + FaItemRec(1).ORGCOST
      DBookTotal# = DBookTotal# + (FaItemRec(1).ORGCOST - FaItemRec(1).DEP2DATE)
      DCDep# = DCDep# + CurDep#
      DYDep# = DYDep# + YtdDep#
    END IF
    IF DFlag = 1 THEN
      DOrigCost#(1) = DOrigCost#(1) + FaItemRec(1).ORGCOST
      DBookTotal#(1) = DBookTotal#(1) + (FaItemRec(1).ORGCOST - FaItemRec(1).DEP2DATE)
      DCDep#(1) = DCDep#(1) + CurDep#
      DYDep#(1) = DYDep#(1) + YtdDep#
    END IF
    IF AFlag = 1 THEN
      DOrigCost#(2) = DOrigCost#(2) + FaItemRec(1).ORGCOST
      DBookTotal#(2) = DBookTotal#(2) + (FaItemRec(1).ORGCOST - FaItemRec(1).DEP2DATE)
      DCDep#(2) = DCDep#(2) + CurDep#
      DYDep#(2) = DYDep#(2) + YtdDep#
    END IF
    
    
    
SkipEm1:
  NEXT Cnt&
  'First Print Subtotals
  PRINT #RptHandle, STRING$(120, "-")
  PRINT #RptHandle, "Current for Dept Number: "; DeptNumber;
  PRINT #RptHandle, TAB(63); USING "#######,#.##"; DOrigCost#;
  PRINT #RptHandle, TAB(79); USING "#######,#.##"; DBookTotal#;
  PRINT #RptHandle, TAB(93); USING "#######,#.##"; DCDep#;
  PRINT #RptHandle, TAB(107); USING "#######,#.##"; DYDep#
  
  PRINT #RptHandle, "Additions for Dept Number: "; DeptNumber;
  PRINT #RptHandle, TAB(63); USING "#######,#.##"; DOrigCost#(2);
  PRINT #RptHandle, TAB(79); USING "#######,#.##"; DBookTotal#(2);
  PRINT #RptHandle, TAB(93); USING "#######,#.##"; DCDep#(2);
  PRINT #RptHandle, TAB(107); USING "#######,#.##"; DYDep#(2)
  
  PRINT #RptHandle, "Deletions for Dept Number: "; DeptNumber;
  PRINT #RptHandle, TAB(63); USING "#######,#.##"; DOrigCost#(1);
  PRINT #RptHandle, TAB(79); USING "#######,#.##"; DBookTotal#(1);
  PRINT #RptHandle, TAB(93); USING "#######,#.##"; DCDep#(1);
  PRINT #RptHandle, TAB(107); USING "#######,#.##"; DYDep#(1)
  
  PRINT #RptHandle, "Totals for Dept Number: "; DeptNumber;
  PRINT #RptHandle, TAB(63); USING "#######,#.##"; DOrigCost# + DOrigCost#(1) + DOrigCost#(2);
  PRINT #RptHandle, TAB(79); USING "#######,#.##"; DBookTotal# + DBookTotal#(1) + DBookTotal#(2);
  PRINT #RptHandle, TAB(93); USING "#######,#.##"; DCDep# + DCDep#(1) + DCDep#(2);
  PRINT #RptHandle, TAB(107); USING "#######,#.##"; DYDep# + DYDep#(1) + DYDep#(2)
  LineCnt& = LineCnt& + 1
  
  PRINT #RptHandle, "": LineCnt& = LineCnt& + 1
  PRINT #RptHandle, "": LineCnt& = LineCnt& + 1
  
  GOSUB PrintMasterValueEnding1
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi
  CLOSE         'Close all open files now
  
  IF DevSpec$ = "P" THEN
    EntryPoint = 4
  ELSEIF DevSpec$ = "S" THEN
    EntryPoint = 2
  ELSE
    EntryPoint = 1
  END IF
  
  ERASE Array, Frm, Form$, Fld, FaItemRec
  
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  
  KILL ReportFile$
  
  EXIT SUB
  
PrintMasterHeader1:
  Page = Page + 1
  PRINT #RptHandle, TAB(20); "Master Asset Listing : Additions and Deletions"
  PRINT #RptHandle,
  PRINT #RptHandle, "Dept # "; Dept$
  PRINT #RptHandle, "Assets Acquired From "; BDate$; " to "; EDate$
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(75); "Page #"; Page
  PRINT #RptHandle, TAB(94); "[---   Calculated   ---]"
  PRINT #RptHandle, TAB(94); "[--- Not Yet Posted ---]"
  PRINT #RptHandle, "Asset Number"; TAB(22); "Description"; TAB(51); "Dept"; TAB(58); "Life"; TAB(64); "Original Cost"; TAB(81); "Book Value"; TAB(94); "Curr Deprec"; TAB(109); "YTD Deprec"
  PRINT #RptHandle, STRING$(120, "=")
  LineCnt& = 8
  RETURN
  
PrintMasterValueEnding1:
  PRINT #RptHandle, STRING$(120, "-")
  PRINT #RptHandle, "Current ";
  PRINT #RptHandle, TAB(63); USING "#######,#.##"; OrigCost#;
  PRINT #RptHandle, TAB(79); USING "#######,#.##"; BookTotal#;
  PRINT #RptHandle, TAB(93); USING "#######,#.##"; CDep#;
  PRINT #RptHandle, TAB(107); USING "#######,#.##"; YDep#
  
  PRINT #RptHandle, "Additions ";
  PRINT #RptHandle, TAB(63); USING "#######,#.##"; OrigCost#(2);
  PRINT #RptHandle, TAB(79); USING "#######,#.##"; BookTotal#(2);
  PRINT #RptHandle, TAB(93); USING "#######,#.##"; CDep#(2);
  PRINT #RptHandle, TAB(107); USING "#######,#.##"; YDep#(2)
  
  PRINT #RptHandle, "Deletions ";
  PRINT #RptHandle, TAB(63); USING "#######,#.##"; OrigCost#(1);
  PRINT #RptHandle, TAB(79); USING "#######,#.##"; BookTotal#(1);
  PRINT #RptHandle, TAB(93); USING "#######,#.##"; CDep#(1);
  PRINT #RptHandle, TAB(107); USING "#######,#.##"; YDep#(1)
  
  PRINT #RptHandle, "Grand Totals ";
  PRINT #RptHandle, TAB(63); USING "#######,#.##"; OrigCost# + OrigCost#(1) + OrigCost#(2);
  PRINT #RptHandle, TAB(79); USING "#######,#.##"; BookTotal# + BookTotal#(1) + BookTotal#(2);
  PRINT #RptHandle, TAB(93); USING "#######,#.##"; CDep# + CDep#(1) + CDep#(2);
  PRINT #RptHandle, TAB(107); USING "#######,#.##"; YDep# + YDep#(1) + YDep#(2)
  PRINT #RptHandle, FF$
  RETURN
  
GetIndex1:
  REDIM Array(1 TO NumOfFaRecs) AS Struct
  FOR Cnt& = 1 TO NumOfFaRecs
    GET FAFile, Cnt&, FaItemRec(1)
    IF LEFT$(Index$, 3) = "TAG" THEN
      Array(Cnt&).Who = LTRIM$(LEFT$(UCASE$(FaItemRec(1).ITEMTAG), 14))
    ELSE
      RSET Array(Cnt&).Who = LTRIM$(RTRIM$(UCASE$(FaItemRec(1).IDEPT)) + RTRIM$(FaItemRec(1).AssetCode))
    END IF
    Array(Cnt&).RecNum = Cnt&
  NEXT
  
  'Sort Them Here
  SortT Array(1), NumOfFaRecs, 0, LEN(Array(1)), 0, 14
  RETURN
  
END SUB

SUB ItemSummaryReport

  SHARED Choice$()
  REDIM CodeRec(1)  AS FAAssetCodeRecType
  CodeRecLen = LEN(CodeRec(1))

  DIM DeptNumber AS STRING

  REDIM FundSum(1 TO 1) AS FundTotals

  REDIM Array(1 TO 1) AS Struct 'Template for the sort array
  ReportFile$ = "FADEPREP.PRN"  'Report File Name
  Dash80$ = STRING$(80, "=")
  FF$ = CHR$(12)

  MaxLines = 50
  LineCnt& = 0
  ItemCnt& = 0

  LibName$ = "FA"
  ScrnName$ = "SUMDEP"

  FAItemRecLen = LEN(FaItemRec(1))

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

  REDIM Choice$(0 TO 2, 0 TO 1)

  Choice$(0, 0) = "1"
'  Choice$(1, 0) = "TAG NUMBER"
'  Choice$(2, 0) = "DEPT"
  Choice$(0, 1) = "4"
  Choice$(1, 1) = "Screen"
  Choice$(2, 1) = "Printer"
  'Form$(2, 0) = "01-01-1900"
  'Form$(3, 0) = "12-31-2008"

  Action = 1

  BlockClear
  ShowCursor

  DisplayFAScrn ScrnName$
  'LOCATE 6, 22
  'COLOR 14
  'PRINT "Depreciation"
  FirstTime = -1
  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      MinAmt# = Value(Form$(1, 0), ECode)
      BDate$ = Form$(2, 0)
      BDate = Date2Num(BDate$)
      EDate$ = Form$(3, 0)
      EDate = Date2Num(EDate$)
      DevSpec$ = LEFT$(Form$(4, 0), 1)
      ExitFlag = True

    CASE EscKey
      AbortFlag = True
      ExitFlag = True           'EXIT DO
    END SELECT
  LOOP UNTIL ExitFlag

  IF AbortFlag THEN EXIT SUB

  CodeFileSize& = FileSize(FACodeFile)
  NumCodeRec = CodeFileSize& \ CodeRecLen
  REDIM CodeRec(1 TO NumCodeRec) AS FAAssetCodeRecType

  CodeFile = FREEFILE
  OPEN FACodeFile FOR RANDOM AS CodeFile LEN = CodeRecLen

  FOR zz = 1 TO NumCodeRec
    GET CodeFile, zz, CodeRec(zz)
  NEXT
  CLOSE CodeFile

  RptHandle = FREEFILE

  OPEN ReportFile$ FOR OUTPUT AS #RptHandle

  GOSUB PrintMasterHeader33

  FAFile = FREEFILE
  OPEN FAItemFile FOR RANDOM AS FAFile LEN = FAItemRecLen
  NumOfFaRecs = LOF(FAFile) / FAItemRecLen

  GOSUB GetIndex33
  DeptNumber = QPTrim$(Array(1).Who)

  FOR Cnt& = 1 TO NumOfFaRecs

    ItemRecNo = Array(Cnt&).RecNum

    GET FAFile, ItemRecNo, FaItemRec(1)
    IF FaItemRec(1).ORGCOST < MinAmt# THEN
      GOTO SkipEm33
    END IF

    PValue# = 0

    IF LineCnt& >= MaxLines THEN
      PRINT #RptHandle, FF$
      GOSUB PrintMasterHeader33
    END IF

'    IF Dept$ = "ALL" THEN
'    ELSE
'      IF VAL(Dept$) <> VAL(FaItemRec(1).IDEPT) THEN GOTO SkipEm33
'    END IF

    'Check For Disposed Date
    DisposeDate = FaItemRec(1).DISPDATE
    'Check for Acquired Date
    AcquireDate = FaItemRec(1).AQURDATE

    IF (DisposeDate >= BDate AND DisposeDate <= EDate) OR (AcquireDate >= BDate AND AcquireDate <= EDate) THEN
    ELSE
      GOTO SkipEm33
    END IF

    IF DeptNumber <> QPTrim$(Array(Cnt&).Who) THEN
      'Print Subtotals and Clear
      PRINT #RptHandle, STRING$(105, "-")
      PRINT #RptHandle, "Totals for Fund/Group: "; DeptNumber; "  ";
      FOR zz = 1 TO NumCodeRec
        IF RIGHT$(DeptNumber, 4) = CodeRec(zz).AssetCode THEN
          PRINT #RptHandle, CodeRec(zz).AssetDesc;
          EXIT FOR
        END IF
      NEXT
      PRINT #RptHandle, TAB(62); USING "########,#.##"; DOrigCost#;
      PRINT #RptHandle, TAB(78); USING "########,#.##"; DBookTotal#;
      PRINT #RptHandle, TAB(92); USING "########,#.##"; DYDep#
      PRINT #RptHandle, ""
      PRINT #RptHandle, ""
      LineCnt& = LineCnt& + 3

      'GOSUB PrintMasterHeader1
      DeptNumber = QPTrim$(Array(Cnt&).Who)
      DOrigCost# = 0
      DBookTotal# = 0
      DYDep# = 0
    END IF

    'Figure Values
    'Calc Depreciation for This Period
    YtdDep# = FaItemRec(1).DEP2DATE
    BookVal# = FaItemRec(1).ORGCOST - FaItemRec(1).DEP2DATE
    IF BookVal# < 0 THEN BookVal# = 0

    PRINT #RptHandle, FaItemRec(1).ITEMTAG; TAB(22); LEFT$(FaItemRec(1).IDESC1, 28);
    PRINT #RptHandle, TAB(51); FaItemRec(1).IDEPT;
    PRINT #RptHandle, TAB(58); USING "###"; FaItemRec(1).ILIFE;
    PRINT #RptHandle, TAB(62); USING "########,#.##"; FaItemRec(1).ORGCOST;
    PRINT #RptHandle, TAB(78); USING "########,#.##"; BookVal#;
    PRINT #RptHandle, TAB(92); USING "########,#.##"; YtdDep#

    ItemFund = VAL(LEFT$(Array(Cnt&).Who, 2))

    FOR zz = 1 TO FundCnt
      IF FundSum(zz).Fund = ItemFund THEN
        FundSum(zz).OrgVal = Round(FundSum(zz).OrgVal + FaItemRec(1).ORGCOST)
        FundSum(zz).CurVal = Round(FundSum(zz).CurVal + BookVal#)
        FundSum(zz).YtdDep = Round(FundSum(zz).YtdDep + YtdDep#)
        EXIT FOR
      END IF
    NEXT


    'SubTotal Here
    LineCnt& = LineCnt& + 1
    ItemCnt& = ItemCnt& + 1
    'Grand Totals Here
    OrigCost# = OrigCost# + FaItemRec(1).ORGCOST
    BookTotal# = BookTotal# + BookVal#
    YDep# = YDep# + YtdDep#
    'Dept Totals Here
    DOrigCost# = DOrigCost# + FaItemRec(1).ORGCOST
    DBookTotal# = DBookTotal# + BookVal#
    DYDep# = DYDep# + YtdDep#


SkipEm33:
  NEXT Cnt&
  'First Print Subtotals
  PRINT #RptHandle, STRING$(105, "-")
  'PRINT #RptHandle, "Totals for Dept Number: "; DeptNumber;
  PRINT #RptHandle, "Totals for Fund/Group: "; DeptNumber; "  ";
  FOR zz = 1 TO NumCodeRec
    IF RIGHT$(DeptNumber, 4) = CodeRec(zz).AssetCode THEN
      PRINT #RptHandle, CodeRec(zz).AssetDesc;
      EXIT FOR
    END IF
  NEXT

  PRINT #RptHandle, TAB(62); USING "########,#.##"; DOrigCost#;
  PRINT #RptHandle, TAB(78); USING "########,#.##"; DBookTotal#;
  PRINT #RptHandle, TAB(92); USING "########,#.##"; DYDep#
  PRINT #RptHandle, ""
  PRINT #RptHandle, ""
  LineCnt& = LineCnt& + 3

  GOSUB PrintDepRepEnding11

  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi
  CLOSE         'Close all open files now

  IF DevSpec$ = "P" THEN
    EntryPoint = 4
  ELSEIF DevSpec$ = "S" THEN
    EntryPoint = 2
  ELSE
    EntryPoint = 1
  END IF

  ERASE Array, Frm, Form$, Fld, FaItemRec

  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint

  KILL ReportFile$

  EXIT SUB

PrintMasterHeader33:
  Page = Page + 1
  PRINT #RptHandle, TAB(16); "Fund/Group Asset Listing : Depreciation Status Report"
  'PRINT #RptHandle, "Depreciated Thru Year: "; CYear$
  'PRINT #RptHandle, "Dept # "; Dept$
  PRINT #RptHandle,
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
  PRINT #RptHandle, "Asset Number"; TAB(22); "Description"; TAB(51); "Dept"; TAB(58); "Life"; TAB(64); "Original Cost"; TAB(81); "Book Value"; TAB(94); "YTD Deprec"
  PRINT #RptHandle, STRING$(105, "=")
  LineCnt& = 6
  RETURN

PrintDepRepEnding11:
  PRINT #RptHandle, FF$
  PRINT #RptHandle, TAB(62); "Original Cost"; TAB(78); "Current Value"; TAB(93); "Depreciation"
  PRINT #RptHandle, STRING$(105, "-")
  PRINT #RptHandle, "Grand Totals ";
  PRINT #RptHandle, TAB(62); USING "########,#.##"; OrigCost#;
  PRINT #RptHandle, TAB(78); USING "########,#.##"; BookTotal#;
  PRINT #RptHandle, TAB(92); USING "########,#.##"; YDep#
  PRINT #RptHandle,
  PRINT #RptHandle, "Fund Summary"
  
  FOR zz = 1 TO FundCnt
    IF FundSum(zz).Fund = 0 THEN
      PRINT #RptHandle, TAB(33); "00";
    ELSE
      PRINT #RptHandle, TAB(32); FundSum(zz).Fund;
    END IF
    PRINT #RptHandle, TAB(62); USING "########,#.##"; FundSum(zz).OrgVal;
    PRINT #RptHandle, TAB(78); USING "########,#.##"; FundSum(zz).CurVal#;
    PRINT #RptHandle, TAB(92); USING "########,#.##"; FundSum(zz).YtdDep
  NEXT

  PRINT #RptHandle, FF$
  RETURN

GetIndex33:


  REDIM Array(1 TO NumOfFaRecs) AS Struct

  FOR Cnt& = 1 TO NumOfFaRecs
    GET FAFile, Cnt&, FaItemRec(1)

    FundCode$ = QPTrim$(LEFT$(FaItemRec(1).GLACCT, 2))
    IF LEN(FundCode$) < 2 THEN
      FundCode$ = "00"
    END IF
    AssetCode$ = QPTrim$(FaItemRec(1).AssetCode)

    Array(Cnt&).Who = FundCode$ + AssetCode$
    Array(Cnt&).RecNum = Cnt&

SkipThisin:
  NEXT

  'Sort Them Here
  SortT Array(1), NumOfFaRecs, 0, LEN(Array(1)), 0, 14
  
  FundCnt = 1
  FundSum(FundCnt).Fund = VAL(LEFT$(Array(1).Who, 2))
  FOR Cnt& = 2 TO NumOfFaRecs
    ThisFund = VAL(LEFT$(Array(Cnt&).Who, 2))
    IF ThisFund <> FundSum(FundCnt).Fund THEN
      FundCnt = FundCnt + 1
      REDIM PRESERVE FundSum(1 TO FundCnt) AS FundTotals
      FundSum(FundCnt).Fund = ThisFund
    END IF
  NEXT
  RETURN

END SUB

SUB ItemValuation
  
END SUB

SUB MasterList
  SHARED Choice$()
  
  REDIM Array(1 TO 1) AS Struct 'Template for the sort array
  ReportFile$ = "FAMaster.PRN"  'Report File Name
  Dash80$ = STRING$(80, "=")
  FF$ = CHR$(12)
  
  MaxLines = 50
  LineCnt& = 0
  ItemCnt& = 0
  
  LibName$ = "FA"
  ScrnName$ = "MASTRPT"
  
  FAItemRecLen = LEN(FaItemRec(1))
  
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
  
  REDIM Choice$(0 TO 2, 0 TO 1)
  
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "TAG NUMBER"
  Choice$(2, 0) = "DEPT"
  Choice$(0, 1) = "5"
  Choice$(1, 1) = "Screen"
  Choice$(2, 1) = "Printer"
  
  Form$(2, 0) = "N"
  Form$(3, 0) = "ALL"
  CYear = VAL(RIGHT$(DATE$, 4))
  CYear$ = LTRIM$(STR$(CYear))
  Form$(4, 0) = "N/A"           'CYear$
  Fld(4).Protected = 1
  
  Action = 1
  
  BlockClear
  ShowCursor
  
  DisplayFAScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      Index$ = Form$(1, 0)
      Dispose$ = Form$(2, 0)
      Dept$ = RTRIM$(Form$(3, 0))
      DevSpec$ = LEFT$(Form$(5, 0), 1)
      ExitFlag = True
    CASE EscKey
      AbortFlag = True
      ExitFlag = True           'EXIT DO
    END SELECT
  LOOP UNTIL ExitFlag
  
  IF AbortFlag THEN EXIT SUB
  
  RptHandle = FREEFILE
  
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  GOSUB PrintMasterHeader
  
  FAFile = FREEFILE
  OPEN FAItemFile FOR RANDOM AS FAFile LEN = FAItemRecLen
  NumOfFaRecs = LOF(FAFile) / FAItemRecLen
  
  GOSUB GetIndex
  
  FOR Cnt& = 1 TO NumOfFaRecs
    ItemRecNo = Array(Cnt&).RecNum
    GET FAFile, ItemRecNo, FaItemRec(1)
    IF Cnt& = 1 THEN DeptNumber = VAL(FaItemRec(1).IDEPT)
    PValue# = 0

    IF LineCnt& >= MaxLines THEN
      PRINT #RptHandle, FF$
      GOSUB PrintMasterHeader
    END IF
    
    'Check For Disposed Of
    
    IF Dispose$ = "N" THEN
      IF FaItemRec(1).DISPDATE > 0 THEN GOTO SkipEm
    END IF
    
    IF Dept$ = "ALL" THEN
      
    ELSE
      IF VAL(Dept$) <> VAL(FaItemRec(1).IDEPT) THEN GOTO SkipEm
    END IF
    
    IF RTRIM$(Index$) = "DEPT" AND RTRIM$(Dept$) = "ALL" THEN
      IF DeptNumber <> VAL(FaItemRec(1).IDEPT) THEN
        'Print Subtotals and Clear
        PRINT #RptHandle, "Subtotal for Dept Number: "; DeptNumber
        PRINT #RptHandle, "Total Items Printed: "; USING "#####"; DItemCnt&
        PRINT #RptHandle, "Total Original Cost: "; USING "$$########,#.##"; DOrigCost#
        PRINT #RptHandle, "Total Deprc to Date: "; USING "$$########,#.##"; DDepTot#
        PRINT #RptHandle, FF$
        GOSUB PrintMasterHeader
        DeptNumber = VAL(FaItemRec(1).IDEPT)
        DItemCnt& = 0
        DOrigCost# = 0
        DDepTot# = 0
      END IF
      'May Need to Skip Them
      DItemCnt& = DItemCnt& + 1
      DOrigCost# = DOrigCost# + FaItemRec(1).ORGCOST
      DDepTot# = DDepTot# + FaItemRec(1).DEP2DATE
    END IF
    
    'Figure Values
    
    IF FaItemRec(1).DISPDATE > 0 THEN Disp$ = "Y" ELSE Disp$ = "N"
    PRINT #RptHandle, "Tag #"; FaItemRec(1).ITEMTAG;
    PRINT #RptHandle, TAB(40); "Ser #"; FaItemRec(1).SERIALNO
    PRINT #RptHandle, "Desc "; FaItemRec(1).IDESC1;
    PRINT #RptHandle, TAB(40); "Mfg  "; FaItemRec(1).ITEMMFG
    PRINT #RptHandle, TAB(5); FaItemRec(1).IDESC2;
    PRINT #RptHandle, TAB(40); "Cont "; FaItemRec(1).CONTACT
    PRINT #RptHandle, "Loc  "; QPTrim$(FaItemRec(1).ITEMLOC);
    PRINT #RptHandle, TAB(40); "Dept "; FaItemRec(1).IDEPT
    PRINT #RptHandle, "Orig Cost: "; USING "$$######,#.##"; FaItemRec(1).ORGCOST;
    PRINT #RptHandle, TAB(40); "Acquired On "; Num2Date(FaItemRec(1).AQURDATE)
    PRINT #RptHandle, "Asset Life: "; USING "###"; FaItemRec(1).ILIFE;
    PRINT #RptHandle, TAB(40); "Depreciation To-Date "; USING "$$######,#.##"; FaItemRec(1).DEP2DATE
    PRINT #RptHandle, "Status: "; FaItemRec(1).ISTATUS;
    PRINT #RptHandle, TAB(40); "G/L Dep Exp Acct "; FaItemRec(1).GLACCT
    PRINT #RptHandle, "Asset Group: "; FaItemRec(1).AssetCode
    PRINT #RptHandle, STRING$(79, "-")
    
    'SubTotal Here
    LineCnt& = LineCnt& + 8
    ItemCnt& = ItemCnt& + 1
    OrigCost# = OrigCost# + FaItemRec(1).ORGCOST
    DepTot# = DepTot# + FaItemRec(1).DEP2DATE
    
SkipEm:
  NEXT
  
  GOSUB PrintMasterValueEnding
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi
  CLOSE         'Close all open files now
  
  IF DevSpec$ = "P" THEN
    EntryPoint = 4
  ELSEIF DevSpec$ = "S" THEN
    EntryPoint = 2
  ELSE
    EntryPoint = 1
  END IF
  
  ERASE Array, Frm, Form$, Fld, FaItemRec
  
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  
  KILL ReportFile$
  
  EXIT SUB
  
PrintMasterHeader:
  Page = Page + 1
  PRINT #RptHandle, TAB(29); "Master Asset Listing"
  PRINT #RptHandle, "Dept # "; Dept$
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
  PRINT #RptHandle, Dash80$
  LineCnt& = 6
RETURN
  
PrintMasterValueEnding:
  PRINT #RptHandle, Dash80$
  PRINT #RptHandle, "Dept Number: "; Dept$
  PRINT #RptHandle, "Total Items Printed: "; USING "#####"; ItemCnt&
  PRINT #RptHandle, "Total Original Cost: "; USING "$$########,#.##"; OrigCost#
  PRINT #RptHandle, "Total Deprc to Date: "; USING "$$########,#.##"; DepTot#
  PRINT #RptHandle, FF$
RETURN
  
GetIndex:
  TagFlag = LEFT$(Index$, 3) = "TAG"
  REDIM Array(1 TO NumOfFaRecs) AS Struct
  FOR Cnt& = 1 TO NumOfFaRecs
    GET FAFile, Cnt&, FaItemRec(1)
    Array(Cnt&).RecNum = Cnt&
    IF TagFlag THEN
      Array(Cnt&).Who = LTRIM$(LEFT$(UCASE$(FaItemRec(1).ITEMTAG), 14))
    ELSE
      RSET Array(Cnt&).Who = LTRIM$(RTRIM$(UCASE$(FaItemRec(1).IDEPT)) + RTRIM$(FaItemRec(1).AssetCode))
    END IF
  NEXT
  
  'Sort Them Here
  SortT Array(1), NumOfFaRecs, 0, LEN(Array(1)), 0, 14
RETURN
  
END SUB

