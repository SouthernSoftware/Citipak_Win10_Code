DEFINT A-Z
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
    who AS STRING * 14
    RecNum AS INTEGER
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
  REDIM MChoice$(1 TO 6)

  MChoice$(1) = " Master Item Listing - (8.5 Format)"
  MChoice$(2) = " Item Check List Report"
  MChoice$(3) = " New/Deleted Item Report"
  MChoice$(4) = " Item Depreciation Report "
  MChoice$(5) = " Exit to OS"

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
  ReportFile$ = "FACHK.PRN"   'Report File Name
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
  Form$(4, 0) = "N/A"'CYear$
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

  GOSUB PrintMasterHeader2

  FAFile = FREEFILE
  OPEN FAItemFile FOR RANDOM AS FAFile LEN = FAItemRecLen
  NumOfFaRecs = LOF(FAFile) / FAItemRecLen

  GOSUB GetIndex2


  FOR Cnt& = 1 TO NumOfFaRecs



    ItemRecNo = Array(Cnt&).RecNum
    GET FAFile, ItemRecNo, FaItemRec(1)
    IF Cnt& = 1 THEN DeptNumber = VAL(FaItemRec(1).IDept)
    PValue# = 0

      IF LineCnt& >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintMasterHeader2
      END IF

      'Check For Disposed Of

      IF Dispose$ = "N" THEN
       IF FaItemRec(1).DispDate > 0 THEN GOTO SkipEm2
      END IF

      IF Dept$ = "ALL" THEN
       ELSE
       IF VAL(Dept$) <> VAL(FaItemRec(1).IDept) THEN GOTO SkipEm2
      END IF

       IF RTRIM$(Index$) = "DEPT" AND RTRIM$(Dept$) = "ALL" THEN
        IF DeptNumber <> VAL(FaItemRec(1).IDept) THEN
          PRINT #RptHandle, FF$
          GOSUB PrintMasterHeader2
          DeptNumber = VAL(FaItemRec(1).IDept)
        END IF
       END IF


      
        'Figure Values

        IF FaItemRec(1).DispDate > 0 THEN Disp$ = "Y" ELSE Disp$ = "N"
              
        PRINT #RptHandle, FaItemRec(1).ItemTag;
         PRINT #RptHandle, TAB(22); RTRIM$(FaItemRec(1).IDesc1);
         PRINT #RptHandle, TAB(53); LEFT$(QPTrim$(FaItemRec(1).ItemLoc), 24);
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
  PRINT #RptHandle, "Dept # "; Dept$
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
    Array(Cnt&).who = LTRIM$(LEFT$(UCASE$(FaItemRec(1).ItemTag), 14))
     ELSE
    Array(Cnt&).who = LTRIM$(LEFT$(UCASE$(FaItemRec(1).IDept), 4))
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
  ReportFile$ = "FADEPREP.PRN"   'Report File Name
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
  Form$(2, 0) = "Y"
  Form$(3, 0) = "ALL"
  CYear = VAL(RIGHT$(DATE$, 4))
  CYear$ = LTRIM$(STR$(CYear))
  Form$(4, 0) = CYear$
  Fld(1).Protected = 1
  Fld(2).Protected = 1
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
      Dispose$ = Form$(2, 0)
      Dept$ = RTRIM$(Form$(3, 0))
      CYear$ = RTRIM$(Form$(4, 0))
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
    IF Cnt& = 1 THEN DeptNumber = VAL(FaItemRec(1).IDept)
    PValue# = 0

      IF LineCnt& >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintMasterHeader3
      END IF

      
      IF Dept$ = "ALL" THEN
       ELSE
       IF VAL(Dept$) <> VAL(FaItemRec(1).IDept) THEN GOTO SkipEm3
      END IF

       IF RTRIM$(Index$) = "DEPT" AND RTRIM$(Dept$) = "ALL" THEN
        IF DeptNumber <> VAL(FaItemRec(1).IDept) THEN
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
           DeptNumber = VAL(FaItemRec(1).IDept)
           DOrigCost# = 0
           DBookTotal# = 0
           DYDep# = 0
          END IF
       END IF

      'Figure Values
      'Calc Depreciation for This Period
         YTDDep# = CurDep# + FaItemRec(1).Dep2Date
         BookVal# = FaItemRec(1).OrgCost - FaItemRec(1).Dep2Date
         IF BookVal# < 0 THEN BookVal# = 0

        PRINT #RptHandle, FaItemRec(1).ItemTag; TAB(22); LEFT$(FaItemRec(1).IDesc1, 28);
        PRINT #RptHandle, TAB(51); FaItemRec(1).IDept;
        PRINT #RptHandle, TAB(58); USING "###"; FaItemRec(1).ILife;
        PRINT #RptHandle, TAB(62); USING "########,#.##"; FaItemRec(1).OrgCost;
        PRINT #RptHandle, TAB(78); USING "########,#.##"; BookVal#;
        PRINT #RptHandle, TAB(92); USING "########,#.##"; YTDDep#


       'SubTotal Here
         LineCnt& = LineCnt& + 1
         ItemCnt& = ItemCnt& + 1
       'Grand Totals Here
         OrigCost# = OrigCost# + FaItemRec(1).OrgCost
         BookTotal# = BookTotal# + BookVal#
         YDep# = YDep# + YTDDep#
       'Dept Totals Here
         DOrigCost# = DOrigCost# + FaItemRec(1).OrgCost
         DBookTotal# = DBookTotal# + BookVal#
         DYDep# = DYDep# + YTDDep#
         
       
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
    IF LEFT$(Index$, 3) = "TAG" THEN
    Array(Cnt&).who = LTRIM$(LEFT$(UCASE$(FaItemRec(1).ItemTag), 14))
     ELSE
    Array(Cnt&).who = LTRIM$(LEFT$(UCASE$(FaItemRec(1).IDept), 4))
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
  ReportFile$ = "FAMaster.PRN"   'Report File Name
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
  Form$(2, 0) = "Y"
  Form$(3, 0) = "ALL"
  CYear = VAL(RIGHT$(DATE$, 4))
  CYear$ = LTRIM$(STR$(CYear))
  Form$(4, 0) = CYear$
  Fld(1).Protected = 1
  Fld(2).Protected = 1
  Action = 1

  BlockClear
  ShowCursor

  DisplayFAScrn ScrnName$
  LOCATE 6, 22: COLOR 14: PRINT "New/Deleted "


  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      Index$ = Form$(1, 0)
      Dispose$ = Form$(2, 0)
      Dept$ = RTRIM$(Form$(3, 0))
      CYear$ = RTRIM$(Form$(4, 0))
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
    IF Cnt& = 1 THEN DeptNumber = VAL(FaItemRec(1).IDept)
    PValue# = 0

      IF LineCnt& >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintMasterHeader1
      END IF

      'Check For Disposed Date
      DisposeDate$ = Num2Date$(FaItemRec(1).DispDate)
      IF LEN(DisposeDate$) > 4 THEN DisposeDate$ = RIGHT$(DisposeDate$, 4)
      'Check for Acquired Date
      AcquireDate$ = Num2Date$(FaItemRec(1).AQURDate)
      IF LEN(AcquireDate$) > 4 THEN AcquireDate$ = RIGHT$(AcquireDate$, 4)



      IF DisposeDate$ = CYear$ OR AcquireDate$ = CYear$ THEN
       IF DisposeDate$ = CYear$ THEN DFlag = 1
       IF AcquireDate$ = CYear$ THEN AFlag = 1
       ELSE
       GOTO SkipEm1
      END IF

      IF Dept$ = "ALL" THEN
       ELSE
       IF VAL(Dept$) <> VAL(FaItemRec(1).IDept) THEN GOTO SkipEm1
      END IF

       IF RTRIM$(Index$) = "DEPT" AND RTRIM$(Dept$) = "ALL" THEN
        IF DeptNumber <> VAL(FaItemRec(1).IDept) THEN
         'Print Subtotals and Clear
          PRINT #RptHandle, STRING$(120, "-")
          PRINT #RptHandle, "Current for Dept Number: "; DeptNumber;
          PRINT #RptHandle, TAB(64); USING "######,#.##"; DOrigCost#;
          PRINT #RptHandle, TAB(80); USING "######,#.##"; DBookTotal#;
          PRINT #RptHandle, TAB(94); USING "######,#.##"; DCDep#;
          PRINT #RptHandle, TAB(108); USING "######,#.##"; DYDep#

          PRINT #RptHandle, "Additions for Dept Number: "; DeptNumber;
          PRINT #RptHandle, TAB(64); USING "######,#.##"; DOrigCost#(2);
          PRINT #RptHandle, TAB(80); USING "######,#.##"; DBookTotal#(2);
          PRINT #RptHandle, TAB(94); USING "######,#.##"; DCDep#(2);
          PRINT #RptHandle, TAB(108); USING "######,#.##"; DYDep#(2)

          PRINT #RptHandle, "Deletions for Dept Number: "; DeptNumber;
          PRINT #RptHandle, TAB(64); USING "######,#.##"; DOrigCost#(1);
          PRINT #RptHandle, TAB(80); USING "######,#.##"; DBookTotal#(1);
          PRINT #RptHandle, TAB(94); USING "######,#.##"; DCDep#(1);
          PRINT #RptHandle, TAB(108); USING "######,#.##"; DYDep#(1)

          PRINT #RptHandle, "Totals for Dept Number: "; DeptNumber;
          PRINT #RptHandle, TAB(64); USING "######,#.##"; DOrigCost# + DOrigCost#(1) + DOrigCost#(2);
          PRINT #RptHandle, TAB(80); USING "######,#.##"; DBookTotal# + DBookTotal#(1) + DBookTotal#(2);
          PRINT #RptHandle, TAB(94); USING "######,#.##"; DCDep# + DCDep#(1) + DCDep#(2);
          PRINT #RptHandle, TAB(108); USING "######,#.##"; DYDep# + DYDep#(1) + DYDep#(2)
          LineCnt& = LineCnt& + 4
          
          PRINT #RptHandle, "": LineCnt& = LineCnt& + 1
          PRINT #RptHandle, "": LineCnt& = LineCnt& + 1


          'GOSUB PrintMasterHeader1
           DeptNumber = VAL(FaItemRec(1).IDept)
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
        IF FaItemRec(1).ILife > 0 THEN
         CurDep# = FaItemRec(1).OrgCost / FaItemRec(1).ILife
         MaxDep# = FaItemRec(1).OrgCost - FaItemRec(1).Dep2Date
         IF CurDep# > MaxDep# THEN CurDep# = MaxDep#
         YTDDep# = CurDep# + FaItemRec(1).Dep2Date
        ELSE
         CurDep# = 0
         YTDDep# = 0
        END IF

        PRINT #RptHandle, FaItemRec(1).ItemTag; TAB(22); LEFT$(FaItemRec(1).IDesc1, 28);
        PRINT #RptHandle, TAB(51); FaItemRec(1).IDept;
        PRINT #RptHandle, TAB(58); USING "###"; FaItemRec(1).ILife;
        PRINT #RptHandle, TAB(64); USING "######,#.##"; FaItemRec(1).OrgCost;
        PRINT #RptHandle, TAB(80); USING "######,#.##"; FaItemRec(1).OrgCost - FaItemRec(1).Dep2Date;
        PRINT #RptHandle, TAB(94); USING "######,#.##"; CurDep#;
        PRINT #RptHandle, TAB(108); USING "######,#.##"; YTDDep#
       

       'SubTotal Here
         LineCnt& = LineCnt& + 1
         ItemCnt& = ItemCnt& + 1
       'Grand Totals Here
         IF DFlag = 0 AND AFlag = 0 THEN
         OrigCost# = OrigCost# + FaItemRec(1).OrgCost
         BookTotal# = BookTotal# + (FaItemRec(1).OrgCost - FaItemRec(1).Dep2Date)
         CDep# = CDep# + CurDep#
         YDep# = YDep# + YTDDep#
         END IF
         IF DFlag = 1 THEN
          OrigCost#(1) = OrigCost#(1) + FaItemRec(1).OrgCost
          BookTotal#(1) = BookTotal#(1) + (FaItemRec(1).OrgCost - FaItemRec(1).Dep2Date)
          CDep#(1) = CDep#(1) + CurDep#
          YDep#(1) = YDep#(1) + YTDDep#
         END IF
         IF AFlag = 1 THEN
          OrigCost#(2) = OrigCost#(2) + FaItemRec(1).OrgCost
          BookTotal#(2) = BookTotal#(2) + (FaItemRec(1).OrgCost - FaItemRec(1).Dep2Date)
          CDep#(2) = CDep#(2) + CurDep#
          YDep#(2) = YDep#(2) + YTDDep#
         END IF

       'Dept Totals Here
         IF DFlag = 0 AND AFlag = 0 THEN
          DOrigCost# = DOrigCost# + FaItemRec(1).OrgCost
          DBookTotal# = DBookTotal# + (FaItemRec(1).OrgCost - FaItemRec(1).Dep2Date)
          DCDep# = DCDep# + CurDep#
          DYDep# = DYDep# + YTDDep#
         END IF
         IF DFlag = 1 THEN
          DOrigCost#(1) = DOrigCost#(1) + FaItemRec(1).OrgCost
          DBookTotal#(1) = DBookTotal#(1) + (FaItemRec(1).OrgCost - FaItemRec(1).Dep2Date)
          DCDep#(1) = DCDep#(1) + CurDep#
          DYDep#(1) = DYDep#(1) + YTDDep#
         END IF
         IF AFlag = 1 THEN
          DOrigCost#(2) = DOrigCost#(2) + FaItemRec(1).OrgCost
          DBookTotal#(2) = DBookTotal#(2) + (FaItemRec(1).OrgCost - FaItemRec(1).Dep2Date)
          DCDep#(2) = DCDep#(2) + CurDep#
          DYDep#(2) = DYDep#(2) + YTDDep#
         END IF



SkipEm1:
      NEXT Cnt&
      'First Print Subtotals
          PRINT #RptHandle, STRING$(120, "-")
          PRINT #RptHandle, "Current for Dept Number: "; DeptNumber;
          PRINT #RptHandle, TAB(64); USING "######,#.##"; DOrigCost#;
          PRINT #RptHandle, TAB(80); USING "######,#.##"; DBookTotal#;
          PRINT #RptHandle, TAB(94); USING "######,#.##"; DCDep#;
          PRINT #RptHandle, TAB(108); USING "######,#.##"; DYDep#

          PRINT #RptHandle, "Additions for Dept Number: "; DeptNumber;
          PRINT #RptHandle, TAB(64); USING "######,#.##"; DOrigCost#(2);
          PRINT #RptHandle, TAB(80); USING "######,#.##"; DBookTotal#(2);
          PRINT #RptHandle, TAB(94); USING "######,#.##"; DCDep#(2);
          PRINT #RptHandle, TAB(108); USING "######,#.##"; DYDep#(2)

          PRINT #RptHandle, "Deletions for Dept Number: "; DeptNumber;
          PRINT #RptHandle, TAB(64); USING "######,#.##"; DOrigCost#(1);
          PRINT #RptHandle, TAB(80); USING "######,#.##"; DBookTotal#(1);
          PRINT #RptHandle, TAB(94); USING "######,#.##"; DCDep#(1);
          PRINT #RptHandle, TAB(108); USING "######,#.##"; DYDep#(1)

          PRINT #RptHandle, "Totals for Dept Number: "; DeptNumber;
          PRINT #RptHandle, TAB(64); USING "######,#.##"; DOrigCost# + DOrigCost#(1) + DOrigCost#(2);
          PRINT #RptHandle, TAB(80); USING "######,#.##"; DBookTotal# + DBookTotal#(1) + DBookTotal#(2);
          PRINT #RptHandle, TAB(94); USING "######,#.##"; DCDep# + DCDep#(1) + DCDep#(2);
          PRINT #RptHandle, TAB(108); USING "######,#.##"; DYDep# + DYDep#(1) + DYDep#(2)
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
  PRINT #RptHandle, "Accounting Year: "; CYear$
  PRINT #RptHandle, "Dept # "; Dept$
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
  PRINT #RptHandle, "Asset Number"; TAB(22); "Description"; TAB(51); "Dept"; TAB(58); "Life"; TAB(64); "Original Cost"; TAB(81); "Book Value"; TAB(94); "Curr Deprec"; TAB(109); "YTD Deprec"
  PRINT #RptHandle, STRING$(120, "=")
  LineCnt& = 6
  RETURN

PrintMasterValueEnding1:
          PRINT #RptHandle, STRING$(120, "-")
          PRINT #RptHandle, "Current ";
          PRINT #RptHandle, TAB(64); USING "######,#.##"; OrigCost#;
          PRINT #RptHandle, TAB(80); USING "######,#.##"; BookTotal#;
          PRINT #RptHandle, TAB(94); USING "######,#.##"; CDep#;
          PRINT #RptHandle, TAB(108); USING "######,#.##"; YDep#

          PRINT #RptHandle, "Additions ";
          PRINT #RptHandle, TAB(64); USING "######,#.##"; OrigCost#(2);
          PRINT #RptHandle, TAB(80); USING "######,#.##"; BookTotal#(2);
          PRINT #RptHandle, TAB(94); USING "######,#.##"; CDep#(2);
          PRINT #RptHandle, TAB(108); USING "######,#.##"; YDep#(2)

          PRINT #RptHandle, "Deletions ";
          PRINT #RptHandle, TAB(64); USING "######,#.##"; OrigCost#(1);
          PRINT #RptHandle, TAB(80); USING "######,#.##"; BookTotal#(1);
          PRINT #RptHandle, TAB(94); USING "######,#.##"; CDep#(1);
          PRINT #RptHandle, TAB(108); USING "######,#.##"; YDep#(1)

          PRINT #RptHandle, "Grand Totals ";
          PRINT #RptHandle, TAB(64); USING "######,#.##"; OrigCost# + OrigCost#(1) + OrigCost#(2);
          PRINT #RptHandle, TAB(80); USING "######,#.##"; BookTotal# + BookTotal#(1) + BookTotal#(2);
          PRINT #RptHandle, TAB(94); USING "######,#.##"; CDep# + CDep#(1) + CDep#(2);
          PRINT #RptHandle, TAB(108); USING "######,#.##"; YDep# + YDep#(1) + YDep#(2)
          PRINT #RptHandle, FF$
          RETURN

GetIndex1:
  REDIM Array(1 TO NumOfFaRecs) AS Struct
  FOR Cnt& = 1 TO NumOfFaRecs
    GET FAFile, Cnt&, FaItemRec(1)
    IF LEFT$(Index$, 3) = "TAG" THEN
    Array(Cnt&).who = LTRIM$(LEFT$(UCASE$(FaItemRec(1).ItemTag), 14))
     ELSE
    Array(Cnt&).who = LTRIM$(LEFT$(UCASE$(FaItemRec(1).IDept), 4))
    END IF
    Array(Cnt&).RecNum = Cnt&
  NEXT

 'Sort Them Here
  SortT Array(1), NumOfFaRecs, 0, LEN(Array(1)), 0, 14
  RETURN

END SUB

SUB ItemValuation

END SUB

SUB MasterList
  SHARED Choice$()

  REDIM Array(1 TO 1) AS Struct 'Template for the sort array
  ReportFile$ = "FAMaster.PRN"   'Report File Name
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
  Form$(4, 0) = "N/A"'CYear$
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
    IF Cnt& = 1 THEN DeptNumber = VAL(FaItemRec(1).IDept)
    PValue# = 0
    
      IF LineCnt& >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintMasterHeader
      END IF
      
      'Check For Disposed Of

      IF Dispose$ = "N" THEN
       IF FaItemRec(1).DispDate > 0 THEN GOTO SkipEm
      END IF
      
      IF Dept$ = "ALL" THEN
       ELSE
       IF VAL(Dept$) <> VAL(FaItemRec(1).IDept) THEN GOTO SkipEm
      END IF

       IF RTRIM$(Index$) = "DEPT" AND RTRIM$(Dept$) = "ALL" THEN
        IF DeptNumber <> VAL(FaItemRec(1).IDept) THEN
         'Print Subtotals and Clear
          PRINT #RptHandle, "Subtotal for Dept Number: "; DeptNumber
          PRINT #RptHandle, "Total Items Printed: "; USING "#####"; DItemCnt&
          PRINT #RptHandle, "Total Original Cost: "; USING "$$########,#.##"; DOrigCost#
          PRINT #RptHandle, "Total Deprc to Date: "; USING "$$########,#.##"; DDepTot#
          PRINT #RptHandle, FF$
          GOSUB PrintMasterHeader
          DeptNumber = VAL(FaItemRec(1).IDept)
          DItemCnt& = 0
          DOrigCost# = 0
          DDepTot# = 0

        END IF
         'May Need to Skip Them
          DItemCnt& = DItemCnt& + 1
          DOrigCost# = DOrigCost# + FaItemRec(1).OrgCost
          DDepTot# = DDepTot# + FaItemRec(1).Dep2Date
       END IF





        'Figure Values

        IF FaItemRec(1).DispDate > 0 THEN Disp$ = "Y" ELSE Disp$ = "N"
        PRINT #RptHandle, "Tag #"; FaItemRec(1).ItemTag;
         PRINT #RptHandle, TAB(40); "Ser #"; FaItemRec(1).SerialNo
        PRINT #RptHandle, "Desc "; FaItemRec(1).IDesc1;
         PRINT #RptHandle, TAB(40); "Mfg  "; FaItemRec(1).ItemMfg
        PRINT #RptHandle, TAB(5); FaItemRec(1).IDesc2;
         PRINT #RptHandle, TAB(40); "Cont "; FaItemRec(1).Contact
        PRINT #RptHandle, "Loc  "; QPTrim$(FaItemRec(1).ItemLoc);
        PRINT #RptHandle, TAB(40); "Dept "; FaItemRec(1).IDept
        PRINT #RptHandle, "Orig Cost: "; USING "$$######,#.##"; FaItemRec(1).OrgCost;
         PRINT #RptHandle, TAB(40); "Acquired On "; Num2Date(FaItemRec(1).AQURDate)
        PRINT #RptHandle, "Asset Life: "; USING "###"; FaItemRec(1).ILife;
         PRINT #RptHandle, TAB(40); "Depreciation To-Date "; USING "$$######,#.##"; FaItemRec(1).Dep2Date
        PRINT #RptHandle, "Status: "; FaItemRec(1).IStatus;
         PRINT #RptHandle, TAB(40); "G/L Dep Exp Acct "; FaItemRec(1).GLAcct
        PRINT #RptHandle, STRING$(79, "-")

       'SubTotal Here
         LineCnt& = LineCnt& + 8
         ItemCnt& = ItemCnt& + 1
         OrigCost# = OrigCost# + FaItemRec(1).OrgCost
         DepTot# = DepTot# + FaItemRec(1).Dep2Date
         
        
        
SkipEm:
      NEXT Cnt&

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
  REDIM Array(1 TO NumOfFaRecs) AS Struct
  FOR Cnt& = 1 TO NumOfFaRecs
    GET FAFile, Cnt&, FaItemRec(1)
    IF LEFT$(Index$, 3) = "TAG" THEN
    Array(Cnt&).who = LTRIM$(LEFT$(UCASE$(FaItemRec(1).ItemTag), 14))
     ELSE
    Array(Cnt&).who = LTRIM$(LEFT$(UCASE$(FaItemRec(1).IDept), 4))
    END IF
    Array(Cnt&).RecNum = Cnt&
  NEXT

 'Sort Them Here
  SortT Array(1), NumOfFaRecs, 0, LEN(Array(1)), 0, 14
  RETURN

END SUB

