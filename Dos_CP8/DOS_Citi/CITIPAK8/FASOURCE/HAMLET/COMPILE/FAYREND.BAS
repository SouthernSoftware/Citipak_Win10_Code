DEFINT A-Z
DECLARE SUB YearEndPrint ()
DECLARE SUB YearEndPost ()
DECLARE SUB YearEndProcessing ()
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
DECLARE FUNCTION QPValL& (Number$)
DECLARE FUNCTION QPStrI$ (Num%)
DECLARE FUNCTION QPStrL$ (Num&)
DECLARE FUNCTION QPTrim$ (Text$)
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
  DIM SHARED FaYear(1) AS FAYearEndType
  DIM SHARED FaDep(1) AS FADepFileType
  DIM SHARED FASetup AS FASetupRecType

  STACK 5000
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 5)
  
  MChoice$(1) = " Build Depreciation File "
  MChoice$(2) = " Print Depreciation Build File"
  MChoice$(3) = " Post Depreciation to Assets"
  MChoice$(4) = " Exit to OS"
  
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
    
    TitleBox 2, Col, MaxLen + 3, "Fixed Assets Year End Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      YearEndProcessing
    CASE 2
      YearEndPrint
    CASE 3
      YearEndPost
    CASE 4
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

SUB YearEndPost
  
  SHARED Choice$()
  
  REDIM Choice$(0 TO 4, 0)
  
  ItemRecLen = LEN(FaItemRec(1))
  
  YrFile = FREEFILE
  OPEN "FAYEAR.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS YrFile LEN = LEN(FaYear(1))
  IF LOF(YrFile) / LEN(FaYear(1)) = 0 THEN
    LastYr$ = "N/A"
  ELSE
    GET YrFile, 1, FaYear(1)
    LastYr$ = FaYear(1).CurYear
  END IF
  CLOSE YrFile
  
  LibName$ = "FA"
  ScrnName$ = "YEAREND"
  
  '--define the multi-choice fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  '--Set screen number to one and display screen
  DisplayFAScrn "BAKCLEAR"
  DisplayFAScrn ScrnName$
  
  ShowCursor
  FirstTime = True
  
  
  Form$(1, 0) = LastYr$
  
  Action = 1
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE -68    'F10Key
      CurYear$ = Form$(2, 0)
      GOSUB ProcessDepreciation1
      DisplayFAScrn "UPDATEOK"
      WaitForAction
      
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 22   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 52 TO 61           '--Save Button
          PressButton -68, 22, 52, 61
        CASE 63 TO 74           '--cancel button
          PressButton 27, 22, 63, 74
        END SELECT
      END SELECT                'row
    END IF
    '--Check screen page
  LOOP UNTIL Frm(1).KeyCode = 27 OR ExitFlag
  
  EXIT SUB
  
ProcessDepreciation1:
  ExitFlag = True               'assume all is well
  
  'put some validation here
  IF LEN(QPTrim$(CurYear$)) <> 4 THEN
    GOTO BadParm1
  END IF
  DepFile = FREEFILE
  OPEN "FADEPEDT.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #DepFile LEN = LEN(FaDep(1))
  NumOfDepRecs = LOF(DepFile) / LEN(FaDep(1))
  IF NumOfDepRecs = 0 THEN
    FldNo = 2
    GOTO BadParm2
  END IF
  
  
  'Build Deprecition File
  'Open Deprec Edit File
  DepFile = FREEFILE
  OPEN "FADEPEDT.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #DepFile LEN = LEN(FaDep(1))
  NumOfDepRecs = LOF(DepFile) / LEN(FaDep(1))
  'Open Item File
  FAItemRecLen = LEN(FaItemRec(1))
  FaFile = FREEFILE
  OPEN FAItemFile FOR RANDOM AS FaFile LEN = FAItemRecLen
  NumOfFaRecs = LOF(FaFile) / FAItemRecLen
  
  FOR Cnt& = 1 TO NumOfDepRecs
    GET DepFile, Cnt&, FaDep(1)
    ItemRecNo = FaDep(1).AssetRecord
    GET FaFile, ItemRecNo, FaItemRec(1)
    FaItemRec(1).DEP2DATE = FaItemRec(1).DEP2DATE + FaDep(1).CurYrDep
    FaItemRec(1).CDEPDATE = Date2Num%("12-31-" + CurYear$)
    PUT FaFile, ItemRecNo, FaItemRec(1)
  NEXT Cnt&
  CLOSE
  'Now Update Processing Year
  YrFile = FREEFILE
  OPEN "FAYEAR.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS YrFile LEN = LEN(FaYear(1))
  GET YrFile, 1, FaYear(1)
  FaYear(1).LastYear = FaYear(1).CurYear
  FaYear(1).CurYear = CurYear$
  PUT YrFile, 1, FaYear(1)
  CLOSE YrFile
  'Now Clear Edit File
  KILL "FADEPEDT.DAT"
  RETURN
  
BadParm1:
  Ok = MsgBox("FA", "NOBLANK")
  Action = 1
  ExitFlag = False
  Frm(1).FldNo = FldNo
  EXIT SUB
  
BadParm2:
  Ok = MsgBox("FA", "NOBLANK")
  Action = 1
  ExitFlag = False
  Frm(1).FldNo = FldNo
  
END SUB

SUB YearEndPrint
  
  SHARED Choice$()
  DIM DOrigCost#(2), DBookTotal#(2), DCDep#(2), DYDep#(2), OrigCost#(2), BookTotal#(2), CDep#(2), YDep#(2)
  
  YrFile = FREEFILE
  OPEN "FAYEAR.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS YrFile LEN = LEN(FaYear(1))
  IF LOF(YrFile) / LEN(FaYear(1)) = 0 THEN
    LastYr$ = "N/A"
  ELSE
    GET YrFile, 1, FaYear(1)
    LastYr$ = FaYear(1).CurYear
  END IF
  CLOSE YrFile
  
  
  REDIM Array(1 TO 1) AS Struct 'Template for the sort array
  ReportFile$ = "FADEPEDT.PRN"  'Report File Name
  Dash80$ = STRING$(80, "=")
  FF$ = CHR$(12)
  
  MaxLines = 50
  LineCnt& = 0
  ItemCnt& = 0
  
  FAItemRecLen = LEN(FaItemRec(1))
  
  RptHandle = FREEFILE
  
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  GOSUB PrintMasterHeader3
  
  FaFile = FREEFILE
  OPEN FAItemFile FOR RANDOM AS FaFile LEN = FAItemRecLen
  NumOfFaRecs = LOF(FaFile) / FAItemRecLen
  
  'Open Deprec Edit File
  DepFile = FREEFILE
  OPEN "FADEPEDT.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #DepFile LEN = LEN(FaDep(1))
  NumOfDepRecs = LOF(DepFile) / LEN(FaDep(1))
  
  FOR Cnt& = 1 TO NumOfDepRecs
    GET DepFile, Cnt&, FaDep(1)
    ItemRecNo = FaDep(1).AssetRecord
    GET FaFile, ItemRecNo, FaItemRec(1)
    IF Cnt& = 1 THEN
      DeptNumber = VAL(FaItemRec(1).IDEPT)
    END IF
    PValue# = 0
    
    IF LineCnt& >= MaxLines THEN
      PRINT #RptHandle, FF$
      GOSUB PrintMasterHeader3
    END IF
    
    IF DeptNumber <> VAL(FaItemRec(1).IDEPT) THEN
      'Print Subtotals and Clear
      PRINT #RptHandle, STRING$(105, "-")
      PRINT #RptHandle, "Totals for Dept Number: "; DeptNumber;
      PRINT #RptHandle, TAB(62); USING "########,#.##"; DOrigCost#;
      PRINT #RptHandle, TAB(78); USING "########,#.##"; DCurDep#;
      'IF FaDep(1).PctFlag THEN
      '  PRINT #RptHandle, "*";
      'END IF
      PRINT #RptHandle, TAB(92); USING "########,#.##"; DYDep#
      LineCnt& = LineCnt& + 1
      
      PRINT #RptHandle, "": LineCnt& = LineCnt& + 1
      PRINT #RptHandle, "": LineCnt& = LineCnt& + 1
      
      'GOSUB PrintMasterHeader1
      DeptNumber = VAL(FaItemRec(1).IDEPT)
      DOrigCost# = 0
      DCurDep# = 0
      DYDep# = 0
    END IF
    
    
    'Figure Values
    'Calc Depreciation for This Period
    YTDDep# = FaDep(1).CurYrDep + FaItemRec(1).DEP2DATE
    PRINT #RptHandle, FaItemRec(1).ITEMTAG; TAB(22); LEFT$(FaItemRec(1).IDESC1, 28);
    PRINT #RptHandle, TAB(51); FaItemRec(1).IDEPT;
    PRINT #RptHandle, TAB(58); USING "###"; FaItemRec(1).ILIFE;
    PRINT #RptHandle, TAB(62); USING "########,#.##"; FaItemRec(1).ORGCOST;
    PRINT #RptHandle, TAB(78); USING "########,#.##"; FaDep(1).CurYrDep;
    IF FaDep(1).PctFlag = -1 THEN
      PRINT #RptHandle, "*";
    ELSEIF FaDep(1).PctFlag = -1 THEN
      PRINT #RptHandle, "D";
    END IF
    PRINT #RptHandle, TAB(92); USING "########,#.##"; YTDDep#
    'SubTotal Here
    LineCnt& = LineCnt& + 1
    ItemCnt& = ItemCnt& + 1
    'Grand Totals Here
    OrigCost# = OrigCost# + FaItemRec(1).ORGCOST
    CurDep# = CurDep# + FaDep(1).CurYrDep
    YDep# = YDep# + YTDDep#
    'Dept Totals Here
    DOrigCost# = DOrigCost# + FaItemRec(1).ORGCOST
    DCurDep# = DCurDep# + FaDep(1).CurYrDep
    DYDep# = DYDep# + YTDDep#
    
    
SkipEm3:
  NEXT Cnt&
  'First Print Subtotals
  PRINT #RptHandle, STRING$(105, "-")
  PRINT #RptHandle, "Totals for Dept Number: "; DeptNumber;
  PRINT #RptHandle, TAB(62); USING "########,#.##"; DOrigCost#;
  PRINT #RptHandle, TAB(78); USING "########,#.##"; DCurDep#;
  PRINT #RptHandle, TAB(92); USING "########,#.##"; DYDep#
  
  LineCnt& = LineCnt& + 1
  
  PRINT #RptHandle, "": LineCnt& = LineCnt& + 1
  PRINT #RptHandle, "": LineCnt& = LineCnt& + 1
  
  GOSUB PrintDepRepEnding1
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi
  CLOSE         'Close all open files now
  
  EntryPoint = 2

  CursorOff
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  'KILL ReportFile$
  
  EXIT SUB
  
PrintMasterHeader3:
  Page = Page + 1
  PRINT #RptHandle, TAB(20); "Master Asset Listing : Depreciation Edit Report"
  PRINT #RptHandle,
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(68); "Page #"; Page
  PRINT #RptHandle, "Asset Number"; TAB(22); "Description"; TAB(51); "Dept"; TAB(58); "Life"; TAB(64); "Original Cost"; TAB(80); "Cur Yr Deprec"; TAB(95); "YTD Deprec"
  PRINT #RptHandle, STRING$(105, "=")
  LineCnt& = 6
  RETURN
  
PrintDepRepEnding1:
  PRINT #RptHandle, STRING$(105, "-")
  PRINT #RptHandle, "Grand Totals ";
  PRINT #RptHandle, TAB(62); USING "########,#.##"; OrigCost#;
  PRINT #RptHandle, TAB(78); USING "########,#.##"; CurDep#;
  PRINT #RptHandle, TAB(92); USING "########,#.##"; YDep#
  
  PRINT #RptHandle, FF$
  RETURN
  
  
  
END SUB

SUB YearEndProcessing
  
  SHARED Choice$()
  REDIM Choice$(0 TO 4, 0)
  
  FASetupLen = LEN(FASetup)
  FASetupFile = FREEFILE
  OPEN "fasetup.dat" FOR RANDOM SHARED AS FASetupFile LEN = FASetupLen
  SetupSize = LOF(FASetupFile)
  IF SetupSize > 0 THEN
    GET FASetupFile, 1, FASetup
    PRate1StY$ = FASetup.PRate1St
    Pct1StYear = FASetup.Pct1St
    IF PRate1StY$ = "Y" THEN
      PRate = True
      Pct1StYear = 100
    ELSEIF Pct1StYear < 0 THEN
      Pct1StYear = 100
    END IF
  ELSE
    Pct1StYear = 100
  END IF
  CLOSE FASetupFile

'  IF Pct1StYear < 100 THEN
'    UsePct = True
'    Pct1St# = Round#(Pct1StYear * .01)
'  END IF

  ItemRecLen = LEN(FaItemRec(1))
  YrFile = FREEFILE
  OPEN "FAYEAR.DAT" FOR RANDOM SHARED AS YrFile LEN = LEN(FaYear(1))
  IF LOF(YrFile) / LEN(FaYear(1)) = 0 THEN
    LastYr$ = "N/A"
  ELSE
    GET YrFile, 1, FaYear(1)
    LastYr$ = FaYear(1).CurYear
  END IF
  CLOSE YrFile
  
  LibName$ = "FA"
  ScrnName$ = "YEAREND"
  
  '--define the multi-choice fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  '--Set screen number to one and display screen
  DisplayFAScrn "BAKCLEAR"
  DisplayFAScrn ScrnName$
  
  ShowCursor
  FirstTime = True
  
  Form$(1, 0) = LastYr$
  Action = 1
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE -68    'F10Key
      IF LEN(QPTrim$(Form$(2, 0))) = 4 THEN
        CurYear$ = Form$(2, 0)
        ProcessThru$ = QPTrim$(Form$(3, 0))
        ProcThruMonth = VAL(LEFT$(ProcessThru$, 2))
        ProcessFrom$ = LEFT$(ProcessThru$, 6) + QPTrim$(STR$(VAL(RIGHT$(ProcessThru$, 4)) - 1))

        PDate = Date2Num(ProcessThru$)
        ThruDate = PDate
        FromDate = Date2Num(ProcessFrom$)

        GOSUB ProcessDepreciation
        DisplayFAScrn "UPDATEOK"
        WaitForAction
      END IF
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 22   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 52 TO 61           '--Save Button
          PressButton -68, 22, 52, 61
        CASE 63 TO 74           '--cancel button
          PressButton 27, 22, 63, 74
        END SELECT
      END SELECT                'row
    END IF
    '--Check screen page
  LOOP UNTIL Frm(1).KeyCode = 27 OR ExitFlag
  
  EXIT SUB
  
ProcessDepreciation:
  ExitFlag = True               'assume all is well
  'put some validation here
  IF LEN(CurYear$) <> 4 THEN
    FldNo = 2
    GOTO BadParm
  END IF
  
  PSDate = Date2Num("01-01-" + RIGHT$(ProcessThru$, 4))
  PEDate = PDate

  'Build Deprecition File
  'Build Work File
  DepFile = FREEFILE
  OPEN "FADEPEDT.DAT" FOR RANDOM SHARED AS #DepFile LEN = LEN(FaDep(1))
  CLOSE DepFile
  KILL "FADEPEDT.DAT"
  'Open Deprec Edit File
  DepFile = FREEFILE
  OPEN "FADEPEDT.DAT" FOR RANDOM SHARED AS #DepFile LEN = LEN(FaDep(1))
  'Open Item File
  FAItemRecLen = LEN(FaItemRec(1))
  FaFile = FREEFILE
  OPEN FAItemFile FOR RANDOM AS FaFile LEN = FAItemRecLen
  NumOfFaRecs = LOF(FaFile) / FAItemRecLen
  
  'Sort Items by Dept
  REDIM Array(1 TO NumOfFaRecs) AS Struct

  FOR Cnt& = 1 TO NumOfFaRecs
    GET FaFile, Cnt&, FaItemRec(1)
    IF FaItemRec(1).AQURDATE <= PDate THEN
      Array(Cnt&).who = LTRIM$(LEFT$(UCASE$(FaItemRec(1).IDEPT), 4))
      Array(Cnt&).RecNum = Cnt&
    ELSE
      Array(Cnt&).who = "N/A"
      Array(Cnt&).RecNum = 0
    END IF
  NEXT
  
  'Sort Them Here
  SortT Array(1), NumOfFaRecs, 0, LEN(Array(1)), 0, 14
  
  FOR Cnt& = 1 TO NumOfFaRecs
    ItemRecNo = Array(Cnt&).RecNum
    IF ItemRecNo = 0 THEN
      GOTO SkipThisAsset               'Skip Assets Not in Purchase Range
    END IF
    GET FaFile, ItemRecNo, FaItemRec(1)
    'Calc Depreciation for This Period
    IF FaItemRec(1).ILIFE > 0 AND FaItemRec(1).DEPYN <> "N" THEN
      CurDep# = Round#(FaItemRec(1).ORGCOST / FaItemRec(1).ILIFE)
      IAqurDate = FaItemRec(1).AQURDATE
      IF IAqurDate > FromDate AND IAqurDate <= ThruDate THEN
        IADate$ = Num2Date(IAqurDate)
        IAMonth = VAL(LEFT$(IADate$, 2))
        SELECT CASE IAMonth
        CASE IS < ProcThruMonth
          DepMonths = (ProcThruMonth - IAMonth) + 1
        CASE IS > ProcThruMonth
          DepMonths = ((12 - IAMonth) + ProcThruMonth) + 1
        END SELECT
        DepPerMonth# = Round(CurDep# / 12)
        CurDep# = Round#(DepPerMonth# * DepMonths)
        FaDep(1).PctFlag = True
      ELSE
        FaDep(1).PctFlag = False
      END IF

      MaxDep# = Round#(FaItemRec(1).ORGCOST - FaItemRec(1).DEP2DATE)
      IF Round#(MaxDep# + FaItemRec(1).DEP2DATE) > FaItemRec(1).ORGCOST THEN
        MaxDep# = Round#(MaxDep# - (Round#(MaxDep# + FaItemRec(1).DEP2DATE) - FaItemRec(1).ORGCOST))
        'FaDep(1).PctFlag = 1
      END IF

      IF MaxDep# < 0 THEN MaxDep# = 0

      IF CurDep# > MaxDep# THEN CurDep# = MaxDep#
      IF Round#(CurDep#) >= .01 THEN
        FaDep(1).AssetRecord = ItemRecNo
        FaDep(1).CurYrDep = Round#(CurDep#)
        NextEditRecord! = NextEditRecord! + 1
        PUT DepFile, NextEditRecord!, FaDep(1)
      END IF
    END IF
    
SkipThisAsset:
  NEXT Cnt&
  CLOSE
RETURN
  
BadParm:
  Ok = MsgBox("FA", "NOBLANK")
  Action = 1
  ExitFlag = False
  Frm(1).FldNo = FldNo
RETURN

END SUB

