DEFINT A-Z
DECLARE SUB CodeLookUp (RecNo%, CLSFlag%, ActiveOnly%)
DECLARE SUB AddEditAssetCode (EditFlag%)
DECLARE SUB PrintCodeListing ()
DECLARE SUB FALookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, ActiveOnly%)
DECLARE SUB DisplayFAScrn (ScrnName$)
DECLARE SUB ItemMaint ()
DECLARE SUB AddEditItem (RecNo&)
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
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION Num2Date$ (TheDate%)
DECLARE SUB ubGetAcctStruct (GLFundLen%, GLAcctLen%, GLDetLen%)
DECLARE FUNCTION ubFindAcct% (AcctNum$)
DECLARE FUNCTION ubFmtAcct$ (AN$, FundLen%, AcctLen%, DetLen%)
DECLARE FUNCTION ubStripAcct$ (AcctNum$)
DECLARE FUNCTION ubAcctLookUp% (Acct$)
DECLARE SUB WaitForAction ()
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB QPrintRC (Text$, Row, Col, TheColor)


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
  '$INCLUDE: 'GL.BI'

  STACK 5000

  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 5)

  MChoice$(1) = " Add New Asset Code  "
  MChoice$(2) = " Edit Existing Code  "
  MChoice$(3) = " Print Code Listing  "
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
  Col = ((80 - MaxLen) \ 2) - 2

  DO

    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear

    TitleBox 2, Col, MaxLen + 3, "Assets Codes Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

    ShowCursor
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf

    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0

    SELECT CASE Choice
    CASE 1   'Adding
      AddEditAssetCode False
    CASE 2   'Editing
      CodeLookUp RecNo, True, ActiveOnly%
      IF RecNo > 0 THEN
        AddEditAssetCode RecNo
      END IF
    CASE 3
      PrintCodeListing
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

SUB AddEditAssetCode (RecNo%)

  SHARED Choice$()

  REDIM Choice$(0 TO 4, 0)
  REDIM CodeRec(1) AS FAAssetCodeRecType
  CodeRecLen = LEN(CodeRec(1))

  Choice$(0, 0) = "2"
  Choice$(1, 0) = "Active"
  Choice$(2, 0) = "Inactive"

  LibName$ = "FA"
  ScrnName$ = "FACODE"

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

  Action = 1

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF FirstTime THEN
      FirstTime = False
      IF RecNo > 0 THEN
        Mode$ = "Editing"
        Fld(1).Protected = True
        GOSUB LoadCodeRec
      ELSE
        Mode$ = " Adding"
      END IF
      QPrintRC Mode$, 6, 23, -1
      Action = 1
    END IF

    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE -68    'F10Key
      GOSUB CheckAndSave
      IF ExitFlag THEN
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

CheckAndSave:
  ExitFlag = True    'assume all is well

'put some validation here

  REDIM CodeRec(1) AS FAAssetCodeRecType
  CodeRecLen = LEN(CodeRec(1))

  AssetCode$ = QPTrim$(Form$(1, 0))
  IF LEN(AssetCode$) = 0 THEN
    FldNo = 1
    GOTO BadParm
  END IF
  AssetStatus$ = QPTrim$(Form$(2, 0))
  IF LEN(AssetStatus$) = 0 THEN
    FldNo = 2
    GOTO BadParm
  END IF
  
  CodeRec(1).AssetCode = AssetCode$
  CodeRec(1).AssetStatus = QPTrim$(Form$(2, 0))
  CodeRec(1).AssetDesc = QPTrim$(Form$(3, 0))

  CodeFile = FREEFILE
  OPEN FACodeFile FOR RANDOM AS CodeFile LEN = CodeRecLen
  IF RecNo = 0 THEN
    RecNo = (LOF(CodeFile) \ CodeRecLen) + 1
  END IF
  PUT CodeFile, RecNo, CodeRec(1)
  CLOSE
  GOTO SaveReturn

BadParm:
  Ok = MsgBox("FA", "NOBLANK")
  Action = 1
  ExitFlag = False
  Frm(1).FldNo = FldNo

SaveReturn:
  RETURN

LoadCodeRec:

  REDIM CodeRec(1) AS FAAssetCodeRecType
  CodeRecLen = LEN(CodeRec(1))

  CodeFile = FREEFILE
  OPEN FACodeFile FOR RANDOM AS CodeFile LEN = CodeRecLen
  GET CodeFile, RecNo, CodeRec(1)
  CLOSE

  LSET Form$(1, 0) = QPTrim$(CodeRec(1).AssetCode)
  LSET Form$(2, 0) = QPTrim$(CodeRec(1).AssetStatus)
  LSET Form$(3, 0) = CodeRec(1).AssetDesc

RETURN
END SUB

SUB PrintCodeListing



  REDIM Array(1 TO 1) AS Struct 'Template for the sort array
  REDIM CodeRec(1) AS FAAssetCodeRecType
  CodeRecLen = LEN(CodeRec(1))


  ReportFile$ = "FACODE.PRN"   'Report File Name
  Dash80$ = STRING$(80, "=")
  FF$ = CHR$(12)

  MaxLines = 50
  LineCnt& = 0
  ItemCnt& = 0

  


  RptHandle = FREEFILE

  OPEN ReportFile$ FOR OUTPUT AS #RptHandle

  GOSUB PrintHeader

  FAFile = FREEFILE
  OPEN FACodeFile FOR RANDOM AS FAFile LEN = CodeRecLen
  NumOfFaRecs = LOF(FAFile) / CodeRecLen

  GOSUB GetIndex


  FOR Cnt& = 1 TO NumOfFaRecs

    CodeRecNo = Array(Cnt&).RecNum
    GET FAFile, CodeRecNo, CodeRec(1)

    
      IF LineCnt& >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintHeader
      END IF

      'Check For Disposed Of

         PRINT #RptHandle, CodeRec(1).AssetCode;
          PRINT #RptHandle, TAB(25); CodeRec(1).AssetDesc;
           PRINT #RptHandle, TAB(70); CodeRec(1).AssetStatus
        LineCnt& = LineCnt& + 1
SkipEm:
      NEXT Cnt&

  GOSUB PrintEnding
  CLOSE         'Close all open files now

  EntryPoint = 1
  
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  KILL ReportFile$
  EXIT SUB

PrintHeader:
  Page = Page + 1
  PRINT #RptHandle, TAB(29); "Master Code Listing"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
  PRINT #RptHandle,
  PRINT #RptHandle, "Asset Catagory Code"; TAB(25); "Description"; TAB(70); "Status"
  PRINT #RptHandle, Dash80$
  LineCnt& = 6
  RETURN

PrintEnding:
  PRINT #RptHandle, FF$
  RETURN

GetIndex:
  REDIM Array(1 TO NumOfFaRecs) AS Struct
  FOR Cnt& = 1 TO NumOfFaRecs
    GET FAFile, Cnt&, CodeRec(1)
    Array(Cnt&).who = LTRIM$(CodeRec(1).AssetCode)
    Array(Cnt&).RecNum = Cnt&
  NEXT

 'Sort Them Here
  SortT Array(1), NumOfFaRecs, 0, LEN(Array(1)), 0, 14
  RETURN

END SUB

