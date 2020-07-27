DEFINT A-Z
DECLARE SUB SmallPause ()
DECLARE SUB SelectEmps ()
DECLARE SUB PrintEmps ()
DECLARE SUB BlockClear ()
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE SUB CursorOff ()
DECLARE SUB HideCursor ()
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
DECLARE SUB ShowCursor ()
DECLARE SUB VertMenu (Item$(), Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE FUNCTION ASCII% (Strng$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB ShowProcessingScrn (Text$)
DECLARE SUB ShowPctComp (BYVAL Count, BYVAL TotCount)
DECLARE SUB DisplayMiscScrn (ScrnName$)
DECLARE SUB WaitForAction ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)

'$INCLUDE: 'PRFiles.bi'
'$INCLUDE: 'PREmpRec.bi'
'$INCLUDE: 'DefCnf.bi'                  'Defines TYPE for monitor/color info.
'$INCLUDE: 'FieldInf.bi'                '        "        field information
'$INCLUDE: 'FormEdit.bi'                '        "        form editing info.
'$INCLUDE: 'qscr.bi'                  'Defines TYPE for monitor/color info.
'$INCLUDE: 'setCnf.bi'                  'Defines TYPE for monitor/color info.
'$INCLUDE: 'PRUNIT.Bi'

 TYPE EmpSortType
   EmpNo AS STRING * 14
   RecNo AS INTEGER
 END TYPE

 CONST False = 0, True = NOT False
 
 DIM Menu$(1 TO 3)

 Menu$(1) = "Select employee's for testing"
 Menu$(2) = "Print employee's to test listing"
 Menu$(3) = "Exit to OS"
 MenuTitle$ = "Random Drug Testing Menu"

 FOR Cnt = 1 TO 3
   ILen = LEN(Menu$(Cnt))
   IF MaxLen < ILen THEN
     MaxLen = ILen
   END IF
 NEXT

Top:

  Ky$ = ""
  ExitFlag = False
  VAction = 1
  BoxBot = 17
  BoxTop = 9

  LeftCol = ((80 - MaxLen) \ 2) - 1
  TitleCol = ((80 - LEN(MenuTitle$)) \ 2) + 1

  CursorOff
  HideCursor
  LOCATE BoxTop, LeftCol
  BlockClear
  WazzWind 3, LeftCol, 5, LeftCol + MaxLen + 3, Cnf.MenBox, 1, True
  WazzWind 21, LeftCol, 23, LeftCol + MaxLen + 3, Cnf.MenBox, 1, True
  QPrintRC MenuTitle$, 4, TitleCol, Cnf.ActivCh
  QPrintRC " Use (" + CHR$(24) + "-" + CHR$(25) + ") to select", 22, LeftCol + (MaxLen \ 2) - 8, Cnf.ActivCh
  ShowCursor

  DO
    
    VertMenu Menu$(), Choice, MaxLen, BoxBot, Ky$, VAction, Cnf

    SELECT CASE ASCII(Ky$)
      CASE EscKey
        Choice = EscKey
        ExitFlag = True
      CASE 13
        ExitFlag = True
      CASE ELSE
    END SELECT
  LOOP UNTIL ExitFlag
  SELECT CASE Choice
  CASE 1
    SelectEmps
  CASE 2
    PrintEmps
  CASE 3
    HideCursor
    CLS
    END
  CASE EscKey
    RUN "cpmenu"
  END SELECT
  GOTO Top

END

SUB PrintEmps

  REDIM Unit(1)      AS UnitFileRecType
  FGetAH UnitFileName, Unit(1), LEN(Unit(1)), 1
  WhatUnit$ = QPTrim$(Unit(1).UFEMPR)
  UnitLen = LEN(WhatUnit$)
  OffSet = 40 - (UnitLen / 2)
  RptName$ = "PRRPTS\PRDTLIST.RPT"
  DTFileName$ = "PRDATA\PREMPDTL.DAT"

  DIM EmpRec AS EmpData2Type
  EmpLen = LEN(EmpRec)
  DIM DTEmp AS INTEGER

  ShowProcessingScrn "Processing Report"

  LSTFile = FREEFILE
  OPEN DTFileName$ FOR RANDOM AS LSTFile LEN = 2
  DTEmpCnt = LOF(LSTFile) / 2

  EmpFile = FREEFILE
  OPEN EmpData2Name FOR RANDOM AS EmpFile LEN = EmpLen

  REDIM ERecs(1 TO DTEmpCnt) AS EmpSortType

  FOR Cnt = 1 TO DTEmpCnt
    GET LSTFile, Cnt, DTEmp
    GET EmpFile, DTEmp, EmpRec
    RSET ERecs(Cnt).EmpNo = EmpRec.EmpNo
    ERecs(Cnt).RecNo = DTEmp
  NEXT
  CLOSE

  SortT ERecs(1), DTEmpCnt, 0, 16, 0, 14

  RPTFile = FREEFILE
  OPEN RptName$ FOR OUTPUT AS #RPTFile
  GOSUB DTLHeader

  EmpFile = FREEFILE
  OPEN EmpData2Name FOR RANDOM AS EmpFile LEN = EmpLen
  'NumEmp = LOF(EmpFile) / EmpLen
  FOR Cnt = 1 TO DTEmpCnt
    GET EmpFile, ERecs(Cnt).RecNo, EmpRec
    PRINT #RPTFile, EmpRec.EmpNo; TAB(20); QPTrim$(EmpRec.EmpFName); " "; QPTrim$(EmpRec.EmpLName)
    ShowPctComp Cnt, DTEmpCnt
    SmallPause
  NEXT
  GOSUB DTLFooter
  CLOSE

  PrintRptFile "Random Selection Report.", RptName$, 1, RetCode, 0

  EXIT SUB

DTLHeader:

  Page = Page + 1
  UTemp$ = SPACE$(79)
  LSET UTemp$ = "Random Drug Testing Selection Report"
  MID$(UTemp$, 66) = "Page:" + STR$(Page)
  PRINT #RPTFile, TAB(OffSet); WhatUnit$
  PRINT #RPTFile, QPTrim$(UTemp$)
  PRINT #RPTFile, "Report Date: " + DATE$
  PRINT #RPTFile, "    Emp #          Name"
  PRINT #RPTFile, STRING$(79, "-")
  LineCnt = 6
RETURN

DTLFooter:
  PRINT #RPTFile, STRING$(79, "-")
  PRINT #RPTFile, "   Employee's Selected:"; DTEmpCnt
RETURN

END SUB

SUB SelectEmps

  RANDOMIZE TIMER

  DTFileName$ = "PRDATA\PREMPDTL.DAT"

  DIM EmpRec AS EmpData2Type
  EmpLen = LEN(EmpRec)

  ScrnName$ = "PRRNDPCK"
  NumScrns = 1

  '--Initialize the form name array
  NumFlds = LibNumberOfFields(MiscQLib, ScrnName$) + 1

  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef MiscQLib, ScrnName$, StartEl, Fld(), Form$(), ErrCode

  '--Clear all fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT
  Action = 1
  'Fld(2).Protected = True
  'Fld(3).Protected = True

  '--Set screen number to one and display screen
  BlockClear

  LibFile2Scrn MiscQLib, ScrnName$, MonoCode%, Attribute%, ErrorCode%

  ShowCursor


  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE -68    'F10Key
      GOSUB CheckCnt
      IF EmpCntOK THEN
        EXIT DO
      END IF
    CASE EscKey
      ExitFlag = True
    END SELECT

    '--Check screen page
  LOOP UNTIL ExitFlag
  IF ExitFlag THEN
    GOTO ExitSelect
  END IF

  ShowProcessingScrn "Processing Employee's"

  REDIM EmpRecs(1 TO DTEmpCnt) AS INTEGER

  CALL KillFile(DTFileName$)

  LSTFile = FREEFILE
  OPEN DTFileName$ FOR RANDOM AS LSTFile LEN = 2

  EmpFile = FREEFILE
  OPEN EmpData2Name FOR RANDOM AS EmpFile LEN = EmpLen
  NumEmp = LOF(EmpFile) / EmpLen

  DO
    EmpNum = INT(RND * NumEmp) + 1
    GET EmpFile, EmpNum, EmpRec
    EmpNo$ = QPTrim$(EmpRec.EmpNo)
    ENOLen = LEN(EmpNo$)
    IF (EmpRec.EMPTDATE > 0) OR (EmpRec.Deleted <> 0) OR (ENOLen = 0) THEN
      GOTO NoGood
    ELSE
      ECnt = ECnt + 1
      IF ECnt > 1 THEN
        FOR TCnt = 1 TO ECnt - 1
          IF EmpRecs(TCnt) = EmpNum THEN
            ECnt = ECnt - 1
            GOTO NoGood
          END IF
        NEXT
      END IF
      EmpRecs(ECnt) = EmpNum
    END IF
NoGood:
  LOOP UNTIL ECnt >= DTEmpCnt

  SortT SEG EmpRecs(1), DTEmpCnt, 0, 2, 0, -1

  FOR Cnt = 1 TO DTEmpCnt
    PUT LSTFile, Cnt, EmpRecs(Cnt)
    SmallPause
    ShowPctComp Cnt, DTEmpCnt
  NEXT
  CLOSE
  BlockClear
  DisplayMiscScrn "UPDATEOK"

  WaitForAction

ExitSelect:
EXIT SUB

CheckCnt:
  DTEmpCnt = VAL(Form$(1, 0))
'  DTEmpCnt = 5
  IF DTEmpCnt > 0 AND DTEmpCnt < 99 THEN
    EmpCntOK = True
  ELSE
    EmpCntOK = False
    Frm(1).FldNo = 1
    Action = 1
  END IF

RETURN


END SUB

SUB SmallPause
st# = TIMER
DO
LOOP UNTIL TIMER > st# + .01
END SUB

