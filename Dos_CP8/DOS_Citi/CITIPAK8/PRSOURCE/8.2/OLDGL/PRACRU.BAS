DEFINT A-Z
DECLARE SUB RPTSetupPRN (RPTNum%, Handle%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB ProcessAccrual (AccDate%)
DECLARE SUB PostTransactions ()
DECLARE FUNCTION PPDelFromPay% ()
DECLARE SUB PrintChecks (First%, last%)
DECLARE SUB ButtonPress (ButNo%, Down%, Presses%, X%, Y%)
DECLARE SUB PPrintRegister ()
DECLARE SUB PrintCheckMenu ()
DECLARE SUB ViewFile (FileName$, WTopRow, WLeftCol, NumChars, Rows, Cnf AS ANY, Action)
DECLARE SUB PPLoadPFreq ()
DECLARE FUNCTION Round# (DblNum#)
DECLARE SUB PPGetEmp2Rec (EmpRecNo%, ErrorCode%)
DECLARE SUB PPLoadSystemFiles ()
DECLARE SUB PPEdScrnCalc (TransRec AS ANY, TransRecNo%, EmpName$, EmpRecNo%)
DECLARE FUNCTION Value# (E$, ErCode%)
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB WaitForAction ()

DECLARE SUB FCreate (FileName$)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FClose (Handle%)

DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FPutRT (Handle%, Source AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutRTA (Handle%, SEG Source AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutT (Handle%, Source AS ANY, NumBytes%)
DECLARE SUB FPut (Handle%, ToDisk$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetRT (Handle%, Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FGetT (Handle%, Dest AS ANY, NumBytes%)

DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB FSeek (Handle%, Location&)

DECLARE SUB FFlush (Handle%)

DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FEof% (Handle%)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION FLoc& (Handle%)
DECLARE FUNCTION FLof& (Handle%)
DECLARE SUB InsertT (SEG StartElement AS ANY, ElSize%, NumEls%)

DECLARE SUB MScrnSave (UlRow, UlCol, LRRow, LRCol, SEG Address)
DECLARE SUB MScrnRest (UlRow, UlCol, LRRow, LRCol, SEG Address)
DECLARE SUB StuffBuf (Ky$)
DECLARE SUB GetCursor (X, Y, Button)
DECLARE SUB MQPrint (X$, Colr)
DECLARE SUB MPaintBox (UlRow, UlCol, LRRow, LRCol, Colr)
DECLARE SUB ButtonPress (ButNo%, Down%, Presses%, X%, Y%)
DECLARE SUB HideCursor ()
DECLARE SUB TextCursor (Fg, Bg)
DECLARE SUB ShowCursor ()
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)

DECLARE FUNCTION LibNumberOfFields (LibName$, FrmName$)
DECLARE FUNCTION LibScrName$ (Scr, ScrnLib())
DECLARE FUNCTION ScrnLibSize (LibName$)
DECLARE SUB EditForm (Form$(), Fld() AS ANY, Frm AS ANY, Cnf AS ANY, Action)
DECLARE SUB Lib2Scrn (NameInLib$, ScrnLib(), MonoCode, Attribute, ErrorCode)
DECLARE SUB LibGetFldDef (LibName$, FrmName$, StartEl, Fld() AS ANY, Form$(), ErrCode)
DECLARE SUB LoadScrnLib (LibName$, ScrnLib(), ErrorCode)
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB PrintArray (FirstFld, LastFld, Form$(), Fld() AS ANY)
DECLARE SUB UnPackBuffer (FirstFld, LastFld, Form$(), Fld() AS ANY)
DECLARE FUNCTION FldNum (FldName$, Fld() AS ANY)
DECLARE FUNCTION EndOfForms (Fld() AS ANY)
DECLARE FUNCTION PDQTimer& ()
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE SUB SearchT2 (SEG Array AS ANY, NumEls%, Match, Found%, Dir%, Code%, StructSize%, MemberOff%, MemberSize%)
DECLARE FUNCTION QPRTrim$ (Text$)
DECLARE FUNCTION PromptSaveData% ()
'DECLARE SUB PayMenu (MenuNum%, Choice%)
DECLARE SUB PayMenu (BYVAL MenuNum, Choice, NumOfItems%)
DECLARE SUB VertMenu (Item$(), Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB DisplayMiscScrn (ScrnName$)
DECLARE SUB CursorOff ()
DECLARE SUB RestScrn (ScrnArray%())
DECLARE SUB SaveScrn (ScrnArray%())
DECLARE FUNCTION ASCII% (Strng$)
DECLARE SUB Pause (Ticks)
DECLARE FUNCTION Compare3% (BYVAL Seg1, BYVAL Adr1, BYVAL Seg2, BYVAL Adr2, NumBytes)
DECLARE SUB ClearScrn ()
DECLARE SUB BlockClear ()
DECLARE SUB SaveField (FldNo%, Form$(), Fld() AS ANY, BadFld%)
DECLARE SUB FEInitForm (Form$(), Fld() AS ANY, Frm AS ANY)
DECLARE SUB CalcFields (StartOfForm, FldNo, Form$(), Fld() AS ANY)
DECLARE FUNCTION FldNum (FldName$, Fld() AS ANY)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION Num2Date$ (Number%)
DECLARE FUNCTION Date2Num (DateIn$)
DECLARE SUB PrintRptFile1 (FileName$, LPTPort%, RetCode%)
DECLARE SUB PrintArray (FirstFld%, LastFld%, Form$(), Fld() AS ANY)

'$INCLUDE: 'DefCnf.bi'                  'Defines TYPE for monitor/color info.

'$INCLUDE: 'FieldInf.bi'                '        "        field information
'$INCLUDE: 'FormEdit.bi'                '        "        form editing info.
'$INCLUDE: 'PRFiles.bi'
'$INCLUDE: 'PREmpRec.bi'
'$INCLUDE: 'PageInfo.bi'    '        "          Page information
'$INCLUDE: 'Retire.Bi'
'$INCLUDE: 'Leave.Bi'
'$INCLUDE: 'PREIC.Bi'
'$INCLUDE: 'PRUNIT.Bi'
'$INCLUDE: 'PRSYSCTR.Bi'
'$INCLUDE: 'PRPPDEF.Bi'
'$INCLUDE: 'PRCHECK.Bi'

CONST False = 0
CONST True = NOT False

SUB AccruLeave (SaveScrnFlag)

  DIM FrmAL AS FormInfo

  REDIM TempScrn(0)

  IF SaveScrnFlag THEN
    SaveScrn TempScrn()
  END IF

  'LibName$ = AccruQLib     'Specify a Screen Library file name.
  FormName$ = "ACRULEAV"
  LibSize = ScrnLibSize(AccruQLib)     'Get scrn lib array size
  REDIM ScrnLib(LibSize)              'Create scrn lib array
  LoadScrnLib AccruQLib, ScrnLib(), ErrorCode   'load scrn lib
  NumFlds = LibNumberOfFields(AccruQLib, FormName$)
  REDIM FormAL$(NumFlds, 2)             'DIM the form data array
  REDIM FldAL(NumFlds) AS FieldInfo     'DIM the field information array
  StartEl = 0                         'Load first form at array start
  LibGetFldDef AccruQLib, FormName$, StartEl, FldAL(), FormAL$(), ErrCode


  '----- Setup TYPE for setting and reading form editing information.
  FrmAL.FldNo = 1                 'Start editing on field #1
  FrmAL.InsStat = False            'Set the insert state (-1 = Insert on)
  FrmAL.StartEl = 0               'Set form starting element to 0 and

  FormAL$(0, 0) = SPACE$(FldAL(0).Row)

  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.

  Action = 1
  BlockClear

  IF FileSize(AccrueFileName) > 0 THEN
    LastAccrue = 0
    FGetAH AccrueFileName, LastAccrue, 2, 1
    FormAL$(1, 0) = Num2Date$(LastAccrue)
    SaveField 1, FormAL$(), FldAL(), BadFld
  ELSE
    NoneFlag = True
  END IF

  DO                 'Pole the editing procedure
    Lib2Scrn FormName$, ScrnLib(), MonoCode, -1, ErrCode

    DO
      EditForm FormAL$(), FldAL(), FrmAL, Cnf, Action

      IF NoneFlag THEN
        NoneFlag = NOT NoneFlag
        HideCursor
        QPrintRC "  None    ", 8, 44, 112
        ShowCursor
      END IF
'IF FrmAL.KeyCode <> 0 THEN STOP
     SELECT CASE FrmAL.KeyCode
       CASE F0Key, EscKey
         SELECT CASE FrmAL.KeyCode
           CASE EscKey
             SaveFlag = False
           CASE F0Key
             SaveFlag = True
         END SELECT
         SELECT CASE SaveFlag
           CASE True
             'LastAccrue = Date2Num(FormAL$(2, 0))
             ThisAccrue = Date2Num(FormAL$(2, 0))
             IF ThisAccrue < LastAccrue THEN

               'BEEP: BEEP: STOP
               'error screen
             ELSE
               FPutAH AccrueFileName, ThisAccrue, 2, 1
               BlockClear
               Lib2Scrn "ACCRUING", ScrnLib(), MonoCode, -1, ErrCode
               ProcessAccrual ThisAccrue
               ExitFlag = True
             END IF
           CASE False
             ExitFlag = True
         END SELECT
     CASE ELSE
     END SELECT
    LOOP UNTIL ExitFlag = True

  LOOP UNTIL ExitFlag = True '  the Escape key.
'  ExitFlag = False

  CursorOff

  IF SaveScrnFlag THEN
    RestScrn TempScrn()
    ERASE TempScrn
  END IF

END SUB

SUB ProcessAccrual (AccrualDate)

  SHARED PrnDef$()
  CursorOff
        
  REDIM Unit(1)          AS UnitFileRecType
  REDIM EmpRec2(1)       AS EmpData2Type

  REDIM ToPrint(1) AS STRING * 79
  'REDIM Hrs(1)  AS STRING * 6

  REDIM Tot(1)  AS STRING * 8
  REDIM LeaveRec(1)      AS LeaveRecType

  LRecLen = LEN(LeaveRec(1))
  NumLeaveRec = FileSize(LeaveFileName) \ LRecLen

  REDIM LeaveRec(1 TO NumLeaveRec)      AS LeaveRecType

  FGetAH LeaveFileName, LeaveRec(1), LRecLen, NumLeaveRec
  FGetAH UnitFileName, Unit(1), LEN(Unit(1)), 1

  Image$ = "###.##"
  TImage$ = "#####.##"
  Image1$ = "##"
  Image2$ = "###"

  TblPos = 41
  YrsPos = 46
  BenPos = 53
  VacPos = 63
  SickPos = 72

'Number     Name                            Benefit Pct   Vaction     Sick"
'                                              ##.##       ###.##   ###.##
  LineCnt = 0
  MaxLines = 50
  
  EmpRecSize = LEN(EmpRec2(1))

  NumOfRecs = FileSize(EmpData2Name) \ EmpRecSize

  IdxRecLen = 2
  IdxFileSize& = FileSize(EmpIdxNName)
  INumOfRecs = IdxFileSize& \ IdxRecLen

  REDIM IdxBuf(1 TO INumOfRecs)

  FGetAH EmpIdxNName, IdxBuf(1), IdxRecLen, INumOfRecs

  FOpenS EmpData2Name, Handle       'open employee data file

  FCreate AccrualRptFile
  FOpenS AccrualRptFile, RHandle
  RPTSetupPRN 5, RHandle

  'FClose RHandle

  'OPEN AccrualRptFile FOR APPEND AS #1

  GOSUB PrintLeaveHeader

  FOR RecNo = 1 TO NumOfRecs
    FGetRTA Handle, EmpRec2(1), CLNG(IdxBuf(RecNo)), EmpRecSize
    'IF INSTR(EmpRec2(1).EmpNo, "2") > 0 THEN STOP
    'IF DosError THEN BEEP
    IF EmpRec2(1).EMPTDATE < 0 AND EmpRec2(1).EMPBCODE > 0 AND NOT EmpRec2(1).Deleted THEN

    'if employee not terminated AND they get benefits.
      EmpTotal = EmpTotal + 1
      HireDate = EmpRec2(1).EMPHDATE
      'IF HireDate = -32768 THEN
      IF HireDate <= -32767 THEN
        'STOP
        GOSUB BadHireDate
        GOTO BadDateSkip
        'this returns to BadDateSkip
      END IF

      WhatLeaveTbl = EmpRec2(1).LeaveTbl

      IF WhatLeaveTbl < 1 THEN
        WhatLeaveTbl = 1
      END IF

      AccrualDays = AccrualDate - HireDate
      IF AccrualDays > 365 THEN
        YearsOfService = INT(AccrualDays / 365)
        'STOP
      ELSE
        YearsOfService = 0
      END IF

      FOR Cnt = 1 TO 20
        IF YearsOfService <= LeaveRec(WhatLeaveTbl).VEntry(Cnt).YEARS THEN
          EXIT FOR
        END IF
      NEXT
      IF Cnt > 20 THEN Cnt = 20
      IF YearsOfService = LeaveRec(WhatLeaveTbl).VEntry(Cnt).YEARS THEN
        VTableEntry = Cnt
      ELSE
        VTableEntry = Cnt - 1
      END IF

      IF VTableEntry = 0 THEN VTableEntry = 1
      VAmt# = Round#(LeaveRec(WhatLeaveTbl).VEntry(VTableEntry).EARN * (EmpRec2(1).EMPBCODE * .01))

      IF VAmt# > 0 THEN           ' if there is amount to add
        IF EmpRec2(1).EMPVBAL + VAmt# > LeaveRec(WhatLeaveTbl).VACMAX THEN     'if > max amt
          VAmt# = LeaveRec(WhatLeaveTbl).VACMAX - EmpRec2(1).EMPVBAL   'set amt to max
          VADJFlag = True
        END IF                                             '
        EmpRec2(1).EMPVBAL = Round#(EmpRec2(1).EMPVBAL + VAmt#)
        EmpRec2(1).EMPVACE = Round(EmpRec2(1).EMPVBAL + EmpRec2(1).EMPVUSED)         'set emp to amt
      END IF

      FOR Cnt = 1 TO 20
        IF YearsOfService <= LeaveRec(WhatLeaveTbl).SEntry(Cnt).YEARS THEN
          EXIT FOR
        END IF
      NEXT
      IF Cnt > 20 THEN Cnt = 20
      IF YearsOfService = LeaveRec(WhatLeaveTbl).SEntry(Cnt).YEARS THEN
        STableEntry = Cnt
      ELSE
        STableEntry = Cnt - 1
      END IF

      IF STableEntry = 0 THEN STableEntry = 1
      SAmt# = Round#(LeaveRec(WhatLeaveTbl).SEntry(STableEntry).EARN * (EmpRec2(1).EMPBCODE * .01))

      IF SAmt# > 0 THEN           ' if there is amount to add
        IF EmpRec2(1).EMPSLBAL + SAmt# > LeaveRec(WhatLeaveTbl).SICKMAX THEN   'if > max amt
          SADJFlag = True
          SAmt# = LeaveRec(WhatLeaveTbl).SICKMAX - EmpRec2(1).EMPSLBAL
        END IF
        EmpRec2(1).EMPSLBAL = Round#(EmpRec2(1).EMPSLBAL + SAmt#)
        EmpRec2(1).EMPSLE = Round(EmpRec2(1).EMPSLBAL + EmpRec2(1).EMPSLUSE)         'set emp to amt
      END IF

      GOSUB UpDateReport
      FPutRTA Handle, EmpRec2(1), CLNG(IdxBuf(RecNo)), EmpRecSize 'put to disk
    END IF

BadDateSkip:
    GOSUB PctComplete
  NEXT

  GOSUB PrintLeaveSummary

  RPTSetupPRN 0, RHandle
  FClose Handle
  FClose RHandle

  RptTitle$ = "Employee Leave Accrual Report"

  PrintRptFile RptTitle$, AccrualRptFile, 1, RetCode, 1

  EXIT SUB

BadHireDate:

  EmpName$ = SPACE$(28)

  LSET EmpName$ = QPRTrim$(EmpRec2(1).EMPLNAME) + ", " + QPRTrim$(EmpRec2(1).EMPFNAME)

  LSET ToPrint(1) = LTRIM$(EmpRec2(1).EmpNo)    'set number

  MID$(ToPrint(1), 13, 28) = EmpName$  'set name
  MID$(ToPrint(1), YrsPos) = "Invalid hire date."

  FPut RHandle, ToPrint(1) + CrLf$
  LineCnt = LineCnt + 1
  GOSUB Check4NewPage
  RETURN

PrintLeaveHeader:
  'RSET Pg(1) = STR$(Page)
  FPut RHandle, Unit(1).UFEMPR + CrLf$   '; SPACE$(87); "Page:"; Pg(1)
  FPut RHandle, "Leave Benefits Earned" + CrLf$
  FPut RHandle, "Accrual Date: " + Num2Date$(AccrualDate) + CrLf$
  FPut RHandle, CrLf$
  FPut RHandle, "Number      Name                       Tbl   Yrs   Benefit%  Vaction     Sick" + CrLf$
  FPut RHandle, "-----------------------------------------------------------------------------" + CrLf$

  LineCnt = LineCnt + 6
  'Page = Page + 1
  GOSUB Check4NewPage
RETURN

UpDateReport:

  TotalSick# = Round#(TotalSick# + SAmt#)
  TotalVac# = Round#(TotalVac# + VAmt#)

  EmpName$ = SPACE$(28)

  LSET EmpName$ = QPRTrim$(EmpRec2(1).EMPLNAME) + ", " + QPRTrim$(EmpRec2(1).EMPFNAME)

  'LSET ToPrint(1) = EmpRec2(1).EmpNo    'set number
  LSET ToPrint(1) = LTRIM$(EmpRec2(1).EmpNo)    'set number

  MID$(ToPrint(1), 13, 28) = EmpName$  'set name

  MID$(ToPrint(1), TblPos, 3) = FUsing(STR$(WhatLeaveTbl), Image1$)'set benefit
  MID$(ToPrint(1), YrsPos, 6) = FUsing(STR$(AccrualDays \ 365), Image1$)'set benefit

  MID$(ToPrint(1), BenPos, 6) = FUsing(STR$(EmpRec2(1).EMPBCODE), Image$) 'set benefit
  MID$(ToPrint(1), VacPos, 6) = FUsing(STR$(VAmt#), Image$)               'set vac
  IF VADJFlag THEN
    VADJFlag = False
    MID$(ToPrint(1), VacPos + 6) = "*"
  END IF

  MID$(ToPrint(1), SickPos, 6) = FUsing(STR$(SAmt#), Image$)             'set sick
  IF SADJFlag THEN
    SADJFlag = False
    MID$(ToPrint(1), SickPos + 6) = "*"
  END IF

  FPut RHandle, ToPrint(1) + CrLf$
  LineCnt = LineCnt + 1
  GOSUB Check4NewPage

RETURN


PrintLeaveSummary:
  LSET ToPrint(1) = "Totals          Employees"

  MID$(ToPrint(1), 13, 3) = FUsing(STR$(EmpTotal), Image2$)
  MID$(ToPrint(1), VacPos - 2, 8) = FUsing(STR$(TotalVac#), TImage$)             'set vac
  MID$(ToPrint(1), SickPos - 2, 8) = FUsing(STR$(TotalSick#), TImage$)           'set sick

  FPut RHandle, "-----------------------------------------------------------------------------" + CrLf$
  FPut RHandle, ToPrint(1) + CrLf$
  FPut RHandle, CrLf$
  LSET ToPrint(1) = "NOTE: " + CHR$(34) + "*" + CHR$(34) + " Indicates maxium balance reached."
  FPut RHandle, ToPrint(1) + CrLf$
  FPut RHandle, CHR$(12)
RETURN



Check4NewPage:
  IF LineCnt > MaxLines THEN
    LineCnt = 0
    FPut RHandle, CHR$(12)
    GOSUB PrintLeaveHeader
  END IF
RETURN


PctComplete:
  HideCursor
  QPrintRC STR$(INT((RecNo / NumOfRecs) * 100)), 12, 34, 112
  ShowCursor
  RETURN
END SUB

