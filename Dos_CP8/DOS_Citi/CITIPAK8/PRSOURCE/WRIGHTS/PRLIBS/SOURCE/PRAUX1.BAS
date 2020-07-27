DEFINT A-Z
DECLARE SUB BlockClear ()
DECLARE SUB ExplodBox (UlRow%, UlCol%, BrRow%, BrCol%, Scr(), El%)
'DECLARE FUNCTION NotFromPRMain% ()
DECLARE SUB FCreate (FileName$)
DECLARE SUB FOpenS (FileName$, Handle%)
'DECLARE SUB FOpenAll (FileName$, AccessMode%, ShareMode%, Handle%)
'DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FClose (Handle%)
DECLARE SUB InitMenuText ()
DECLARE SUB ClearScrn ()

DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FPutRT (Handle%, Source AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutRTA (Handle%, SEG Source AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutT (Handle%, Source AS ANY, NumBytes%)
DECLARE SUB FPut (Handle%, ToFile$)

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

DECLARE SUB ReplaceString (Work$, Old$, New$)
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
DECLARE SUB PayMenu (BYVAL MenuNum%, Choice%)

DECLARE SUB VertMenu (Item$(), Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB DisplayMiscScrn (ScrnName$)
DECLARE SUB CursorOff ()
DECLARE SUB RestScrn (ScrnArray%())
DECLARE SUB SaveScrn (ScrnArray%())
DECLARE FUNCTION ASCII% (Strng$)
DECLARE SUB Pause (Ticks)
DECLARE SUB Pause3 (MillaSecs%, ProcessorSpeed&)
DECLARE SUB MPaintBox (UlRow, UlCol, LRRow, LRCol, Colr)
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
DECLARE FUNCTION FudgeFactor& ()

'$INCLUDE: 'DefCnf.bi'                  'Defines TYPE for monitor/color info.
'$INCLUDE: 'FieldInf.bi'                '        "        field information
'$INCLUDE: 'FormEdit.bi'                '        "        form editing info.
'$INCLUDE: 'PRFiles.bi'
'$INCLUDE: 'PageInfo.bi'    '        "          Page information
'$INCLUDE: 'Retire.Bi'

CONST False = 0
CONST True = NOT False

'REDIM SHARED Menu$(0)


SUB DisplayMiscScrn (ScrnName$) STATIC

  LibFile2Scrn MiscQLib, ScrnName$, MonoCode, 0, 0

END SUB

SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
  
  Cols = ButtonRCol - ButtonLCol + 1
  REDIM Btn(Cols)
  MScrnSave ButtonRow, ButtonLCol, ButtonRow, ButtonRCol, Btn(1)

  MPaintBox ButtonRow, ButtonLCol, ButtonRow, ButtonRCol, 15
  
'  Pause 2
  DO
    GetCursor 0, 0, i
  LOOP WHILE i
  ButtonPress 0, 0, 0, 0, 0

  MScrnRest ButtonRow, ButtonLCol, ButtonRow, ButtonRCol, Btn(1)
  
  IF KeyCode < 0 THEN
    StuffBuf CHR$(0) + CHR$(ABS(KeyCode))
  ELSEIF KeyCode > 0 THEN
    StuffBuf CHR$(KeyCode)
  END IF

  ERASE Btn

END SUB

FUNCTION PromptSaveData%

  DIM FrmPSD AS FormInfo

  REDIM TempScrn(0)
  SaveScrn TempScrn()

  SaveFlag = 2

    FormName$ = "SAVE1ST"
    NumFlds = LibNumberOfFields(MiscQLib, FormName$)
    REDIM Form$(NumFlds, 2)             'DIM the form data array
    REDIM Fld(NumFlds) AS FieldInfo     'DIM the field information array
    StartEl = 0                         'Load first form at array start
    LibGetFldDef MiscQLib, FormName$, StartEl, Fld(), Form$(), ErrCode


  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1

  '----- Setup TYPE for setting and reading form editing information.
  FrmPSD.FldNo = 1                 'Start editing on field #1
  FrmPSD.InsStat = False           'Set insert state (True = Insert on)
  FrmPSD.StartEl = 0               'Set form starting element to 0 and

  DisplayMiscScrn FormName$

  DO
    EditForm Form$(), Fld(), FrmPSD, Cnf, Action
    SELECT CASE FrmPSD.KeyCode
      CASE F0Key
        SaveFlag = True
      CASE EscKey
        SaveFlag = 1
      CASE 88, 120              'X Key
        SaveFlag = False
    END SELECT
   
  LOOP WHILE SaveFlag = 2   'proper key not set

  PromptSaveData = SaveFlag
  CursorOff

  RestScrn TempScrn()

  ERASE TempScrn, Form$, Fld

END FUNCTION

SUB ReplaceString (Work$, Old$, New$) STATIC

    OldLen = LEN(Old$)
    NewLen = LEN(New$)
    Pointer = 1

Again:
    'change the QInstr2 to QInstr to make searching case sensitive
    Pointer = INSTR(Pointer, Work$, Old$)

    IF Pointer THEN
       Work$ = LEFT$(Work$, Pointer - 1) + New$ + MID$(Work$, Pointer + OldLen)
       Pointer = Pointer + NewLen
       GOTO Again
    END IF

END SUB

SUB RestScrn (ScrnArray())
  'MScrnRest 1, 1, 25, 80, ScrnArray(0)
  HideCursor
  ExplodBox 1, 1, 25, 80, ScrnArray(), 0
  ShowCursor
END SUB

SUB RPTSetupPRN (RPTNum, Handle)

  SHARED PrnDef$()

  IF RPTNum > 0 THEN
    REDIM PRNSet(1 TO 16) AS INTEGER
    FGetAH PrinterSetUPFile, PRNSet(1), 2, 16
    RPTPitch = PRNSet(RPTNum)
    ERASE PRNSet
    SELECT CASE RPTPitch
      CASE 10
        ToPrint$ = PrnDef$(2)
      CASE 12
        ToPrint$ = PrnDef$(3)
      CASE 17
        ToPrint$ = PrnDef$(4)
      CASE ELSE
    END SELECT

  ELSE 'send reset to printer
    ToPrint$ = PrnDef$(1)
  END IF

  IF LEN(ToPrint$) THEN
    FPut Handle, ToPrint$
  END IF

END SUB

SUB SaveScrn (ScrnArray())
 REDIM ScrnArray(2000)
 MScrnSave 1, 1, 25, 80, ScrnArray(0)
END SUB

SUB WaitForAction STATIC
  DO
    ButtonPress 1, Status, 0, 0, 0
'    call QPrintRC(STR$(FRE("")), 25, 1, 31)
  LOOP UNTIL LEN(INKEY$) OR Status      'then loop while status
  IF Status THEN
    DO
      GetCursor 0, 0, Status
    LOOP WHILE Status
    ButtonPress 0, 0, 0, 0, 0
  END IF
END SUB

