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
''$INCLUDE: 'PRFiles.bi'
''$INCLUDE: 'PREmpRec.bi'
'$INCLUDE: 'PageInfo.bi'    '        "          Page information
''$INCLUDE: 'Retire.Bi'

CONST False = 0
CONST True = NOT False

'REDIM SHARED Menu$(0)

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

