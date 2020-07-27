DECLARE SUB AddHHCode ()
DECLARE SUB EditHHCode ()
DECLARE SUB PrintHHCode ()
DECLARE SUB SendHHCode2Husky ()
DEFINT A-Z
DECLARE SUB HHCodeMaint ()
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE FUNCTION nGetPassword$ (Row%, Col%, Max%, Abort%)
DECLARE SUB UBLog (Text$)
DECLARE SUB BlockClear ()
DECLARE SUB CursorOff ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE FUNCTION WEnvTest% ()
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB WaitForAction ()
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB KillFile (File2Kill$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Round# (DblNum#)
DECLARE SUB SaveScrn (ScrnArray())
DECLARE SUB RestScrn (ScrnArray())
DECLARE SUB QPrintRC (Text$, Row%, Col%, Kolor)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION Num2Date$ (DateNum)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION PromptSaveData% ()
DECLARE FUNCTION ASCII (Character$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValI% (Number$)
DECLARE FUNCTION QPValL& (Number$)
DECLARE FUNCTION QPStrL$ (Number&)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB HideCursor ()
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB ClearScrn ()
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetupLen%)

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'Qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  'INCLUDE: 'ubdraft.BI'
  '$INCLUDE: 'ubsetup.bi'
  'INCLUDE: 'ubtrans.bi'
  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubrate.BI'

  CONST False = 0, True = NOT False

SUB AddHHCode

  LibName$ = "UBHH"
  ScrnName$ = "HHCODEED"
  NumScrns = 1

  '--define the multi-choice fields
  NumFlds = -1
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1

  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode

  '--Clear all fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT
  
  BlockClear
  LibFile2Scrn LibName$, ScrnName$, MonoCode, Attribute%, ErrCode

  ShowCursor

  Action = 1
  FirstTime = True

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      QPrintRC "[                       ]", 7, 17, 10
      QPrintRC "Add New HandHeld Code", 7, 19, 14
    END IF

  LOOP UNTIL Frm(1).KeyCode = 27 OR ExitFlag

END SUB

SUB EditHHCode
END SUB

SUB HHCodeMaint

'  UBLog " IN: Edit HH Codes."

  REDIM HHCodeRec(1) AS HHCodeRecType
  HHCodeLen = LEN(HHCodeRec(1))
  IF Exist("UBHHCODE.DAT") THEN

  END IF

  REDIM MChoice$(1 TO 5)

  MChoice$(1) = "Add a New HandHeld Code"
  MChoice$(2) = "Edit an Existing HandHeld Code"
  MChoice$(3) = "Print HandHeld Codes Listing"
  MChoice$(4) = "Send HandHeld Codes to Husky"
  MChoice$(5) = "Exit to OS"
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 18   'limit the box length to go no lower than line 20
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
    BlockClear
    TitleBox 2, Col, MaxLen + 3, "HandHeld Code Maintenance", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

    ShowCursor

    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf

    IF Ky$ = CHR$(27) THEN
      Choice = 0
      ExitFlag = True
    END IF

    SELECT CASE Choice
    CASE 1
      AddHHCode
    CASE 2
      EditHHCode
    CASE 3
      PrintHHCode
    CASE 4
      SendHHCode2Husky
    CASE 5
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP UNTIL ExitFlag

END SUB

SUB PrintHHCode
END SUB

SUB SendHHCode2Husky

END SUB

