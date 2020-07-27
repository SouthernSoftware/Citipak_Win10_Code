DEFINT A-Z

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
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION Num2Date$ (TheDate%)
DECLARE SUB WaitForAction ()
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB SaveScrn (Array())
DECLARE SUB RestScrn (Array())

  CONST False = 0, True = NOT False
  CONST BlankInt% = -32767      'Specifies blank integer

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'newfa.BI'
  '$INCLUDE: 'Qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'GL.BI'

  DIM FASetup AS FASetupRecType
  SetupLen = LEN(FASetup)

  LibName$ = "FA"
  ScrnName$ = "FASETUP"

  '--define the multi-choice fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1

  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode

  Frm(1).PageDisable = True


  '--Set screen number to one and display screen
  'ClearBak
  BlockClear
  DisplayFAScrn ScrnName$

  ShowCursor
  FirstTime = True

  Action = 1

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      GOSUB LoadSetupFile
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

  LOOP UNTIL Frm(1).KeyCode = 27 OR ExitFlag
  HideCursor
  RUN "famenu"
  END

LoadSetupFile:
  IF Exist("FASETUP.DAT") THEN
    SetupFile = FREEFILE
    OPEN "fasetup.dat" FOR RANDOM AS SetupFile LEN = SetupLen
    GET SetupFile, 1, FASetup
    CLOSE SetupFile
    LSET Form$(1, 0) = FASetup.TownName
    LSET Form$(2, 0) = QPTrim$(STR$(FASetup.Pct1St))
  END IF
RETURN

CheckAndSave:
  LSET FASetup.TownName = QPTrim$(Form$(1, 0))
  FASetup.Pct1St = Value(Form$(2, 0), ECode)
  SetupFile = FREEFILE
  OPEN "fasetup.dat" FOR RANDOM AS SetupFile LEN = SetupLen
  PUT SetupFile, 1, FASetup
  CLOSE SetupFile
  ExitFlag = True
RETURN

