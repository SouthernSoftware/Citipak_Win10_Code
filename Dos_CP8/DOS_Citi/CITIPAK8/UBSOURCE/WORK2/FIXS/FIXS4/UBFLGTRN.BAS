DEFINT A-Z
DECLARE SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE FUNCTION AreYouSure% (FlagDate%)
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB FlagTractions ()
DECLARE FUNCTION nGetPassword$ (Row%, Col%, Max%, Abort%)
DECLARE SUB MakeBKGround ()
DECLARE SUB ShowMsg ()
DECLARE SUB WaitForKey ()
DECLARE FUNCTION GetPassword$ (Row%, Col%, Max%, Abort%)
DECLARE SUB QPrintRC (t$, r%, c%, clr%)
DECLARE SUB QPSound (Frequency%, Duration%)
DECLARE FUNCTION WaitKey
DECLARE FUNCTION BiosKey
DECLARE SUB WazzWind2 (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE SUB Pause3 (MillaSecs%, ProcessorSpeed&)
DECLARE SUB MPaintBox (UlRow, UlCol, LRRow, LRCol, Colr)
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE FUNCTION FudgeFactor& ()
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION Num2Date$ (TheDate%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE SUB WaitForAction
DECLARE SUB SaveScrn (ScrnArray())
DECLARE SUB RestScrn (ScrnArray())

'$INCLUDE: 'DefCnf.BI'

DECLARE FUNCTION Monitor% ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)
DECLARE SUB HideCursor ()
DECLARE SUB ShowCursor ()
DECLARE SUB BlockClear ()
DECLARE SUB UBLog (Text2Log$)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE FUNCTION QPStrL$ (Number&)

  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'Qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubtrans.BI'

  CONST False = 0, True = NOT False

  CrLf$ = CHR$(13) + CHR$(10)

'ON ERROR GOTO GetOut
LOCATE 1, 1, 0
COLOR 14

GetPW:
  DO
    MakeBKGround
    WazzWind 10, 15, 14, 65, 14, 1, 1
    QPrintRC "[ Flag Transactions ]", 10, 18, -1
    QPrintRC "Enter password:", 12, 22, 15
    
    PW$ = nGetPassword(12, 38, 10, Abort)
    
    IF Abort THEN GOTO GetOut

    IF PW$ = "SOSOFT" THEN
      OkFlag = -1
      EXIT DO
    END IF
    
    Try = Try + 1
    
    IF Try = 4 THEN
      ExitFlag = -1
      MakeBKGround
      WazzWind 8, 15, 16, 65, 15, 1, 1
      QPrintRC "Get the password from the system operator.", 11, 20, 14
      QPrintRC "Press any key to return to the Main Menu.", 13, 20, 14
      WaitForAction
      EXIT DO
    END IF

    QPrintRC "Invalid Password. Press any key to try again.", 12, 18, 31

    QPSound 1000, 2
    QPSound 800, 2
    
    WaitForAction
    
    IF K = 27 THEN
      GOTO GetOut
    END IF
  LOOP
  
  IF ExitFlag THEN
    RUN "UBMISC"
  ELSEIF OkFlag THEN
    FlagTractions
  ELSE
    GOTO GetPW
  END IF

GetOut:

IF INSTR(COMMAND$, "TEST") THEN
  CALL ClearScrn
  END
ELSE
  RUN "UBMISC"
END IF

FUNCTION AreYouSure (FlagDate)

  REDIM TempScrn(0)
  SaveScrn TempScrn()
  LibName$ = "UBSETUP"
  ScrnName$ = "AREUSURE"
  NumScrns = 1

  '--Initialize the form name array
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1

  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode

  '--Set screen number to one and display screen

  BlockClear

  LibFile2Scrn LibName$, ScrnName$, MonoCode, Attribute%, ErrCode
  QPrintRC Num2Date$(FlagDate), 13, 49, -1

  Action = 1

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE -68    'F10Key
      AreYouSure = True
      ExitFlag = True
    CASE ESCKey
      AreYouSure = False
      ExitFlag = True
      RestScrn TempScrn()
    END SELECT
  LOOP UNTIL ExitFlag
  'RestScrn TempScrn()

END FUNCTION

SUB FlagTractions

  TodayDate = Date2Num(DATE$)

  UBLog " IN: Flag Transactions"

RateTop:
  REDIM ScrnArray(0)

  LibName$ = "UBSETUP"
  ScrnName$ = "FLAGDATE"

  NumScrns = 1

  '--define the multi-choice fields
  'SHARED Choice$()
  'REDIM Choice$(0 TO 4, 0)

  '--Initialize the form name array
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

  '--Set screen number to one and display screen
  Scr = 1

  BlockClear

  LibFile2Scrn LibName$, ScrnName$, MonoCode, Attribute%, ErrCode

  ShowCursor

  Action = 1

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE -68    'F10Key
      GOSUB ValidateDate.AskOK
      IF DateOK THEN
        DoItFlag = True
      ELSE
        Frm(1).FldNo = 1
      END IF
    CASE ESCKey
      UBLog "ABORTED FLAG TRANSACTIONS"
      ExitFlag = True
    END SELECT

    '--Check screen page
  LOOP UNTIL ExitFlag OR DoItFlag

  IF DoItFlag THEN
    GOSUB FlagTransactions
  END IF

ExitFlagTran:
  UBLog "OUT: Flag Transactions" + CrLf$
  HideCursor
  EXIT SUB

ValidateDate.AskOK:

  FlagDate = Date2Num(Form$(1, 0))
  IF FlagDate < 0 OR FlagDate > TodayDate THEN
    ok = MsgBox("UBSETUP", "BADFDATE")
    Action = 1
  ELSE
    IF AreYouSure(FlagDate) THEN
      DateOK = True
    ELSE
      Action = 1
    END IF
  END IF

  'RestScrn ScrnArray()

RETURN

FlagTransactions:
  REDIM UBTransRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTransRec(1))

  BlockClear
  ShowProcessingScrn "Flaging Transactions as Interfaced"

  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen
  NumOfTRecs& = LOF(UBTran) \ UBTranRecLen
  
  FOR TCnt& = 1 TO NumOfTRecs&
    GET UBTran, TCnt&, UBTransRec(1)
    IF UBTransRec(1).TransDate <= FlagDate THEN
      UBTransRec(1).Posted2GL = "Y"
      PUT UBTran, TCnt&, UBTransRec(1)
    END IF
    ShowPctCompL TCnt&, NumOfTRecs&
  NEXT

  BlockClear
  DisplayUBScrn "UPDATEOK"
  WaitForAction


RETURN
END SUB

SUB MakeBKGround
CALL BlockClear
'BK$ = STRING$(80, 177)
'FOR RowCnt = 1 TO 25
'  QPrintRC BK$, RowCnt, 1, 3
'NEXT

END SUB

FUNCTION nGetPassword$ (Row%, Col%, Max%, Abort%)
  
Start:
  
  CurPos = Col
  
  DO
    
    LOCATE Row, CurPos, 1
    P$ = INPUT$(1)
    
    SELECT CASE P$
      
    CASE CHR$(8)                'Backspace
      CurPos = CurPos - 1
      
      IF CurPos <= Col THEN     'Check for backspace key at
        LOCATE Row, Col         ' start of input
        PRINT " "
        PW$ = ""
        GOTO Start
      END IF
      
      LOCATE Row, CurPos: PRINT " "             'Clear the last *
      
    CASE CHR$(13)               'Enter
      EXIT DO
      
    CASE CHR$(27)               'Esc
      Abort = -1
      EXIT FUNCTION
      
    CASE CHR$(32) TO CHR$(126)  'ignore all other keys
      LOCATE Row, CurPos        ' except keyboard chars
      PRINT "*"
      PW$ = PW$ + P$
      CurPos = CurPos + 1
      
    END SELECT
    
  LOOP UNTIL LEN(PW$) = Max
  
  nGetPassword$ = UCASE$(PW$)
  LOCATE , , 0
END FUNCTION

SUB oWazzWind2 (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
  STATIC CPUSpeed&

  IF CPUSpeed& = 0 THEN
    CPUSpeed& = FudgeFactor&
  END IF

  REDIM Frame$(1 TO 6)

  IF FrameType < 1 OR FrameType > 5 THEN FrameType = 5

  Style$ = MID$("⁄ƒø≥¿Ÿ’Õ∏≥‘æ÷ƒ∑∫”Ω…Õª∫»º      ", ((FrameType - 1) * 6) + 1, 6)

  FOR A = 1 TO 6
    Frame$(A) = MID$(Style$, A, 1)
  NEXT

  LDCol = LeftCol + ((RghtCol - LeftCol) \ 2)
  RDCol = RghtCol - ((RghtCol - LeftCol) \ 2)
  TDRow = TopRow + ((BotRow - TopRow) \ 2)
  BDRow = BotRow - ((BotRow - TopRow) \ 2)

  CALL HideCursor
  DO
    LDCol = LDCol - 3
    IF LDCol < LeftCol THEN LDCol = LeftCol
    RDCol = RDCol + 3
    IF RDCol > RghtCol THEN RDCol = RghtCol

    IF TDRow > TopRow THEN TDRow = TDRow - 1
    IF BDRow < BotRow THEN BDRow = BDRow + 1

    Cnt = (RDCol - LDCol) - 1

    T1$ = STRING$(Cnt, Frame$(2))
    Middle$ = Frame$(4) + STRING$(Cnt, 32) + Frame$(4)
    QPrintRC Frame$(1) + T1$ + Frame$(3), TDRow, LDCol, FrameColor
    FOR i = (TDRow + 1) TO (BDRow - 1)
      QPrintRC Middle$, i, LDCol, FrameColor
    NEXT
    QPrintRC Frame$(5) + T1$ + Frame$(6), BDRow, LDCol, FrameColor
    IF Shadow THEN
      MPaintBox BDRow + 1, LDCol + 2, BDRow + 1, LDCol + Cnt + 3, 8
      MPaintBox TDRow + 1, RDCol + 1, BDRow, RDCol + 2, 8
    END IF

    Pause3 7, CPUSpeed&

  LOOP UNTIL LDCol = LeftCol AND BDRow = BotRow

  CALL ShowCursor
  ERASE Frame$

END SUB

