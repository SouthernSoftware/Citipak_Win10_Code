DEFINT A-Z
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

ON ERROR GOTO GetOut

LOCATE 1, 1, 0

GetPW:
  DO
    MakeBKGround
    WazzWind2 8, 15, 16, 65, 15, 1, 1
    QPrintRC "Southern Software Systems, Inc.", 10, 22, 15
    QPrintRC "Tax Billing System v8.20", 12, 22, 15
    QPrintRC "Enter the password to begin:", 14, 22, 15
    
    COLOR 14
    
    PW$ = nGetPassword(14, 51, 10, Abort)
    
    IF Abort THEN RUN "CPMENU"

    IF PW$ = "TAXTIME" THEN
      OkFlag = -1
      EXIT DO
    END IF
    
    Try = Try + 1
    
    IF Try = 4 THEN
      ExitFlag = -1
      MakeBKGround
      WazzWind2 8, 15, 16, 65, 15, 1, 1
      QPrintRC "Get the password from the system operator.", 11, 20, 14
      QPrintRC "Press any key to return to the Main Menu.", 13, 20, 14
      WaitForKey
      EXIT DO
    END IF
    
    'LOCATE 1, 1, 0
    'QPSound 100, 3
    'QPSound 40, 3
    QPrintRC "Invalid Password. Press any key to try again.", 14, 18, 31

    WaitForKey
    
    IF K = 27 THEN
      GOTO GetOut
    END IF
  LOOP
  
  IF ExitFlag THEN
    RUN "CPMenu"
  ELSEIF OkFlag THEN
    MakeBKGround
    RUN "TAXMENU"
  ELSE
    GOTO GetPW
  END IF

GetOut:
CLS
END

RETURN

SUB MakeBKGround

BK$ = STRING$(80, 177)
FOR RowCnt = 1 TO 25
  QPrintRC BK$, RowCnt, 1, 3
NEXT

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

SUB WaitForKey
LabelForKeyTrap:
DO
  a$ = INKEY$
LOOP UNTIL LEN(a$) > 0
END SUB

SUB WazzWind2 (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
  STATIC CPUSpeed&

  IF CPUSpeed& = 0 THEN
    CPUSpeed& = FudgeFactor&
  END IF

  REDIM Frame$(1 TO 6)

  IF FrameType < 1 OR FrameType > 5 THEN FrameType = 5

  Style$ = MID$("⁄ƒø≥¿Ÿ’Õ∏≥‘æ÷ƒ∑∫”Ω…Õª∫»º      ", ((FrameType - 1) * 6) + 1, 6)

  FOR a = 1 TO 6
    Frame$(a) = MID$(Style$, a, 1)
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

