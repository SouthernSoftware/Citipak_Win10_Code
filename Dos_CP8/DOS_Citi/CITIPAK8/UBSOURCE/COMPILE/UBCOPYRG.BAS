DEFINT A-Z
DECLARE SUB MakeBKGround ()
DECLARE SUB WaitForKey ()
DECLARE FUNCTION GetPassword$ (Row%, Col%, MaxLen%, Abort%)
DECLARE SUB QPrintRC (t$, r%, c%, clr%)
DECLARE SUB QPSound (Frequency%, Duration%)
DECLARE FUNCTION WaitKey
'DECLARE FUNCTION BiosKey
DECLARE SUB WazzWind2 (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE SUB MPaintBox (UlRow, UlCol, LRRow, LRCol, Colr)
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
DECLARE FUNCTION Exist% (FileName$)

CONST False = 0, True = NOT False

LOCATE 1, 1, 0

PassCol = 40

  Pz$ = ""
  z$ = "*==63361"    'Copyright Password is: UBBILLIN
  FOR Cnt = 1 TO LEN(z$)
    Pz$ = Pz$ + CHR$(ASC(MID$(z$, Cnt, 1)) XOR 127)
  NEXT

GetPW:
  DO
    MakeBKGround
    WazzWind2 7, 15, 17, 65, 10, 1, 1
    QPrintRC "Southern Software Systems, Inc.", 9, 24, 15
    QPrintRC "Utility Billing v9.01 (C)97-99", 11, 24, 15
    QPrintRC "Enter PASSWORD:", 14, PassCol - 16, 14
    QPrintRC "               ", 14, PassCol, 15

    IF Try > 0 THEN
      IF Try <> 3 THEN
        QPrintRC "Attempt:" + STR$(Try), 16, 35, 7
      ELSE
        QPrintRC "LAST CHANCE", 16, 34, 143
      END IF
    END IF

    COLOR 15
    
    PW$ = GetPassword(14, PassCol, 10, Abort)
    
    IF Abort THEN
      ExitFlag = True
      EXIT DO
    END IF

    IF PW$ = Pz$ THEN
      OkFlag = True
      EXIT DO
    END IF
    
    Try = Try + 1

    IF Try = 4 THEN
      ExitFlag = -1
      MakeBKGround
      WazzWind2 8, 15, 16, 65, 15, 4, 1
      QPrintRC "Get the password from the system operator.", 11, 20, 14
      QPrintRC "Press any key to return to the Main Menu.", 13, 20, 14
      WaitForKey
      EXIT DO
    END IF
    
    QPrintRC "Invalid Password. Press any key to try again.", 14, 18, 31

  LOOP
  

  IF ExitFlag THEN
    IF Exist("CPMENU.EXE") THEN
      RUN "CPMenu"
    ELSE
      GOTO GetOut
    END IF
  ELSEIF OkFlag THEN
    MakeBKGround
    RUN "UBMENU"
  ELSE
    GOTO GetPW
  END IF

GetOut:
CLS
END

FUNCTION GetPassword$ (Row%, Col%, MaxLen%, Abort%)
  
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
      
      LOCATE Row, CurPos
      PRINT " "             'Clear the last *
      
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
    
  LOOP UNTIL LEN(PW$) = MaxLen
  
  GetPassword$ = UCASE$(PW$)
  LOCATE , , 0
END FUNCTION

SUB MakeBKGround

BK$ = STRING$(80, 177)
FOR RowCnt = 1 TO 25
  QPrintRC BK$, RowCnt, 1, 3
NEXT

END SUB

SUB WaitForKey
LabelForKeyTrap:
DO
  a$ = INKEY$
LOOP UNTIL LEN(a$) > 0
END SUB

SUB WazzWind2 (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)

  REDIM Frame$(1 TO 6)

  IF FrameType < 1 OR FrameType > 5 THEN FrameType = 1

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

  LOOP UNTIL LDCol = LeftCol AND BDRow = BotRow

  CALL ShowCursor
  ERASE Frame$

END SUB

