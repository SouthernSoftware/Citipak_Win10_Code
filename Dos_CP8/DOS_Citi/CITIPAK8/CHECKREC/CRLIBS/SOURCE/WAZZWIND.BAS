DEFINT A-Z

DECLARE SUB Pause3 (MillaSecs%, ProcessorSpeed&)
DECLARE SUB MPaintBox (UlRow, UlCol, LRRow, LRCol, Colr)
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE FUNCTION FudgeFactor& ()

SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)

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

