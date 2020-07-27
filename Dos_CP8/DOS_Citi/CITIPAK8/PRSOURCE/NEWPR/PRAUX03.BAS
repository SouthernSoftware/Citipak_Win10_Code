DEFINT A-Z

DECLARE FUNCTION NovellThere% ()
DECLARE FUNCTION CheckFldEmpty% (FldData$)
DECLARE FUNCTION QPStrI$ (Num%)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION FudgeFactor& ()
DECLARE SUB BlockClear ()
DECLARE SUB CursorOff ()
DECLARE SUB HideCursor ()
DECLARE SUB ClearScrn ()
DECLARE SUB Pause3 (MillaSecs%, ProcessorSpeed&)
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
DECLARE SUB ShowCursor ()
DECLARE SUB DisplayMiscScrn (ScrnName$)
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE FUNCTION Exist% (FileName$)
DECLARE SUB MScrnSave (UlRow, UlCol, LRRow, LRCol, SEG Address)
DECLARE SUB MScrnRest (UlRow, UlCol, LRRow, LRCol, SEG Address)
DECLARE SUB DisplayScrn (BTmp%(), Element%, MonoCode%, WipeType%)
DECLARE SUB ExplodBox (UlRow%, UlCol%, BrRow%, BrCol%, Scr(), El%)

'$INCLUDE: 'DefCnf.bi'      'Defines a TYPE for monitor and color info..
'$INCLUDE: 'FORMEDIT.bi'
'$INCLUDE: 'PRFiles.bi'
'INCLUDE: 'RegType.bi'

CONST False = 0
CONST True = NOT False

DIM SHARED PctC(1) AS STRING * 3

SUB BlockClear

  REDIM Scrn(2000)
  CALL BlkClear(Wipe, SEG Scrn(0))
  HideCursor
  ExplodBox 1, 1, 25, 80, Scrn(), 0
  ShowCursor
  ERASE Scrn

'  STATIC CPUSpeed&
'
'  IF CPUSpeed& = 0 THEN
'    CPUSpeed& = FudgeFactor&
'  END IF
'
'  TopRow = 1
'  LeftCol = 1
'  BotRow = 25
'  RghtCol = 80
'  FrameColor = 23
'
'  LDCol = LeftCol + ((RghtCol - LeftCol) \ 2)
'  RDCol = RghtCol - ((RghtCol - LeftCol) \ 2)
'  TDRow = TopRow + ((BotRow - TopRow) \ 2)
'  BDRow = BotRow - ((BotRow - TopRow) \ 2)
'
'  HideCursor
'
'  DO
'    LDCol = LDCol - 3
'    IF LDCol < LeftCol THEN LDCol = LeftCol
'    RDCol = RDCol + 3
'    IF RDCol > RghtCol THEN RDCol = RghtCol
'
'    IF TDRow > TopRow THEN TDRow = TDRow - 1
'    IF BDRow < BotRow THEN BDRow = BDRow + 1
'
'    t1$ = STRING$(RDCol - LDCol + 1, "±")
'
'    FOR i = TDRow TO BDRow
'      QPrintRC t1$, i, LDCol, FrameColor
'    NEXT
'
'    Pause3 7, CPUSpeed&
'
'  LOOP UNTIL LDCol = LeftCol AND BDRow = BotRow
'  t1$ = ""
'  ShowCursor

END SUB

FUNCTION CheckFldEmpty (FldData$)
 DO
   OffSet = INSTR(FldData$, "-")
   IF OffSet > 0 THEN
     FldData$ = MID$(FldData$, 1, OffSet - 1) + MID$(FldData$, OffSet + 1)
   END IF
 LOOP WHILE OffSet
 FldLen = LEN(QPTrim$(FldData$))
 IF FldLen = 0 THEN
   CheckFldEmpty = True
 ELSE
   CheckFldEmpty = False
 END IF
END FUNCTION

SUB ClearScrn STATIC
  WazzWind 1, 1, 25, 80, 7, 0, 0
END SUB

SUB CursorOff STATIC
  LOCATE , , 0
END SUB

SUB ShowPctComp (BYVAL RecNo, BYVAL NumOfRecs) STATIC
  RSET PctC(1) = QPStrI$(INT((RecNo / NumOfRecs) * 100))
  QPrintRC PctC(1), 13, 40, Cnf.HiLite
'  QPrintRC STR$(FRE("")), 25, 1, Cnf.HiLite
END SUB

SUB ShowProcessingScrn (RptTitle$)

  TitleRow = 9
  TitleCol = 40 - (LEN(RptTitle$) \ 2) + 1

  CursorOff
  BlockClear
  DisplayMiscScrn ProcessReport

  QPrintRC RptTitle$, TitleRow, TitleCol, 126
  QPrintRC "Processing:    % Completed.", 13, 28, Cnf.HiLite

END SUB

'FUNCTION NovellThere
'  REDIM Registers(1) AS RegType
'  Registers(1).AX = 31232                       ' hex &H7A00
'  CALL Interrupt(47, Registers(1), Registers(1))   ' hex &H2F
'  AL = Registers(1).AX AND 255
'  IF AL = 255 THEN NovellThere = True
'END FUNCTION

