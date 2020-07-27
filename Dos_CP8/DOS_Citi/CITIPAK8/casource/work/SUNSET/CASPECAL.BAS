DECLARE SUB WaitForThem ()
DEFINT A-Z
DECLARE FUNCTION LotSize# (Desc$)
DECLARE SUB WInput (Edit$, GoodKey$, Row%, Col%, ExitCode%)
DECLARE FUNCTION RoundDbl# (DblNum#)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION WEnvTest% ()
DECLARE SUB ClearBack ()
DECLARE SUB ClearScrn ()
DECLARE SUB CursorOff ()
DECLARE SUB HideCursor ()
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB ShowCursor ()
DECLARE FUNCTION QPStrI$ (Num%)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValI% (Num$)
DECLARE FUNCTION QPValL& (Num$)
DECLARE FUNCTION Date2Num (DateIn$)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (Number%)
DECLARE FUNCTION BiosKey% ()

  CONST False = 0, True = NOT False

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'CACust.BI'

  STACK 5000


  Image1$ = "######.##"
  Image2$ = "######"

  NumKeys$ = "1234567890."
  DrvKeys$ = "AaBbcC"
  q$ = CHR$(34)

  COLOR 15, 1

top:
  DO
    CLS
    LOCATE 2, 15
    PRINT "Special LOT calculation"
    Edit$ = "      "
    LOCATE 6, 9
    PRINT "LOT Size: ";
    WInput Edit$, NumKeys$, 6, 19, ExitCode
    IF ExitCode = -27 OR LEN(Edit$) = 0 THEN
      ExitFlag = True
      EXIT DO
    END IF
    Lot2Find# = VAL(Edit$)

GetBucksPerFoot:
    Edit$ = "       "
    LOCATE 8, 9
    PRINT "Per Foot:     ";
    WInput Edit$, NumKeys$, 8, 19, ExitCode
    IF ExitCode = -27 OR LEN(Edit$) = 0 THEN
      ExitFlag = True
      EXIT DO
    END IF

    BuckPer# = VAL(Edit$)
    IF BuckPer# <= 0 OR BuckPer# > 9999 THEN
      LOCATE 12, 10
      PRINT "Invalid 'Per Foot' Specification."
      LOCATE 14, 11
      PRINT "Press any key to continue."
      dodo = BiosKey
      CTop = 8
      GOSUB ClearArea
      GOTO GetBucksPerFoot
    END IF
    '-----

GetWhatRev:
    COLOR 14, 1
    LOCATE 11, 7: PRINT "1) Street         2) SideWalk      3) Driveway"
    LOCATE 12, 7: PRINT "4) Demolition     5) Utility       6) Dredging"
    COLOR 15, 1
    LOCATE 11, 7:  PRINT "1)"
    LOCATE 11, 25: PRINT "2)"
    LOCATE 11, 42: PRINT "3)"
    LOCATE 12, 7:  PRINT "4)"
    LOCATE 12, 25: PRINT "5)"
    LOCATE 12, 42: PRINT "6)"

    LOCATE 14, 7
    PRINT "Rev Number:     ";
    Edit$ = " "
    WInput Edit$, "1234567", 14, 19, ExitCode
    IF ExitCode = -27 OR LEN(Edit$) = 0 THEN
      ExitFlag = True
      EXIT DO
    END IF

    RevNo = VAL(Edit$)
    IF RevNo < 1 OR RevNo > 6 THEN
      LOCATE 16, 10
      PRINT "Invalid 'Rev Number' Specification."
      LOCATE 18, 11
      PRINT "Press any key to continue."
      dodo = BiosKey
      CTop = 10
      GOSUB ClearArea
      GOTO GetWhatRev
    END IF
    '-----
    EXIT DO
  LOOP
  IF ExitFlag THEN
    LOCATE 20, 7
    PRINT "Processing Aborted!!"
    LOCATE 22, 7
    PRINT "Press any Key.";
    WaitForThem
    RUN "camenu"
  END IF

  DIM CACust AS CACUSTType
  CustLen = LEN(CACust)
  CustFile = FREEFILE
  OPEN "camaster.dat" FOR RANDOM AS CustFile LEN = CustLen
  NumCust& = LOF(CustFile) / CustLen
  FOR Cnt& = 1 TO NumCust&
    LOCATE 16, 1
    PRINT "Processing:"; Cnt&; "of"; NumCust&;
    GET CustFile, Cnt&, CACust
    Desc$ = QPTrim$(CACust.LEGALDES)
    ThisSize# = LotSize#(Desc$)
    IF (ThisSize# = 0) OR (ThisSize# <> Lot2Find#) THEN
      GOTO SkipThisCust
    ELSE
      Amt# = RoundDbl#(ThisSize# * BuckPer#)
      Matched = Matched + 1

      SELECT CASE RevNo
      CASE 1
        CACust.STRAMT = Amt#
      CASE 2
        CACust.SIDEAMT = Amt#
      CASE 3
        CACust.DRIVEAMT = Amt#
      CASE 4
        CACust.DEMOAMT = Amt#
      CASE 5
        CACust.WATERAMT = Amt#
      CASE 6
        CACust.SEWERAMT = Amt#
      END SELECT
      PUT CustFile, Cnt&, CACust
    END IF


SkipThisCust:
  NEXT
  CLOSE
  LOCATE 20, 1
  PRINT "   Changed:"; Matched
  LOCATE 22, 1
  PRINT "Press any Key.";

  WaitForThem
  RUN "camenu"
END

ClearArea:
  T$ = SPACE$(60)
  FOR Cnt = CTop TO 24
    LOCATE Cnt, 1: PRINT T$;
  NEXT
RETURN

FUNCTION LotSize# (Desc$)
  FootPos = INSTR(Desc$, "'")
  IF FootPos <= 0 THEN
    LotSize# = 0
    GOTO ExitLotSize
  END IF
  LotSize# = VAL(LEFT$(Desc$, FootPos))


ExitLotSize:

END FUNCTION

FUNCTION RoundDbl# (DblNum#)
  RoundDbl# = (INT((DblNum# * 100) + .5) / 100)
END FUNCTION

SUB WaitForThem
  DO
  LOOP WHILE LEN(INKEY$) = 0
END SUB

SUB WInput (Edit$, GoodKey$, Row, Col, ExitCode)
  
  'Return Codes:
  
  ' Enter Key Exit
  '     0
  '
  '
  ' Escape Key Exit
  '   -27 Escape Key was Pressed to exit sub
  '       Edit string reset to original string
  '
  
  Original$ = Edit$
  ExitCode = 0
  Ptr = 0
  
  MaxLength = LEN(Edit$)        'max length of input string
  
  LOCATE , , 1  'turn cursor on
  
  DO            'main processing loop
    
    LOCATE Row, Col             'update cursor position.
    PRINT Edit$;                'and display the line
    LOCATE Row, Col + Ptr       'update cursor position.
    
    PressedKey = BiosKey        'wait for a key press
    
    SELECT CASE PressedKey      'process the key
      
      '****  Exit keys
    CASE 13     'ENTER - Accept line and, exit
      Edit$ = RTRIM$(LTRIM$(Edit$))
      IF LEN(Edit$) = 0 THEN
        ExitCode = -1
      ELSE
        ExitCode = 0
      END IF
      EXIT DO
      
    CASE 27     'ESCAPE - Abort operation exit
      ExitCode = -27
      Edit$ = Original$
      EXIT DO
      'add alt exits here
      ' i.e. downarrow
      'End of Exit keys
      
      '****  Cursor keys
    CASE -82    'INSERT - Toggle insert or, overtype mode
      InsertMode = NOT InsertMode
      IF InsertMode THEN
        LOCATE , , , 0, 6
      ELSE      'show appropriate cursor type
        LOCATE , , , 5, 6
      END IF
      
    CASE -75    'LEFTARROW - Moves cursor left
      IF Ptr THEN Ptr = Ptr - 1
      
    CASE -77    'RIGHTARROW - Moves cursor right
      IF Ptr < MaxLength - 1 THEN Ptr = Ptr + 1
      
    CASE -71    'HOME - Moves cursor to first position
      Ptr = 0
      
    CASE -79    'END - Moves cursor to last position
      Ptr = LEN(RTRIM$(Edit$))
      IF Ptr = MaxLength THEN Ptr = MaxLength - 1
      'End of Cursor keys
      
      '****   Editing Keys & Normal Input
    CASE -117   'CONTROLEND - Erase from cursor to end of Edit$
      IF Ptr THEN
        LSET Edit$ = LEFT$(Edit$, Ptr)
      ELSE
        Edit$ = SPACE$(LEN(Edit$))
      END IF
      
    CASE -22    'ALT-U     - Restore to orginial string
      Edit$ = Original$
      Ptr = 0
      
    CASE 8      'BACKSPACE - Moves cursor left and erase characters
      IF Ptr THEN
        Edit$ = LEFT$(Edit$, Ptr - 1) + MID$(Edit$, Ptr + 1) + " "
        Ptr = Ptr - 1
      END IF
      
    CASE -83    'DELETE - Erases characters at cursor
      Edit$ = LEFT$(Edit$, Ptr) + MID$(Edit$, Ptr + 2) + " "
      
    CASE ELSE   'ALL OTHER NORMAL KEYS
      IF PressedKey > 1 AND PressedKey < 256 THEN
        K$ = CHR$(PressedKey)
        IF INSTR(GoodKey$, K$) THEN             'if key is in the keylist
          IF InsertMode THEN    'Insert or Overtype the key
            Edit$ = LEFT$(Edit$, Ptr) + K$ + MID$(Edit$, Ptr + 1)
            Edit$ = LEFT$(Edit$, MaxLength)
          ELSE
            IF Ptr < MaxLength THEN MID$(Edit$, Ptr + 1, 1) = K$
          END IF
          IF Ptr < MaxLength - 1 THEN Ptr = Ptr + 1             'Are we at max length
        END IF
      END IF
      'End of Editing & Input keys
    END SELECT
  LOOP
  
  LOCATE , , 0  'turn cursor off
  
END SUB

