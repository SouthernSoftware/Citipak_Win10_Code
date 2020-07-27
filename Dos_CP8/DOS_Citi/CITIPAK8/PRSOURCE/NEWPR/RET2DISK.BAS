DEFINT A-Z
DECLARE SUB WInput (Edit$, NumKeys$, Row%, Col%, ExitCode%)
DECLARE FUNCTION BiosKey%
  
DECLARE FUNCTION Date2Num (DateIn$)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (Number%)
DECLARE FUNCTION QPStrI$ (Num%)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValI% (Num$)
DECLARE FUNCTION QPValL& (Num$)
DECLARE FUNCTION RoundDbl# (DblNum#)
DECLARE FUNCTION DosError ()
  
DECLARE SUB FClose (Handle%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FPut (Handle%, ToDisk$)
DECLARE SUB FPutA (Handle%, SEG Element AS ANY, NumBytes AS ANY)
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
  
  '$INCLUDE: 'PRFiles.bi'
  '$INCLUDE: 'PREmpRec.bi'
  '$INCLUDE: 'PRTRANS.Bi'
  
  CONST False = 0, True = NOT False
  
  REDIM TRec(1)     AS TransRecType
  REDIM E2Rec(1)    AS EmpData2Type
  
  TRecSize = LEN(TRec(1))
  EmpRecSize = LEN(E2Rec(1))
  
  NumKeys$ = "1234567890"
  DrvKeys$ = "AaBbCc"
  q$ = CHR$(34)
  
  COLOR 15, 1

top:
  DO
    CLS
    LOCATE 2, 15
    PRINT "Mag Media Retirement Report"
    Edit$ = " "
    LOCATE 6, 9
    PRINT "Enter Drive(A-B): ";
    WInput Edit$, DrvKeys$, 6, 29, ExitCode
    IF ExitCode = -27 OR LEN(Edit$) = 0 THEN
      ExitFlag = True
      EXIT DO
    END IF
    Drive$ = Edit$

GetStartMonth:
    Edit$ = "  "
    LOCATE 7, 10
    PRINT "Beginning Month:     ";
    WInput Edit$, NumKeys$, 7, 29, ExitCode
    IF ExitCode = -27 OR LEN(Edit$) = 0 THEN
      ExitFlag = True
      EXIT DO
    END IF
    BMonth = QPValI(Edit$)
    IF BMonth < 1 OR BMonth > 12 THEN
      LOCATE 12, 10
      PRINT "Invalid Month Specification."
      LOCATE 14, 11
      PRINT "Press any key to continue."
      dodo = BiosKey
      GOSUB ClearArea
      GOTO GetStartMonth
    END IF
    '-----

GetLastMonth:
    Edit$ = "  "
    LOCATE 8, 10
    PRINT "   Ending Month:     ";
    WInput Edit$, NumKeys$, 8, 29, ExitCode
    IF ExitCode = -27 OR LEN(Edit$) = 0 THEN
      ExitFlag = True
      EXIT DO
    END IF
    EMonth = QPValI(Edit$)
    IF (EMonth < 1 OR EMonth > 12) OR BMonth > EMonth THEN
      LOCATE 12, 10
      PRINT "Invalid Month Specification."
      LOCATE 14, 11
      PRINT "Press any key to continue."
      dodo = BiosKey
      GOSUB ClearArea
      GOTO GetLastMonth
    END IF
    '-----

GetYear:
    Edit$ = RIGHT$(DATE$, 4)
    LOCATE 9, 11
    PRINT "Enter the Year: ";
    WInput Edit$, NumKeys$, 9, 29, ExitCode
    IF ExitCode = -27 THEN
      ExitFlag = True
      EXIT DO
    END IF
    Year = QPValI(Edit$)
    IF Year < 1994 OR Year > 2065 THEN
      LOCATE 12, 10
      PRINT "Invalid Year Specifcation."
      LOCATE 14, 10
      PRINT "Press any key to continue."
      dodo = BiosKey
      GOSUB ClearArea
      GOTO GetYear
    ELSE
      OkFlag = True
    END IF
    
  LOOP UNTIL OkFlag
  
  IF ExitFlag THEN
    GOTO EndTheProg
  END IF
  
  IF BMonth < 10 THEN
    BMonth$ = "0" + LTRIM$(STR$(BMonth))
  ELSE
    BMonth$ = LTRIM$(STR$(BMonth))
  END IF

  IF EMonth < 10 THEN
    EMonth$ = "0" + LTRIM$(STR$(EMonth))
  ELSE
    EMonth$ = LTRIM$(STR$(EMonth))
  END IF
  
  Year$ = LTRIM$(STR$(Year))
  
  LowDate = Date2Num(BMonth$ + "-" + "01" + "-" + Year$)
  
  SELECT CASE EMonth
  CASE 2
    HiDate = Date2Num(EMonth$ + "-" + "28" + "-" + Year$)
  CASE 4, 6, 9, 11
    HiDate = Date2Num(EMonth$ + "-" + "30" + "-" + Year$)
  CASE 1, 3, 5, 7, 8, 10, 12
    HiDate = Date2Num(EMonth$ + "-" + "31" + "-" + Year$)
  END SELECT
  
  IdxRecLen = 2
  IdxFileSize& = FileSize(EmpIdxNName)
  NumOfRecs = IdxFileSize& \ IdxRecLen

  IF DosError THEN
    LOCATE 15, 10
    PRINT "Unable to Find/Open Transaction History file!"
    LOCATE 16, 10
    PRINT "Press any key to return to system."
    dodo = BiosKey
    GOTO EndTheProg
  END IF
  
  REDIM IdxBuff(1 TO NumOfRecs)
  FGetAH EmpIdxNName, IdxBuff(1), IdxRecLen, NumOfRecs
  
  RptName$ = Drive$ + ":\MAGRET" + ".RPT"

  'RptName$ = "ASARET.RPT"

  FCreate RptName$
  IF DosError THEN
    LOCATE 15, 10
    PRINT "Unable to Open/Create report file!"
    LOCATE 16, 10
    PRINT "Press any key to return to system."
    dodo = BiosKey
    GOTO EndTheProg
  END IF
  
  OPEN RptName$ FOR OUTPUT AS #1 LEN = 16384

  HFile = FREEFILE
  OPEN TransHistFileName FOR RANDOM SHARED AS HFile LEN = TRecSize
  EFile = FREEFILE
  OPEN EmpData2Name FOR RANDOM AS EFile LEN = EmpRecSize
  
  GOSUB PrintHeader
  
  FOR RecNo = 1 TO NumOfRecs
    GET #EFile, IdxBuff(RecNo), E2Rec(1)

    IF E2Rec(1).LastTransRec <= 0 THEN
      GOTO SkipEm2
    END IF
    TransRecNum& = CLNG(E2Rec(1).LastTransRec)
    DO
      GET #HFile, TransRecNum&, TRec(1)
      'FGetRTA THandle, TRec(1), TransRecNum&, TRecSize

      SELECT CASE TRec(1).CheckDate
      CASE LowDate TO HiDate
        EGro# = RoundDbl(EGro# + TRec(1).GROSSPAY)
        EHrs# = RoundDbl(EHrs# + TRec(1).RegHrsWork + TRec(1).SickUsed + TRec(1).CompUsed + TRec(1).VacUsed)
        OTHrs# = RoundDbl(OTHrs# + TRec(1).OTHrsPaid)
        
  '      ERet# = RoundDbl(ERet# + TRec(1).RetireAmt)
        
        UsingThisOne = True
      CASE ELSE
      END SELECT
      IF TRec(1).PrevTransRec <= 0 THEN
        IF UsingThisOne THEN
          GOSUB PrintThisOne
        END IF
        EXIT DO
      ELSE
        TransRecNum& = CLNG(TRec(1).PrevTransRec)
      END IF
    LOOP
SkipEm2:
    LOCATE 10, 10:
    PRINT FUsing(STR$((RecNo / NumOfRecs) * 100), "###"); "% Completed.";
  NEXT
  
  CLOSE
  
  LOCATE 10, 1
  PRINT SPACE$(79);
  LOCATE 12, 9
  PRINT "Report Completed."
  LOCATE 14, 2
  PRINT "Press any key to Exit."
  aa$ = INPUT$(1)

EndTheProg:
  RUN "PR"

END
  
PrintThisOne:
  IF E2Rec(1).EMPTDATE > 0 THEN
    Term$ = Num2Date(E2Rec(1).EMPTDATE)
  ELSE
    Term$ = "  -  -    "
  END IF
  
  IF E2Rec(1).EMPBDAY > -29219 THEN
    Birth$ = Num2Date(E2Rec(1).EMPBDAY)
  ELSE
    Birth$ = "  -  -    "
  END IF
  IF E2Rec(1).EMPHDATE > -29219 THEN
    HDay$ = Num2Date(E2Rec(1).EMPHDATE)
  ELSE
    HDay$ = "  -  -    "
  END IF
  City$ = SPACE$(30)
  LSET City$ = E2Rec(1).EMPCITY
  MID$(City$, 20) = (E2Rec(1).EMPSTATE + E2Rec(1).EMPZIP)
  WRITE #1, QPTrim$(E2Rec(1).EMPSSN), (QPTrim$(E2Rec(1).EMPLNAME) + " " + QPTrim$(E2Rec(1).EMPFNAME)), QPTrim$(E2Rec(1).EMPADDR1), City$, Birth$, HDay$, Term$, QPTrim$(FUsing$(STR$(EHrs#), "#####.##")), QPTrim$(FUsing$(STR$(OTHrs#), "#####.##")),  _
QPTrim$(FUsing$(STR$(EGro#), "######.##"))
  UsingThisOne = False

  EHrs# = 0
  OTHrs# = 0
  EGro# = 0
  ERet# = 0
  RETURN
  
PrintHeader:
  WRITE #1, "SSN", "LastName  FirstName", "Addrs", "City    ST     Zip", "BirthDate", "HireDate", "Term Date", "REGHours", "OT Hours", "Gross"
  RETURN
  
ClearArea:
  T$ = SPACE$(60)
  FOR Cnt = 11 TO 18
    LOCATE Cnt, 1: PRINT T$;
  NEXT
RETURN

FUNCTION RoundDbl# (DblNum#)
  RoundDbl# = (INT((DblNum# * 100) + .5) / 100)
END FUNCTION

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

