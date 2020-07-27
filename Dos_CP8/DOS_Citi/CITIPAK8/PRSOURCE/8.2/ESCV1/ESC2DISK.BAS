DEFINT A-Z
DECLARE SUB DoWRecs ()
DECLARE FUNCTION RSet0$ (Amt#, StrLen%)
DECLARE FUNCTION LSet0$ (Amt#, StrLen%)
DECLARE SUB DoABERecs ()
  
DECLARE SUB WInput (Edit$, GoodKey$, Row%, Col%, ExitCode%)
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
  

  ''$INCLUDE: 'MAGtype.Bi'
  
  '$INCLUDE: 'DefCnf.bi'                  'Defines TYPE for monitor/color info.
  '$INCLUDE: 'PRFiles.bi'
  '$INCLUDE: 'PREmpRec.bi'
  '$INCLUDE: 'PRUNIT.Bi'
  '$INCLUDE: 'PRSYSCTR.Bi'
  '$INCLUDE: 'PRTRANS.Bi'
  '$INCLUDE: 'PRRpts.BI'
  
  CONST False = 0, True = NOT False
  
  COLOR 15, 1

top:
  DO
    CLS
    LOCATE 2, 15
    PRINT "ESC Mag Media Reporting "

    Edit$ = " "
    LOCATE 6, 11
    PRINT "Enter Drive(A-B):"
    WInput Edit$, "ABabCc", 6, 29, ExitCode
    Edit$ = QPTrim$(UCASE$(Edit$))

    IF ExitCode = -27 THEN
      ExitFlag = True
      EXIT DO
    END IF
    IF Edit$ = "" THEN
      GOTO top
    END IF

    Drive$ = Edit$

    Edit$ = " "
    LOCATE 8, 9
    PRINT "Enter Quarter(1-4):"
    WInput Edit$, "1234", 8, 29, ExitCode
    Edit$ = QPTrim$(UCASE$(Edit$))

    IF ExitCode = -27 THEN
      ExitFlag = True
      EXIT DO
    END IF
    IF Edit$ = "" THEN
      GOTO top
    END IF

    Quarter$ = Edit$
    Qtr = VAL(Quarter$)

    Edit$ = RIGHT$(DATE$, 4)
    LOCATE 10, 17
    PRINT "Enter Year:"
    WInput Edit$, "1234567890", 10, 29, ExitCode
    Edit$ = QPTrim$(UCASE$(Edit$))

    IF ExitCode = -27 THEN
      ExitFlag = True
      EXIT DO
    END IF
    IF LEN(Edit$) < 4 THEN
      GOTO top
    END IF

    Year$ = Edit$

    OkFlag = True

  LOOP UNTIL OkFlag
  
  IF ExitFlag THEN
    GOTO EndTheProg
  END IF
  
  REDIM TransHRec(1)   AS TransRecType
  REDIM Emp2Rec(1)     AS EmpData2Type
  REDIM Unit(1)        AS UnitFileRecType
  REDIM ESCRec(1)      AS ESCMAG2DiskType

  REDIM QtrDates(1 TO 4) AS QtrDateType

  QtrDates(1).LDate = Date2Num("0101" + Year$)
  QtrDates(1).HDate = Date2Num("0331" + Year$)
  QtrDates(2).LDate = Date2Num("0401" + Year$)
  QtrDates(2).HDate = Date2Num("0630" + Year$)
  QtrDates(3).LDate = Date2Num("0701" + Year$)
  QtrDates(3).HDate = Date2Num("0930" + Year$)
  QtrDates(4).LDate = Date2Num("1001" + Year$)
  QtrDates(4).HDate = Date2Num("1231" + Year$)

  ESCRecLen = LEN(ESCRec(1))
  TRecLen = LEN(TransHRec(1))
  EmpRecLen = LEN(Emp2Rec(1))

  IdxRecLen = 2
  IdxFileSize& = FileSize(EmpIdxLName)
  NumOfRecs = IdxFileSize& \ IdxRecLen

  REDIM IdxBuff(1 TO NumOfRecs)
  FGetAH EmpIdxLName, IdxBuff(1), IdxRecLen, NumOfRecs

  FGetAH UnitFileName, Unit(1), LEN(Unit(1)), 1
  
  ESCReport$ = Drive$ + ":\ESCNC.WGS"
  
  ESCFile = FREEFILE
  OPEN ESCReport$ FOR RANDOM AS #ESCFile LEN = ESCRecLen

  TRFile = FREEFILE
  OPEN TransHistFileName FOR RANDOM AS #TRFile LEN = TRecLen

  EMPFile = FREEFILE
  OPEN EmpData2Name FOR RANDOM AS #EMPFile LEN = EmpRecLen

  FOR RecNo = 1 TO NumOfRecs
    GET #EMPFile, RecNo, Emp2Rec(1)
    IF Emp2Rec(1).LastTransRec <= 0 THEN
      GOTO SkipEm
    END IF

    REDIM EQtrWage(1 TO 4) AS DOUBLE
    TransRecNum& = Emp2Rec(1).LastTransRec

    DO
      GET #TRFile, TransRecNum&, TransHRec(1)

      GOSUB SumESCData

      IF TransHRec(1).PrevTransRec <= 0 THEN
        GOSUB PrintESCLine
        EXIT DO
      ELSE
        TransRecNum& = CLNG(TransHRec(1).PrevTransRec)
      END IF
    LOOP

SkipEm:
  NEXT

  CLOSE
PRINT
PRINT
PRINT "Number of Employee Records:"; EmpCnt
PRINT
PRINT
PRINT "Processing Complete."
PRINT
PRINT
PRINT "   Press any key to continue. . ."
a$ = INPUT$(1)


EndTheProg:
  'RUN "CPMENU"
  END

SumESCData:
  FOR Cnt2 = 1 TO 4  'put gross into correct quarter
    IF (TransHRec(1).CheckDate >= QtrDates(Cnt2).LDate) AND (TransHRec(1).CheckDate <= QtrDates(Cnt2).HDate) THEN
      EQtrWage(Cnt2) = RoundDbl(EQtrWage(Cnt2) + TransHRec(1).GrossPay)
      EXIT FOR
    END IF
  NEXT

RETURN

PrintESCLine:
  EmpCnt = EmpCnt + 1
  REDIM ESCRec(1)      AS ESCMAG2DiskType
  CPos = INSTR(Emp2Rec(1).EMPLNAME, ",")
  IF CPos > 0 THEN
    LName$ = LEFT$(Emp2Rec(1).EMPLNAME, CPos - 1)
  ELSE
    LName$ = Emp2Rec(1).EMPLNAME
  END IF
  LName$ = QPTrim$(LName$)

  ESCRec(1).Blank1 = " "
  ESCRec(1).SSN = Emp2Rec(1).EMPSSN
  ESCRec(1).LastName = LName$
  ESCRec(1).EmpInitials = LEFT$(QPTrim$(Emp2Rec(1).EMPFNAME), 1) + LEFT$(QPTrim$(Emp2Rec(1).EMPLNAME), 1)
  ESCRec(1).EmpWages = RSet0$(EQtrWage(Qtr), 9)

  ESCRec(1).SeasInd = "N"
  ESCRec(1).RemitNumb = "009816"
  ESCRec(1).EmplrAcct = "9904993"
  ESCRec(1).BranchAcct = ""
  ESCRec(1).RQuarter = Quarter$
  ESCRec(1).RYear = Year$
  ESCRec(1).EmplrName = UCASE$(QPTrim$(Unit(1).UFEMPR))
  ESCRec(1).Blank2 = " "
  ESCRec(1).CrLf = CHR$(13) + CHR$(10)

  PUT #ESCFile, , ESCRec(1)
RETURN

FUNCTION RoundDbl# (DblNum#)
  RoundDbl# = (INT((DblNum# * 100) + .5) / 100)
END FUNCTION

FUNCTION RSet0$ (Amt#, StrLen)

  Temp$ = STRING$(StrLen, "0")
  NumStr$ = QPTrim$(STR$(Amt#))
  
  Bit$ = RIGHT$(NumStr$, 2)
  
  IF INSTR(Bit$, ".") THEN
    NumStr$ = NumStr$ + "0"
  END IF
  
  ChrPos = INSTR(NumStr$, ".")
  IF ChrPos THEN
    NewStr$ = LEFT$(NumStr$, ChrPos - 1) + MID$(NumStr$, ChrPos + 1)
  ELSE
    NewStr$ = NumStr$ + "00"
  END IF
  
  NumStr$ = QPTrim$(NewStr$)
  
  NumLen = LEN(NumStr$)
  StartPos = (StrLen - NumLen) + 1
  MID$(Temp$, StartPos) = NumStr$
  RSet0$ = Temp$
  
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

