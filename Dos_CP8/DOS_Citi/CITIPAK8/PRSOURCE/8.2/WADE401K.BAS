DEFINT A-Z
DECLARE FUNCTION RSet0$ (Amt#, StrLen%)
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
DECLARE FUNCTION Round# (DblNum#)
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
  '$INCLUDE: 'PRUNIT.Bi'
  '$INCLUDE: 'DedCodes.Bi'
  '$INCLUDE: 'PRRpts.BI'
  '$INCLUDE: 'PR401k.BI'
  
  CONST False = 0, True = NOT False
  
  REDIM TRec(1)     AS TransRecType
  REDIM E2Rec(1)    AS EmpData2Type
  
  REDIM DedCodes(1 TO 12) AS DedCodeRecType
  FGetAH DedCodeFileName, DedCodes(1), LEN(DedCodes(1)), 12

  REDIM Unit(1)              AS UnitFileRecType
  FGetAH UnitFileName, Unit(1), LEN(Unit(1)), 1
  
  VCodeNum = 11
  LCodeNum = 12
  Image1$ = "######.##"
  Image2$ = "######"
  
  
  TRecSize = LEN(TRec(1))
  EmpRecSize = LEN(E2Rec(1))
  
  NumKeys$ = "1234567890"
  DrvKeys$ = "AaBb"
  q$ = CHR$(34)
  
  COLOR 15, 1
top:
  DO
    CLS
    LOCATE 2, 15
    PRINT "401K BB&T Magnetic Media Report"
    Edit$ = " "
    LOCATE 6, 9
    PRINT "Enter Drive(A-B): ";
    WInput Edit$, DrvKeys$, 6, 29, ExitCode
    IF ExitCode = -27 OR LEN(Edit$) = 0 THEN
      ExitFlag = True
      EXIT DO
    END IF
    Drive$ = Edit$
    
GetLastMonth:
    Edit$ = "  "
    LOCATE 8, 10
    PRINT "Reporting Month:     ";
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
    Edit$ = "    "
    LOCATE 9, 11
    PRINT "Enter the Year:       ";
    WInput Edit$, NumKeys$, 9, 27, ExitCode
    IF ExitCode = -27 THEN
      ExitFlag = True
      EXIT DO
    END IF
    Year = QPValI(Edit$)
    IF Year < 1999 THEN
      LOCATE 12, 10
      PRINT "Invalid Year Specifcation."
      LOCATE 14, 10
      PRINT "Press any key to continue."
      dodo = BiosKey
      GOSUB ClearArea
      GOTO GetYear
    ELSE
      OKFlag = True
    END IF
    
  LOOP UNTIL OKFlag
  
  IF ExitFlag THEN
    GOTO EndTheProg
  END IF
  
  IF EMonth < 10 THEN
    EMonth$ = "0" + LTRIM$(STR$(EMonth))
  ELSE
    EMonth$ = LTRIM$(STR$(EMonth))
  END IF
  
  BMonth$ = EMonth$
  
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
  
  RptName$ = Drive$ + ":\NC401K"
  
  FCreate RptName$
  IF DosError THEN
    LOCATE 15, 10
    PRINT "Unable to Open/Create report file!"
    LOCATE 16, 10
    PRINT "Press any key to return to system."
    dodo = BiosKey
    GOTO EndTheProg
  END IF
  '*****************
  'make disk report here
  
  CrLf$ = CHR$(13) + CHR$(10)

  REDIM TransHRec(1)     AS TransRecType
  REDIM Emp2Rec(1)  AS EmpData2Type
  
  EmpRecSize = LEN(Emp2Rec(1))
  TRecSize = LEN(TransHRec(1))
  
  REDIM DedCodes(1 TO 12) AS DedCodeRecType
  FGetAH DedCodeFileName, DedCodes(1), LEN(DedCodes(1)), 12
  
  IdxRecLen = 2
  
  IdxFileSize& = FileSize(EmpIdxLName)
  NumOfRecs = IdxFileSize& \ IdxRecLen
  
  REDIM IdxBuff(1 TO NumOfRecs)
  FGetAH EmpIdxLName, IdxBuff(1), IdxRecLen, NumOfRecs
  
  'RptTitle$ = "Supplemental Retirement Report."
  'ShowProcessingScrn RptTitle$
  
  REDIM D401kRec(1) AS DetailRecType
  REDIM T401kRec(1) AS TrailerRecType
  D401Len = LEN(D401kRec(1))
  T401Len = LEN(T401kRec(1))
  
  'got input here
  
  RptFile = FREEFILE
  OPEN RptName$ FOR OUTPUT AS #RptFile
  CLOSE RptFile
  
  RptFile = FREEFILE
  OPEN RptName$ FOR RANDOM AS #RptFile LEN = D401Len
  HFile = FREEFILE
  OPEN TransHistFileName FOR RANDOM AS #HFile LEN = TRecSize
  EFile = FREEFILE
  OPEN EmpData2Name FOR RANDOM AS #EFile LEN = EmpRecSize
  
  FOR RecNo = 1 TO NumOfRecs
    UsingThisOne = False
    VCalcAmt# = 0
    LCalcAmt# = 0
    GCalcAmt# = 0
    GPct# = 3
    LPct# = 5
    GET #EFile, IdxBuff(RecNo), Emp2Rec(1)
    IF Emp2Rec(1).EMPTDATE > 0 THEN
      IF LowDate > Emp2Rec(1).EMPTDATE THEN
        GOTO SkipEm
      END IF
    END IF
    
    IF INSTR(Emp2Rec(1).EMPLNAME, "CHEWNING, JR") THEN
      IF INSTR(Emp2Rec(1).EMPFNAME, "WILLIAM B.") THEN
        GOTO SkipEm
      END IF
    END IF
    
    IF Emp2Rec(1).LastTransRec <= 0 THEN
      GOTO SkipEm
    END IF
    
    TransRecNum& = Emp2Rec(1).LastTransRec
    
    DO
      GET #HFile, TransRecNum&, TransHRec(1)
      
      SELECT CASE TransHRec(1).CheckDate
        
      CASE LowDate TO HiDate
        
        IF VCodeNum > 0 THEN
          IF TransHRec(1).DAmt(VCodeNum) <> 0 THEN
            VCalcAmt# = RoundDbl#(VCalcAmt# + TransHRec(1).DAmt(VCodeNum))
            UsingThisOne = True
          END IF
        END IF
        IF LCodeNum > 0 THEN
          IF TransHRec(1).DAmt(LCodeNum) > 0 THEN
            LCalcAmt# = RoundDbl#(LCalcAmt# + TransHRec(1).DAmt(LCodeNum))
            UsingThisOne = True
          END IF
        END IF
        EmpRType$ = UCASE$(LEFT$(LTRIM$(Emp2Rec(1).EMPRETTP), 1))
        IF EmpRType$ = "L" OR EmpRType$ = "G" THEN
          GCalcAmt# = RoundDbl#(GCalcAmt# + TransHRec(1).GrossPay)
          UsingThisOne = True
        END IF
      CASE ELSE
      END SELECT
      
      IF TransHRec(1).PrevTransRec <= 0 THEN
        IF UsingThisOne THEN
          IF EmpRType$ = "L" THEN
            EPct# = LPct#
          ELSE
            EPct# = GPct#
          END IF
          GOSUB PrintThisOne
        END IF
        EXIT DO
      ELSE
        TransRecNum& = CLNG(TransHRec(1).PrevTransRec)
      END IF
      
    LOOP
    
SkipEm:
    LOCATE 13, 1
    PRINT "Processing: "; INT((RecNo / NumOfRecs) * 100);
  NEXT
  
  
  GOSUB DoTrailerRec
  
  CLOSE
  
  '*****************
  LOCATE 13, 1
  PRINT SPACE$(79);
  LOCATE 13, 12
  PRINT "Report Completed."
  PRINT
  PRINT "Press any key to continue."
  dodo = BiosKey
  RUN "cpmenu.exe"
  END
  
  
  
PrintThisOne:

  IF EPct# > 0 OR VCalcAmt# > 0 OR LCalcAmt# > 0 THEN
    EPrinted = EPrinted + 1
    REDIM D401kRec(1) AS DetailRecType
    
    TotalVAmt# = RoundDbl#(TotalVAmt# + VCalcAmt#)
    TotalLAmt# = RoundDbl#(TotalLAmt# + LCalcAmt#)
    TMatchAmt# = RoundDbl#((GCalcAmt# * EPct#) * .01)
    TotalMatchAmt# = RoundDbl#(TotalMatchAmt# + TMatchAmt#)
    
    LSET D401kRec(1).ID = "D"
    LSET D401kRec(1).Batch = "01001"
    LSET D401kRec(1).PCN = QPTrim$(Unit(1).BBTCNTNO)
    LSET D401kRec(1).ProcDate = EMonth$ + "31" + RIGHT$(Year$, 2)
    LSET D401kRec(1).SSN = Emp2Rec(1).EMPSSN
    LSET D401kRec(1).EmpName = QPTrim$(Emp2Rec(1).EMPFNAME) + " " + QPTrim$(Emp2Rec(1).EMPLNAME)
    
    VolDed$ = RSet0$(VCalcAmt#, 7)
    LSET D401kRec(1).EmpVolDed = VolDed$        ''AS STRING * 8
    LoanDed$ = RSet0$(LCalcAmt#, 7)
    LSET D401kRec(1).EmpLoanPay = LoanDed$      ''AS STRING * 8
    ContDed$ = RSet0$(TMatchAmt#, 7)
    LSET D401kRec(1).EmpContAmt = ContDed$      ''AS STRING * 8
    D401kRec(1).CrLf = CrLf$

    PUT #RptFile, , D401kRec(1)
    
  END IF

  RETURN

DoTrailerRec:


  LSET T401kRec(1).ID = "T"

  TVolDed$ = RSet0$(TotalVAmt#, 10)
  LSET T401kRec(1).TotVolDED = TVolDed$       ''AS STRING * 11

  TLoanDed$ = RSet0$(TotalLAmt#, 10)
  LSET T401kRec(1).TotLoanAmt = TLoanDed$       ''AS STRING * 11

  TContDed$ = RSet0$(TotalMatchAmt#, 10)
  LSET T401kRec(1).TotContAmt = TContDed$       ''AS STRING * 11
  LSET T401kRec(1).Filler = ""

  TDetRecs$ = FUsing$(STR$(EPrinted), "###")
  TDetRecs$ = "000000" + QPTrim$(TDetRecs$)
  T401kRec(1).TotDRecs = RIGHT$(TDetRecs$, 6)
  LSET T401kRec(1).CrLf = CrLf$
  
  PUT #RptFile, , T401kRec(1)

  RETURN
  
  
  

ClearArea:
  T$ = SPACE$(60)
  FOR Cnt = 11 TO 18
    LOCATE Cnt, 1: PRINT T$;
  NEXT
  RETURN
  
  
EndTheProg:

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

