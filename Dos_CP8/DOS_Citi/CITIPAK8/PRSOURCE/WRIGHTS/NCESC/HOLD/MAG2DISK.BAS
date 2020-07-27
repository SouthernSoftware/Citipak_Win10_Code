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
  
  ''$INCLUDE: 'MAGMEDIA.Bi'

  '$INCLUDE: 'MAGtype.Bi'
  
  '$INCLUDE: 'DefCnf.bi'                  'Defines TYPE for monitor/color info.
  
  '$INCLUDE: 'FieldInf.bi'                '        "        field information
  '$INCLUDE: 'FormEdit.bi'                '        "        form editing info.
  '$INCLUDE: 'PRFiles.bi'
  '$INCLUDE: 'PREmpRec.bi'
  '$INCLUDE: 'PageInfo.bi'    '        "          Page information
  '$INCLUDE: 'Retire.Bi'
  '$INCLUDE: 'DedCodes.Bi'
  '$INCLUDE: 'ErnCodes.Bi'
  '$INCLUDE: 'Leave.Bi'
  '$INCLUDE: 'StateTax.Bi'
  '$INCLUDE: 'FedTax.Bi'
  '$INCLUDE: 'PREIC.Bi'
  '$INCLUDE: 'PRUNIT.Bi'
  '$INCLUDE: 'PRSYSCTR.Bi'
  '$INCLUDE: 'PRPPDEF.Bi'
  '$INCLUDE: 'PRTRANS.Bi'
  '$INCLUDE: 'PRTIMEIN.Bi'
  '$INCLUDE: 'ScrCalc1.Bi'
  '$INCLUDE: 'PRCHECK.Bi'
  '$INCLUDE: 'PROSCHK.bi'
  '$INCLUDE: 'prif.bi'
  '$INCLUDE: 'prw2.bi'
  
  CONST False = 0, True = NOT False
  
  REDIM SHARED TRec(1)     AS TransRecType
  ' REDIM E2Rec(1)    AS EmpData2Type
  
  DIM SHARED A1Rec AS Rec1A
  DIM SHARED A2Rec AS Rec2A
  DIM SHARED B1Rec AS Rec1B
  DIM SHARED B2Rec AS Rec2B
  DIM SHARED E1Rec AS Rec1E

  DIM SHARED E2Rec AS Rec2E
  
  DIM SHARED W1Rec AS Rec1W
  DIM SHARED W2Rec AS Rec2W
  DIM SHARED W3Rec AS Rec3W
  DIM SHARED I1Rec AS Rec1I
  DIM SHARED I2Rec AS Rec2I
  DIM SHARED T1Rec AS Rec1T
  DIM SHARED T2Rec AS Rec2T
  DIM SHARED F1Rec AS Rec1F

  DIM SHARED Unit AS UnitFileRecType
  
  '  PRINT LEN(A1Rec)
  '  PRINT LEN(A2Rec)
  '  PRINT LEN(B1Rec)
  '  PRINT LEN(B2Rec)
  '  PRINT LEN(E1Rec)
  '  PRINT LEN(E2Rec)
  '  PRINT LEN(W1Rec)
  '  PRINT LEN(W2Rec)
  '  PRINT LEN(W3Rec)
  '  PRINT LEN(I1Rec)
  '  PRINT LEN(I2Rec)
  '  PRINT LEN(T1Rec)
  '  PRINT LEN(T2Rec)
  '  PRINT LEN(F1Rec)
  '  END

  
  TRecSize = LEN(TRec(1))

  'EmpRecSize = LEN(E2Rec(1))
  
  GoodKey$ = "1234567890"
  q$ = CHR$(34)
  
  COLOR 15, 1
top:
  DO
    CLS
    LOCATE 2, 15
    PRINT "Mag Media W-2 Reporting "
    Edit$ = " "
    LOCATE 6, 9
    PRINT "Enter Drive(A-B): ";
    WInput Edit$, "ABab", 6, 29, ExitCode
    IF ExitCode = -27 THEN
      ExitFlag = True
      EXIT DO
    END IF
    OkFlag = True
    Drive$ = Edit$
    
  LOOP UNTIL OkFlag
  
  IF ExitFlag THEN
    GOTO EndTheProg
  END IF
  
'  Drive$ = "C"
  W2Report$ = Drive$ + ":\W2REPORT."
  
  OPEN W2Report$ FOR OUTPUT AS #1
  CLOSE
  
  OPEN W2Report$ FOR RANDOM AS #5 LEN = 128
  
  PRINT
  PRINT
  
  DoABERecs
  
  DoWRecs
  
  CLOSE
EndTheProg:
  END

SUB DoABERecs
  'DIM Unit AS UnitFileRecType
  UnitLen = LEN(Unit)
  OPEN UnitFileName FOR RANDOM AS #1 LEN = UnitLen
  GET #1, , Unit
  CLOSE #1
  
  ID$ = UCASE$(Unit.UFFEDID)
  
  DashPos = INSTR(ID$, "-")
  ID$ = LEFT$(ID$, DashPos - 1) + MID$(ID$, DashPos + 1)
  
  A1Rec.IDENT = "1A"
  A1Rec.PYEAR = "1995"
  A1Rec.EMPRNUM = ID$
  A1Rec.BLANK1 = ""
  A1Rec.FORADDR = ""
  A1Rec.EMPRNAME = UCASE$(Unit.UFEMPR)
  A1Rec.EMPRADDR = UCASE$(Unit.UFADDR1)
  A1Rec.BLANK2 = ""
  
  A2Rec.IDENT = "2A"
  A2Rec.EMPRCITY = UCASE$(Unit.UFCITY)
  A2Rec.EMPRSTATE = UCASE$(Unit.UFSTATE)
  A2Rec.BLANK1 = ""
  A2Rec.EMPRZIPEXT = ""
  A2Rec.EMPRZIP = Unit.UFZIP
  A2Rec.BLANK2 = ""
  
  B1Rec.IDENT = "1B"
  B1Rec.PYEAR = "1995"
  B1Rec.EMPRNUM = ID$
  B1Rec.EMPRCOMP = "GATEWAY"
  B1Rec.BLANK1 = ""
  
  B2Rec.IDENT = "2B"
  B2Rec.BLANK1 = ""
  B2Rec.FORADDR = ""
  B2Rec.ORGNAME = UCASE$(Unit.UFEMPR)
  B2Rec.ORGADDR = UCASE$(Unit.UFADDR1)
  B2Rec.ORGCITY = UCASE$(Unit.UFCITY)
  B2Rec.ORGSTATE = UCASE$(Unit.UFSTATE)
  B2Rec.ORGZIPEXT = ""
  B2Rec.ORGZIP = Unit.UFZIP
  B2Rec.BLANK2 = ""
  
  'TYPE Rec1E
  E1Rec.IDENT = "1E"
  E1Rec.PYEAR = "1995"
  E1Rec.EMPRNUM = ID$
  E1Rec.EMPRSTALOC = ""         '???    Unit.UFState
  E1Rec.EMPRNAME = UCASE$(Unit.UFEMPR)
  E1Rec.EMPRADDR = UCASE$(Unit.UFADDR1)
  E1Rec.FORADDR = ""
  E1Rec.BLANK1 = ""
  
  'TYPE Rec2E
  E2Rec.IDENT = "2E"
  E2Rec.EMPRCITY = UCASE$(Unit.UFCITY)
  E2Rec.EMPRSTATE = UCASE$(Unit.UFSTATE)
  E2Rec.EMPRZIPEXT = ""
  E2Rec.EMPRZIP = Unit.UFZIP
  E2Rec.EMPRNAMECODE = "F"
  E2Rec.EMPRETYPE = "R"
  E2Rec.BLANK1 = ""
  E2Rec.ESTABNUM = ""
  E2Rec.LLINDICAT = ""
  E2Rec.BLANK2 = ""
  E2Rec.OTHEREIN = ""
  E2Rec.BLANK3 = ""
  E2Rec.TAXWITHTPP = "000000000000"
  E2Rec.BLANK4 = ""
  
  PUT #5, , A1Rec
  PUT #5, , A2Rec
  PUT #5, , B1Rec
  PUT #5, , B2Rec
  PUT #5, , E1Rec
  PUT #5, , E2Rec
  
  
END SUB

SUB DoWRecs
  DIM Emp2Rec AS EmpData2Type
  E2Len = LEN(Emp2Rec)
  
  DIM PRW2Rec AS W2FormType
  PRW2Len = LEN(PRW2Rec)
  
  
  OPEN EmpData2Name FOR RANDOM AS #2 LEN = E2Len
  
  OPEN W2InfoFile FOR RANDOM AS #1 LEN = PRW2Len
  
  NumOfRec = LOF(1) \ PRW2Len
  
  W1RecCnt = 0
  ICnt = 0
  
  FOR ThisRec = 1 TO NumOfRec
    
    GET #2, , Emp2Rec
    GET #1, , PRW2Rec
    
    IF PRW2Rec.FEDWAGE = 0 AND PRW2Rec.FEDTAXWH = 0 AND PRW2Rec.SOCWAGE = 0 THEN
      IF PRW2Rec.SOCTAXWH = 0 AND PRW2Rec.MEDWAGES = 0 AND PRW2Rec.MEDTAXWH = 0 THEN
        IF PRW2Rec.SOCTIPS = 0 AND PRW2Rec.ALLOCTIP = 0 AND PRW2Rec.ADVEIC = 0 THEN
          IF PRW2Rec.DEPNDCAR = 0 AND PRW2Rec.NQPLAN = 0 AND PRW2Rec.BOX13AMT = 0 THEN
            IF PRW2Rec.BOX13AM1 = 0 THEN
              GOTO DontMagEm
            END IF
          END IF
        END IF
      END IF
    END IF
    
    'ICnt = ICnt + 1
    W1Rec.IDENT = "1W"
    W1Rec.EmpSSN = Emp2Rec.EmpSSN
    W1Rec.EmpName = QPTrim$(UCASE$(Emp2Rec.EMPFNAME)) + " " + QPTrim$(UCASE$(Emp2Rec.EMPLNAME))
    W1Rec.EMPADDR = QPTrim$(UCASE$(Emp2Rec.EmpAddr1))
    W1Rec.EmpCity = QPTrim$(UCASE$(Emp2Rec.EmpCity))
    W1Rec.EmpState = UCASE$(Emp2Rec.EmpState)
    W1Rec.EMPZIPEXT = RIGHT$(Emp2Rec.EmpZip, 4) + " "
    W1Rec.EmpZip = Emp2Rec.EmpZip
    W1Rec.EMPSTATU = ""
    W1Rec.BLANK1 = ""
    
    'TYPE Rec2W
    W2Rec.IDENT = "2W"
    W2Rec.EMPSSW = RSet0$(PRW2Rec.SOCWAGE, 7)
    W2Rec.BLANK1 = ""
    W2Rec.EMPSST = "0000000"
    W2Rec.BLANK2 = ""
    W2Rec.EMPOTHERW = RSet0$(PRW2Rec.FEDWAGE, 9)
    W2Rec.BLANK3 = ""
    W2Rec.EMPSSTAX = RSet0$(PRW2Rec.SOCTAXWH, 6)
    W2Rec.EMPFEDTAX = RSet0$(PRW2Rec.FEDTAXWH, 9)
    W2Rec.EMPNQP = RSet0$(PRW2Rec.NQPLAN, 9)
    W2Rec.BLANK4 = ""
    W2Rec.EMPNQPO = "000000000"
    W2Rec.CTRLNUM = ""
    W2Rec.EMPRPGTLI = RSet0$(PRW2Rec.BENFBOX1, 7)
    W2Rec.EMPUNCTAX = "0000000"
    W2Rec.EMPEIC = RSet0$(PRW2Rec.ADVEIC, 7)
    W2Rec.EMPATIP = RSet0$(PRW2Rec.ALLOCTIP, 7)
    W2Rec.EMPFRINGE = "000000000"               'RSet0$(PRW2Rec.ALLOCTIP, 7)
    W2Rec.BLANK5 = ""
    W2Rec.PENPLANI = "P"
    W2Rec.BLANK6 = ""
    W2Rec.DEFCOMPI = "D"
    W2Rec.BLANK7 = ""
    W2Rec.DEFCOMPC = RSet0$(PRW2Rec.BOX13AMT, 9)
    W2Rec.BLANK8 = ""
    W2Rec.DEPCBEN = RSet0$(PRW2Rec.DEPNDCAR, 7)
    
    'TYPE Rec3W
    W3Rec.IDENT = "3W"
    W3Rec.MEDWAGES = RSet0$(PRW2Rec.MEDWAGES, 9)
    W3Rec.MEDTAX = RSet0$(PRW2Rec.MEDTAXWH, 7)
    W3Rec.BLANK1 = ""
    
    PUT #5, , W1Rec
    PUT #5, , W2Rec
    PUT #5, , W3Rec
    
    GOSUB AddITotals
    
    ICnt = ICnt + 1
    
    W1RecCnt = W1RecCnt + 1
    
    IF ICnt = 41 THEN
      GOSUB WriteIRec
      I1EMPSSW# = 0
      I1EMPSST# = 0
      I1EMPOTHERW# = 0
      I1EMPSSTAX# = 0
      I1EMPFEDTAX# = 0
      I1EMPRPGTLI# = 0
      I1EMPUNCTAX# = 0
      I1EMPEIC# = 0
      I1EMPATIP# = 0
      I1EMPFRINGE# = 0
      I1DEFCOMPC# = 0
      I2DEPCBEN# = 0
      I2NQPLAN# = 0
      I2NQPLANO# = 0
      I2MEDWAGES# = 0
      I2MEDTAX# = 0
      ICnt = 0
    END IF
    
DontMagEm:
    
    'LOOP UNTIL EOF(1)
  NEXT
  'GOSUB AddITotals
  
  GOSUB WriteIRec
  
  GOSUB WriteTRec
  
  CLOSE
  EXIT SUB
  
AddITotals:
  'sum i rec totals
  'TYPE Rec1I
  
  
  I1EMPSSW# = RoundDbl(I1EMPSSW# + PRW2Rec.SOCWAGE)
  
  I1EMPSST# = RoundDbl(I1EMPSST# + PRW2Rec.SOCTIPS)
  I1EMPOTHERW# = RoundDbl(I1EMPOTHERW# + PRW2Rec.FEDWAGE)
  I1EMPSSTAX# = RoundDbl(I1EMPSSTAX# + PRW2Rec.SOCTAXWH)
  I1EMPFEDTAX# = RoundDbl(I1EMPFEDTAX# + PRW2Rec.FEDTAXWH)
  I1EMPRPGTLI# = RoundDbl(I1EMPRPGTLI# + PRW2Rec.BENFBOX1)
  I1EMPUNCTAX# = 0
  I1EMPEIC# = RoundDbl(I1EMPEIC# + PRW2Rec.ADVEIC)
  I1EMPATIP# = RoundDbl(I1EMPATIP# + PRW2Rec.ALLOCTIP)
  I1EMPFRINGE# = 0
  I1DEFCOMPC# = RoundDbl(I1DEFCOMPC# + PRW2Rec.BOX13AMT)
  
  'TYPE Rec2I
  
  I2DEPCBEN# = RoundDbl(I2DEPCBEN# + PRW2Rec.DEPNDCAR)
  I2NQPLAN# = RoundDbl(I2NQPLAN# + PRW2Rec.NQPLAN)
  I2NQPLANO# = 0
  I2MEDWAGES# = RoundDbl(I2MEDWAGES# + PRW2Rec.MEDWAGES)
  I2MEDTAX# = RoundDbl(I2MEDTAX# + PRW2Rec.MEDTAXWH)
  
  RETURN
  
WriteIRec:
  GOSUB AddGrandTotal
  
  'TYPE Rec1I
  I1Rec.IDENT = "1I"
  I1Rec.EMPSSW = RSet0$(I1EMPSSW#, 10)
  I1Rec.BLANK1 = ""
  I1Rec.EMPSST = RSet0$(I1EMPSST#, 10)
  I1Rec.BLANK2 = ""
  I1Rec.EMPOTHERW = RSet0$(I1EMPOTHERW#, 10)
  I1Rec.BLANK3 = ""
  I1Rec.EMPSSTAX = RSet0$(I1EMPSSTAX#, 10)
  I1Rec.BLANK4 = ""
  I1Rec.EMPFEDTAX = RSet0$(I1EMPFEDTAX#, 10)
  I1Rec.CRTLNUM = ""
  I1Rec.EMPRPGTLI = RSet0$(I1EMPRPGTLI#, 10)
  I1Rec.EMPUNCTAX = "0000000000"
  I1Rec.EMPEIC = RSet0$(I1EMPEIC#, 11)
  I1Rec.EMPATIP = RSet0$(I1EMPATIP#, 10)
  I1Rec.EMPFRINGE = "0000000000"
  I1Rec.BLANK5 = ""
  I1Rec.DEFCOMPC = RSet0$(I1DEFCOMPC#, 10)
  I1Rec.BLANK6 = ""
  
  'TYPE Rec2I
  I2Rec.IDENT = "2I"
  I2Rec.DEPCBEN = RSet0$(I2DEPCBEN#, 10)
  I2Rec.BLANK1 = ""
  I2Rec.NQPLAN = RSet0$(I2NQPLAN#, 10)
  I2Rec.BLANK2 = ""
  I2Rec.NQPLANO = "0000000000"
  I2Rec.BLANK3 = ""
  
  I2Rec.MEDWAGES = RSet0$(I2MEDWAGES#, 11)
  I2Rec.BLANK5 = ""
  I2Rec.MEDTAX = RSet0$(I2MEDTAX#, 10)
  I2Rec.BLANK6 = ""
  
  'PRINT I1Rec.EMPSSW, I1Rec.EMPSST, I1Rec.EMPSSTAX
  PUT #5, , I1Rec
  PUT #5, , I2Rec
  
  
  RETURN
  
AddGrandTotal:
  
  TI1EMPSSW# = RoundDbl(I1EMPSSW# + TI1EMPSSW#)
  TI1EMPSST# = RoundDbl(I1EMPSST# + TI1EMPSST#)
  TI1EMPOTHERW# = RoundDbl(I1EMPOTHERW# + TI1EMPOTHERW#)
  TI1EMPSSTAX# = RoundDbl(I1EMPSSTAX# + TI1EMPSSTAX#)
  TI1EMPFEDTAX# = RoundDbl(I1EMPFEDTAX# + TI1EMPFEDTAX#)
  TI1EMPRPGTLI# = RoundDbl(I1EMPRPGTLI# + TI1EMPRPGTLI#)
  TI1EMPUNCTAX# = 0
  TI1EMPEIC# = RoundDbl(I1EMPEIC# + TI1EMPEIC#)
  TI1EMPATIP# = RoundDbl(I1EMPATIP# + TI1EMPATIP#)
  TI1EMPFRINGE# = 0
  TI1DEFCOMPC# = RoundDbl(I1DEFCOMPC# + TI1DEFCOMPC#)
  
  'TYPE Rec2I
  
  TI2DEPCBEN# = RoundDbl(I2DEPCBEN# + TI2DEPCBEN#)
  TI2NQPLAN# = RoundDbl(I2NQPLAN# + TI2NQPLAN#)
  TI2NQPLANO# = 0
  TI2MEDWAGES# = RoundDbl(I2MEDWAGES# + TI2MEDWAGES#)
  TI2MEDTAX# = RoundDbl(I2MEDTAX# + TI2MEDTAX#)
  
  RETURN
  
WriteTRec:
  
  'TYPE Rec1T
  T1Rec.IDENT = "1T"
  T1Rec.NUMOFEMP = "0000185"    'RSet0$(CDBL(W1RecCnt), 7)
  T1Rec.SSWTOTAL = RSet0$(TI1EMPSSW#, 13)
  T1Rec.BLANK1 = ""
  T1Rec.SSTTOTAL = RSet0$(TI1EMPSST#, 12)
  
  T1Rec.OTHERTOTAL = RSet0$(TI1EMPOTHERW#, 13)
  T1Rec.BLANK2 = ""
  T1Rec.SSTAXTOTAL = RSet0$(TI1EMPSSTAX#, 12)
  T1Rec.BLANK3 = ""
  T1Rec.FEDTAXTOTAL = RSet0$(TI1EMPFEDTAX#, 12)
  T1Rec.ETLITOTAL = RSet0$(TI1EMPRPGTLI#, 12)
  T1Rec.UNTAXTOTAL = "000000000000"
  T1Rec.EICTOTAL = RSet0$(TI1EMPEIC#, 12)
  T1Rec.ATIPTOTAL = RSet0$(TI1EMPATIP#, 12)
  T1Rec.BLANK4 = ""
  
  'PRINT T1Rec.SSWTOTAL, T1Rec.OTHERTOTAL, T1Rec.SSTAXTOTAL
  
  
  'TYPE Rec2T
  T2Rec.IDENT = "2T"
  T2Rec.FRINGETOTAL = "000000000000"
  T2Rec.BLANK1 = ""
  T2Rec.DEFCOMPCTOTAL = RSet0$(TI1DEFCOMPC#, 13)
  T2Rec.BLANK2 = ""
  T2Rec.DEPCBENTOTAL = RSet0$(TI2DEPCBEN#, 12)
  T2Rec.BLANK4 = ""
  T2Rec.NQPTOTAL = RSet0$(TI2NQPLAN#, 13)
  T2Rec.BLANK5 = ""
  T2Rec.NQPOTOTAL = "0000000000000"
  T2Rec.BLANK6 = ""
  T2Rec.MEDWAGETOTAL = RSet0$(TI2MEDWAGES#, 13)
  T2Rec.BLANK7 = ""
  T2Rec.MEDTAXTOTAL = RSet0$(TI2MEDTAX#, 12)
  T2Rec.BLANK8 = ""
  
  
  'TYPE Rec1F
  F1Rec.IDENT = "1F"
  F1Rec.NUMOFEMP = RSet0$(CDBL(W1RecCnt), 7)
  F1Rec.SSWGTOTAL = RSet0$(TI1EMPSSW#, 16)
  F1Rec.BLANK1 = ""
  F1Rec.SSTGTOTAL = RSet0$(TI1EMPSST#, 16)
  F1Rec.BLANK2 = ""
  F1Rec.OTHERWGTOTAL = RSet0$(TI1EMPOTHERW#, 16)
  F1Rec.BLANK3 = ""
  F1Rec.SSTAXGTOTAL = RSet0$(TI1EMPSSTAX#, 16)
  F1Rec.BLANK4 = ""
  F1Rec.FEDTAXGTOTAL = RSet0$(TI1EMPFEDTAX#, 16)
  F1Rec.BLANK5 = ""
  F1Rec.EICGTOTAL = RSet0$(TI1EMPEIC#, 16)
  F1Rec.BLANK6 = ""
  
  PUT #5, , T1Rec
  PUT #5, , T2Rec
  PUT #5, , F1Rec
  
END SUB

FUNCTION LSet0$ (Amt#, StrLen%)
  
  Temp$ = STRING$(StrLen, "0")
  NumStr$ = STR$(Amt#)
  ChrPos = INSTR(NumStr$, ".")
  IF ChrPos THEN
    NumStr$ = LEFT$(NumStr$, ChrPos - 1) + RIGHT$(NumStr$, ChrPos + 1)
  END IF
  NumStr$ = QPTrim$(NumStr$)
  NumLen = LEN(NumStr$)
  StartPos = (StrLen - NumLen) + 1
  MID$(Temp$, StartPos) = NumStr$
  LSet0$ = Temp$
  
  
END FUNCTION

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

