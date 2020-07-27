DEFINT A-Z
DECLARE SUB InitStateTaxDef ()
DECLARE FUNCTION InitializeSystem% ()
DECLARE SUB FCreate (FileName$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPValI% (NumText$)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE FUNCTION Monitor ()
DECLARE SUB WaitForAction ()
DECLARE SUB HideCursor ()
DECLARE SUB TextCursor (Fg, Bg)
DECLARE SUB ShowCursor ()
DECLARE SUB InitMouse (MooseFlag%)
DECLARE FUNCTION Round# (DblNum#)
DECLARE FUNCTION FLInput$ (Handle%, Buff$)
DECLARE FUNCTION WhichError% ()
DECLARE FUNCTION DOSError% ()
DECLARE SUB FClose (Handle%)
DECLARE SUB FOpenS (FileName$, Handle%)
  
'$INCLUDE: 'DefCnf.bi'      'Defines a TYPE for monitor and color info..

'$INCLUDE: 'PRFILES.BI'
'$INCLUDE: 'FormEdit.BI'
'$INCLUDE: 'PRUNIT.BI'
'$INCLUDE: 'LEAVE.BI'

CONST False = 0
CONST True = NOT False

FUNCTION InitializeSystem

  SHARED PrnDef$()

  IF Exist(UnitFileName) THEN
    REDIM UnitFile(1) AS UnitFileRecType
    FGetAH UnitFileName, UnitFile(1), LEN(UnitFile(1)), 1
    IF Round#(UnitFile(1).FileVer) <> 9.1# THEN
      CLS
      LOCATE , , 1
      PRINT "Payroll Has detected that your DATA files need to be "
      PRINT "converted or, have been damaged!!!"
      PRINT "Please Call Southern Software for Assistance."
      PRINT
      PRINT "Press any key to exit.";
      WaitForAction
      PRINT
      InitializeSystem = False
      EXIT FUNCTION
    END IF
'  ELSE
'    CLS
'    LOCATE , , 1
'    PRINT "Payroll was UNABLE to find your DATA files!!!"
'    PRINT
'    PRINT "Please Call Southern Software for Assistance."
'    PRINT
'    PRINT "Press any key to exit.";
'    WaitForAction
'    PRINT
'    InitializeSystem = False
'    EXIT FUNCTION
  END IF
  

'  IF Exist(PRActiveFile) THEN
'    CLS
'    LOCATE , , 1
'    PRINT "Payroll is already running on another system."
'    PRINT
'    PRINT "Press any key to exit.";
'    WaitForAction
'    PRINT
'    InitializeSystem = False
'    EXIT FUNCTION
'  ELSE
'    FCreate PRActiveFile
'  END IF
  
  IF NOT Exist(EmpIdxLName) THEN                '1
    FCreate EmpIdxLName
  END IF
  
  IF NOT Exist(EmpIdxNName) THEN                '2
    FCreate EmpIdxNName
  END IF
  
  IF NOT Exist(EmpData1Name) THEN               '3
    FCreate EmpData1Name
  END IF
  
  IF NOT Exist(EmpData2Name) THEN               '4
    FCreate EmpData2Name
  END IF
  
  IF NOT Exist(EmpData3Name) THEN               '5
    FCreate EmpData3Name
  END IF
  
  IF NOT Exist(EMPNumFileName) THEN             '6
    FCreate EMPNumFileName
  END IF
  
  IF NOT Exist(EMPPinFileName) THEN             '7
    FCreate EMPPinFileName
  END IF
  
  IF NOT Exist(PPDefaultFileName) THEN          '8
    FCreate PPDefaultFileName
  END IF
  
  IF NOT Exist(UnitFileName) THEN               '9
    FCreate UnitFileName
    REDIM UnitFile(1) AS UnitFileRecType
    UnitFile(1).FileVer = 9.1
    FPutAH UnitFileName, UnitFile(1), LEN(UnitFile(1)), 1
    ERASE UnitFile
  END IF
  
  IF NOT Exist(SysFileName) THEN                '10
    FCreate SysFileName
  END IF
  
  IF NOT Exist(DedCodeFileName) THEN            '11
    FCreate DedCodeFileName
  END IF
  
  IF NOT Exist(ErnCodeFileName) THEN            '12
    FCreate ErnCodeFileName
  END IF
  
  IF NOT Exist(FederalTaxFileName) THEN         '13
    FCreate FederalTaxFileName
  END IF
  
  IF NOT Exist(StateTaxFileName) THEN           '14
    FCreate StateTaxFileName
  END IF
  
  IF NOT Exist(StateTaxDefFile) THEN            '14
    FCreate StateTaxDefFile
  END IF
  
  IF NOT Exist(LeaveFileName) THEN              '15
    FCreate LeaveFileName
  END IF

  IF FileSize(LeaveFileName) <= 0 THEN
    REDIM LeaveRec(1) AS LeaveRecType
    LeaveRecLen = LEN(LeaveRec(1))
    FPutAH LeaveFileName, LeaveRec(1), LeaveRecLen, 1
    ERASE LeaveRec
  END IF

  IF NOT Exist(RetireFileName) THEN             '16
    FCreate RetireFileName
  END IF
  
  IF NOT Exist(EICFileName) THEN                '17
    FCreate EICFileName
  END IF
  
  IF NOT Exist(AccrueFileName) THEN             '18
    FCreate AccrueFileName
  END IF
  
  IF NOT Exist(TransWorkFileName) THEN          '19
    FCreate TransWorkFileName
  END IF
  
  IF NOT Exist(TransHistFileName) THEN          '20
    FCreate TransHistFileName
  END IF
  
  IF NOT Exist(ChecksFileName) THEN             '21
    FCreate ChecksFileName
  END IF
  
  IF NOT Exist(W2SetupFile) THEN                '22
    FCreate W2SetupFile
  END IF
  
  IF NOT Exist(W2InfoFile) THEN '23
    FCreate W2InfoFile
  END IF

  IF NOT Exist(MatCodeFileName) THEN
    FCreate MatCodeFileName
  END IF

'07-18-95 *** Added
  IF NOT Exist(PrinterSetUPFile) THEN '24
    REDIM PRNSet(1 TO 16) AS INTEGER
    FOR Cnt = 1 TO 16
      PRNSet(Cnt) = 10
    NEXT
    FCreate PrinterSetUPFile
    FPutAH PrinterSetUPFile, PRNSet(1), 2, 16
    ERASE PRNSet
  END IF

  '--Orig req fields
  'These are the required field numbers in the employee section.
  '2, 3, 4, 20, 21, 22, 23, 24   decimal
  'IF NOT Exist(ReqFldFileName) THEN             '22
  '  REDIM ReqFld(1 TO 8) AS INTEGER
  '  FOR Cnt = 1 TO 3
  '    ReqFld(Cnt) = Cnt + 1
  '  NEXT
  '  FOR Cnt = 4 TO 8
  '    ReqFld(Cnt) = Cnt + 16
  '  NEXT
  '  FCreate ReqFldFileName
  '  FPutAH ReqFldFileName, ReqFld(1), 2, 8
  '  ERASE ReqFld
  'END IF

  'temp for ach
  'These are the required field numbers in the employee section.
  '2, 3, 4, 20, 21, 22, 23, 24   decimal
  IF NOT Exist(ReqFldFileName) THEN             '22
    REDIM ReqFld(1 TO 8) AS INTEGER
    FOR Cnt = 1 TO 3
      ReqFld(Cnt) = Cnt + 1
    NEXT
    FOR Cnt = 0 TO 4
      ReqFld(Cnt + 4) = 20 + Cnt
    NEXT
    FCreate ReqFldFileName
    FPutAH ReqFldFileName, ReqFld(1), 2, 8
    ERASE ReqFld
  END IF
  
  InitStateTaxDef

'Init printer info

  Buffer$ = SPACE$(80)          'buffer
  REDIM PrnDef$(0 TO 4)
  FOpenS PrinterDefFile, Handle                'open state tax def file

  PrnDef$(0) = FLInput$(Handle, Buffer$)        'get the first line of input

  FOR Cnt = 1 TO 4
    PInfo$ = FLInput$(Handle, Buffer$)'get a line of input
    IF DOSError% THEN EXIT FOR                'exit if error (probably end of file)

    PCmd = (INSTR("RESETPRN:10-PITCH:12-PITCH:17-PITCH:", LEFT$(PInfo$, 9)) / 9) + 1
    PrnCmd$ = QPTrim$(MID$(PInfo$, 10))
    CmdLen = LEN(PrnCmd$)

    LastPos = 1

    DO
      ChrPos = INSTR(LastPos, PrnCmd$, ",")
      IF ChrPos > 0 THEN
        Code = QPValI(MID$(PrnCmd$, LastPos, 3))
        PrnDef$(PCmd) = PrnDef$(PCmd) + CHR$(Code)
        LastPos = ChrPos + 1
      ELSE
        Code = QPValI(MID$(PrnCmd$, LastPos))
        IF Code > 0 THEN
          PrnDef$(PCmd) = PrnDef$(PCmd) + CHR$(Code)
        END IF
        EXIT DO
      END IF
    LOOP

  NEXT

  FClose Handle

  Buffer$ = ""                'clear the temp buffer.

'********
  IF INSTR(COMMAND$, "/NM") THEN
  ELSE
    InitMouse Cnf.Mouse
  END IF

  Cnf.MonTyp = Monitor
  
  ClrDisp = Cnf.MonTyp = 3 OR (Cnf.MonTyp >= 5 AND Cnf.MonTyp <= 11)
  IF INSTR(COMMAND$, "/B") THEN ClrDisp = 0
  
  IF ClrDisp THEN
    MonoCode = 0

'Green
   Cnf.PulBar = 3      '23
   Cnf.MenBox = 10     '27
   Cnf.ActivCh = 11   ' 31
   Cnf.InActCh = 3     '31
   Cnf.HiLite = 112    '112
   Cnf.InActHiLt = 64 '112
   Cnf.NonMen = 14    '30
'Blue
'    Cnf.PulBar = 31
'    Cnf.MenBox = 23             ' 49
'    Cnf.ActivCh = 31            '48
'    Cnf.InActCh = 52
'    Cnf.HiLite = 112            '31
'    Cnf.InActHiLt = 64
'    Cnf.NonMen = 30

    Cnf.CsrSize = 7
  ELSE
    MonoCode = 3
    Cnf.PulBar = 112
    Cnf.MenBox = 112
    Cnf.ActivCh = 112
    Cnf.InActCh = 112
    Cnf.HiLite = 15
    Cnf.InActHiLt = 80
    Cnf.NonMen = 7
    Cnf.CsrSize = 12
  END IF
  
  TextCursor 14, 6              'define the Mouse cursor
  '  TextCursor -2, -2             'define the Mouse cursor
  
  'TextCursor 7, 0                       'define the Mouse cursor

  ShowCursor    'Turn it on
  InitializeSystem = True
  
END FUNCTION

SUB InitStateTaxDef

  SHARED TaxText$()

  'Init the state tax info
  REDIM UnitRec(1) AS UnitFileRecType
  FGetAH UnitFileName, UnitRec(1), LEN(UnitRec(1)), 1
  UFSTATE$ = "@" + UnitRec(1).UFSTATE
  ERASE UnitRec

  Buffer$ = SPACE$(80)                          'buffer
  FOpenS StateTaxDefFile, Handle                'open state tax def file
  StateList$ = FLInput$(Handle, Buffer$)        'get the first line of input
  IF INSTR(StateList$, UFSTATE$) THEN
    DO
      StateList$ = FLInput$(Handle, Buffer$)
      IF StateList$ = UFSTATE$ THEN EXIT DO
      IF DOSError THEN
        FClose Handle
        EXIT DO
      END IF
    LOOP
  END IF

  REDIM TaxText$(0 TO 11)
  TaxText$(0) = RIGHT$(UFSTATE$, 2)   '
  FOR Cnt = 1 TO 11
    TaxText$(Cnt) = FLInput$(Handle, Buffer$) 'get a line of input
    WHILE WhichError% = 83             'a "Buffer too small" error means
      TaxText$(Cnt) = TaxText$(Cnt) + FLInput$(Handle, Buffer$)
    WEND                               'we have to go get some more
    IF DOSError% THEN EXIT FOR         'exit if error (probably end of file)
  NEXT
  FClose Handle
  Buffer$ = ""                         'clear the temp buffer.


END SUB

