DEFINT A-Z
DECLARE SUB PCLookUpEmp (EmpNum$, TRecNum%, ERecNum%)
DECLARE FUNCTION LibNumberOfFields (LibName$, FrmName$)
DECLARE SUB PCGetEmpNum (EmpNumber%)
DECLARE SUB PCLoadPayFreqs ()
DECLARE SUB PCSetPeriodDefault ()
DECLARE SUB PostTransactions (EntryType)
DECLARE SUB ManualMenu ()
DECLARE SUB PayMenu (BYVAL MenuNum%, Choice%, NumOfItems%)
DECLARE SUB FPutRTA (Handle%, SEG Source AS ANY, RecNo&, RecSize%)
DECLARE FUNCTION PCDelFromPay% ()
DECLARE SUB PCPickEmpList (TransRecNo%, EmpRecNo%)
DECLARE SUB ButtonPress (ButNo%, Down%, Presses%, x%, Y%)
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE FUNCTION ASCII% (Strng$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE SUB BlockClear ()
DECLARE SUB CursorOff ()
DECLARE SUB DisplayMiscScrn (ScrnName$)
DECLARE SUB FClose (Handle%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FGetT (Handle%, Dest AS ANY, NumBytes%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB HideCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode, Attribute, ErrorCode)
DECLARE SUB PCEnterEdTrans (TransRecNo%, EmpRecNo%)
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
DECLARE SUB RestScrn (ScrnArray%())
DECLARE SUB SaveScrn (ScrnArray%())
DECLARE SUB ShowCursor ()
DECLARE SUB WaitForAction ()
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB VertMenuT (Items() AS ANY, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB MPaintBox (ULRow, ULCol, LRRow, LRCol, Colr)
DECLARE SUB WazzWind (BYVAL TopRow, BYVAL LeftCol, BYVAL BotRow, BYVAL RghtCol, BYVAL FrameColor, BYVAL FrameType, BYVAL Shadow)
DECLARE FUNCTION DOSError ()
DECLARE SUB PCPrintManRegister ()
DECLARE SUB ParseHourly2Trans (TransRec() AS ANY, HourInput() AS ANY)
DECLARE SUB ParseTrans2Hourly (TransRec() AS ANY, HourInput() AS ANY)
DECLARE SUB ParseSalary2Trans (TransRec() AS ANY, SalaryInput() AS ANY)
DECLARE SUB ParseTrans2Salary (TransRec() AS ANY, SalaryInput() AS ANY)
DECLARE SUB ParseScrnCalc2Trans (TransRec() AS ANY, ScrnCalc() AS ANY)
DECLARE SUB ParseTrans2ScrnCalc (TransRec() AS ANY, ScrnCalc() AS ANY)
DECLARE FUNCTION Round# (DblNum#)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE SUB EditForm (Form$(), Fld() AS ANY, Frm AS ANY, Cnf AS ANY, Action)
DECLARE SUB PrintArray (FirstFld%, LastFld%, Form$(), Fld() AS ANY)
DECLARE SUB LibGetFldDef (LibName$, FrmName$, StartEl, Fld() AS ANY, Form$(), ErrCode)

  '$INCLUDE: 'DefCnf.bi'                  'Defines TYPE for monitor/color info.
  '$INCLUDE: 'FieldInf.bi'                '        "        field information
  '$INCLUDE: 'FormEdit.bi'                '        "        form editing info.
  '$INCLUDE: 'PRFiles.bi'
  '$INCLUDE: 'PREmpRec.bi'
  ''$INCLUDE: 'PRSYSCTR.Bi'
  '$INCLUDE: 'PRTRANS.Bi'
  '$INCLUDE: 'PRTIMEIN.Bi'
  '$INCLUDE: 'ScrCalc1.Bi'
  ''$INCLUDE: 'PRCHECK.Bi'
  ''$INCLUDE: 'PRUNIT.Bi'
  ''$INCLUDE: 'PRSYSCTR.Bi'

  TYPE FLen
    v AS STRING * 43
  END TYPE
  
  CONST False = 0, True = NOT False

SUB MakeTransInActive
  
  REDIM TTran(1) AS TransRecType
  TTranSize = LEN(TTran(1))
  NumOfRecs = FileSize(TransWorkFileName) \ TTranSize
  FOpenS TransWorkFileName, TTHandle            'open it
  FOR Cnt = 1 TO NumOfRecs
    RecNo& = Cnt
    FGetRTA TTHandle, TTran(1), RecNo&, TTranSize
    TTran(1).TActive = False
    FPutRTA TTHandle, TTran(1), RecNo&, TTranSize
  NEXT
  
  FClose TTHandle
END SUB

SUB ParseHourly2Trans (TransRec() AS TransRecType, HourInput() AS HourlyInputType)
  
  FOR Cnt = 1 TO 8
    TransRec(1).TDist(Cnt).DAcct = HourInput(1).HDist(Cnt).DAcct
    TransRec(1).TDist(Cnt).DRHrs = HourInput(1).HDist(Cnt).DRHrs
    TransRec(1).TDist(Cnt).DOHrs = HourInput(1).HDist(Cnt).DOHrs
  NEXT
  
  TransRec(1).RegHrsWork = HourInput(1).WORKHRS
  TransRec(1).VacUsed = HourInput(1).VACHRS
  TransRec(1).SickUsed = HourInput(1).SICKHRS
  TransRec(1).HOLHOURS = HourInput(1).HOLHRS
  TransRec(1).CompUsed = HourInput(1).COMPHRS
  
  TransRec(1).PerHours = HourInput(1).PerHRS
  
  TransRec(1).RegHrsPaid = HourInput(1).TOTHRSPD
  TransRec(1).OTHours = HourInput(1).OTWORKED
  TransRec(1).OTHrsPaid = HourInput(1).OTHRSPD
  TransRec(1).OT2COMP = HourInput(1).OT2COMP
  
  TransRec(1).EAmt(1) = HourInput(1).ALTEARN1
  TransRec(1).EAmt(2) = HourInput(1).ALTEARN2
  TransRec(1).EAmt(3) = HourInput(1).ALTEARN3
  TransRec(1).TotAdditEarn = HourInput(1).TOTAERN               'this is sum of (alt earn) fields
  
  TransRec(1).EDist(1).EAcct = HourInput(1).AERNDST1
  TransRec(1).EDist(2).EAcct = HourInput(1).AERNDST2
  TransRec(1).EDist(3).EAcct = HourInput(1).AERNDST3
  TransRec(1).EDist(4).EAcct = HourInput(1).AERNDST4
  TransRec(1).EDist(5).EAcct = HourInput(1).AERNDST5
  TransRec(1).EDist(6).EAcct = HourInput(1).AERNDST6
  
  TransRec(1).EDist(1).EAmt = HourInput(1).AERNAMT1
  TransRec(1).EDist(2).EAmt = HourInput(1).AERNAMT2
  TransRec(1).EDist(3).EAmt = HourInput(1).AERNAMT3
  TransRec(1).EDist(4).EAmt = HourInput(1).AERNAMT4
  TransRec(1).EDist(5).EAmt = HourInput(1).AERNAMT5
  TransRec(1).EDist(6).EAmt = HourInput(1).AERNAMT6
  
  TransRec(1).TaxFring = HourInput(1).TaxFring
END SUB

SUB ParseSalary2Trans (TransRec() AS TransRecType, SalInput() AS SalaryInputType)
  '1
  TransRec(1).PaySFlag = SalInput(1).PAYSAL
  
  TransRec(1).VacUsed = SalInput(1).VACHRS
  TransRec(1).SickUsed = SalInput(1).SICKHRS
  TransRec(1).HOLHOURS = SalInput(1).HOLHRS
  TransRec(1).PerHours = SalInput(1).PerHRS
  
  FOR Cnt = 1 TO 8
    TransRec(1).TDist(Cnt).DAcct = SalInput(1).SDist(Cnt).DAcct
    TransRec(1).TDist(Cnt).DPct = SalInput(1).SDist(Cnt).DPct
  NEXT
  
  TransRec(1).EAmt(1) = SalInput(1).ALTEARN1
  TransRec(1).EAmt(2) = SalInput(1).ALTEARN2
  TransRec(1).EAmt(3) = SalInput(1).ALTEARN3
  
  TransRec(1).EDist(1).EAmt = SalInput(1).AERNAMT1
  TransRec(1).EDist(2).EAmt = SalInput(1).AERNAMT2
  TransRec(1).EDist(3).EAmt = SalInput(1).AERNAMT3
  TransRec(1).EDist(4).EAmt = SalInput(1).AERNAMT4
  TransRec(1).EDist(5).EAmt = SalInput(1).AERNAMT5
  TransRec(1).EDist(6).EAmt = SalInput(1).AERNAMT6
  
  TransRec(1).EDist(1).EAcct = SalInput(1).AERNDST1
  TransRec(1).EDist(2).EAcct = SalInput(1).AERNDST2
  TransRec(1).EDist(3).EAcct = SalInput(1).AERNDST3
  TransRec(1).EDist(4).EAcct = SalInput(1).AERNDST4
  TransRec(1).EDist(5).EAcct = SalInput(1).AERNDST5
  TransRec(1).EDist(6).EAcct = SalInput(1).AERNDST6
  
  TransRec(1).TotAdditEarn = SalInput(1).TOTAERN
  TransRec(1).TaxFring = SalInput(1).TaxFring
  
END SUB

SUB ParseScrnCalc2Trans (TransRec() AS TransRecType, ScrnCalc() AS ScrnCalcType)
  
  TransRec(1).TotRegWage = ScrnCalc(1).REGEARN
  TransRec(1).TotOTWage = ScrnCalc(1).OTEARN
  
  TransRec(1).EAmt(1) = ScrnCalc(1).ALTEARN1
  TransRec(1).EAmt(2) = ScrnCalc(1).ALTEARN2
  TransRec(1).EAmt(3) = ScrnCalc(1).ALTEARN3
  
  TransRec(1).SocTaxAmt = ScrnCalc(1).SOCTAX
  TransRec(1).MedTaxAmt = ScrnCalc(1).MEDTAX
  TransRec(1).FedTaxAmt = ScrnCalc(1).FEDTAX
  TransRec(1).StaTaxAmt = ScrnCalc(1).STATAX
  
  TransRec(1).RetireAmt = ScrnCalc(1).RETIRE
  
  TransRec(1).DAMT(1) = ScrnCalc(1).DED1
  TransRec(1).DAMT(2) = ScrnCalc(1).DED2
  TransRec(1).DAMT(3) = ScrnCalc(1).DED3
  TransRec(1).DAMT(4) = ScrnCalc(1).DED4
  TransRec(1).DAMT(5) = ScrnCalc(1).DED5
  
  TransRec(1).DAMT(6) = ScrnCalc(1).DED6
  TransRec(1).DAMT(7) = ScrnCalc(1).DED7
  TransRec(1).DAMT(8) = ScrnCalc(1).DED8
  TransRec(1).DAMT(9) = ScrnCalc(1).DED9
  TransRec(1).DAMT(10) = ScrnCalc(1).DED10
  TransRec(1).DAMT(11) = ScrnCalc(1).DED11
  TransRec(1).DAMT(12) = ScrnCalc(1).DED12
  
  TransRec(1).GROSSPAY = ScrnCalc(1).GROSSPAY
  
  TransRec(1).TotDedAmt = 0
  FOR Cnt = 1 TO 12
    TransRec(1).TotDedAmt = Round#(TransRec(1).TotDedAmt + TransRec(1).DAMT(Cnt))
  NEXT
  
  'fix from region-d
  TransRec(1).TotDedAmt = Round#(TransRec(1).TotDedAmt + TransRec(1).RetireAmt)
  
  TransRec(1).EICAmt = ScrnCalc(1).EIC
  TransRec(1).NETPAY = ScrnCalc(1).NETPAY
  
  
END SUB

SUB ParseTrans2Hourly (TransRec() AS TransRecType, HourInput() AS HourlyInputType)
  
  FOR Cnt = 1 TO 8
    HourInput(1).HDist(Cnt).DAcct = TransRec(1).TDist(Cnt).DAcct
    HourInput(1).HDist(Cnt).DRHrs = TransRec(1).TDist(Cnt).DRHrs
    HourInput(1).HDist(Cnt).DOHrs = TransRec(1).TDist(Cnt).DOHrs
  NEXT
  
  HourInput(1).WORKHRS = TransRec(1).RegHrsWork
  HourInput(1).VACHRS = TransRec(1).VacUsed
  HourInput(1).SICKHRS = TransRec(1).SickUsed
  HourInput(1).HOLHRS = TransRec(1).HOLHOURS
  HourInput(1).COMPHRS = TransRec(1).CompUsed
  
  HourInput(1).PerHRS = TransRec(1).PerHours
  
  HourInput(1).TOTHRSPD = TransRec(1).RegHrsPaid
  
  
  HourInput(1).OTWORKED = TransRec(1).OTHours
  HourInput(1).OTHRSPD = TransRec(1).OTHrsPaid
  HourInput(1).OT2COMP = TransRec(1).OT2COMP
  
  HourInput(1).ALTEARN1 = TransRec(1).EAmt(1)   '*
  HourInput(1).ALTEARN2 = TransRec(1).EAmt(2)   'actual earning amounts
  HourInput(1).ALTEARN3 = TransRec(1).EAmt(3)   '*
  
  HourInput(1).AERNDST1 = TransRec(1).EDist(1).EAcct            '*
  HourInput(1).AERNDST2 = TransRec(1).EDist(2).EAcct            '
  HourInput(1).AERNDST3 = TransRec(1).EDist(3).EAcct            'Earnings distribution accts.
  HourInput(1).AERNDST4 = TransRec(1).EDist(4).EAcct            '
  HourInput(1).AERNDST5 = TransRec(1).EDist(5).EAcct            '
  HourInput(1).AERNDST6 = TransRec(1).EDist(6).EAcct            '*
  
  HourInput(1).AERNAMT1 = TransRec(1).EDist(1).EAmt             '*
  HourInput(1).AERNAMT2 = TransRec(1).EDist(2).EAmt             '
  HourInput(1).AERNAMT3 = TransRec(1).EDist(3).EAmt             'earnings amounts as distributed
  HourInput(1).AERNAMT4 = TransRec(1).EDist(4).EAmt             'to accts.
  HourInput(1).AERNAMT5 = TransRec(1).EDist(5).EAmt             '
  HourInput(1).AERNAMT6 = TransRec(1).EDist(6).EAmt             '*
  
  HourInput(1).TOTAERN = TransRec(1).TotAdditEarn
  HourInput(1).TaxFring = TransRec(1).TaxFring
  
END SUB

SUB ParseTrans2Salary (TransRec() AS TransRecType, SalInput() AS SalaryInputType)
  
  IF LEN(QPTrim$(TransRec(1).PaySFlag)) = 0 THEN
    SalInput(1).PAYSAL = "Y"
  ELSE
    SalInput(1).PAYSAL = TransRec(1).PaySFlag
  END IF
  
  SalInput(1).VACHRS = TransRec(1).VacUsed
  SalInput(1).SICKHRS = TransRec(1).SickUsed
  SalInput(1).HOLHRS = TransRec(1).HOLHOURS
  SalInput(1).PerHRS = TransRec(1).PerHours
  
  FOR Cnt = 1 TO 8
    SalInput(1).SDist(Cnt).DAcct = TransRec(1).TDist(Cnt).DAcct
    SalInput(1).SDist(Cnt).DPct = TransRec(1).TDist(Cnt).DPct
  NEXT
  
  SalInput(1).ALTEARN1 = TransRec(1).EAmt(1)
  SalInput(1).ALTEARN2 = TransRec(1).EAmt(2)
  SalInput(1).ALTEARN3 = TransRec(1).EAmt(3)
  
  SalInput(1).AERNAMT1 = TransRec(1).EDist(1).EAmt
  SalInput(1).AERNAMT2 = TransRec(1).EDist(2).EAmt
  SalInput(1).AERNAMT3 = TransRec(1).EDist(3).EAmt
  SalInput(1).AERNAMT4 = TransRec(1).EDist(4).EAmt
  SalInput(1).AERNAMT5 = TransRec(1).EDist(5).EAmt
  SalInput(1).AERNAMT6 = TransRec(1).EDist(6).EAmt
  
  SalInput(1).AERNDST1 = TransRec(1).EDist(1).EAcct
  SalInput(1).AERNDST2 = TransRec(1).EDist(2).EAcct
  SalInput(1).AERNDST3 = TransRec(1).EDist(3).EAcct
  SalInput(1).AERNDST4 = TransRec(1).EDist(4).EAcct
  SalInput(1).AERNDST5 = TransRec(1).EDist(5).EAcct
  SalInput(1).AERNDST6 = TransRec(1).EDist(6).EAcct
  
  SalInput(1).TOTAERN = TransRec(1).TotAdditEarn
  SalInput(1).TaxFring = TransRec(1).TaxFring
  
END SUB

SUB ParseTrans2ScrnCalc (TransRec() AS TransRecType, ScrnCalc() AS ScrnCalcType)
  'STOP
  ScrnCalc(1).REGEARN = TransRec(1).TotRegWage
  ScrnCalc(1).OTEARN = TransRec(1).TotOTWage
  
  ScrnCalc(1).ALTEARN1 = TransRec(1).EAmt(1)
  ScrnCalc(1).ALTEARN2 = TransRec(1).EAmt(2)
  ScrnCalc(1).ALTEARN3 = TransRec(1).EAmt(3)
  
  ScrnCalc(1).SOCTAX = TransRec(1).SocTaxAmt
  ScrnCalc(1).MEDTAX = TransRec(1).MedTaxAmt
  ScrnCalc(1).FEDTAX = TransRec(1).FedTaxAmt
  ScrnCalc(1).STATAX = TransRec(1).StaTaxAmt
  ScrnCalc(1).RETIRE = TransRec(1).RetireAmt
  
  ScrnCalc(1).DED1 = TransRec(1).DAMT(1)
  ScrnCalc(1).DED2 = TransRec(1).DAMT(2)
  ScrnCalc(1).DED3 = TransRec(1).DAMT(3)
  ScrnCalc(1).DED4 = TransRec(1).DAMT(4)
  ScrnCalc(1).DED5 = TransRec(1).DAMT(5)
  ScrnCalc(1).DED6 = TransRec(1).DAMT(6)
  ScrnCalc(1).DED7 = TransRec(1).DAMT(7)
  ScrnCalc(1).DED8 = TransRec(1).DAMT(8)
  ScrnCalc(1).DED9 = TransRec(1).DAMT(9)
  ScrnCalc(1).DED10 = TransRec(1).DAMT(10)
  ScrnCalc(1).DED11 = TransRec(1).DAMT(11)
  ScrnCalc(1).DED12 = TransRec(1).DAMT(12)
  
  ScrnCalc(1).GROSSPAY = TransRec(1).GROSSPAY
  
  ScrnCalc(1).TOTDED = Round#(TransRec(1).TotDedAmt + TransRec(1).TotTaxAmt)
  
  ScrnCalc(1).EIC = TransRec(1).EICAmt
  ScrnCalc(1).NETPAY = TransRec(1).NETPAY
  
  
END SUB

FUNCTION PCDelFromPay
  'ok to move
  REDIM TempScrn(0)
  SaveScrn TempScrn()
  DisplayMiscScrn DelFromPay
  DO
    DO
      Ky$ = INKEY$
      ButtonPress 1, 0, MouseButton, XPos, YPos
    LOOP UNTIL LEN(Ky$) OR MouseButton
    IF MouseButton THEN
      MRow = (YPos \ 8) + 1
      MCol = (XPos \ 8)         '14 30
      SELECT CASE MRow
      CASE 16
        SELECT CASE MCol
        CASE 30 TO 40
          PressButton F3Key, 16, 30, 40
        CASE 42 TO 53
          PressButton EscKey, 16, 42, 53
        CASE ELSE
        END SELECT
        Ky$ = INKEY$
      CASE ELSE
      END SELECT
    END IF
    
    SELECT CASE Ky$
    CASE CHR$(0) + "="
      DelPayFlag = -1
      ExitFlag = -1
    CASE CHR$(27)
      DelPayFlag = 0
      ExitFlag = -1
    END SELECT
    
  LOOP UNTIL ExitFlag
  
  PCDelFromPay = DelPayFlag
  
  RestScrn TempScrn()
  ERASE TempScrn
  
END FUNCTION

SUB PCGetEmpNum (EmpNumber)
  
  REDIM FrmGEN(1) AS FormInfo
  
  REDIM TempScrn(0)
  
  FormName$ = "TRANSTRT"
  NumFlds = LibNumberOfFields(CalcQLib, FormName$)
  REDIM FormGEN$(NumFlds, 2)    'DIM the form data array
  REDIM FldGEN(NumFlds) AS FieldInfo            'DIM the field information array
  StartEl = 0   'Load first form at array start
  LibGetFldDef CalcQLib, FormName$, StartEl, FldGEN(), FormGEN$(), ErrCode
  
  '----- Setup TYPE for setting and reading form editing information.
  FrmGEN(1).FldNo = 1           'Start editing on field #1
  FrmGEN(1).InsStat = False     'Set the insert state (-1 = Insert on)
  FrmGEN(1).StartEl = 0         'Set form starting element to 0 and
  
  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1
  BlockClear
  
  DO            'Pole the editing procedure
    'look
    LSET FormGEN$(1, 0) = ""
    LibFile2Scrn CalcQLib, FormName$, MonoCode, -1, ErrCode
    GOSUB DisplayName
    
    DO
      EditForm FormGEN$(), FldGEN(), FrmGEN(1), Cnf, Action
      
      IF FrmGEN(1).Presses AND FrmGEN(1).MRow = 25 THEN
        SELECT CASE FrmGEN(1).MCol
        CASE 3 TO 12            'F1
          PressButton F1Key, 25, 3, 12
          
          
        CASE 47 TO 59           'F7
          PressButton F7KEY, 25, 47, 59
          
        CASE 36 TO 45           'Escape
          PressButton EscKey, 25, 36, 45
        END SELECT
      END IF
      
      SELECT CASE FrmGEN(1).KeyCode
      CASE EnterKey
        WorkNum$ = LTRIM$(RTRIM$(FormGEN$(1, 0)))
        IF LEN(WorkNum$) THEN
          PCLookUpEmp WorkNum$, TransRecNo, EmpRecNo
          IF TransRecNo > 0 THEN
            PCEnterEdTrans TransRecNo, EmpRecNo
            Action = 1
            EXIT DO
          ELSE
            SaveScrn TempScrn()
            DisplayMiscScrn EMPNotFound
            WaitForAction
            LSET FormGEN$(1, 0) = ""
            RestScrn TempScrn()
            PrintArray 1, 1, FormGEN$(), FldGEN()
          END IF
        ELSE
          ExitFlag = True
        END IF
        '04-07-96 Added F7 pick trans list for Mr. Dew
      CASE F7KEY
        PCPickEmpList TransRecNo, EmpRecNo
        Action = 1
        EXIT DO
        '-=-=-=-=
      CASE EscKey
        ExitFlag = True
      END SELECT
      
    LOOP UNTIL ExitFlag = True
    
  LOOP UNTIL ExitFlag = True    '  the Escape key.
  CursorOff
  
  ERASE TempScrn, FormGEN$, FldGEN, FrmGEN
  
  EXIT SUB
  
DisplayName:
  HideCursor
  IF EntryType = Normal THEN
    QPrintRC "Payroll", 1, 3, -1
  ELSE
    QPrintRC "Manual", 1, 4, -1
  END IF
  IF LEN(EmpName$) > 0 THEN QPrintRC EmpName$, 1, 45, -1
  ShowCursor
  RETURN
  
END SUB

SUB PCLookUpEmp (EmpNum$, TRecNum, ERecNum)
  
  REDIM EmpRec1(1)     AS EmpData1Type
  REDIM LUEmpRec2(1)   AS EmpData2Type
  
  TRecNum = 0
  ERecNum = 0
  
  RecLen = LEN(EmpRec1(1))
  RecLen2 = LEN(LUEmpRec2(1))
  
  NumOfRecs = FileSize(EmpData1Name) \ RecLen
  
  FOpenS EmpData1Name, Handle
  FOpenS EmpData2Name, Handle2  'open employee data file
  
  FOR Cnt = 1 TO NumOfRecs
    IF Cnt > NumOfRecs - 3 THEN STOP
    FGetT Handle, EmpRec1(1), RecLen
    TestNum$ = QPTrim$(EmpRec1(1).EmpNo)
    IF TestNum$ = EmpNum$ THEN
      FGetRTA Handle2, LUEmpRec2(1), CLNG(Cnt), RecLen2
      IF LUEmpRec2(1).EMPTDATE <= -32767 AND NOT LUEmpRec2(1).Deleted = True THEN
        TRecNum = Cnt           'EmpRec1(1).TransRecNum
        ERecNum = Cnt           'EmpRec1(1).Data1RecNum
      END IF
      EXIT FOR
    END IF
  NEXT
  
  FClose Handle
  FClose Handle2
  ERASE EmpRec1, LUEmpRec2
  
END SUB

SUB PCPickEmpList (TransRecNo%, EmpRecNo%)
  
  CursorOff
  
  REDIM TempScrn(0)
  SaveScrn TempScrn()
  
  REDIM Emp2DataRec(1)     AS EmpData2Type
  REDIM TempTransRec(1)    AS TransRecType
  REDIM TempData(1)        AS STRING * 43
  
  Data2RecLen = LEN(Emp2DataRec(1))
  TransRecLen = LEN(TempTransRec(1))
  NumOfRecs = FileSize(EmpIdxLName) \ 2
  
  REDIM Items(1 TO NumOfRecs) AS FLen
  
  MatchCnt = 0
  
  FOpenS EmpData2Name, D2Handle
  FOpenS TransWorkFileName, WHandle             'open transaction history file
  
  FOR Cnt = 1 TO NumOfRecs
    FGetRTA D2Handle, Emp2DataRec(1), CLNG(Cnt), Data2RecLen
    IF NOT Emp2DataRec(1).Deleted = True THEN   'put deleted test here
      IF Emp2DataRec(1).EMPTDATE = -32767 THEN
        FGetRTA WHandle, TempTransRec(1), CLNG(Cnt), TransRecLen
        MatchCnt = MatchCnt + 1
        LSET Items(MatchCnt).v = RIGHT$(Emp2DataRec(1).EmpNo, 7)
        MID$(Items(MatchCnt).v, 9, 15) = QPTrim$(Emp2DataRec(1).EMPLNAME)
        MID$(Items(MatchCnt).v, 25, 10) = QPTrim$(Emp2DataRec(1).EMPFNAME)
        IF TempTransRec(1).TActive THEN
          MID$(Items(MatchCnt).v, 39) = "Y"
        ELSE
          MID$(Items(MatchCnt).v, 39) = "N"
        END IF
        MID$(Items(MatchCnt).v, 42) = MKI$(Cnt)
      END IF
    END IF
  NEXT
  
  FClose D2Handle
  FClose WHandle                'open transaction history file
  
  IF MatchCnt > 1 THEN
    REDIM PRESERVE Items(1 TO MatchCnt) AS FLen
    SortT Items(1), MatchCnt, 0, LEN(Items(1).v), 0, 43
    'END IF
    IF LEN(QPTrim$(EmpName$)) > 0 THEN
      SPos = INSTR(EmpName$, " ")
      ENumber$ = QPTrim$(LEFT$(EmpName$, SPos))
      FOR Cnt = 1 TO MatchCnt
        IF INSTR(Items(Cnt).v, ENumber$) THEN
          Choice = Cnt
          EXIT FOR
        END IF
      NEXT
    END IF
  END IF
  
  VAction = 1
  BoxBot = 17
  BoxTop = 7
  MaxLen = 40
  LeftCol = ((80 - MaxLen) \ 2) - 1
  
  DO
    HideCursor
    LOCATE BoxTop, LeftCol
    BlockClear
    LibFile2Scrn CalcQLib, "PICKPREV", MonoCode, -2, ErrCode
    GOSUB DisplayPrevName
    
    QPrintRC " Number   Last Name       First       Trans ", BoxTop - 1, LeftCol, Cnf.ActivCh
    
    MPaintBox BoxTop, LeftCol + MaxLen + 4, BoxTop, LeftCol + MaxLen + 5, 8
    WazzWind 21, LeftCol, 23, LeftCol + MaxLen + 3, Cnf.MenBox, 2, True
    QPrintRC " Use (" + CHR$(24) + "-" + CHR$(25) + ") to select", 22, LeftCol + (MaxLen \ 2) - 8, Cnf.ActivCh
    
    ShowCursor
    '   CursorOff
    
    DO
      
      VertMenuT Items(), Choice, MaxLen, BoxBot, Ky$, VAction, Cnf
      
      SELECT CASE ASCII(Ky$)
      CASE EscKey
        TransRecNo = 0
        Choice = EscKey
        ExitFlag = True
        EXIT DO
      CASE 13
        EmpRecNo = CVI(RIGHT$(Items(Choice).v, 2))
        TransRecNo = EmpRecNo
        PCEnterEdTrans TransRecNo, EmpRecNo
        FOpenS TransWorkFileName, WHandle       'open transaction  file
        FGetRTA WHandle, TempTransRec(1), CLNG(TransRecNo), TransRecLen
        FClose WHandle          'close transaction file
        IF TempTransRec(1).TActive THEN
          MID$(Items(Choice).v, 39) = "Y"
        ELSE
          MID$(Items(Choice).v, 39) = "N"
        END IF
        EXIT DO
      CASE ELSE
      END SELECT
    LOOP
    
  LOOP UNTIL ExitFlag
  
  CursorOff
  
  RestScrn TempScrn()
  
  ERASE Items, TempScrn, TempData, Emp2DataRec, TempTransRec
  EXIT SUB
  
DisplayPrevName:
  HideCursor
  IF EntryType = Normal THEN
    QPrintRC "Payroll", 1, 3, 112
  ELSE
    QPrintRC "Manual", 1, 4, 112
  END IF
  IF LEN(EmpName$) > 0 THEN QPrintRC EmpName$, 1, 45, 112
  ShowCursor
  RETURN
  
END SUB

SUB UpDateTransFile (TransRec() AS TransRecType, TransRecNo)
  FOpenS TransWorkFileName, Handle              'open it
  FPutRTA Handle, TransRec(1), CLNG(TransRecNo), LEN(TransRec(1))
  FClose Handle
END SUB

