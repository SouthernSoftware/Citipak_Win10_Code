DEFINT A-Z
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

DECLARE SUB ParseHourly2Trans (TransRec() AS ANY, HourInput() AS ANY)
DECLARE SUB ParseTrans2Hourly (TransRec() AS ANY, HourInput() AS ANY)

DECLARE SUB ParseSalary2Trans (TransRec() AS ANY, SalaryInput() AS ANY)
DECLARE SUB ParseTrans2Salary (TransRec() AS ANY, SalaryInput() AS ANY)

DECLARE SUB ParseScrnCalc2Trans (TransRec() AS ANY, ScrnCalc() AS ANY)
DECLARE SUB ParseTrans2ScrnCalc (TransRec() AS ANY, ScrnCalc() AS ANY)

DECLARE FUNCTION Round# (DblNum#)
DECLARE FUNCTION QPTrim$ (Text$)

'$INCLUDE: 'DefCnf.bi'                  'Defines TYPE for monitor/color info.

'$INCLUDE: 'FieldInf.bi'                '        "        field information
'$INCLUDE: 'FormEdit.bi'                '        "        form editing info.
'$INCLUDE: 'PRFiles.bi'
'$INCLUDE: 'PREmpRec.bi'
'$INCLUDE: 'PRSYSCTR.Bi'
'$INCLUDE: 'PRTRANS.Bi'
'$INCLUDE: 'PRTIMEIN.Bi'
'$INCLUDE: 'ScrCalc1.Bi'

TYPE FLen
   v AS STRING * 43
END TYPE

CONST False = 0, True = NOT False

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

  TransRec(1).RegHrsPaid = HourInput(1).TOTHRSPD

  TransRec(1).OTHours = HourInput(1).OTWORKED
  TransRec(1).OTHrsPaid = HourInput(1).OTHRSPD
  TransRec(1).OT2COMP = HourInput(1).OT2COMP


  TransRec(1).EAmt(1) = HourInput(1).ALTEARN1
  TransRec(1).EAmt(2) = HourInput(1).ALTEARN2
  TransRec(1).EAmt(3) = HourInput(1).ALTEARN3
  TransRec(1).TotAdditEarn = HourInput(1).TOTAERN  'this is sum of (alt earn) fields

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

 TransRec(1).DAmt(1) = ScrnCalc(1).DED1
 TransRec(1).DAmt(2) = ScrnCalc(1).DED2
 TransRec(1).DAmt(3) = ScrnCalc(1).DED3
 TransRec(1).DAmt(4) = ScrnCalc(1).DED4
 TransRec(1).DAmt(5) = ScrnCalc(1).DED5

 TransRec(1).DAmt(6) = ScrnCalc(1).DED6
 TransRec(1).DAmt(7) = ScrnCalc(1).DED7
 TransRec(1).DAmt(8) = ScrnCalc(1).DED8
 TransRec(1).DAmt(9) = ScrnCalc(1).DED9
 TransRec(1).DAmt(10) = ScrnCalc(1).DED10
 TransRec(1).DAmt(11) = ScrnCalc(1).DED11
 TransRec(1).DAmt(12) = ScrnCalc(1).DED12

 TransRec(1).GROSSPAY = ScrnCalc(1).GROSSPAY

 TransRec(1).TotDedAmt = 0
 FOR Cnt = 1 TO 12
   TransRec(1).TotDedAmt = Round#(TransRec(1).TotDedAmt + TransRec(1).DAmt(Cnt))
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
   HourInput(1).TOTHRSPD = TransRec(1).RegHrsPaid


   HourInput(1).OTWORKED = TransRec(1).OTHours
   HourInput(1).OTHRSPD = TransRec(1).OTHrsPaid
   HourInput(1).OT2COMP = TransRec(1).OT2COMP

   HourInput(1).ALTEARN1 = TransRec(1).EAmt(1) '*
   HourInput(1).ALTEARN2 = TransRec(1).EAmt(2) 'actual earning amounts
   HourInput(1).ALTEARN3 = TransRec(1).EAmt(3) '*

    HourInput(1).AERNDST1 = TransRec(1).EDist(1).EAcct '*
    HourInput(1).AERNDST2 = TransRec(1).EDist(2).EAcct '
    HourInput(1).AERNDST3 = TransRec(1).EDist(3).EAcct 'Earnings distribution accts.
    HourInput(1).AERNDST4 = TransRec(1).EDist(4).EAcct '
    HourInput(1).AERNDST5 = TransRec(1).EDist(5).EAcct '
    HourInput(1).AERNDST6 = TransRec(1).EDist(6).EAcct '*

    HourInput(1).AERNAMT1 = TransRec(1).EDist(1).EAmt '*
    HourInput(1).AERNAMT2 = TransRec(1).EDist(2).EAmt '
    HourInput(1).AERNAMT3 = TransRec(1).EDist(3).EAmt 'earnings amounts as distributed
    HourInput(1).AERNAMT4 = TransRec(1).EDist(4).EAmt 'to accts.
    HourInput(1).AERNAMT5 = TransRec(1).EDist(5).EAmt '
    HourInput(1).AERNAMT6 = TransRec(1).EDist(6).EAmt '*

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

ScrnCalc(1).DED1 = TransRec(1).DAmt(1)
ScrnCalc(1).DED2 = TransRec(1).DAmt(2)
ScrnCalc(1).DED3 = TransRec(1).DAmt(3)
ScrnCalc(1).DED4 = TransRec(1).DAmt(4)
ScrnCalc(1).DED5 = TransRec(1).DAmt(5)
ScrnCalc(1).DED6 = TransRec(1).DAmt(6)
ScrnCalc(1).DED7 = TransRec(1).DAmt(7)
ScrnCalc(1).DED8 = TransRec(1).DAmt(8)
ScrnCalc(1).DED9 = TransRec(1).DAmt(9)
ScrnCalc(1).DED10 = TransRec(1).DAmt(10)
ScrnCalc(1).DED11 = TransRec(1).DAmt(11)
ScrnCalc(1).DED12 = TransRec(1).DAmt(12)

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
      MCol = (XPos \ 8)        '14 30
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
    IF NOT Emp2DataRec(1).Deleted = True THEN        'put deleted test here
      IF Emp2DataRec(1).EMPTDATE = -32767 THEN
        FGetRTA WHandle, TempTransRec(1), CLNG(Cnt), TransRecLen
        MatchCnt = MatchCnt + 1
        LSET Items(MatchCnt).v = RIGHT$(Emp2DataRec(1).EMPNO, 6)
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
  FClose WHandle             'open transaction history file

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
        FOpenS TransWorkFileName, WHandle    'open transaction  file
        FGetRTA WHandle, TempTransRec(1), CLNG(TransRecNo), TransRecLen
        FClose WHandle             'close transaction file
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

