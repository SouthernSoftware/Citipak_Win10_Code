DEFINT A-Z

DECLARE SUB BLPrint (LPTNo%, Work$, ErrCount%)
DECLARE SUB RPTSetupPRN (RPTNum%, Handle%)
DECLARE FUNCTION GetStartEmp% (FirstBadChkNum&)
DECLARE SUB RestScrn (ScrnArray%())
DECLARE SUB SaveScrn (ScrnArray%())
DECLARE FUNCTION Date2Num (DateIn$)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION LibNumberOfFields (LibName$, FrmName$)
DECLARE FUNCTION Num2Date$ (Number%)
DECLARE FUNCTION QPStrL$ (LongInt&)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValI% (IntegerNum$)
DECLARE FUNCTION QPValL& (LongNum$)
DECLARE FUNCTION Round# (DblNum#)
DECLARE FUNCTION SpellNumber$ (Number$)

DECLARE SUB BlockClear ()
DECLARE SUB CreateCheckRegister ()
DECLARE SUB CursorOff ()
DECLARE SUB DisplayMiscScrn (ScrnName$)
DECLARE SUB EditForm (Form$(), Fld() AS ANY, Frm AS ANY, Cnf AS ANY, Action)
DECLARE SUB FClose (Handle%)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FPut (Handle%, ToDisk$)
DECLARE SUB FPutRTA (Handle%, SEG Source AS ANY, RecNo&, RecSize%)
DECLARE SUB GetCheckInfo (BYVAL RestartFlag%)
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB LibGetFldDef (LibName$, FrmName$, StartEl, Fld() AS ANY, Form$(), ErrCode)
DECLARE SUB PRCheckMenu ()
DECLARE SUB PayMenu (BYVAL MenuNum%, Choice%, NumOfItems%)
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB PrintCheckRegister ()
DECLARE SUB PrintChecks (FirstEmp%, StartChkNum&, Num2Print, LPTPort)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB WaitForAction ()
DECLARE FUNCTION Monitor ()

'$INCLUDE: 'DefCnf.bi'                  'Defines TYPE for monitor/color info.

'$INCLUDE: 'FieldInf.bi'                '        "        field information
'$INCLUDE: 'FormEdit.bi'                '        "        form editing info.

'$INCLUDE: 'PRFiles.bi'
'$INCLUDE: 'PREmpRec.bi'
'$INCLUDE: 'DedCodes.Bi'
'$INCLUDE: 'ErnCodes.Bi'
'$INCLUDE: 'PRUNIT.Bi'
'$INCLUDE: 'PRSYSCTR.Bi'
'$INCLUDE: 'PRPPDEF.Bi'
'$INCLUDE: 'PRTRANS.Bi'
''$INCLUDE: 'oPRtypes.bi'

''$INCLUDE: 'oPRTRANS.Bi'
'$INCLUDE: 'PRCHECK.Bi'
'$INCLUDE: 'PRdraft.BI'
'$INCLUDE: 'setcnf.BI'

CONST False = 0
CONST True = -1

DIM SHARED CheckNum&, CheckDate, LPTPort, TCheckNum&

 PRCheckMenu

'DIM OCheck   AS PRCheckRecType2
'DIM Check    AS PRCheckRecType
'
'OLen = LEN(OCheck)
'NLen = LEN(Check)
'
'NAME ChecksFileName AS "prdata\ocheck.dat"
'
'OPEN ChecksFileName FOR RANDOM AS #1 LEN = NLen
'NumOfRec = 328 \ 2 'LOF(1) \ OLen
'CLOSE #1
'
'OPEN ChecksFileName FOR RANDOM AS #2 LEN = NLen
'
'FOR Cnt = 1 TO NumOfRec
'''  GET #1, Cnt, OCheck
'  PUT #2, Cnt, Check
'NEXT
'CLOSE

'KILL "prdata\ochecks.dat"

SUB CreateCheckRegister

  Title$ = "Updating Check Register"
  ShowProcessingScrn Title$

  NumOfChecks = 0
  TotChecksAmt# = 0

  REDIM Check(1)     AS PRCheckRecType
  REDIM Unit(1)      AS UnitFileRecType
  REDIM Pg(1)        AS STRING * 3

  IdxRecLen = 2
  CRecLen = LEN(Check(1))

  IdxFileSize& = FileSize(EmpIdxNName)
  NumOfRecs = IdxFileSize& \ IdxRecLen

  REDIM IdxBuff(1 TO NumOfRecs)
  FGetAH EmpIdxNName, IdxBuff(1), IdxRecLen, NumOfRecs

  FGetAH UnitFileName, Unit(1), LEN(Unit(1)), 1

  REDIM ChkNum(1) AS STRING * 10
  REDIM ChkNet(1) AS STRING * 14
  REDIM Dash(1) AS STRING * 70

  Image1$ = "$$,######.##"
  Image2$ = "######"

  Dash(1) = STRING$(70, "-")

  LineCnt = 0
  NextChkNum& = -1
  MaxLines = 55
  Page = 0

  FCreate CheckRegisterRptName
  FOpenS CheckRegisterRptName, RHandle

  RPTSetupPRN 16, RHandle

  FOpenS ChecksFileName, CHandle

  Cnt = 0
  DO
    Cnt = Cnt + 1
    FGetRTA CHandle, Check(1), CLNG(Cnt), CRecLen
    IF DosError THEN EXIT DO
  LOOP UNTIL Check(1).CActive = True OR (Cnt > NumOfRecs)

  CheckDate = Check(1).CheckDate

  GOSUB PrintCheckHeader

  FOR Cnt = 1 TO NumOfRecs
    GRecNum& = CLNG(IdxBuff(Cnt))
    FGetRTA CHandle, Check(1), GRecNum&, CRecLen

    IF Check(1).CActive = True THEN
      ToPrint$ = SPACE$(70)
      IF NextChkNum& = -1 THEN
        NextChkNum& = Check(1).CheckNum
      ELSE
        NextChkNum& = NextChkNum& + 1
      END IF
DoThisOne:
      IF NextChkNum& = Check(1).CheckNum THEN
        NumOfChecks = NumOfChecks + 1
        TotChecksAmt# = Round(TotChecksAmt# + Check(1).NetPay)
        RSET ChkNum(1) = FUsing$(STR$(Check(1).CheckNum), Image2$)
        LSET ToPrint$ = ChkNum(1)
        IF Check(1).DDFlag = True THEN
          MID$(ToPrint$, 15) = Check(1).EmpName + "Draft"
        ELSE
          MID$(ToPrint$, 15) = Check(1).EmpName
        END IF
        MID$(ToPrint$, 59) = FUsing$(STR$(Check(1).NetPay), Image1$)
        FPut RHandle, ToPrint$ + CrLf$
        LineCnt = LineCnt + 1
        IF LineCnt >= MaxLines THEN
          FPut RHandle, FF$
          GOSUB PrintCheckHeader
        END IF
      ELSE
        'LSET ToPrint$ = FUsing$(STR$(NextChkNum&), Image2$) + "<-->" + QPStrL(Check(1).CheckNum - 1) + "              ****REPRINTED****"
        LSET ToPrint$ = "              ****REPRINTED****"
        FPut RHandle, ToPrint$ + CrLf$
        LineCnt = LineCnt + 1
        NextChkNum& = Check(1).CheckNum
        GOTO DoThisOne
      END IF
    END IF
    ShowPctComp Cnt, NumOfRecs

  NEXT
  GOSUB PrintTotChecksLine
  RPTSetupPRN 0, RHandle
  FClose CHandle
  FClose RHandle

  UTemp$ = "": ToPrint$ = ""
  ERASE Check, Unit, Pg, IdxBuff, ChkNum, ChkNet, Dash

EXIT SUB

PrintCheckHeader:
  Page = Page + 1
  RSET Pg(1) = STR$(Page)
  UTemp$ = SPACE$(70)
  LSET UTemp$ = Unit(1).UFEMPR
  MID$(UTemp$, 62) = "Page:" + Pg(1)
  FPut RHandle, UTemp$ + CrLf$
  FPut RHandle, "Payroll Check Register" + CrLf$
  FPut RHandle, "Check Date: " + Num2Date$(CheckDate) + CrLf$
  FPut RHandle, CrLf$
  FPut RHandle, " Check No.    Employee Name                               Check Amount" + CrLf$
  FPut RHandle, Dash(1) + CrLf$
  LineCnt = 6
RETURN

PrintTotChecksLine:
  LSET ChkNet(1) = LTRIM$(FUsing(STR$(TotChecksAmt#), Image1$))
  LSET ChkNum(1) = LTRIM$(FUsing(STR$(NumOfChecks), Image2$))
  FPut RHandle, Dash(1) + CrLf$
  FPut RHandle, "Number of checks printed: " + ChkNum(1) + "Total Amount of Checks: " + ChkNet(1) + CrLf$
  FPut RHandle, FF$
RETURN

END SUB

SUB GetCheckInfo (BYVAL RestartFlag%)

  SHARED PrnDef$()

  REDIM TempScrn(0)
  REDIM Frm(1) AS FormInfo

  FormName$ = "CHKINFO"
  BtnScrn$ = "CHKINFOB"

  IF RestartFlag THEN
    FormName$ = FormName$ + "2"
  ELSE
    FormName$ = FormName$ + "1"
  END IF

  FChkNumFld = 1
  LChkNumFld = 2
  RChkNumFld = 3
  ChkDateFld = 4
  PPortFld = 5

  NumFlds = LibNumberOfFields(CalcQLib, FormName$)
  REDIM Form$(NumFlds, 2)             'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo     'DIM the field information array

  StartEl = 0                            'Load first form at array start
  LibGetFldDef CalcQLib, FormName$, StartEl, Fld(), Form$(), ErrCode

  '----- Setup TYPE for setting and reading form editing information.
  Frm(1).FldNo = 1                 'Start editing on field #1
  Frm(1).InsStat = False           'Set the insert state (-1 = Insert on)
  Frm(1).StartEl = 0               'Set form starting element to 0 and

  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.

  CheckDate = Date2Num(DATE$)
  Action = 1
  'LSET Form$(ChkDateFld, 0) = Num2Date(CheckDate)
  'LSET Form$(PPortFld, 0) = "1"

  FirstTime = True

  DO                 'Pole the editing procedure
    BlockClear
    RetryFlag = False
    LibFile2Scrn CalcQLib, FormName$, MonoCode, -1, ErrCode
    LibFile2Scrn CalcQLib, BtnScrn$, MonoCode, -1, ErrCode
    DO
      EditForm Form$(), Fld(), Frm(1), Cnf, Action
      IF FirstTime THEN
        FirstTime = NOT FirstTime
        LSET Form$(ChkDateFld, 0) = Num2Date(CheckDate)
        LSET Form$(PPortFld, 0) = "1"
        'SaveField 1, Form$(), Fld(), BadFld
        Action = 1
        Frm(1).FldNo = 1
      END IF

      IF Frm(1).Presses AND Frm(1).MRow = 25 THEN
        SELECT CASE Frm(1).MCol
          CASE 2 TO 13          'F1
            PressButton F1Key, 25, 2, 13
        END SELECT
      END IF

      SELECT CASE Frm(1).KeyCode
        CASE F0Key
          IF RestartFlag THEN

            FChkNum& = QPValL(Form$(FChkNumFld, 0))
            LChkNum& = QPValL(Form$(LChkNumFld, 0))

            StartEmp = GetStartEmp(FChkNum&)

            IF (FChkNum& > LChkNum&) OR StartEmp = 0 THEN
              CursorOff
              SaveScrn TempScrn()
              DisplayMiscScrn InvalidCheckNum
              WaitForAction
              FirstTime = True
              OK2Print = False
              RestScrn TempScrn()
            ELSE
              Num2Print = (LChkNum& - FChkNum&) + 1
              OK2Print = True
            END IF
          ELSE
            StartEmp = 1
            OK2Print = True
          END IF

          IF OK2Print THEN
            CheckDate = Date2Num(Form$(ChkDateFld, 0))
            CheckNum& = QPValL(Form$(RChkNumFld, 0))
            LPTPort = QPValI(Form$(PPortFld, 0))
            IF CheckNum& <= 0 THEN
              CursorOff
              SaveScrn TempScrn()
              DisplayMiscScrn InvalidCheckNum
              WaitForAction
              'FirstTime = True
              Action = 2
              RestScrn TempScrn()
              OK2Print = False
            END IF
          END IF

          IF OK2Print THEN
            PrintChecks StartEmp, CheckNum&, Num2Print, LPTPort
            ExitFlag = True
          END IF

        CASE F5KEY
          CursorOff
          BlockClear
          LPTPort = QPValI(Form$(PPortFld, 0))
          REDIM PRNSet(1 TO 16) AS INTEGER
          FGetAH PrinterSetUPFile, PRNSet(1), 2, 16
          RPTPitch = PRNSet(15)
          ERASE PRNSet

          SELECT CASE RPTPitch
            CASE 10
              ToPrint$ = PrnDef$(2)
            CASE 12
              ToPrint$ = PrnDef$(3)
            CASE 17
              ToPrint$ = PrnDef$(4)
            CASE ELSE
          END SELECT

          BLPrint LPTPort, PrnDef$(1), ErrCount
          BLPrint LPTPort, ToPrint$, ErrCount
          PrintRptFile "", CheckAlignMask, LPTPort, RetCode, 5
          BLPrint LPTPort, PrnDef$(1), ErrCount

          CheckNum& = QPValL(Form$(RChkNumFld, 0))
          CheckDate = Date2Num(Form$(ChkDateFld, 0))
                                                                                                               
          IF RetCode = 0 AND CheckNum& > 0 THEN
            CheckNum& = CheckNum& + 1
            LSET Form$(RChkNumFld, 0) = QPStrL$(CheckNum&)
'            CALL SaveField(2, Form$(), Fld(), BadFld)
'            CALL UnPackBuffer(0, 0, Form$(), Fld())
          END IF
          RetryFlag = True
          Action = 2
          'FirstTime = True
        CASE EscKey
          ExitFlag = True
       END SELECT

    LOOP UNTIL ExitFlag OR RetryFlag

  LOOP UNTIL ExitFlag      '

  CursorOff

  ERASE Form$, Fld, Frm, TempScrn

END SUB

'this function returns a pointer into
'the employee index array. whose record we restart
'check printing with
FUNCTION GetStartEmp% (FirstBadChkNum&)

  IF FirstBadChkNum& <= 0 THEN GOTO SkipChkScrh

  REDIM Check(1)        AS PRCheckRecType
  CheckRecLen = LEN(Check(1))
  IdxRecLen = 2
  IdxFileSize& = FileSize(EmpIdxNName)
  NumOfRecs = IdxFileSize& \ IdxRecLen
  REDIM IdxBuff(1 TO NumOfRecs)
  FGetAH EmpIdxNName, IdxBuff(1), IdxRecLen, NumOfRecs

  FOpenS ChecksFileName, CHandle

  FOR Cnt = 1 TO NumOfRecs
    GRecNum& = CLNG(IdxBuff(Cnt))
    FGetRTA CHandle, Check(1), GRecNum&, CheckRecLen
    IF Check(1).CheckNum = FirstBadChkNum& THEN
      GotEmp = True
      EXIT FOR
    END IF
  NEXT

  FClose CHandle

SkipChkScrh:
  IF GotEmp THEN
    GetStartEmp% = Cnt
  ELSE
    GetStartEmp% = 0
  END IF

  ERASE Check, IdxBuff

END FUNCTION

SUB PRCheckMenu

  PPrintChecksMenu = 9
  BlockClear
  Choice = 1

  DO
    PayMenu PPrintChecksMenu, Choice, 3
    SELECT CASE Choice
      CASE 1
        GetCheckInfo False
      CASE 2
        GetCheckInfo True
      CASE 3
        PrintCheckRegister
    END SELECT
  LOOP UNTIL Choice = EscKey

END SUB

SUB PrintCheckRegister
  CursorOff
  CreateCheckRegister
  RptTitle$ = "Payroll Check Register Report"
  PrintRptFile RptTitle$, CheckRegisterRptName, 1, RetCode, 1

END SUB

SUB PrintChecks (FirstEmp, StartChkNum&, Num2Print, LPTPort)

  SHARED PrnDef$()
  REDIM PayFreq$(1 TO 7)

  PayFreq$(1) = "Weekly          "
  PayFreq$(2) = "Bi-Weekly       "
  PayFreq$(3) = "Semi-Monthly    "
  PayFreq$(4) = "Monthly         "
  PayFreq$(5) = "Quarterly       "
  PayFreq$(6) = "Semi-Annually   "
  PayFreq$(7) = "Annually        "

  FF$ = CHR$(12)

  IF FirstEmp = 0 THEN FirstEmp = 1

  Title$ = "Updating Check Information"
  ShowProcessingScrn Title$

  TCheckNum& = CheckNum&

  REDIM Check(1)        AS PRCheckRecType

  REDIM PPDFInfo(1)     AS PRPPDraftInfoType
  
  REDIM DedCodes(1 TO 12) AS DedCodeRecType
  FGetAH DedCodeFileName, DedCodes(1), LEN(DedCodes(1)), 12

  REDIM ErnCodes(1 TO 3) AS ErnCodeRecType
  FGetAH ErnCodeFileName, ErnCodes(1), LEN(ErnCodes(1)), 3

  REDIM TransRec(1)   AS TransRecType
  REDIM EmpRec2(1)    AS EmpData2Type
  REDIM EmpRec3(1)    AS EmpData3Type

  TransRecLen = LEN(TransRec(1))
  Emp2RecLen = LEN(EmpRec2(1))
  Emp3RecLen = LEN(EmpRec3(1))
  CRecLen = LEN(Check(1))
  CheckRecLen = LEN(Check(1))
  PPDFLen = LEN(PPDFInfo(1))

  IdxRecLen = 2
  IdxFileSize& = FileSize(EmpIdxNName)
  NumOfRecs = IdxFileSize& \ IdxRecLen
  REDIM IdxBuff(1 TO NumOfRecs)
  FGetAH EmpIdxNName, IdxBuff(1), IdxRecLen, NumOfRecs

  FOpenS ChecksFileName, CHandle
  FOpenS EmpData2Name, EHandle2   'open it
  FOpenS EmpData3Name, EHandle3   'open it
  FOpenS TransWorkFileName, THandle   'open it

  FCreate CheckPrintFile
  'Employe DD Payperiod Info
  FCreate PPDraftInfoFileName

'Set printer to corect pitch
  FOpenS CheckPrintFile, RHandle
  RPTSetupPRN 15, RHandle
  FClose RHandle
  
  '1 = FREEFILE
  'this must be opened as #1
  OPEN CheckPrintFile FOR APPEND AS #1

  PPDFFile = FREEFILE
  OPEN PPDraftInfoFileName FOR RANDOM AS PPDFFile LEN = PPDFLen

  FOR Cnt = FirstEmp TO NumOfRecs
    IF Num2Print > 0 THEN
      IF DoneCnt = Num2Print THEN
        GOTO DoneEM
      END IF
    END IF

    GRecNum& = CLNG(IdxBuff(Cnt))
    FGetRTA THandle, TransRec(1), GRecNum&, TransRecLen

    IF TransRec(1).TActive = True AND TransRec(1).NetPay > 0 THEN
      FGetRTA EHandle2, EmpRec2(1), GRecNum&, Emp2RecLen
      FGetRTA EHandle3, EmpRec3(1), GRecNum&, Emp3RecLen

      IF (EmpRec2(1).DRAFTCOD = "C" OR EmpRec2(1).DRAFTCOD = "S") THEN 'AND EmpRec2(1).PreNoted = "Y" THEN
        Check(1).DDFlag = True
        PPDFInfo(1).EmpRec = GRecNum&
        PPDFInfo(1).DraftDate = CheckDate
        PPDFInfo(1).NetPay = TransRec(1).NetPay
        PUT #PPDFFile, , PPDFInfo(1)
      ELSE
        Check(1).DDFlag = False
      END IF

      Check(1).CActive = True
      Check(1).EmpName = QPTrim$(EmpRec2(1).EMPFNAME) + " " + QPTrim$(EmpRec2(1).EMPLNAME)
      Check(1).EMPNO = EmpRec2(1).EMPNO
      Check(1).EmpSSN = LEFT$(EmpRec2(1).EmpSSN, 3) + "-" + MID$(EmpRec2(1).EmpSSN, 4, 2) + "-" + MID$(EmpRec2(1).EmpSSN, 6, 4)
      Check(1).EmpAddr1 = EmpRec2(1).EmpAddr1
      Check(1).EmpCity = EmpRec2(1).EmpCity
      Check(1).EmpState = EmpRec2(1).EmpState
      Check(1).EmpZip = EmpRec2(1).EmpZip

      Check(1).PayEndDate = TransRec(1).PayPdEnd

      Check(1).CheckDate = CheckDate
      Check(1).BaseRate = TransRec(1).BaseRate
      Check(1).GrossPay = TransRec(1).GrossPay
      Check(1).FedTaxAmt = TransRec(1).FedTaxAmt
      Check(1).StaTaxAmt = TransRec(1).StaTaxAmt
      Check(1).MedTaxAmt = TransRec(1).MedTaxAmt
      Check(1).SocTaxAmt = TransRec(1).SocTaxAmt
      Check(1).TotDedAmt = TransRec(1).TotDedAmt
      Check(1).RetireAmt = TransRec(1).RetireAmt
      Check(1).TaxFring = TransRec(1).TaxFring
      Check(1).EICAmt = TransRec(1).EICAmt

      FOR Cnt2 = 1 TO 3
        Check(1).AEarn(Cnt2).DAmt = TransRec(1).EAmt(Cnt2)
        Check(1).AEarn(Cnt2).DCode = ErnCodes(Cnt2).ERNCODE1
      NEXT

      Check(1).TotAdditEarn = TransRec(1).TotAdditEarn
      Check(1).NetPay = TransRec(1).NetPay

'02-16-95 ?????????
      Check(1).TotOTWage = TransRec(1).TotOTWage
'was NEVER printed on the CHECK correctly!!!!!!!!

      Check(1).TotRegWage = TransRec(1).TotRegWage

      Check(1).RegHrsWork = TransRec(1).RegHrsWork
      Check(1).RegHrsPaid = TransRec(1).RegHrsPaid
      Check(1).OTHrsPaid = TransRec(1).OTHrsPaid

      Check(1).YTDGrossPay = Round#(EmpRec3(1).YTDGrossPay + TransRec(1).GrossPay)
      Check(1).YTDFederal = Round#(EmpRec3(1).YTDFederal + TransRec(1).FedTaxAmt)
      Check(1).YTDState = Round#(EmpRec3(1).YTDState + TransRec(1).StaTaxAmt)
      Check(1).YTDSocial = Round#(EmpRec3(1).YTDSocial + TransRec(1).SocTaxAmt)
      Check(1).YTDMedicare = Round#(EmpRec3(1).YTDMedicare + TransRec(1).MedTaxAmt)
      Check(1).YTDTotDed = Round#(EmpRec3(1).YTDDAmtT + EmpRec3(1).YTDRetire + TransRec(1).TotDedAmt)
      Check(1).YTDNetPay = Round#(EmpRec3(1).YTDNet + TransRec(1).NetPay)
      Check(1).YTDRetire = Round#(EmpRec3(1).YTDRetire + TransRec(1).RetireAmt)

      Check(1).VactBal = Round(EmpRec2(1).EMPVBAL - TransRec(1).VacUsed)
      IF Check(1).VactBal < -20 THEN Check(1).VactBal = 0

      Check(1).SickBal = Round(EmpRec2(1).EMPSLBAL - TransRec(1).SickUsed)
      IF Check(1).SickBal < -20 THEN Check(1).SickBal = 0

      Check(1).CompEarn = Round(EmpRec2(1).EMPCTE + TransRec(1).OT2Comp)
      Check(1).CompBal = Round#(Check(1).CompEarn - (EmpRec2(1).EMPCTUSE + TransRec(1).CompUsed))
      IF Check(1).CompBal < -20 THEN Check(1).CompBal = 0

      Check(1).VacUsed = TransRec(1).VacUsed
      Check(1).SickUsed = TransRec(1).SickUsed
      Check(1).CompUsed = TransRec(1).CompUsed
      Check(1).HolUsed = TransRec(1).HOLHOURS

      'Check(1).PERUSED = TransRec(1).PerHours

      HolBal# = Round#(EmpRec2(1).HOLERN - (EmpRec2(1).HolUsed + TransRec(1).HOLHOURS))
      'PERBAL# = Round#(EmpRec2(1).PERERN - (EmpRec2(1).PERUSED + TransRec(1).PerHours))

      FOR Cnt2 = 1 TO 12
        Check(1).CDED(Cnt2).DCode = DedCodes(Cnt2).DCDESC1
        Check(1).CDED(Cnt2).DAmt = TransRec(1).DAmt(Cnt2)
        Check(1).CDED(Cnt2).YTDDAmt = Round#(EmpRec3(1).YTDDAmt(Cnt2) + TransRec(1).DAmt(Cnt2))
      NEXT

      Check(1).CheckDate = CheckDate
      Check(1).CheckNum = TCheckNum&
      'TCheckNum& = TCheckNum& + 1

'*NEW 01-16-96  Added calc here for (Vac, Sick, Hol, and Comp pay)

      PayType$ = LEFT$(UCASE$(QPTrim$(EmpRec2(1).EMPPTYPE)), 1)

      SELECT CASE PayType$
        CASE "S"
          FOR FreqCnt = 1 TO 7
            IF EmpRec2(1).EMPPFREQ = PayFreq$(FreqCnt) THEN
              EXIT FOR
            END IF
          NEXT
          SELECT CASE FreqCnt
            CASE 1
              PayFreq = 52
            CASE 2
              PayFreq = 26
            CASE 3
              PayFreq = 24
            CASE 4
              PayFreq = 12
            CASE 5
              PayFreq = 4
            CASE 6
              PayFreq = 2
            CASE 7
              PayFreq = 1
          END SELECT
          YGross# = Round(EmpRec2(1).EMPPRATE * PayFreq)
          HRate# = Round(YGross# / 2080)
          WorkPay# = Check(1).BaseRate
        CASE "H"
          HRate# = EmpRec2(1).EMPPRATE
          WorkPay# = Round#(Check(1).BaseRate * Check(1).RegHrsWork)
      END SELECT

      IF TransRec(1).VacUsed > 0 THEN
        VacPay# = Round(HRate# * TransRec(1).VacUsed)
      ELSE
        VacPay# = 0
      END IF

      IF TransRec(1).SickUsed > 0 THEN
        SickPay# = Round(HRate# * TransRec(1).SickUsed)
      ELSE
        SickPay# = 0
      END IF

      IF TransRec(1).CompUsed > 0 THEN
        CompPay# = Round(HRate# * TransRec(1).CompUsed)
      ELSE
        CompPay# = 0
      END IF

      IF TransRec(1).HOLHOURS > 0 THEN
        HolPay# = Round(HRate# * TransRec(1).HOLHOURS)
      ELSE
        HolPay# = 0
      END IF

      'IF TransRec(1).PerHours > 0 THEN
      '  PerPay# = Round(HRate# * TransRec(1).PerHours)
      'ELSE
      '  PerPay# = 0
      'END IF

'-------
''$INCLUDE: 'GROVER.BI'
''$INCLUDE: 'GROVER5.BI'
''$INCLUDE: 'PR9007.BI'
''$INCLUDE: 'PR9013.BI'
''$INCLUDE: 'PRTROUT.BI'
''$INCLUDE: 'APPALACH.BI'
''$INCLUDE: 'BALDHEAD.BI'
''$INCLUDE: 'BEAVER.BI'
''$INCLUDE: 'BEECH.BI'             'This is the New Standard
''$INCLUDE: 'BURNS.BI'
''$INCLUDE: 'CARMEN.BI'
''$INCLUDE: 'CARTHAGE.BI'
''$INCLUDE: 'CLYDE.BI'
''$INCLUDE: 'DOBBINS.BI'
''$INCLUDE: 'ELOREE.BI'
''$INCLUDE: 'FAIRMONT.BI'
''$INCLUDE: 'GRAYSON.BI'
''$INCLUDE: 'HARRIS.BI'
''$INCLUDE: 'HOLLY.BI'
''$INCLUDE: 'HOLNWAL2.BI'
''$INCLUDE: 'LOUISA.BI'
''$INCLUDE: 'MCCORE.BI'
''$INCLUDE: 'MTGILEAD.BI'
''$INCLUDE: 'NEWLAND.BI'
''$INCLUDE: 'NORLINA.BI'
''$INCLUDE: 'PONDCREK.BI'
''$INCLUDE: 'REGD.BI'
''$INCLUDE: 'RICHLAND.BI'
''$INCLUDE: 'ROLESCHK.BI'
'$INCLUDE: 'STDCHK.BI'
''$INCLUDE: 'oldSYLVA.BI'
''$INCLUDE: 'SYLVA2.BI'
''$INCLUDE: 'TROUTMAN.BI'
''$INCLUDE: 'TROYCHK.BI'
''$INCLUDE: 'WHITAKER.BI'
''$INCLUDE: 'WHITELAK.BI'
''$INCLUDE: 'PRWADES.BI'    'wadesboro
''$INCLUDE: 'PRLARNCE.BI'    '
''$INCLUDE: 'PRGILES.BI'    '
''$INCLUDE: 'PRGILBRT.BI'   'Gilbert
''$INCLUDE: 'PRWRIGHT.BI'    'Wrightsville Beech
'--
      TCheckNum& = TCheckNum& + 1
      IF Num2Print > 0 THEN
        DoneCnt = DoneCnt + 1
      END IF

    ELSE
      Check(1).CActive = False
    END IF

    FPutRTA CHandle, Check(1), GRecNum&, CheckRecLen

DoneEM:
    ShowPctComp Cnt, NumOfRecs

  NEXT
'**UNREM
  PRINT #1, PrnDef$(1);
  CLOSE #1

  CLOSE PPDFFile

  FClose EHandle2
  FClose EHandle3
  FClose THandle
  FClose CHandle

  ERASE DedCodes, ErnCodes, TransRec, EmpRec2, EmpRec3, Check, IdxBuff

  PrintRptFile "", CheckPrintFile, LPTPort, RetCode, 1

END SUB

