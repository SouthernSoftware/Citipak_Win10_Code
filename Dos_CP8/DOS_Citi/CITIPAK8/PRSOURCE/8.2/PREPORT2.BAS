DEFINT A-Z
DECLARE SUB CheckIssuedRpt ()
DECLARE SUB RePrintRpts ()
DECLARE SUB VARetirementRpt ()
DECLARE SUB RPTSetupPRN (RPTNum%, FHandle%)
DECLARE SUB BenefitReport ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB SEPPContRpt ()
DECLARE SUB YTDFundReport ()

DECLARE FUNCTION NotFromPRMain% ()
DECLARE SUB InitMouse (MooseFlag%)
DECLARE FUNCTION Date2Num (DateIn$)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION LibNumberOfFields (LibName$, FrmName$)
DECLARE FUNCTION MonthName$ (MonthNum%)
DECLARE FUNCTION Num2Date$ (Number%)
DECLARE FUNCTION QPStrI$ (Num%)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValI% (Num$)
DECLARE FUNCTION QPValL& (Num$)
DECLARE FUNCTION Round# (DblNum#)
DECLARE SUB BlockClear ()
DECLARE SUB CursorOff ()
DECLARE SUB DisplayMiscScrn (ScrnName$)
DECLARE SUB ESCQuartlyWageRpt ()
DECLARE SUB EditForm (Form$(), Fld() AS ANY, Frm AS ANY, Cnf AS ANY, Action)

DECLARE SUB UnPackBufer (FirstFld, LastFld, Form$(), Fld() AS ANY)

DECLARE SUB EMPrintTermEmpList ()
DECLARE SUB EMPrintEmpData ()
DECLARE SUB EMPrintEmpList ()
DECLARE SUB EmpEarnHistRpt ()
DECLARE SUB FClose (Handle%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FPut (Handle%, ToDisk$)
DECLARE SUB FPutA (Handle%, SEG Element AS ANY, NumBytes AS ANY)
DECLARE SUB GrossWageReport ()
DECLARE SUB LibGetFldDef (LibName$, FrmName$, StartEl, Fld() AS ANY, Form$(), ErrCode)
DECLARE SUB PayDedReport ()
DECLARE SUB PayMenu (BYVAL MenuNum%, Choice%, NumOfItems%)
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB PrintArray (FirstFld, LastFld, Form$(), Fld() AS ANY)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
DECLARE SUB RetirementRpt ()
DECLARE SUB RestScrn (ScrnArray%())
DECLARE SUB SaveScrn (ScrnArray%())
DECLARE SUB ShowCursor ()
DECLARE SUB HideCursor ()
DECLARE SUB TextCursor (Fg, Bg)
DECLARE SUB SupRetReport ()
DECLARE SUB WaitForAction ()
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION WEnvTest% ()
DECLARE SUB BCopy (FromSeg%, FromAdr%, ToSeg%, ToAdr%, NumBytes%, Dir%)

DECLARE FUNCTION IMinI% (SEG Element%, NumEls%)
DECLARE FUNCTION IMaxI% (SEG Element%, NumEls%)

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
''$INCLUDE: 'oPRtypes.Bi'
'$INCLUDE: 'PRRpts.BI'
'$INCLUDE: 'PRTIMEIN.BI'

CONST False = 0
CONST True = NOT False

SUB EmpEarnHistRpt

  'SHARED PrnDef$()

  REDIM TempScrn(1)
  REDIM Emp2Rec(1)           AS EmpData2Type
  REDIM Emp1Rec(1)           AS EmpData1Type
  REDIM TransRec(1)          AS TransRecType
  REDIM Unit(1)              AS UnitFileRecType
  REDIM EMPHIST(1 TO 3)      AS EmpHistoryRptType
  REDIM DedCodes(1 TO 12)    AS DedCodeRecType
  REDIM ErnCodes(1 TO 3)     AS ErnCodeRecType
  REDIM EmpHistRpt(1)        AS EmpHistFormType

  REDIM Dash(1)              AS STRING * 132
  REDIM TotDeds(1 TO 12)     AS DOUBLE
  REDIM TotErns(1 TO 3)      AS DOUBLE
  REDIM ESubDeds(1 TO 12)    AS DOUBLE
  REDIM ESubErns(1 TO 3)     AS DOUBLE
  REDIM EmpNo(1)             AS STRING * 14
  REDIM RErnP(1)             AS STRING * 11
  REDIM EICP(1)              AS STRING * 11
  REDIM GPayP(1)             AS STRING * 11
  REDIM SSTaxP(1)            AS STRING * 11
  REDIM MTaxP(1)             AS STRING * 11
  REDIM FTaxP(1)             AS STRING * 11
  REDIM STaxP(1)             AS STRING * 11
  REDIM RetirP(1)            AS STRING * 11
  REDIM NetPayP(1)           AS STRING * 11
  REDIM OErnP(1)             AS STRING * 11
  REDIM Ded(1)               AS STRING * 11
  REDIM Ern(1)               AS STRING * 11
  REDIM Pg(1)                AS STRING * 5
  REDIM Fill11(1)            AS STRING * 11
  REDIM RHrs(1)              AS STRING * 11
  REDIM VHrs(1)              AS STRING * 11
  REDIM SHrs(1)              AS STRING * 11
  REDIM HHrs(1)              AS STRING * 11
  REDIM CHrs(1)              AS STRING * 11
  REDIM THrs(1)              AS STRING * 11
  REDIM OTHrs(1)             AS STRING * 11
  REDIM OTPaid(1)            AS STRING * 11
  REDIM EICP(1)              AS STRING * 11
  REDIM RErnP(1)             AS STRING * 11
  REDIM EChkDate(1)          AS STRING * 11
  REDIM EChkNo(1)            AS STRING * 11
  REDIM TFedGrs(1)           AS STRING * 11
  REDIM TStaGrs(1)           AS STRING * 11
  REDIM TSocGrs(1)           AS STRING * 11
  REDIM TMedGrs(1)           AS STRING * 11
  'REDIM TRetGrs(1)          AS STRING * 11


'  FF$ = CHR$(12)
  Image2$ = "####.##"
  Image3$ = "########.##"

  LSET Fill11(1) = ""

  LSET Dash(1) = STRING$(132, "-")
  
  FGetAH UnitFileName, Unit(1), LEN(Unit(1)), 1
  FGetAH DedCodeFileName, DedCodes(1), LEN(DedCodes(1)), 12
  FGetAH ErnCodeFileName, ErnCodes(1), LEN(ErnCodes(1)), 3

'*** Create the voluntary deduction description line
  DTitle$ = ""
  FOR Cnt = 1 TO 12
    TDed$ = QPTrim$(DedCodes(Cnt).DCDESC1)
    IF LEN(TDed$) > 0 THEN
      LastDed = LastDed + 1
      RSET Ded(1) = TDed$
      DTitle$ = DTitle$ + Ded(1)
    ELSE
      EXIT FOR
    END IF
  NEXT
  IF LastDed < 12 THEN
    DTitle$ = SPACE$(11 * (12 - LastDed)) + DTitle$
  END IF

'*** Create the alternate earnings description line
  ETitle$ = ""
  FOR Cnt = 1 TO 3
    TErn$ = QPTrim$(ErnCodes(Cnt).ERNCODE1)
    IF LEN(TErn$) > 0 THEN
      LastErn = LastErn + 1
      RSET Ern(1) = TErn$
      ETitle$ = ETitle$ + Ern(1)
    ELSE
      EXIT FOR
    END IF
  NEXT
  IF LastErn < 3 THEN
    ETitle$ = SPACE$(11 * (3 - LastErn)) + ETitle$
  END IF

  SumHeader2$ = "  Reg Wages  O/T Wages" + ETitle$
  ETitle$ = "   Reg Earn   O/T Earn" + ETitle$ + "  Gross Pay    Soc Sec   Medicare        FWT        SWT     Retire    Net Pay"

'------------------------------------------------------------------

  EmpRecSize = LEN(Emp2Rec(1))

  TransRecLen = LEN(TransRec(1))

  LineCnt = 0
  MaxLines = 50
  Page = 1

  IdxRecLen = 2
  IdxFileSize& = FileSize(EmpIdxNName)
  NumOfRecs = IdxFileSize& \ IdxRecLen

  REDIM IdxBuff(1 TO NumOfRecs)
  FGetAH EmpIdxNName, IdxBuff(1), IdxRecLen, NumOfRecs


  REDIM FrmEHP(1) AS FormInfo
  NumFlds = LibNumberOfFields(MiscQLib, EmpHistory)
  REDIM FormEHP$(NumFlds, 2)             'DIM the form data array
  REDIM FldEHP(NumFlds) AS FieldInfo     'DIM the field information array
  StartEl = 0                            'Load first form at array start
  LibGetFldDef MiscQLib, EmpHistory, StartEl, FldEHP(), FormEHP$(), ErrCode

'----- Setup TYPE for setting and reading form editing information.
  FrmEHP(1).FldNo = 1                 'Start editing on field #1
  FrmEHP(1).InsStat = False           'Set the insert state (-1 = Insert on)
  FrmEHP(1).StartEl = 0               'Set form starting element to 0 and

  Emp1RecLen = LEN(Emp1Rec(1))

  FOpenS EmpData1Name, EHandle1
  FGetRTA EHandle1, Emp1Rec(1), CLNG(IdxBuff(1)), Emp1RecLen
  EmpHistRpt(1).FirstEmp& = VAL(Emp1Rec(1).EmpNo)
  FGetRTA EHandle1, Emp1Rec(1), CLNG(IdxBuff(NumOfRecs)), Emp1RecLen
  EmpHistRpt(1).LastEmp& = VAL(Emp1Rec(1).EmpNo)
  FClose EHandle1

  Today$ = DATE$
  EmpHistRpt(1).EndDate = Date2Num(Today$)
  EmpHistRpt(1).StartDate = Date2Num("01-01-20" + RIGHT$(Today$, 2))
  EmpHistRpt(1).SumOnly = "Y"

  HistFormLen = LEN(EmpHistRpt(1))

  FormEHP$(0, 0) = SPACE$(HistFormLen)

  BCopy VARSEG(EmpHistRpt(1)), VARPTR(EmpHistRpt(1)), SSEG(FormEHP$(0, 0)), SADD(FormEHP$(0, 0)), HistFormLen, 0

  CALL UnPackBuffer(1, NumFlds, FormEHP$(), FldEHP())

'----- Set the "Action" flag to force the editor to initialize itself and
'      display the data on the form.

  Action = 1
'******************************************************************

    BlockClear
    DisplayMiscScrn EmpHistory

    DO           'Pole the editing procedure
      EditForm FormEHP$(), FldEHP(), FrmEHP(1), Cnf, Action

      IF FrmEHP(1).Presses AND FrmEHP(1).MRow = 15 THEN
        SELECT CASE FrmEHP(1).MCol
        CASE 28 TO 38           'F10
          PressButton EscKey, 15, 28, 38
        CASE 41 TO 53           'Escape
          PressButton F0Key, 15, 41, 53
        END SELECT
      END IF
      
      SELECT CASE FrmEHP(1).KeyCode
      CASE F0Key
        FirstEmp& = QPValL(FormEHP$(1, 0))
        LastEmp& = QPValL(FormEHP$(2, 0))
        StartDate = Date2Num(FormEHP$(3, 0))
        EndDate = Date2Num(FormEHP$(4, 0))
        FromToDate$ = "Report Date: " + FormEHP$(3, 0) + " to " + FormEHP$(4, 0)

        IF INSTR("Yy", QPTrim$(FormEHP$(5, 0))) THEN
          SumFlag = True
        ELSE
          SumFlag = False
        END IF

        IF LastEmp& < FirstEmp& OR EndDate < StartDate THEN
          CursorOff
          SaveScrn TempScrn()
          DisplayMiscScrn ReportParmError
          WaitForAction
          RestScrn TempScrn()
          Action = 2
          IF LastEmp& < FirstEmp& THEN
            FrmEHP(1).FldNo = 1
          ELSE
            FrmEHP(1).FldNo = 3
          END IF
        ELSE
          OKFlag = True
        END IF
      CASE EscKey
        ExitFlag = True
      END SELECT
  LOOP UNTIL ExitFlag = True OR OKFlag = True   '  the Escape key.

'  CursorOff
  
  ERASE FormEHP$, FldEHP, FrmEHP, Emp1Rec, TempScrn
  
  IF ExitFlag GOTO AltExit

'**********************************************************
'------------------------------------------------------------------
  RptTitle$ = "Employee Earnings History Report"
  ShowProcessingScrn RptTitle$

  FCreate EmpHistoryRpt

  FOpenS EmpHistoryRpt, RHandle

  RPTSetupPRN 4, RHandle
  
  FOpenS TransHistFileName, THandle
'***
'  THistRecs = FileSize(TransHistFileName) \ TransRecLen
'***
  FOpenS EmpData2Name, DHandle   'open it
  EmpHistHeader = False

  FOR RecNo = 1 TO NumOfRecs
    UsingThisOne = False
    IF NOT SumFlag THEN
      EmpHistHeader = False
    END IF
    DAmt# = 0
    FGetRTA DHandle, Emp2Rec(1), CLNG(IdxBuff(RecNo)), EmpRecSize

    IF QPValL(Emp2Rec(1).EmpNo) >= FirstEmp& AND QPValL(Emp2Rec(1).EmpNo) <= LastEmp& THEN
    'if employee number is in range
      IF Emp2Rec(1).LastTransRec > 0 THEN         'if there are any
        TransRecNum& = Emp2Rec(1).LastTransRec
      ELSE
        GOTO Skip2NextEmp
      END IF
      DidDate = 0
      DO
        FGetRTA THandle, TransRec(1), TransRecNum&, TransRecLen
'***
'HistRecs = HistRecs + 1
'***
        IF (TransRec(1).CheckDate >= StartDate) AND (TransRec(1).CheckDate <= EndDate) THEN
         'IF NOT DidDate THEN
           DidDate = -1
          'if this is in the date range
           UsingThisOne = True
           GOSUB PrintAndSumEmp
           IF LineCnt >= MaxLines THEN
             FPut RHandle, FF$
             LineCnt = 0
             GOSUB PrintEmpHistoryHeader
           END IF
         'ELSE
         '  TransRec(1).CheckDate = -32767
         '  FPutRTA THandle, TransRec(1), TransRecNum&, TransRecLen
         'END IF
        END IF 'ELSE
          IF TransRec(1).PrevTransRec > 0 THEN
            TransRecNum& = TransRec(1).PrevTransRec
          ELSE
            IF UsingThisOne THEN
              GOSUB PrintSubTotal
              EXIT DO
            ELSE
              GOTO Skip2NextEmp
            END IF
          END IF
        'END IF
      LOOP
      EMPHIST(1) = EMPHIST(2)
      REDIM ESubDeds(1 TO 12) AS DOUBLE
      REDIM ESubErns(1 TO 3) AS DOUBLE

      TTaxFring# = Round#(TTaxFring# + TaxFring#)
      TFedGross# = Round#(TFedGross# + FedGross#)
      TStaGross# = Round#(TStaGross# + STAGROSS#)
      TSocGross# = Round#(TSocGross# + SocGross#)
      TMedGross# = Round#(TMedGross# + MedGross#)

      TaxFring# = 0
      FedGross# = 0
      STAGROSS# = 0
      SocGross# = 0
      MedGross# = 0

    END IF

Skip2NextEmp:
    ShowPctComp RecNo, NumOfRecs

  NEXT

  GOSUB PrintGrandTotals

  RPTSetupPRN 0, RHandle

  
  FClose THandle
  FClose DHandle
  FClose RHandle
'***
'  PRINT "History count;   File:"; THistRecs; " Processed:"; HistRecs
'  zz$ = INPUT$(1)
'***
  PrintRptFile RptTitle$, EmpHistoryRpt, 1, RetCode, 0
  
AltExit:

  ERASE EmpNo, Emp2Rec, TransRec, Unit, EMPHIST, DedCodes, ErnCodes, EmpHistRpt
  ERASE Dash, TotDeds, TotErns, ESubDeds, ESubErns, RErnP, EICP, GPayP
  ERASE SSTaxP, MTaxP, FTaxP, STaxP, RetirP, NetPayP, OErnP, Ded, Ern, Pg
  ERASE Fill11, RHrs, VHrs, SHrs, HHrs, CHrs, THrs, OTHrs, OTPaid, EICP, RErnP
  ERASE EChkDate, EChkNo, IdxBuff
  ERASE TFedGrs, TStaGrs, TSocGrs, TMedGrs

EXIT SUB
  
  
PrintEmpHistoryHeader:
  RSET Pg(1) = STR$(Page)
  
  FPut RHandle, Unit(1).UFEMPR + SPACE$(86) + "Page:" + Pg(1) + CrLf$
  FPut RHandle, "Employee Earnings History Report" + SPACE$(63) + FromToDate$ + CrLf$
  IF SumFlag THEN
    FPut RHandle, Dash(1) + CrLf$
    LineCnt = 4
  ELSE
    FPut RHandle, CrLf$
    LSET EmpNo(1) = QPTrim$(Emp2Rec(1).EmpNo)
    FPut RHandle, EmpNo(1) + QPTrim$(Emp2Rec(1).EMPLNAME) + ", " + QPTrim$(Emp2Rec(1).EMPFNAME) + CrLf$
    FPut RHandle, " Trans Date   Check No  Tax Fring   Reg Hrs      Vacat       Sick        Hol       Comp      Total    O/T Hrs   O/T Paid        EIC" + CrLf$
    FPut RHandle, ETitle$ + CrLf$
    FPut RHandle, DTitle$ + CrLf$
    FPut RHandle, Dash(1) + CrLf$
    LineCnt = 9
  END IF
  Page = Page + 1
  RETURN
  
PrintAndSumEmp:
  IF NOT EmpHistHeader THEN
    EmpHistHeader = True
    GOSUB PrintEmpHistoryHeader
  END IF

  EMPHIST(1).RegHrs = Round#(EMPHIST(1).RegHrs + TransRec(1).RegHrsWork)
  EMPHIST(1).VACHRS = Round#(EMPHIST(1).VACHRS + TransRec(1).VacUsed)
  EMPHIST(1).SICKHRS = Round#(EMPHIST(1).SICKHRS + TransRec(1).SickUsed)
  EMPHIST(1).HOLHRS = Round#(EMPHIST(1).HOLHRS + TransRec(1).HOLHOURS)
  EMPHIST(1).COMPHRS = Round#(EMPHIST(1).COMPHRS + TransRec(1).CompUsed)
  EMPHIST(1).TotalHrs = Round(EMPHIST(1).TotalHrs + TransRec(1).RegHrsWork + TransRec(1).VacUsed + TransRec(1).SickUsed + TransRec(1).HOLHOURS + TransRec(1).CompUsed)
  
  EMPHIST(1).TOTHrs = Round#(EMPHIST(1).TOTHrs + TransRec(1).OTHours)
  EMPHIST(1).TOTPaid = Round#(EMPHIST(1).TOTPaid + TransRec(1).OTHrsPaid)
  EMPHIST(1).TOTEIC = Round#(EMPHIST(1).TOTEIC + TransRec(1).EICAmt)
  
  EMPHIST(1).TRegWage = Round#(EMPHIST(1).TRegWage + TransRec(1).TotRegWage)
  EMPHIST(1).TOTWage = Round#(EMPHIST(1).TOTWage + TransRec(1).TotOTWage)
  
  EMPHIST(1).GPay = Round#(EMPHIST(1).GPay + TransRec(1).GrossPay)
  EMPHIST(1).SSTax = Round#(EMPHIST(1).SSTax + TransRec(1).SocTaxAmt)
  EMPHIST(1).MTax = Round#(EMPHIST(1).MTax + TransRec(1).MedTaxAmt)
  EMPHIST(1).FTax = Round#(EMPHIST(1).FTax + TransRec(1).FedTaxAmt)
  EMPHIST(1).STax = Round#(EMPHIST(1).STax + TransRec(1).StaTaxAmt)

  TaxFring# = Round#(TaxFring# + TransRec(1).TaxFring)

  FedGross# = Round#(FedGross# + TransRec(1).FedGrossPay)
  STAGROSS# = Round#(STAGROSS# + TransRec(1).StaGrossPay)
  SocGross# = Round#(SocGross# + TransRec(1).SocGrossPay)
  MedGross# = Round#(MedGross# + TransRec(1).MedGrossPay)

'*****
  IF TransRec(1).TaxFring > 0 THEN
    FedGross# = Round#(FedGross# + TransRec(1).TaxFring)
    STAGROSS# = Round#(STAGROSS# + TransRec(1).TaxFring)
    SocGross# = Round#(SocGross# + TransRec(1).TaxFring)
    MedGross# = Round#(MedGross# + TransRec(1).TaxFring)
  END IF

  EMPHIST(1).RETTOT = Round(EMPHIST(1).RETTOT + TransRec(1).RetireAmt)
  'TOTDED# = Round(TOTDED# + TransRec(1).TotDedAmt)
  
  EMPHIST(1).TNetPay = Round#(EMPHIST(1).TNetPay + TransRec(1).NetPay)

  'IF TransRec(1).OTHrsPaid <> 0 THEN STOP

  IF NOT SumFlag THEN
    LSET EChkDate(1) = Num2Date$(TransRec(1).CheckDate)           'LTRIM$(EmpRec1(1).EmpNo)
    'RSET EChkNo(1) = STR$(TransRec(1).CheckNum)   'QPRTrim$(EmpRec1(1).EMPLNAME) + ", " + QPRTrim$(EmpRec1(1).EMPFNAME)
    RSET EChkNo(1) = STR$(TransRecNum&)
    RSET RHrs(1) = FUsing(STR$(TransRec(1).RegHrsWork), Image2$)
    RSET VHrs(1) = FUsing(STR$(TransRec(1).VacUsed), Image2$)
    RSET SHrs(1) = FUsing(STR$(TransRec(1).SickUsed), Image2$)
    RSET HHrs(1) = FUsing(STR$(TransRec(1).HOLHOURS), Image2$)
    RSET CHrs(1) = FUsing(STR$(TransRec(1).CompUsed), Image2$)
    RSET THrs(1) = FUsing(STR$(TransRec(1).RegHrsPaid), Image2$)
    RSET OTHrs(1) = FUsing(STR$(TransRec(1).OTHours), Image2$)
    RSET OTPaid(1) = FUsing(STR$(TransRec(1).OTHrsPaid), Image2$)
    RSET EICP(1) = FUsing(STR$(TransRec(1).EICAmt), Image2$)
    RSET Fill11(1) = FUsing(STR$(TransRec(1).TaxFring), Image3$)
    RSET RErnP(1) = FUsing(STR$(TransRec(1).TotRegWage), Image3$)
    RSET OErnP(1) = FUsing(STR$(TransRec(1).TotOTWage), Image3$)
    RSET GPayP(1) = FUsing(STR$(TransRec(1).GrossPay), Image3$)
    RSET SSTaxP(1) = FUsing(STR$(TransRec(1).SocTaxAmt), Image3$)
    RSET MTaxP(1) = FUsing(STR$(TransRec(1).MedTaxAmt), Image3$)
    RSET FTaxP(1) = FUsing(STR$(TransRec(1).FedTaxAmt), Image3$)
    RSET STaxP(1) = FUsing(STR$(TransRec(1).StaTaxAmt), Image3$)
    RSET RetirP(1) = FUsing(STR$(TransRec(1).RetireAmt), Image3$)
    RSET NetPayP(1) = FUsing(STR$(TransRec(1).NetPay), Image3$)
  END IF

  SELECT CASE TransRec(1).PayType
  CASE "S"
    RSET RHrs(1) = "Salaried"
    SalCnt = SalCnt + 1
  CASE ELSE
    HrlCnt = HrlCnt + 1
  END SELECT
  
  SumDed$ = ""
  FOR Cnt2 = 1 TO LastDed
    ESubDeds(Cnt2) = Round#(ESubDeds(Cnt2) + TransRec(1).DAmt(Cnt2))
    TotDeds(Cnt2) = Round#(TotDeds(Cnt2) + TransRec(1).DAmt(Cnt2))
    RSET Ded(1) = FUsing(STR$(TransRec(1).DAmt(Cnt2)), Image3$)
    SumDed$ = SumDed$ + Ded(1)
  NEXT
  IF LastDed < 12 THEN
    SumDed$ = SPACE$(11 * (12 - LastDed)) + SumDed$
  END IF
  
  '----------------------------------------------
  SumErn$ = ""
  FOR Cnt2 = 1 TO LastErn
    ESubErns(Cnt2) = Round#(ESubErns(Cnt2) + TransRec(1).EAmt(Cnt2))
    TotErns(Cnt2) = Round#(TotErns(Cnt2) + TransRec(1).EAmt(Cnt2))
    RSET Ern(1) = FUsing(STR$(TransRec(1).EAmt(Cnt2)), Image3$)
    SumErn$ = SumErn$ + Ern(1)
  NEXT
  IF LastErn < 3 THEN
    SumErn$ = SPACE$(11 * (3 - LastErn)) + SumErn$
  END IF
  '-------------------------------------------------------
'  RSET EL2(1).SumEarn = SumErn$

  IF NOT SumFlag THEN
    FPut RHandle, EChkDate(1) + EChkNo(1) + Fill11(1) + RHrs(1) + VHrs(1) + SHrs(1) + HHrs(1)
    FPut RHandle, CHrs(1) + THrs(1) + OTHrs(1) + OTPaid(1) + EICP(1) + CrLf$
    FPut RHandle, RErnP(1) + OErnP(1) + SumErn$ + GPayP(1) + SSTaxP(1) + MTaxP(1) + FTaxP(1) + STaxP(1)
    FPut RHandle, RetirP(1) + NetPayP(1) + CrLf$
    FPut RHandle, SumDed$ + CrLf$ + CrLf$
    LineCnt = LineCnt + 4
  END IF

'JumpThis:
RETURN
  
PrintSubTotal:
  
  RSET THrs(1) = FUsing(STR$(EMPHIST(1).TotalHrs), Image3$)
  RSET RHrs(1) = FUsing(STR$(EMPHIST(1).RegHrs), Image3$)
  RSET VHrs(1) = FUsing(STR$(EMPHIST(1).VACHRS), Image3$)
  RSET SHrs(1) = FUsing(STR$(EMPHIST(1).SICKHRS), Image3$)
  RSET HHrs(1) = FUsing(STR$(EMPHIST(1).HOLHRS), Image3$)
  RSET CHrs(1) = FUsing(STR$(EMPHIST(1).COMPHRS), Image3$)
  RSET OTHrs(1) = FUsing(STR$(EMPHIST(1).TOTHrs), Image3$)
  RSET OTPaid(1) = FUsing(STR$(EMPHIST(1).TOTPaid), Image3$)
  RSET EICP(1) = FUsing(STR$(EMPHIST(1).TOTEIC), Image3$)
  
  RSET RErnP(1) = FUsing(STR$(EMPHIST(1).TRegWage), Image3$)
  RSET OErnP(1) = FUsing(STR$(EMPHIST(1).TOTWage), Image3$)
  
  RSET GPayP(1) = FUsing(STR$(EMPHIST(1).GPay), Image3$)
  RSET SSTaxP(1) = FUsing(STR$(EMPHIST(1).SSTax), Image3$)
  RSET MTaxP(1) = FUsing(STR$(EMPHIST(1).MTax), Image3$)
  RSET FTaxP(1) = FUsing(STR$(EMPHIST(1).FTax), Image3$)
  RSET STaxP(1) = FUsing(STR$(EMPHIST(1).STax), Image3$)
  RSET RetirP(1) = FUsing(STR$(EMPHIST(1).RETTOT), Image3$)
  RSET NetPayP(1) = FUsing(STR$(EMPHIST(1).TNetPay), Image3$)

  RSET Fill11(1) = FUsing(STR$(TaxFring#), Image3$)
  RSET TFedGrs(1) = FUsing(STR$(FedGross#), Image3$)
  RSET TStaGrs(1) = FUsing(STR$(STAGROSS#), Image3$)
  RSET TSocGrs(1) = FUsing(STR$(SocGross#), Image3$)
  RSET TMedGrs(1) = FUsing(STR$(MedGross#), Image3$)

  '---------------------------------------------------------------
  
  EMPHIST(3).TotalHrs = Round(EMPHIST(3).TotalHrs + EMPHIST(1).TotalHrs)
  EMPHIST(3).RegHrs = Round(EMPHIST(3).RegHrs + EMPHIST(1).RegHrs)
  EMPHIST(3).VACHRS = Round(EMPHIST(3).VACHRS + EMPHIST(1).VACHRS)
  EMPHIST(3).SICKHRS = Round(EMPHIST(3).SICKHRS + EMPHIST(1).SICKHRS)
  EMPHIST(3).HOLHRS = Round(EMPHIST(3).HOLHRS + EMPHIST(1).HOLHRS)
  EMPHIST(3).COMPHRS = Round(EMPHIST(3).COMPHRS + EMPHIST(1).COMPHRS)
  EMPHIST(3).TOTHrs = Round(EMPHIST(3).TOTHrs + EMPHIST(1).TOTHrs)
  EMPHIST(3).TOTPaid = Round(EMPHIST(3).TOTPaid + EMPHIST(1).TOTPaid)
  EMPHIST(3).TOTEIC = Round(EMPHIST(3).TOTEIC + EMPHIST(1).TOTEIC)
  EMPHIST(3).TRegWage = Round(EMPHIST(3).TRegWage + EMPHIST(1).TRegWage)
  EMPHIST(3).TOTWage = Round(EMPHIST(3).TOTWage + EMPHIST(1).TOTWage)
  EMPHIST(3).GPay = Round(EMPHIST(3).GPay + EMPHIST(1).GPay)
  EMPHIST(3).SSTax = Round(EMPHIST(3).SSTax + EMPHIST(1).SSTax)
  EMPHIST(3).MTax = Round(EMPHIST(3).MTax + EMPHIST(1).MTax)
  EMPHIST(3).FTax = Round(EMPHIST(3).FTax + EMPHIST(1).FTax)
  EMPHIST(3).STax = Round(EMPHIST(3).STax + EMPHIST(1).STax)
  EMPHIST(3).RETTOT = Round(EMPHIST(3).RETTOT + EMPHIST(1).RETTOT)
  EMPHIST(3).TNetPay = Round(EMPHIST(3).TNetPay + EMPHIST(1).TNetPay)
  '---------------------------------------------------------------
  
  SumDed$ = ""
  FOR Cnt2 = 1 TO LastDed
    RSET Ded(1) = FUsing(STR$(ESubDeds(Cnt2)), Image3$)
    SumDed$ = SumDed$ + Ded(1)
  NEXT
  IF LastDed < 12 THEN
    SumDed$ = SPACE$(11 * (12 - LastDed)) + SumDed$
  END IF
  '---------------------------------------------------------
  SumErn$ = ""
  FOR Cnt2 = 1 TO LastErn
    RSET Ern(1) = FUsing(STR$(ESubErns(Cnt2)), Image3$)
    SumErn$ = SumErn$ + Ern(1)
  NEXT
  IF LastErn < 3 THEN
    SumErn$ = SPACE$(11 * (3 - LastErn)) + SumErn$
  END IF
  
  '--------------NEW----------------------------
  RSET Pg(1) = STR$(Page)
  IF NOT SumFlag THEN
    FPut RHandle, Dash(1) + CrLf$
  END IF
  IF SumFlag THEN
    LSET EmpNo(1) = QPTrim$(Emp2Rec(1).EmpNo)
    FPut RHandle, EmpNo(1) + QPTrim$(Emp2Rec(1).EMPLNAME) + ", " + QPTrim$(Emp2Rec(1).EMPFNAME) + CrLf$
    FPut RHandle, "                        Tax Fring    Reg Hrs      Vacat       Sick        Hol       Comp      Total    O/T Hrs   O/T Paid        EIC" + CrLf$
  ELSE
    FPut RHandle, "Employee Totals:        Tax Fring    Reg Hrs      Vacat       Sick        Hol       Comp      Total    O/T Hrs   O/T Paid        EIC" + CrLf$
  END IF

  FPut RHandle, SPACE$(22) + Fill11(1) + RHrs(1) + VHrs(1) + SHrs(1) + HHrs(1)
'  FPut RHandle, "Employee Totals:      " + Fill11(1) + RHrs(1) + VHrs(1) + SHrs(1) + HHrs(1)
  
  FPut RHandle, CHrs(1) + THrs(1) + OTHrs(1) + OTPaid(1) + EICP(1) + CrLf$
  FPut RHandle, CrLf$
  FPut RHandle, SumHeader2$ + "  Gross Pay    Soc Sec   Medicare        FWT        SWT  Ret Total    Net Pay" + CrLf$
  FPut RHandle, RErnP(1) + OErnP(1) + SumErn$ + GPayP(1) + SSTaxP(1) + MTaxP(1) + FTaxP(1)
  FPut RHandle, STaxP(1) + RetirP(1) + NetPayP(1) + CrLf$
  FPut RHandle, CrLf$
  FPut RHandle, DTitle$ + CrLf$
  FPut RHandle, SumDed$ + CrLf$
  FPut RHandle, CrLf$
  FPut RHandle, "  Fed Gross  Sta Gross  Soc Gross  Med Gross" + CrLf$
  FPut RHandle, TFedGrs(1) + TStaGrs(1) + TSocGrs(1) + TMedGrs(1) + CrLf$
  FPut RHandle, CrLf$
  FPut RHandle, Dash(1) + CrLf$
  IF NOT SumFlag THEN
    
    FPut RHandle, FF$
    LineCnt = 0
  ELSE
    LineCnt = LineCnt + 14
  END IF
RETURN
  
  
  '-----------------------------------------------------------------------
PrintGrandTotals:
  RSET Fill11(1) = FUsing(STR$(TTaxFring#), Image3$)
  RSET THrs(1) = FUsing(STR$(EMPHIST(3).TotalHrs), Image3$)
  RSET RHrs(1) = FUsing(STR$(EMPHIST(3).RegHrs), Image3$)
  RSET VHrs(1) = FUsing(STR$(EMPHIST(3).VACHRS), Image3$)
  RSET SHrs(1) = FUsing(STR$(EMPHIST(3).SICKHRS), Image3$)
  RSET HHrs(1) = FUsing(STR$(EMPHIST(3).HOLHRS), Image3$)
  RSET CHrs(1) = FUsing(STR$(EMPHIST(3).COMPHRS), Image3$)
  RSET OTHrs(1) = FUsing(STR$(EMPHIST(3).TOTHrs), Image3$)
  RSET OTPaid(1) = FUsing(STR$(EMPHIST(3).TOTPaid), Image3$)
  RSET EICP(1) = FUsing(STR$(EMPHIST(3).TOTEIC), Image3$)
  
  RSET RErnP(1) = FUsing(STR$(EMPHIST(3).TRegWage), Image3$)
  RSET OErnP(1) = FUsing(STR$(EMPHIST(3).TOTWage), Image3$)
  
  RSET GPayP(1) = FUsing(STR$(EMPHIST(3).GPay), Image3$)
  RSET SSTaxP(1) = FUsing(STR$(EMPHIST(3).SSTax), Image3$)
  RSET MTaxP(1) = FUsing(STR$(EMPHIST(3).MTax), Image3$)
  RSET FTaxP(1) = FUsing(STR$(EMPHIST(3).FTax), Image3$)
  RSET STaxP(1) = FUsing(STR$(EMPHIST(3).STax), Image3$)
  RSET RetirP(1) = FUsing(STR$(EMPHIST(3).RETTOT), Image3$)
  RSET NetPayP(1) = FUsing(STR$(EMPHIST(3).TNetPay), Image3$)

  RSET TFedGrs(1) = FUsing(STR$(TFedGross#), Image3$)
  RSET TStaGrs(1) = FUsing(STR$(TStaGross#), Image3$)
  RSET TSocGrs(1) = FUsing(STR$(TSocGross#), Image3$)
  RSET TMedGrs(1) = FUsing(STR$(TMedGross#), Image3$)

  '---------------------------------------------------------------
  SumDed$ = ""
  FOR Cnt2 = 1 TO LastDed
    RSET Ded(1) = FUsing(STR$(TotDeds(Cnt2)), Image3$)
    SumDed$ = SumDed$ + Ded(1)
  NEXT
  IF LastDed < 12 THEN
    SumDed$ = SPACE$(11 * (12 - LastDed)) + SumDed$
  END IF
  '---------------------------------------------------------
  SumErn$ = ""
  FOR Cnt2 = 1 TO LastErn
    RSET Ern(1) = FUsing(STR$(TotErns(Cnt2)), Image3$)
    SumErn$ = SumErn$ + Ern(1)
  NEXT
  IF LastErn < 3 THEN
    SumErn$ = SPACE$(11 * (3 - LastErn)) + SumErn$
  END IF
  
  '--------------NEW----------------------------
  IF SumFlag THEN
    FPut RHandle, FF$
  END IF
  RSET Pg(1) = STR$(Page)
  FPut RHandle, Unit(1).UFEMPR + SPACE$(86) + "Page:" + Pg(1) + CrLf$
  FPut RHandle, "Employee Earnings History Report" + SPACE$(63) + FromToDate$ + CrLf$
  FPut RHandle, Dash(1) + CrLf$
  FPut RHandle, CrLf$
  FPut RHandle, "Report Totals:          Tax Fring    Reg Hrs      Vacat       Sick        Hol       Comp      Total    O/T Hrs   O/T Paid        EIC" + CrLf$
  FPut RHandle, SPACE$(22) + Fill11(1) + RHrs(1) + VHrs(1) + SHrs(1) + HHrs(1)
  FPut RHandle, CHrs(1) + THrs(1) + OTHrs(1) + OTPaid(1) + EICP(1) + CrLf$
  FPut RHandle, CrLf$
  FPut RHandle, SumHeader2$ + "  Gross Pay    Soc Sec   Medicare        FWT        SWT  Ret Total    Net Pay" + CrLf$

'add grand totals here
  FPut RHandle, RErnP(1) + OErnP(1) + SumErn$ + GPayP(1) + SSTaxP(1) + MTaxP(1) + FTaxP(1)
  FPut RHandle, STaxP(1) + RetirP(1) + NetPayP(1) + CrLf$

  FPut RHandle, CrLf$
  FPut RHandle, DTitle$ + CrLf$
  FPut RHandle, SumDed$ + CrLf$ + CrLf$
  FPut RHandle, "  Fed Gross  Sta Gross  Soc Gross  Med Gross" + CrLf$
  FPut RHandle, TFedGrs(1) + TStaGrs(1) + TSocGrs(1) + TMedGrs(1) + CrLf$
  
  FPut RHandle, FF$
  
  RETURN
  
END SUB

SUB ESCQuartlyWageRpt

  FirstTime = True
  REDIM TempScrn(1)

  REDIM Frm(1) AS FormInfo

  FormName$ = "QESCRPT"
  NumFlds = LibNumberOfFields(MiscQLib, FormName$)
  REDIM Form$(NumFlds, 2)             'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo     'DIM the field information array
  StartEl = 0                         'Load first form at array start
  LibGetFldDef MiscQLib, FormName$, StartEl, Fld(), Form$(), ErrCode


  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1

  '----- Setup TYPE for setting and reading form editing information.

  Frm(1).FldNo = 1                 'Start editing on field #1
  Frm(1).InsStat = False           'Set insert state (True = Insert on)
  Frm(1).StartEl = 0               'Set form starting element to 0 and

  BlockClear
  DisplayMiscScrn FormName$

  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      Form$(2, 0) = RIGHT$(DATE$, 4)
      Action = 2
    END IF
    SELECT CASE Frm(1).KeyCode
      CASE F0Key
        RptQuarter = QPValI(Form$(1, 0))
        Year$ = Form$(2, 0)

        IF (RptQuarter > 0 AND RptQuarter < 5) THEN
          OKFlag = True
        ELSE
          CursorOff
          SaveScrn TempScrn()
          DisplayMiscScrn ReportParmError
          WaitForAction
          RestScrn TempScrn()
          Frm(1).FldNo = 1
          Action = 2
          OKFlag = False
        END IF
      CASE EscKey
        OKFlag = True
        ExitFlag = True
    END SELECT

  LOOP UNTIL OKFlag             'proper key not set

  ERASE Form$, Fld, Frm, TempScrn

  IF ExitFlag THEN EXIT SUB

  CrLf2$ = CrLf$ + CrLf$
  CrLf5$ = CrLf2$ + CrLf2$ + CrLf$
  CrLf6$ = CrLf2$ + CrLf2$ + CrLf2$

  REDIM TransHRec(1)   AS TransRecType
  REDIM Emp2Rec(1)     AS EmpData2Type
  REDIM Unit(1)        AS UnitFileRecType
  REDIM Pg(1)          AS STRING * 3
  REDIM GrsRpt(1 TO 3) AS ESCGrossWageRptType
  REDIM Line2(1)       AS STRING * 80

  REDIM EQtrWage(1 TO 4) AS DOUBLE
  REDIM QtrDates(1 TO 4) AS QtrDateType

  'get the unit file here
  FGetAH UnitFileName, Unit(1), LEN(Unit(1)), 1

  IF Unit(1).UFSTATE = "NC" THEN
    NCFlag = True
  END IF

  REDIM ToDisk1(1)      AS ESC2DiskRecType1
  ToDisk1(1).Fill1 = ""
  ToDisk1(1).CrLf = CrLf$

  Fmt$ = "#######,.##"
  MaxLines = 25

  LineCnt = 0

  EmpRecSize = LEN(Emp2Rec(1))
  TRecSize = LEN(TransHRec(1))

  IdxRecLen = 2
  IdxFileSize& = FileSize(EmpIdxLName)
  NumOfRecs = IdxFileSize& \ IdxRecLen

  REDIM IdxBuff(1 TO NumOfRecs)
  FGetAH EmpIdxLName, IdxBuff(1), IdxRecLen, NumOfRecs
  
  QtrDates(1).LDate = Date2Num("0101" + Year$)
  QtrDates(1).HDate = Date2Num("0331" + Year$)
  QtrDates(2).LDate = Date2Num("0401" + Year$)
  QtrDates(2).HDate = Date2Num("0630" + Year$)
  QtrDates(3).LDate = Date2Num("0701" + Year$)
  QtrDates(3).HDate = Date2Num("0930" + Year$)
  QtrDates(4).LDate = Date2Num("1001" + Year$)
  QtrDates(4).HDate = Date2Num("1231" + Year$)

  Qtr$ = QPStrI(RptQuarter) + " " + Year$

  RptTitle$ = "ESC Quarterly Wage Report"
  ShowProcessingScrn RptTitle$

  RptName$ = "PRRPTS\ESCQTR" + QPStrI$(RptQuarter) + ".RPT"

  FCreate RptName$
  FOpenS RptName$, RHandle

  RPTSetupPRN 7, RHandle

  FOpenS TransHistFileName, THandle
  FOpenS EmpData2Name, DHandle   'open employee data file

  GOSUB PrintESCHeader

  FOR RecNo = 1 TO NumOfRecs
    FGetRTA DHandle, Emp2Rec(1), CLNG(IdxBuff(RecNo)), EmpRecSize
    IF (Emp2Rec(1).LastTransRec <= 0) THEN
    'IF Emp2Rec(1).LastTransRec <= 0 THEN
      GOTO SkipEm5
    END IF
    TransRecNum& = CLNG(Emp2Rec(1).LastTransRec)

    DO
      FGetRTA THandle, TransHRec(1), TransRecNum&, TRecSize
      GOSUB SumEmpESCData
      IF TransHRec(1).PrevTransRec <= 0 THEN
        GOSUB PrintEmpESCLine
        IF LineCnt >= MaxLines THEN
          GOSUB PrintSubTotals
          FPut RHandle, FF$ + CrLf$
          GOSUB PrintESCHeader
        END IF
        EXIT DO
      ELSE
        TransRecNum& = CLNG(TransHRec(1).PrevTransRec)
      END IF
    LOOP
SkipEm5:
    ShowPctComp RecNo, NumOfRecs
  NEXT

  GOSUB PrintESCGTotals
  FPut RHandle, FF$
  Page = Page + 1
  LSET Pg(1) = QPStrI(Page)
  FPut RHandle, CrLf6$
  FPut RHandle, "Employer's Quarterly Tax and Wage Report Summary" + CrLf2$
  FPut RHandle, "  Total Wages:  " + FUsing$(STR$(GrandTotal#), Fmt$) + CrLf$
  IF Unit(1).ESCRTYPE = 2 THEN
    FPut RHandle, " Excess Wages:  " + FUsing$(STR$(TotalGrossOvr#), Fmt$) + CrLf$
    FPut RHandle, "Taxable Wages:  " + FUsing$(STR$(Round(GrandTotal# - TotalGrossOvr#)), Fmt$) + CrLf$
  END IF
  FPut RHandle, FF$
  RPTSetupPRN 0, RHandle
  
  FClose DHandle
  FClose THandle
  FClose RHandle

  CLOSE

  PrintRptFile RptTitle$, RptName$, 1, RetCode, 0

  ERASE TransHRec, Emp2Rec, Unit, Pg, GrsRpt
  ERASE IdxBuff, Line2, ToDisk1, EQtrWage


EXIT SUB

PrintESCHeader:
  LineCnt = 0
  Page = Page + 1
  LSET Pg(1) = QPStrI(Page)
  '---
  FPut RHandle, CrLf6$
  '---
  FPut RHandle, "   " + Unit(1).UFEMPR + SPACE$(33) + Pg(1) + SPACE$(10) + Unit(1).UFSTAID + CrLf$
  FPut RHandle, CrLf5$
  DoQtrLine = True
  'set a flag and todisk$ to qtr$    HERE
RETURN

SumEmpESCData:
  FOR Cnt2 = 1 TO 4  'put gross into correct quarter
    IF (TransHRec(1).CheckDate >= QtrDates(Cnt2).LDate) AND (TransHRec(1).CheckDate <= QtrDates(Cnt2).HDate) THEN
      IF NCFlag THEN
        EQtrWage(Cnt2) = Round(EQtrWage(Cnt2) + TransHRec(1).SocGrossPay)
      ELSE
        EQtrWage(Cnt2) = Round(EQtrWage(Cnt2) + TransHRec(1).GrossPay)
      END IF
      EXIT FOR
    END IF
  NEXT
RETURN

PrintEmpESCLine:

  IF Unit(1).ESCRTYPE = 2 THEN
    IF RptQuarter > 1 THEN            'if Not 1st qtr, we will have to
      FOR Cnt2 = 1 TO RptQuarter - 1  'examine gross thru prior qtr
        YTD2PrevQtr# = Round(YTD2PrevQtr# + EQtrWage(Cnt2))
      NEXT
      FOR Cnt2 = 1 TO RptQuarter
        YTD2ThisQtr# = Round(YTD2ThisQtr# + EQtrWage(Cnt2))
      NEXT
      IF YTD2PrevQtr# > Unit(1).TAXWBASE THEN     'if the prior qtr was
        GrossOvr# = EQtrWage(RptQuarter)          'over also TaxBase also
      ELSEIF YTD2ThisQtr# > Unit(1).TAXWBASE THEN         'else if gross thru
        GrossOvr# = Round(YTD2ThisQtr# - Unit(1).TAXWBASE)'this qtr is over
      ELSE   'still not over
        GrossOvr# = 0
      END IF
'*** This deals with the 1ST quarter only!!
    ELSEIF EQtrWage(1) > Unit(1).TAXWBASE THEN  'else this is 1st qtr report
      GrossOvr# = Round(EQtrWage(1) - Unit(1).TAXWBASE)
    ELSE
      GrossOvr# = 0
    END IF
'*** 1ST Quarter end
    GrsRpt(1).GrossPay = EQtrWage(RptQuarter) 'Round(EQtrWage(RptQuarter) - GrossOvr#)
  ELSE       'not type 2 report
    GrsRpt(1).GrossPay = EQtrWage(RptQuarter)
  END IF

  'IF GrossOvr# > 0 THEN STOP

  TotalGrossOvr# = Round(TotalGrossOvr# + GrossOvr#)

  IF GrsRpt(1).GrossPay = 0 GOTO SkipEMPPrint

  RSET ToDisk1(1).GPay = FUsing$(STR$(GrsRpt(1).GrossPay), Fmt$)
  LSET ToDisk1(1).ESSN = LEFT$(Emp2Rec(1).EMPSSN, 3) + "-" + MID$(Emp2Rec(1).EMPSSN, 4, 2) + "-" + MID$(Emp2Rec(1).EMPSSN, 6, 4)
  LSET ToDisk1(1).EName = LEFT$(Emp2Rec(1).EMPFNAME, 1) + " " + LEFT$(Emp2Rec(1).EMPLNAME, 18)

  IF DoQtrLine THEN
    RSET ToDisk1(1).Qtr = Qtr$
    DoQtrLine = False
  ELSE
    RSET ToDisk1(1).Qtr = " "
  END IF

  FPutA RHandle, ToDisk1(1), 80
  FPut RHandle, CrLf$
  LineCnt = LineCnt + 1     'employeesprinted = employeesprinted + 1
  SubTotal# = Round(SubTotal# + GrsRpt(1).GrossPay)
SkipEMPPrint:

  GrsRpt(1) = GrsRpt(2)
  YTD2PrevQtr# = 0
  YTD2ThisQtr# = 0
  FOR Cnt2 = 1 TO 4
    EQtrWage(Cnt2) = 0
  NEXT

RETURN

PrintSubTotals:
  RSET Line2(1) = (FUsing$(STR$(SubTotal#), Fmt$) + CrLf$)
  FPut RHandle, Line2(1)
  GrandTotal# = Round(GrandTotal# + SubTotal#)
  SubTotal# = 0
RETURN

PrintESCGTotals:
  IF LineCnt < MaxLines THEN
    FOR Cnt = LineCnt TO MaxLines - 1
      FPut RHandle, CrLf2$
    NEXT
  END IF
  GOSUB PrintSubTotals
RETURN

END SUB

SUB PRReportMenu
  SHARED TaxText$()

  IF TaxText$(0) = "NC" THEN
    VAFlag = False
    MenuCnt = 14
  ELSE
    VAFlag = True
    MenuCnt = 12
  END IF
  
  Choice = 1

  DO
    PayMenu ReportsMenu, Choice, MenuCnt

    SELECT CASE Choice
    CASE 1
      'EMPrintEmpData               'unrem
    CASE 2
      'EMPrintEmpList               'unrem
    CASE 3
      'EMPrintTermEmpList          'unrem
    CASE 4
      EmpEarnHistRpt
    CASE 5
      'GrossWageReport              'unrem
    CASE 6
      'PayDedReport                 'unrem
    CASE 7
      ESCQuartlyWageRpt
    CASE 8
      'BenefitReport                'unrem
    CASE 9
      'YTDFundReport                'unrem
    CASE 10
      'CheckIssuedRpt               'unrem
    CASE 11
      'RePrintRpts                  'unrem
    CASE 12
       'VARetirementRpt                'unrem
        'RetirementRpt                'unrem
        'SCRetirementRpt                'unrem
    CASE 13
      'SupRetReport                 'unrem
    CASE 14
      SEPPContRpt                  'unrem
    END SELECT
  LOOP UNTIL Choice = EscKey

END SUB

SUB SEPPContRpt

  FirstTime = True

  REDIM TempScrn(1)
  REDIM Frm(1) AS FormInfo

  FormName$ = "SEPPRPT"
  NumFlds = LibNumberOfFields(MiscQLib, FormName$)
  REDIM Form$(NumFlds, 2)             'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo     'DIM the field information array
  StartEl = 0                         'Load first form at array start
  LibGetFldDef MiscQLib, FormName$, StartEl, Fld(), Form$(), ErrCode


  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1

  '----- Setup TYPE for setting and reading form editing information.

  Frm(1).FldNo = 1                 'Start editing on field #1
  Frm(1).InsStat = False           'Set insert state (True = Insert on)
  Frm(1).StartEl = 0               'Set form starting element to 0 and

  BlockClear
  DisplayMiscScrn FormName$

  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      Form$(2, 0) = DATE$
      Form$(1, 0) = LEFT$(DATE$, 3) + "01-" + RIGHT$(DATE$, 4)
      Action = 2
    END IF
    SELECT CASE Frm(1).KeyCode
      CASE F0Key
        LowDate = Date2Num(Form$(1, 0))
        HighDate = Date2Num(Form$(2, 0))
        IF LowDate <= HighDate THEN
          MonthNum = QPValI(LEFT$(Form$(2, 0), 2))
          EPct# = VAL(Form$(3, 0))
          OKFlag = True
        ELSE
          CursorOff
          SaveScrn TempScrn()
          DisplayMiscScrn ReportParmError
          WaitForAction
          RestScrn TempScrn()
          Frm(1).FldNo = 1
          Action = 2
          OKFlag = False
        END IF
      CASE EscKey
        OKFlag = True
        ExitFlag = True
    END SELECT

  LOOP UNTIL OKFlag           'proper key not set

  CursorOff


  ERASE Form$, Fld, TempScrn, Frm
  
  IF ExitFlag THEN EXIT SUB

  CursorOff

  RptName$ = "PRRPTS\SEPPCONT.RPT"

'  FF$ = CHR$(12)

  'Image1$ = "$$,####.##"
  'Image2$ = "$$###,###.##"

  REDIM TransHRec(1)     AS TransRecType
  REDIM Emp2Rec(1)  AS EmpData2Type
  REDIM Unit(1)        AS UnitFileRecType
  REDIM Pg(1)          AS STRING * 3
  REDIM Dash(1)     AS STRING * 80

  Dash(1) = STRING$(80, "-")

  EmpRecSize = LEN(Emp2Rec(1))
  TRecSize = LEN(TransHRec(1))

  IdxRecLen = 2

  IdxFileSize& = FileSize(EmpIdxLName)
  NumOfRecs = IdxFileSize& \ IdxRecLen

  REDIM IdxBuff(1 TO NumOfRecs)
  FGetAH EmpIdxLName, IdxBuff(1), IdxRecLen, NumOfRecs
  FGetAH UnitFileName, Unit(1), LEN(Unit(1)), 1
  MaxLines = 55

  RptTitle$ = "SEPP Contribution Report."
  ShowProcessingScrn RptTitle$
  
  OPEN RptName$ FOR OUTPUT AS #1

  FOpenS TransHistFileName, THandle
  FOpenS EmpData2Name, DHandle   'open employee data file

  GOSUB SEPPRptHeader

  FOR RecNo = 1 TO NumOfRecs
    UsingThisOne = False
    EMatchAmt# = 0
    GrossAmt# = 0
    RETAMT# = 0
    FGetRTA DHandle, Emp2Rec(1), CLNG(IdxBuff(RecNo)), EmpRecSize

    IF Emp2Rec(1).LastTransRec <= 0 THEN
      GOTO SkipEm8
    END IF

    TransRecNum& = CLNG(Emp2Rec(1).LastTransRec)

    DO
      FGetRTA THandle, TransHRec(1), TransRecNum&, TRecSize
      SELECT CASE TransHRec(1).CheckDate
      CASE LowDate TO HighDate
'        IF UCASE$(LEFT$(LTRIM$(Emp2Rec(1).EMPRETTP), 1)) = "L" THEN
         IF LEN(QPTrim$(Emp2Rec(1).EMPRETNO)) THEN
          GrossAmt# = Round(GrossAmt# + TransHRec(1).GrossPay)
          EMatchAmt# = Round((EMatchAmt# + (TransHRec(1).GrossPay * EPct#) * .01))
          TEMatch# = Round(TEMatch# + EMatchAmt#)
          UsingThisOne = True
        END IF
      CASE ELSE
      END SELECT

      IF TransHRec(1).PrevTransRec <= 0 THEN
        IF UsingThisOne THEN
          GOSUB PrintEmpSEPPLine
          IF LineCnt >= MaxLines THEN
            PRINT #1, FF$
            GOSUB SEPPRptHeader
          END IF
        END IF
        EXIT DO
      ELSE
        TransRecNum& = CLNG(TransHRec(1).PrevTransRec)
      END IF
    LOOP

SkipEm8:
    ShowPctComp RecNo, NumOfRecs
  NEXT
  GOSUB SEPPTotals

  FClose THandle
  FClose DHandle   'open employee data file
  CLOSE
'****************************************************************

  PrintRptFile RptTitle$, RptName$, 1, RetCode, 0

EXIT SUB


PrintEmpSEPPLine:
  PRINT #1, LEFT$(Emp2Rec(1).EMPSSN, 3); "-"; MID$(Emp2Rec(1).EMPSSN, 4, 2); "-"; MID$(Emp2Rec(1).EMPSSN, 6, 4);
  PRINT #1, TAB(14); LEFT$(Emp2Rec(1).EMPRETNO, 15);
  PRINT #1, TAB(30); QPTrim$(Emp2Rec(1).EMPLNAME); ", "; QPTrim$(Emp2Rec(1).EMPFNAME);
  PRINT #1, TAB(55); USING "#####,.##"; GrossAmt#; TAB(69); EMatchAmt#
  'PRINT #1,
  LineCnt = LineCnt + 1     'employeesprinted = employeesprinted + 1
  GTotal# = Round(GTotal# + GrossAmt#)
  RTotal# = Round(RTotal# + RETAMT#)
  RETURN

SEPPTotals:
  PRINT #1, Dash(1)
  PRINT #1, TAB(28); "Totals:";
  PRINT #1, TAB(55); USING "#####,.##"; GTotal#; TAB(69); TEMatch#
  PRINT #1, FF$
  RETURN

'EmpRetNo


SEPPRptHeader:
'PRINT #1, "North Carolina Department of State Treasurer"
'PRINT #1, "Law Enforcement Officers Benefit and Retirement Fund"
'PRINT #1, "Monthly Report of Subject Wages And Retirement Contributions"
'PRINT #1,
'PRINT #1, "Unit Code: "; Unit(1).UFRETID
'PRINT #1, "Month: "; MonthName$(MonthNum); " 1994"
'PRINT #1,
'PRINT #1, "                                                   Wages Subject   Retirement"
'PRINT #1, "Soc Sec #    Ret #         Employee Name           To Retirem't    Deductions"
'PRINT #1, Dash(1)
'LineCnt = 10
'RETURN
  Page = Page + 1
  RSET Pg(1) = STR$(Page)
  UTemp$ = SPACE$(70)
  LSET UTemp$ = Unit(1).UFEMPR
  MID$(UTemp$, 62) = "Page:" + Pg(1)
  PRINT #1, UTemp$
  PRINT #1, "SEPP Contribution Report."
  PRINT #1, "Month: "; MonthName$(MonthNum); " 1994"
  'PRINT #1, "                                                   Wages Subject     Employer  "
  PRINT #1, "Soc Sec #    Ret #           Employee Name           Gross Wage  Contribution"
  PRINT #1, Dash(1)
  'PRINT #1, ""
  LineCnt = 5
RETURN

END SUB

