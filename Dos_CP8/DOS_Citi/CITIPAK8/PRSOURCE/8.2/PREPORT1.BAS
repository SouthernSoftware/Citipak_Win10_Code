DEFINT A-Z
DECLARE SUB RPTSetupPRN (RPTNum%, Handle%)
DECLARE SUB PRReportMenu ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB SEPPContRpt ()
DECLARE SUB KillFile (FileName$)
DECLARE SUB InitMouse (MooseFlag%)
DECLARE FUNCTION Value# (E$, ErCode%)
DECLARE FUNCTION Date2Num (DateIn$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION LibNumberOfFields (LibName$, FrmName$)
DECLARE FUNCTION MonthName$ (MonthNum%)
DECLARE FUNCTION Num2Date$ (Number%)
DECLARE FUNCTION QPStrI$ (Num%)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPRTrim$ (Text$)
DECLARE FUNCTION QPValI% (Num$)
DECLARE FUNCTION QPValL& (Num$)
DECLARE FUNCTION Round# (DblNum#)
DECLARE SUB BlockClear ()
DECLARE SUB CursorOff ()
DECLARE SUB DisplayMiscScrn (ScrnName$)
DECLARE SUB ESCQuartlyWageRpt ()
DECLARE SUB EditForm (Form$(), Fld() AS ANY, Frm AS ANY, Cnf AS ANY, Action)
  
DECLARE SUB UnPackBufer (FirstFld, LastFld, Form$(), Fld() AS ANY)
  
DECLARE SUB EmpEarnHistRpt ()
  
DECLARE SUB FClose (Handle%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
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
DECLARE FUNCTION QPStrL$ (Num&)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOf%, MemberSize%)
  
  
  
  '$INCLUDE: 'DefCnf.bi'                  'Defines TYPE for monitor/color info.
  '$INCLUDE: 'FieldInf.bi'                '        "        field information
  '$INCLUDE: 'FormEdit.bi'                '        "        form editing info.
  '$INCLUDE: 'PRFiles.bi'
  '$INCLUDE: 'PREmpRec.bi'
  '$INCLUDE: 'Retire.Bi'
  '$INCLUDE: 'DedCodes.Bi'
  '$INCLUDE: 'ErnCodes.Bi'
  '$INCLUDE: 'Leave.Bi'
  '$INCLUDE: 'PRUNIT.Bi'
  '$INCLUDE: 'PRSYSCTR.Bi'
  '$INCLUDE: 'PRPPDEF.Bi'
  '$INCLUDE: 'PRTRANS.Bi'
  '$INCLUDE: 'oPRtypes.Bi'
  '$INCLUDE: 'PRTIMEIN.Bi'
  '$INCLUDE: 'PRRpts.BI'
  
  
  CONST False = 0
  CONST True = NOT False
  
  CrLf$ = CHR$(13) + CHR$(10)
  
  FF$ = CHR$(12)

SUB BenefitReport
  
  MinHrs = -10000
  Image1$ = ",####.##"
  Image2$ = "###,###.##"
  
  REDIM Dash(1) AS STRING * 79
  
  REDIM EMP2Rec(1)     AS EmpData2Type
  
  REDIM Unit(1)        AS UnitFileRecType
  REDIM Pg(1)          AS STRING * 3
  REDIM ENumb(1)    AS STRING * 11
  REDIM EName(1)    AS STRING * 19
  REDIM VBal(1)    AS STRING * 13
  REDIM SBal(1)    AS STRING * 13
  REDIM CBal(1)    AS STRING * 13
  REDIM LTbl(1)    AS STRING * 6
  
  MaxLines = 55
  LineCnt = 0
  Dash(1) = STRING$(77, "-") + CrLf$
  
  EmpRecSize = LEN(EMP2Rec(1))
  
  IdxRecLen = 2
  IdxFileSize& = FileSize(EmpIdxNName)
  NumOfRecs = IdxFileSize& \ IdxRecLen
  
  REDIM IdxBuff(1 TO NumOfRecs)
  FGetAH EmpIdxNName, IdxBuff(1), IdxRecLen, NumOfRecs
  
  FGetAH UnitFileName, Unit(1), LEN(Unit(1)), 1
  'Unit(1).UFEMPR
  RptTitle$ = "Benefits Accrual Report"
  
  OKFlag = False

  BlockClear
  CursorOff
  ShowProcessingScrn RptTitle$
  
  RptName$ = "PRRPTS\BENEACCR.RPT"
  LastEmp& = 999999999
  FirstEmp& = 0
  
  FCreate RptName$
  FOpenS RptName$, RHandle
  RPTSetupPRN 8, RHandle
  FOpenS EmpData2Name, DHandle  'open employee data file
  
  GOSUB PrintBenefitHeader
  
  FOR RecNo = 1 TO NumOfRecs
    FGetRTA DHandle, EMP2Rec(1), CLNG(IdxBuff(RecNo)), EmpRecSize
    IF EMP2Rec(1).Deleted OR EMP2Rec(1).EMPTDATE > -32767 THEN
      GOTO SkipEmBene
    END IF
    GOSUB PrintEmpBalance
    IF (LineCnt > MaxLines) AND RecNo < NumOfRecs THEN          'bottom of page?
      FPut RHandle, FF$         'yes; form feed
      GOSUB PrintBenefitHeader  'write title lines
    END IF
SkipEmBene:
    ShowPctComp RecNo, NumOfRecs
  NEXT
  
  FPut RHandle, FF$             'yes; form feed
  GOSUB PrintBenefitFooter
  FPut RHandle, FF$             'yes; form feed
  
  RPTSetupPRN 0, RHandle
  FClose DHandle
  FClose RHandle
  
  ERASE Dash, EMP2Rec, Unit, Pg, IdxBuff
  ERASE ENumb, EName, VBal, SBal, CBal, LTbl
  
  PrintRptFile RptTitle$, RptName$, 1, RetCode, 0

  EXIT SUB
  
PrintBenefitHeader:
  Page = Page + 1
  RSET Pg(1) = STR$(Page)
  UTemp$ = SPACE$(71)
  LSET UTemp$ = Unit(1).UFEMPR
  MID$(UTemp$, 62) = "Page:" + Pg(1) + CrLf$
  FPut RHandle, UTemp$
  
  FPut RHandle, "Benefits Accrual Report" + CrLf$
  FPut RHandle, "Report Date: " + DATE$ + CrLf$
  FPut RHandle, "Emp #      Name                    Vac Hrs.    Sick Hrs.    Comp Hrs.  Table" + CrLf$
  FPut RHandle, "                                    Balance      Balance      Balance       " + CrLf$
  
  FPut RHandle, Dash(1)
  LineCnt = 6
  RETURN
  
PrintBenefitFooter:
  Page = Page + 1
  RSET Pg(1) = STR$(Page)
  UTemp$ = SPACE$(71)
  LSET UTemp$ = Unit(1).UFEMPR
  MID$(UTemp$, 62) = "Page:" + Pg(1) + CrLf$
  FPut RHandle, UTemp$
  
  FPut RHandle, "Benefits Accrual Report" + CrLf$
  FPut RHandle, "Report Date: " + DATE$ + CrLf$
  FPut RHandle, "Report Totals.                     Vac Hrs.    Sick Hrs.    Comp Hrs." + CrLf$
  FPut RHandle, "                                    Balance      Balance      Balance" + CrLf$
  FPut RHandle, Dash(1)
  
  LSET ENumb(1) = ""
  LSET EName(1) = ""
  RSET VBal(1) = FUsing(STR$(TotalVacHrs#), Image2$)
  RSET SBal(1) = FUsing(STR$(TotalSicHrs#), Image2$)
  RSET CBal(1) = FUsing(STR$(TotalCmpHrs#), Image2$)
  RSET LTbl(1) = ""
  FPut RHandle, ENumb(1) + EName(1) + VBal(1) + SBal(1) + CBal(1) + LTbl(1) + CrLf$
  
  RSET VBal(1) = FUsing(STR$(TotalVacBal#), Image2$)
  RSET SBal(1) = FUsing(STR$(TotalSicBal#), Image2$)
  RSET CBal(1) = FUsing(STR$(TotalCmpBal#), Image2$)
  
  FPut RHandle, ENumb(1) + EName(1) + VBal(1) + SBal(1) + CBal(1) + LTbl(1) + CrLf$
  
  RETURN

PrintEmpBalance:
  IF EMP2Rec(1).EMPCTBAL < MinHrs THEN
    EMP2Rec(1).EMPCTBAL = 0
  END IF
  IF EMP2Rec(1).EMPVBAL < MinHrs THEN
    EMP2Rec(1).EMPVBAL = 0
  END IF
  IF EMP2Rec(1).EMPSLBAL < MinHrs THEN
    EMP2Rec(1).EMPSLBAL = 0
  END IF
  
  SELECT CASE UCASE$(QPTrim$(EMP2Rec(1).EMPPFREQ))
  CASE "WEEKLY"
    PayFreq = 52
  CASE "BI-WEEKLY"
    PayFreq = 26
  CASE "SEMI-MONTHLY"
    PayFreq = 24
  CASE "MONTHLY"
    PayFreq = 12
  CASE "QUARTERLY"
    PayFreq = 4
  CASE "SEMI-ANNUALLY"
    PayFreq = 2
  CASE "ANNUALLY"
    PayFreq = 1
  END SELECT
  
  SELECT CASE LEFT$(UCASE$(QPTrim$(EMP2Rec(1).EMPPTYPE)), 1)
  CASE "S"
    HrWage# = Round#((EMP2Rec(1).EMPPRATE * PayFreq) / 2048)
  CASE "H"
    HrWage# = EMP2Rec(1).EMPPRATE
  END SELECT
  
  LSET ENumb(1) = QPTrim$(EMP2Rec(1).EmpNo)

  'F INSTR(ENumb(1), "56") THEN STOP

  LSET EName(1) = QPTrim$(EMP2Rec(1).EMPLNAME) + ", " + QPTrim$(EMP2Rec(1).EMPFNAME)
  TotalVacHrs# = Round#(TotalVacHrs# + EMP2Rec(1).EMPVBAL)
  TotalSicHrs# = Round#(TotalSicHrs# + EMP2Rec(1).EMPSLBAL)
  TotalCmpHrs# = Round#(TotalCmpHrs# + EMP2Rec(1).EMPCTBAL)
  
  RSET VBal(1) = FUsing(STR$(EMP2Rec(1).EMPVBAL), Image2$)
  RSET SBal(1) = FUsing(STR$(EMP2Rec(1).EMPSLBAL), Image2$)
  RSET CBal(1) = FUsing(STR$(EMP2Rec(1).EMPCTBAL), Image2$)
  RSET LTbl(1) = STR$(EMP2Rec(1).LeaveTbl)
  
  FPut RHandle, ENumb(1) + EName(1) + VBal(1) + SBal(1) + CBal(1) + LTbl(1) + CrLf$
  TotalVacBal# = Round#(TotalVacBal# + Round#(HrWage# * EMP2Rec(1).EMPVBAL))
  TotalSicBal# = Round#(TotalSicBal# + Round#(HrWage# * EMP2Rec(1).EMPSLBAL))
  TotalCmpBal# = Round#(TotalCmpBal# + Round#(HrWage# * EMP2Rec(1).EMPCTBAL))
  LSET ENumb(1) = ""
  LSET EName(1) = ""

  RSET VBal(1) = FUsing(STR$(Round#(HrWage# * EMP2Rec(1).EMPVBAL)), Image2$)
  RSET SBal(1) = FUsing(STR$(Round#(HrWage# * EMP2Rec(1).EMPSLBAL)), Image2$)
  RSET CBal(1) = FUsing(STR$(Round#(HrWage# * EMP2Rec(1).EMPCTBAL)), Image2$)
  RSET LTbl(1) = ""
  FPut RHandle, ENumb(1) + EName(1) + VBal(1) + SBal(1) + CBal(1) + LTbl(1) + CrLf$
  FPut RHandle, CrLf$
  LineCnt = LineCnt + 3
RETURN
  
END SUB

SUB GrossWageReport
  
  REDIM TempScrn(1)
  
  ExitFlag = False
  FirstTime = True
  
  OKFlag = False
  
  REDIM Frm(1) AS FormInfo
  
  FormName$ = "GWAGERPT"
  NumFlds = LibNumberOfFields(MiscQLib, FormName$)
  REDIM Form$(NumFlds, 2)       'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo               'DIM the field information array
  StartEl = 0   'Load first form at array start
  LibGetFldDef MiscQLib, FormName$, StartEl, Fld(), Form$(), ErrCode
  
  
  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1
  
  '----- Setup TYPE for setting and reading form editing information.
  
  Frm(1).FldNo = 1              'Start editing on field #1
  Frm(1).InsStat = False        'Set insert state (True = Insert on)
  Frm(1).StartEl = 0            'Set form starting element to 0 and
  
  BlockClear
  DisplayMiscScrn FormName$
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      LSET Form$(1, 0) = "12-01-2000"
      LSET Form$(2, 0) = "12-31-2000"
      'Form$(2, 0) = DATE$
      Action = 2
    END IF
    
    SELECT CASE Frm(1).KeyCode
    CASE F0Key
      LowDate = Date2Num(Form$(1, 0))
      HiDate = Date2Num(Form$(2, 0))
      IF HiDate >= LowDate THEN
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
  
  CursorOff
  
  ERASE Form$, Fld, TempScrn, Frm
  
  IF ExitFlag THEN EXIT SUB
  Image1$ = ",####.##"
  Image2$ = ",######.##"
  Image3$ = "#######.##"
  
  REDIM Dash(1) AS STRING * 133
  REDIM TransHRec(1)   AS TransRecType
  REDIM EMP2Rec(1)     AS EmpData2Type
  REDIM Unit(1)        AS UnitFileRecType
  REDIM Pg(1)          AS STRING * 3
  
  REDIM GrsRpt(1 TO 3) AS GrossWageRptType
  
  REDIM EQWL1(1)       AS EQWRptRecType
  
  MaxLines = 57
  LineCnt = 0
  Dash(1) = STRING$(131, "-") + CrLf$
  
  EQWL1(1).CrLf = CrLf$
  EQWL1Len = LEN(EQWL1(1))
  
  EmpRecSize = LEN(EMP2Rec(1))
  TRecSize = LEN(TransHRec(1))
  
  IdxRecLen = 2
  IdxFileSize& = FileSize(EmpIdxLName)
  NumOfRecs = IdxFileSize& \ IdxRecLen
  
  REDIM IdxBuff(1 TO NumOfRecs)
  FGetAH EmpIdxLName, IdxBuff(1), IdxRecLen, NumOfRecs
  
  FGetAH UnitFileName, Unit(1), LEN(Unit(1)), 1
  
  RptTitle$ = "Employee Gross Wage Report"
  ShowProcessingScrn RptTitle$
  
  RptName$ = "PRRPTS\GROSWAGE.RPT"
  
  FCreate RptName$
  FOpenS RptName$, RHandle
  RPTSetupPRN 5, RHandle
  
  FOpenS TransHistFileName, THandle
  FOpenS EmpData2Name, DHandle  'open employee data file
  
  GOSUB PrintGrossHeader
  
  FOR RecNo = 1 TO NumOfRecs
    UsingThisOne = False
    FGetRTA DHandle, EMP2Rec(1), CLNG(IdxBuff(RecNo)), EmpRecSize
    'IF INSTR(EMP2Rec(1).EmpNo, "22") <= 0 THEN
    '  GOTO SkipEm4
    'END IF
    'IF IdxBuff(RecNo) <> 27 THEN GOTO SkipEm4

    IF EMP2Rec(1).LastTransRec <= 0 THEN
      GOTO SkipEm4
    END IF
    TransRecNum& = CLNG(EMP2Rec(1).LastTransRec)
    
    DO
      FGetRTA THandle, TransHRec(1), TransRecNum&, TRecSize
      
      SELECT CASE TransHRec(1).CheckDate
        
      CASE LowDate TO HiDate
        GOSUB SumEmpGross
        UsingThisOne = True
      CASE ELSE
      END SELECT
      IF TransHRec(1).PrevTransRec <= 0 THEN    'if no more transactions
        IF UsingThisOne THEN    'if there were any at all
          GOSUB PrintEmpGrossTotals             'then write emp totals
          IF LineCnt > MaxLines THEN            'bottom of page?
            FPut RHandle, FF$   'yes; form feed
            GOSUB PrintGrossHeader              'write title lines
          END IF
        END IF
        EXIT DO
      ELSE      'there are more trans to check
        TransRecNum& = CLNG(TransHRec(1).PrevTransRec)
      END IF
    LOOP
SkipEm4:
    ShowPctComp RecNo, NumOfRecs
  NEXT
  
  GOSUB PrintGrossGTotals
  
  RPTSetupPRN 0, RHandle
  
  FClose DHandle
  FClose THandle
  FClose RHandle
  
  ERASE Dash, TransHRec, EMP2Rec, Unit, Pg, GrsRpt, IdxBuff, EQWL1
  
  PrintRptFile RptTitle$, RptName$, 1, RetCode, 0
  
  
  EXIT SUB
  
PrintGrossHeader:
  Page = Page + 1
  RSET Pg(1) = QPStrI$(Page)
  UTemp$ = SPACE$(133)
  LSET UTemp$ = Unit(1).UFEMPR
  MID$(UTemp$, 124) = "Page:" + Pg(1) + CrLf$
  FPut RHandle, UTemp$
  FPut RHandle, "Gross Wage Report" + CrLf$
  FPut RHandle, Num2Date$(LowDate) + " thru " + Num2Date$(HiDate) + CrLf$
  'FPut RHandle, "Emp #      Name                   Gross Pay   FedGross     FedTax   SocGross     SocTax   MedGross     MedTax   StaGross     StaTax" + CrLf$
  FPut RHandle, "Emp #      Name                   Gross Pay   FedGross     FedTax   SocGross     SocTax   MedGross     MedTax        EIC     StaTax" + CrLf$
  FPut RHandle, Dash(1)
  LineCnt = 5
  RETURN
  
SumEmpGross:
    
  GrsRpt(1).GrossPay = Round(GrsRpt(1).GrossPay + TransHRec(1).GrossPay)
  GrsRpt(1).FedGross = Round(GrsRpt(1).FedGross + TransHRec(1).FedGrossPay)
  GrsRpt(1).FedTax = Round(GrsRpt(1).FedTax + TransHRec(1).FedTaxAmt)
  GrsRpt(1).SocGross = Round(GrsRpt(1).SocGross + TransHRec(1).SocGrossPay)
  GrsRpt(1).SocTax = Round(GrsRpt(1).SocTax + TransHRec(1).SocTaxAmt)
  GrsRpt(1).MedGross = Round(GrsRpt(1).MedGross + TransHRec(1).MedGrossPay)
  GrsRpt(1).MedTax = Round(GrsRpt(1).MedTax + TransHRec(1).MedTaxAmt)
  GrsRpt(1).EIC = Round(GrsRpt(1).EIC + TransHRec(1).EICAmt)
  GrsRpt(1).StaTax = Round(GrsRpt(1).StaTax + TransHRec(1).StaTaxAmt)
  IF TransHRec(1).TaxFring > 0 THEN
    GrsRpt(1).GrossPay = Round(GrsRpt(1).GrossPay + TransHRec(1).TaxFring)
    GrsRpt(1).FedGross = Round(GrsRpt(1).FedGross + TransHRec(1).TaxFring)
    GrsRpt(1).SocGross = Round(GrsRpt(1).SocGross + TransHRec(1).TaxFring)
    GrsRpt(1).MedGross = Round(GrsRpt(1).MedGross + TransHRec(1).TaxFring)
  END IF
  RETURN
  
PrintEmpGrossTotals:
  LSET EQWL1(1).ENumb = QPTrim$(EMP2Rec(1).EmpNo)
  LSET EQWL1(1).EName = QPTrim$(EMP2Rec(1).EMPLNAME) + ", " + QPTrim$(EMP2Rec(1).EMPFNAME)
  RSET EQWL1(1).GPay = FUsing(STR$(GrsRpt(1).GrossPay), Image2$)
  RSET EQWL1(1).FedGr = FUsing(STR$(GrsRpt(1).FedGross), Image2$)
  RSET EQWL1(1).FedTx = FUsing(STR$(GrsRpt(1).FedTax), Image2$)
  RSET EQWL1(1).SocGr = FUsing(STR$(GrsRpt(1).SocGross), Image2$)
  RSET EQWL1(1).SocTx = FUsing(STR$(GrsRpt(1).SocTax), Image2$)
  RSET EQWL1(1).MedGr = FUsing(STR$(GrsRpt(1).MedGross), Image2$)
  RSET EQWL1(1).MedTx = FUsing(STR$(GrsRpt(1).MedTax), Image2$)
  RSET EQWL1(1).EIC = FUsing(STR$(GrsRpt(1).EIC), Image2$)
  RSET EQWL1(1).StaTx = FUsing(STR$(GrsRpt(1).StaTax), Image2$)
  
  FPutA RHandle, EQWL1(1), EQWL1Len
  
  LineCnt = LineCnt + 1
  
  'sum to quarter grand totals
  GrsRpt(3).GrossPay = Round(GrsRpt(3).GrossPay + GrsRpt(1).GrossPay)
  GrsRpt(3).FedGross = Round(GrsRpt(3).FedGross + GrsRpt(1).FedGross)
  GrsRpt(3).FedTax = Round(GrsRpt(3).FedTax + GrsRpt(1).FedTax)
  GrsRpt(3).SocGross = Round(GrsRpt(3).SocGross + GrsRpt(1).SocGross)
  GrsRpt(3).SocTax = Round(GrsRpt(3).SocTax + GrsRpt(1).SocTax)
  GrsRpt(3).MedGross = Round(GrsRpt(3).MedGross + GrsRpt(1).MedGross)
  GrsRpt(3).MedTax = Round(GrsRpt(3).MedTax + GrsRpt(1).MedTax)
  GrsRpt(3).EIC = Round(GrsRpt(3).EIC + GrsRpt(1).EIC)
  GrsRpt(3).StaTax = Round(GrsRpt(3).StaTax + GrsRpt(1).StaTax)
  
  'clear for the next transaction record
  GrsRpt(1) = GrsRpt(2)
  
  RETURN
  
PrintGrossGTotals:
  LSET EQWL1(1).ENumb = ""
  LSET EQWL1(1).EName = "        Totals:"
  RSET EQWL1(1).GPay = FUsing(STR$(GrsRpt(3).GrossPay), Image3$)
  RSET EQWL1(1).FedGr = FUsing(STR$(GrsRpt(3).FedGross), Image3$)
  RSET EQWL1(1).FedTx = FUsing(STR$(GrsRpt(3).FedTax), Image3$)
  RSET EQWL1(1).SocGr = FUsing(STR$(GrsRpt(3).SocGross), Image3$)
  RSET EQWL1(1).SocTx = FUsing(STR$(GrsRpt(3).SocTax), Image3$)
  RSET EQWL1(1).MedGr = FUsing(STR$(GrsRpt(3).MedGross), Image3$)
  RSET EQWL1(1).MedTx = FUsing(STR$(GrsRpt(3).MedTax), Image3$)
  RSET EQWL1(1).EIC = FUsing(STR$(GrsRpt(3).EIC), Image3$)
  RSET EQWL1(1).StaTx = FUsing(STR$(GrsRpt(3).StaTax), Image3$)
  
  FPut RHandle, Dash(1)
  FPutA RHandle, EQWL1(1), EQWL1Len
  FPut RHandle, FF$
  
  RETURN
  
END SUB

SUB PayDedReport
  
  SHARED Choice$()
  
  ExitFlag = False
  
  REDIM TempScrn(1)
  REDIM Choice$(12, 1)          '
  
  REDIM DedCodes(1 TO 12) AS DedCodeRecType
  FGetAH DedCodeFileName, DedCodes(1), LEN(DedCodes(1)), 12
  
  Choice$(0, 1) = "3"
  
  FOR Cnt = 1 TO 12
    Choice$(Cnt, 1) = QPTrim$(DedCodes(Cnt).DCDESC1)
  NEXT
  
  FirstTime = True
  REDIM Frm(1) AS FormInfo
  
  FormName$ = "PRDEDRPT"
  NumFlds = LibNumberOfFields(MiscQLib, FormName$)
  REDIM Form$(NumFlds, 2)       'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo               'DIM the field information array
  StartEl = 0   'Load first form at array start
  LibGetFldDef MiscQLib, FormName$, StartEl, Fld(), Form$(), ErrCode
  
  
  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1
  
  '----- Setup TYPE for setting and reading form editing information.
  
  Frm(1).FldNo = 1              'Start editing on field #1
  Frm(1).InsStat = False        'Set insert state (True = Insert on)
  Frm(1).StartEl = 0            'Set form starting element to 0 and
  
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
      TempDed$ = QPTrim$(Form$(3, 0))
      FOR Cnt = 1 TO 12
        IF TempDed$ = QPTrim$(DedCodes(Cnt).DCDESC1) THEN
          DNum = Cnt
          EXIT FOR
        END IF
      NEXT
      
      'process invalid report parms
      IF HighDate >= LowDate AND DNum THEN
        OKFlag = True
      ELSE
        CursorOff
        SaveScrn TempScrn()
        DisplayMiscScrn ReportParmError
        WaitForAction
        RestScrn TempScrn()
        Action = 2
        IF DNum THEN
          Frm(1).FldNo = 1
        ELSE
          Frm(1).FldNo = 3
        END IF
        
      END IF
      '-=-=-=
    CASE EscKey
      OKFlag = True
      ExitFlag = True
    END SELECT
    
  LOOP UNTIL OKFlag             'proper key not set
  
  CursorOff
  
  ERASE Form$, Fld, Frm
  
  IF ExitFlag THEN EXIT SUB
  
  '  FF$ = CHR$(12)
  
  Image1$ = "$$,####.##"
  Image2$ = "$$###,###.##"
  
  REDIM TransHRec(1)     AS TransRecType
  REDIM EMP2Rec(1)  AS EmpData2Type
  REDIM Unit(1)      AS UnitFileRecType
  REDIM Pg(1)       AS STRING * 3
  REDIM DedRpt(1)  AS DedRptType
  
  MaxLines = 55
  EmpRecSize = LEN(EMP2Rec(1))
  TRecSize = LEN(TransHRec(1))
  
  DedLineLen = LEN(DedRpt(1))
  DedRpt(1).CrLf = CrLf$
  
  FGetAH UnitFileName, Unit(1), LEN(Unit(1)), 1
  
  IdxRecLen = 2
  
  IdxFileSize& = FileSize(EmpIdxLName)
  NumOfRecs = IdxFileSize& \ IdxRecLen
  
  REDIM IdxBuff(1 TO NumOfRecs)
  FGetAH EmpIdxLName, IdxBuff(1), IdxRecLen, NumOfRecs
  
  RptTitle$ = "Payroll Deduction Report"
  ShowProcessingScrn RptTitle$
  
  RptName$ = "PRRPTS\DEDUCT" + QPStrI$(DNum) + ".RPT"
  RptDate$ = "Start Date: " + Num2Date(LowDate) + "  Ending Date: " + Num2Date(HighDate) + CrLf$
  RptDed$ = QPTrim$(DedCodes(DNum).DCDESC1) + " Deduction Report." + CrLf$
  
  FCreate RptName$
  FOpenS RptName$, RHandle
  RPTSetupPRN 6, RHandle
  
  FOpenS TransHistFileName, THandle
  FOpenS EmpData2Name, DHandle  'open employee data file
  
  GOSUB DedHeader
  
  FOR RecNo = 1 TO NumOfRecs
    UsingThisOne = False
    DAmt# = 0
    FGetRTA DHandle, EMP2Rec(1), CLNG(IdxBuff(RecNo)), EmpRecSize
    IF EMP2Rec(1).EMPTDATE > 0 THEN
      IF LowDate > EMP2Rec(1).EMPTDATE THEN
        GOTO SkipEm2
      END IF
    END IF
    IF EMP2Rec(1).LastTransRec <= 0 THEN
      GOTO SkipEm2
    END IF
    
    TransRecNum& = CLNG(EMP2Rec(1).LastTransRec)
    
    DO
      FGetRTA THandle, TransHRec(1), TransRecNum&, TRecSize
      SELECT CASE TransHRec(1).CheckDate
      CASE LowDate TO HighDate
        IF TransHRec(1).DAmt(DNum) <> 0 THEN
          DAmt# = Round(DAmt# + TransHRec(1).DAmt(DNum))
          UsingThisOne = True
        END IF
      CASE ELSE
      END SELECT
      IF TransHRec(1).PrevTransRec <= 0 THEN
        IF UsingThisOne THEN
          GOSUB PrintThisOne2
          IF LineCnt > MaxLines THEN
            FPut RHandle, FF$
            GOSUB DedHeader
          END IF
        END IF
        EXIT DO
      ELSE
        TransRecNum& = CLNG(TransHRec(1).PrevTransRec)
      END IF
    LOOP
SkipEm2:
    ShowPctComp RecNo, NumOfRecs
  NEXT
  
  GOSUB GrandTots
  
  RPTSetupPRN 0, RHandle
  
  FClose DHandle
  FClose THandle
  FClose RHandle
  
  ERASE TransHRec, EMP2Rec, Unit, Pg, DedRpt, IdxBuff, DedCodes
  
  PrintRptFile RptTitle$, RptName$, 1, RetCode, 0
  
  EXIT SUB
  
  
PrintThisOne2:
  TotalDAmt# = Round(TotalDAmt# + DAmt#)
  DedRpt(1).EmpName = LTRIM$(RTRIM$(EMP2Rec(1).EMPFNAME)) + " " + LTRIM$(RTRIM$(EMP2Rec(1).EMPLNAME))
  DedRpt(1).SSN = LEFT$(EMP2Rec(1).EMPSSN, 3) + "-" + MID$(EMP2Rec(1).EMPSSN, 4, 2) + "-" + MID$(EMP2Rec(1).EMPSSN, 6, 4)
  RSET DedRpt(1).DAmt = LTRIM$(RTRIM$(FUsing(STR$(DAmt#), Image1$)))
  FPutA RHandle, DedRpt(1), DedLineLen
  LineCnt = LineCnt + 1
  RETURN
  
DedHeader:
  Page = Page + 1
  RSET Pg(1) = STR$(Page)
  UTemp$ = SPACE$(71)
  LSET UTemp$ = Unit(1).UFEMPR
  MID$(UTemp$, 62) = "Page:" + Pg(1) + CrLf$
  FPut RHandle, UTemp$
  FPut RHandle, RptDed$
  FPut RHandle, RptDate$
  FPut RHandle, "Employee Name                       SS#                   Ded. Amount" + CrLf$
  FPut RHandle, CrLf$
  LineCnt = 5
  
  RETURN
  
GrandTots:
  LSET DedRpt(1).EmpName = "Total: "
  LSET DedRpt(1).SSN = ""
  RSET DedRpt(1).DAmt = QPTrim$(FUsing(STR$(TotalDAmt#), Image2$))
  FPut RHandle, CrLf$
  FPutA RHandle, DedRpt(1), DedLineLen
  FPut RHandle, FF$
  RETURN
  
  
END SUB

SUB RetirementRpt
  
  SHARED PrnDef$()
  
  FirstTime = True
  
  REDIM TempScrn(1)
  REDIM Frm(1) AS FormInfo
  
  FormName$ = "ERETIRPT"
  NumFlds = LibNumberOfFields(MiscQLib, FormName$)
  REDIM Form$(NumFlds, 2)       'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo               'DIM the field information array
  StartEl = 0   'Load first form at array start
  LibGetFldDef MiscQLib, FormName$, StartEl, Fld(), Form$(), ErrCode
  
  
  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1
  
  'MaxLines = 60
  '----- Setup TYPE for setting and reading form editing information.
  
  Frm(1).FldNo = 1              'Start editing on field #1
  Frm(1).InsStat = False        'Set insert state (True = Insert on)
  Frm(1).StartEl = 0            'Set form starting element to 0 and
  
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
        YearNum = QPValI(RIGHT$(Form$(2, 0), 4))
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
  
  CursorOff
  
  
  ERASE Form$, Fld, TempScrn, Frm
  
  
  IF ExitFlag THEN EXIT SUB
  
  CursorOff
  
  RptName$ = "PRRPTS\RETIRE.RPT"
  
  '  FF$ = CHR$(12)
  
  'Image1$ = "$$,####.##"
  'Image2$ = "$$###,###.##"
  
  REDIM TransHRec(1)   AS TransRecType
  REDIM EMP2Rec(1)     AS EmpData2Type
  REDIM Unit(1)        AS UnitFileRecType
  REDIM Pg(1)          AS STRING * 3
  REDIM Dash(1)        AS STRING * 91
  
  Dash(1) = STRING$(91, "-")
  
  EmpRecSize = LEN(EMP2Rec(1))
  TRecSize = LEN(TransHRec(1))
  
  IdxRecLen = 2
  
  IdxFileSize& = FileSize(EmpIdxLName)
  NumOfRecs = IdxFileSize& \ IdxRecLen
  
  REDIM IdxBuff(1 TO NumOfRecs)
  FGetAH EmpIdxLName, IdxBuff(1), IdxRecLen, NumOfRecs
  FGetAH UnitFileName, Unit(1), LEN(Unit(1)), 1
  MaxLines = 50
  
  '-=-=-=-=-=-=-=-=-=-=-=
  
  RptTitle$ = "Employee Retirement Report."
  ShowProcessingScrn RptTitle$
  
  TCol = 40 - (LEN(RptTitle$) \ 2) + 1
  PctRow = 11
  
  QPrintRC "Pass 1 of 2", PctRow, TCol, 112
  
  FCreate RptName$
  FOpenS RptName$, RHandle
  RPTSetupPRN 11, RHandle
  FClose RHandle
  
  OPEN RptName$ FOR APPEND AS #1
  
  FOpenS TransHistFileName, THandle
  FOpenS EmpData2Name, DHandle  'open employee data file
  
  GOSUB LRetRptHeader
  
  FOR RecNo = 1 TO NumOfRecs
    UsingThisOne = False
    RetWage# = 0
    RETAMT# = 0
    MatAmt# = 0
    MatchCnt = 0
    FGetRTA DHandle, EMP2Rec(1), CLNG(IdxBuff(RecNo)), EmpRecSize
    
    IF EMP2Rec(1).LastTransRec <= 0 THEN
      GOTO SkipEm6
    END IF
    
    TransRecNum& = CLNG(EMP2Rec(1).LastTransRec)
    
    DO
      FGetRTA THandle, TransHRec(1), TransRecNum&, TRecSize
      SELECT CASE TransHRec(1).CheckDate
      CASE LowDate TO HighDate
        'IF TransHRec(1).CheckNum < 21681 OR TransHRec(1).CheckNum > 21684 THEN
        IF UCASE$(LEFT$(LTRIM$(EMP2Rec(1).EMPRETTP), 1)) = "L" THEN
          'IF TransHRec(1).GrossPay <= 0 THEN STOP
          RetWage# = Round(RetWage# + TransHRec(1).RetGrossPay)
          '                                   was .GrossPay
          RETAMT# = Round(RETAMT# + TransHRec(1).RetireAmt)
          MatAmt# = Round(MatAmt# + TransHRec(1).MatchRetAmt)
          MatchCnt = MatchCnt + 1
          UsingThisOne = True
        END IF
        
        'ELSE
        '  STOP
        'END IF
      CASE ELSE
      END SELECT
      
      IF TransHRec(1).PrevTransRec <= 0 THEN
        IF UsingThisOne THEN
          GOSUB PrintEmpRetLine
          IF LineCnt >= MaxLines THEN
            PRINT #1, FF$
            GOSUB LRetRptHeader
          END IF
        END IF
        EXIT DO
      ELSE
        TransRecNum& = CLNG(TransHRec(1).PrevTransRec)
      END IF
    LOOP
    
SkipEm6:
    ShowPctComp RecNo, NumOfRecs
  NEXT
  GOSUB RetLTotals
  
  'start of PASS 2
  
  QPrintRC "Pass 2 of 2", PctRow, TCol, 112
  
  GTotal# = 0
  RTotal# = 0
  MTotal# = 0
  GOSUB GRetRptHeader
  
  FOR RecNo = 1 TO NumOfRecs
    UsingThisOne = False
    RetWage# = 0
    RETAMT# = 0
    MatAmt# = 0
    FGetRTA DHandle, EMP2Rec(1), CLNG(IdxBuff(RecNo)), EmpRecSize
    
    IF EMP2Rec(1).LastTransRec <= 0 THEN
      GOTO SkipEm7
    END IF
    
    TransRecNum& = CLNG(EMP2Rec(1).LastTransRec)
    
    DO
      FGetRTA THandle, TransHRec(1), TransRecNum&, TRecSize
      SELECT CASE TransHRec(1).CheckDate
      CASE LowDate TO HighDate
        'IF TransHRec(1).CheckNum < 21681 OR TransHRec(1).CheckNum > 21684 THEN
        'IF TransHRec(1).GrossPay <= 0 THEN STOP
        IF UCASE$(LEFT$(LTRIM$(EMP2Rec(1).EMPRETTP), 1)) = "G" THEN
          RetWage# = Round(RetWage# + TransHRec(1).RetGrossPay)
          '                                   was .GrossPay
          RETAMT# = Round(RETAMT# + TransHRec(1).RetireAmt)
          MatAmt# = Round(MatAmt# + TransHRec(1).MatchRetAmt)
          UsingThisOne = True
        END IF
        'END IF
      CASE ELSE
      END SELECT
      
      IF TransHRec(1).PrevTransRec <= 0 THEN
        IF UsingThisOne THEN
          GOSUB PrintEmpRetLine
          IF LineCnt >= MaxLines THEN
            PRINT #1, FF$
            GOSUB GRetRptHeader
          END IF
        END IF
        EXIT DO
      ELSE
        TransRecNum& = CLNG(TransHRec(1).PrevTransRec)
      END IF
    LOOP
    
SkipEm7:
    ShowPctComp RecNo, NumOfRecs
  NEXT
  
  GOSUB RetLTotals
  PRINT #1, PrnDef$(1);
  
  FClose THandle
  FClose DHandle                'open employee data file
  'FClose RHandle                'open employee data file
  
  CLOSE
  
  ERASE TransHRec, EMP2Rec, Unit, Pg, Dash, IdxBuff
  
  PrintRptFile RptTitle$, RptName$, 1, RetCode, 0
  
  EXIT SUB
  
  
PrintEmpRetLine:
  PRINT #1, LEFT$(EMP2Rec(1).EMPSSN, 3); "-"; MID$(EMP2Rec(1).EMPSSN, 4, 2); "-"; MID$(EMP2Rec(1).EMPSSN, 6, 4);
  PRINT #1, TAB(14); LEFT$(EMP2Rec(1).EMPRETNO, 12);
  PRINT #1, TAB(28); QPTrim$(EMP2Rec(1).EMPLNAME); ", "; QPTrim$(EMP2Rec(1).EMPFNAME);
  PRINT #1, TAB(55); USING "#####,.##"; RetWage#; TAB(70); RETAMT#;
  PRINT #1, TAB(83); USING "#####,.##"; MatAmt#
  LineCnt = LineCnt + 1         'employeesprinted = employeesprinted + 1
  GTotal# = Round(GTotal# + RetWage#)
  RTotal# = Round(RTotal# + RETAMT#)
  MTotal# = Round(MTotal# + MatAmt#)
  RETURN
  
RetLTotals:
  PRINT #1, Dash(1)
  PRINT #1, TAB(28); "Totals:";
  PRINT #1, TAB(53); USING "#######,.##"; GTotal#; TAB(68); RTotal#;
  PRINT #1, TAB(81); USING "#######,.##"; MTotal#
  PRINT #1, FF$
  RETURN
  
LRetRptHeader:
  LPage = LPage + 1
  PRINT #1, "Department of State Treasurer"; TAB(70); "     Page:"; LPage
  PRINT #1, "North Carolina Law Enforcement Officers Benefit and Retirement Fund"
  PRINT #1, "Monthly Report of Subject Wages And Retirement Contributions"
  PRINT #1,
  PRINT #1, "Reporting Unit: "; QPTrim$(Unit(1).UFEMPR); TAB(70); "Unit Code: "; Unit(1).UFRETIDL
  PRINT #1, "Month: "; MonthName$(MonthNum); " "; YearNum
  PRINT #1,
  PRINT #1, "                                                   Wages Subject    Retirement     Employer"
  PRINT #1, "Soc Sec #    Ret #         Employee Name           To Retirem't     Deductions        Match"
  PRINT #1, Dash(1)
  LineCnt = 10
  RETURN
  
GRetRptHeader:
  GPage = GPage + 1
  PRINT #1, "Department of State Treasurer"; TAB(70); "     Page:"; GPage
  PRINT #1, "North Carolina Local Government Retirement System."
  PRINT #1, "Monthly Report of Subject Wages And Retirement Contributions"
  PRINT #1,
  PRINT #1, "Reporting Unit: "; QPTrim$(Unit(1).UFEMPR); TAB(70); "Unit Code: "; Unit(1).UFRETID
  PRINT #1, "Month: "; MonthName$(MonthNum); " "; YearNum
  PRINT #1,
  PRINT #1, "                                                   Wages Subject    Retirement     Employer"
  PRINT #1, "Soc Sec #    Ret #         Employee Name           To Retirem't     Deductions        Match"
  PRINT #1, Dash(1)
  LineCnt = 10
  RETURN
  
END SUB

SUB SupRetReport
  
  SHARED Choice$(), PrnDef$()
  
  REDIM TempScrn(1)
  
  REDIM K401Totals(1)  AS K401RptType
  
  REDIM Choice$(12, 1)          '
  REDIM Text$(1 TO 5)
  
  REDIM DedCodes(1 TO 12) AS DedCodeRecType
  FGetAH DedCodeFileName, DedCodes(1), LEN(DedCodes(1)), 12
  
  REDIM Unit(1)              AS UnitFileRecType
  FGetAH UnitFileName, Unit(1), LEN(Unit(1)), 1

  'Unit(1).UFEMPR

  IF INSTR(UCASE$(Unit(1).UFEMPR), "VASS") > 0 THEN
    VassFlag = True
  END IF

  Choice$(0, 1) = "3,4"
  
  FOR Cnt = 1 TO 12
    Choice$(Cnt, 1) = QPTrim$(DedCodes(Cnt).DCDESC1)
  NEXT
  
  FirstTime = True
  REDIM Frm(1) AS FormInfo
  
  FormName$ = "SRETIRPT"
  NumFlds = LibNumberOfFields(MiscQLib, FormName$)
  REDIM Form$(NumFlds, 2)       'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo               'DIM the field information array
  StartEl = 0   'Load first form at array start
  LibGetFldDef MiscQLib, FormName$, StartEl, Fld(), Form$(), ErrCode
  
  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1
  '----- Setup TYPE for setting and reading form editing information.
  
  Frm(1).FldNo = 1              'Start editing on field #1
  Frm(1).InsStat = False        'Set insert state (True = Insert on)
  Frm(1).StartEl = 0            'Set form starting element to 0 and
  
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
      GPct# = VAL(Form$(5, 0))
      LPct# = VAL(Form$(6, 0))
      TempDed$ = QPTrim$(Form$(3, 0))
      FOR Cnt = 1 TO 12
        IF TempDed$ = QPTrim$(DedCodes(Cnt).DCDESC1) THEN
          VCodeNum = Cnt
          EXIT FOR
        END IF
      NEXT
      TempDed$ = QPTrim$(Form$(4, 0))
      FOR Cnt = 1 TO 12
        IF TempDed$ = QPTrim$(DedCodes(Cnt).DCDESC1) THEN
          LCodeNum = Cnt
          EXIT FOR
        END IF
      NEXT
      'IF (HighDate >= LowDate) AND LCodeNum > 0 AND VCodeNum > 0 AND GPct# > 0 THEN
      IF (HighDate >= LowDate) THEN             'AND GPct# > 0 THEN
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
  
  CursorOff
  
  ERASE Form$, Fld, TempScrn, Frm
  
  IF ExitFlag THEN EXIT SUB
  
  RptName$ = "PRRPTS\401K.RPT"
  
  VCalcAmt# = 0: GVTotal# = 0
  LCalcAmt# = 0: GLTotal# = 0
  GCalcAmt# = 0:
  
  'OPEN "dodo.lst" FOR OUTPUT AS #10
  
  FF$ = CHR$(12)
  
  EPRTotal# = 0
  
  Image1$ = ",####.##"
  Image2$ = "###,###.##"
  
  REDIM TransHRec(1)     AS TransRecType
  REDIM EMP2Rec(1)  AS EmpData2Type
  REDIM K401Rec(1 TO 1)       AS K401RptType
  
  EmpRecSize = LEN(EMP2Rec(1))
  TRecSize = LEN(TransHRec(1))
  
  REDIM DedCodes(1 TO 12) AS DedCodeRecType
  FGetAH DedCodeFileName, DedCodes(1), LEN(DedCodes(1)), 12
  
  IdxRecLen = 2
  
  IdxFileSize& = FileSize(EmpIdxNName)
  NumOfRecs = IdxFileSize& \ IdxRecLen
  
  REDIM IdxBuff(1 TO NumOfRecs)
  FGetAH EmpIdxNName, IdxBuff(1), IdxRecLen, NumOfRecs
  
  RptTitle$ = "Supplemental Retirement Report."
  ShowProcessingScrn RptTitle$
  
  'got input here
  
  FCreate RptName$
  FOpenS RptName$, RHandle
  RPTSetupPRN 12, RHandle
  FClose RHandle
  
  OPEN RptName$ FOR APPEND AS #1
  
  FOpenS TransHistFileName, THandle
  FOpenS EmpData2Name, DHandle  'open employee data file
  
  'GOSUB K401Header
  
  FOR RecNo = 1 TO NumOfRecs
    UsingThisOne = False
    VCalcAmt# = 0
    LCalcAmt# = 0
    GCalcAmt# = 0
    FGetRTA DHandle, EMP2Rec(1), CLNG(IdxBuff(RecNo)), EmpRecSize
    IF EMP2Rec(1).EMPTDATE > 0 THEN
      IF LowDate > EMP2Rec(1).EMPTDATE THEN
        GOTO SkipEm
      END IF
    END IF
    IF INSTR(EMP2Rec(1).EMPLNAME, "CHEWNING, JR") THEN
      IF INSTR(EMP2Rec(1).EMPFNAME, "WILLIAM B.") THEN
        GOTO SkipEm
      END IF
    END IF
    IF EMP2Rec(1).LastTransRec <= 0 THEN
      GOTO SkipEm
    END IF
    
    TransRecNum& = CLNG(EMP2Rec(1).LastTransRec)
    
    DO
      FGetRTA THandle, TransHRec(1), TransRecNum&, TRecSize
      SELECT CASE TransHRec(1).CheckDate
      CASE LowDate TO HighDate
       IF VCodeNum > 0 THEN
          IF TransHRec(1).DAmt(VCodeNum) <> 0 THEN
            VCalcAmt# = Round(VCalcAmt# + TransHRec(1).DAmt(VCodeNum))
            UsingThisOne = True
          END IF
        END IF
        
        IF LCodeNum > 0 THEN
          IF TransHRec(1).DAmt(LCodeNum) > 0 THEN
            LCalcAmt# = Round(LCalcAmt# + TransHRec(1).DAmt(LCodeNum))
            UsingThisOne = True
          END IF
        END IF
        EmpRType$ = UCASE$(LEFT$(LTRIM$(EMP2Rec(1).EMPRETTP), 1))
        IF EmpRType$ = "L" OR EmpRType$ = "G" THEN
          GCalcAmt# = Round(GCalcAmt# + TransHRec(1).GrossPay)
          UsingThisOne = True
        END IF
      CASE ELSE
      END SELECT
      
      IF TransHRec(1).PrevTransRec <= 0 THEN
        IF UsingThisOne THEN
          IF EmpRType$ = "L" THEN
            EPct# = LPct#
            TVLawAmt# = Round(TVLawAmt# + VCalcAmt#)
            TLLawAmt# = Round(TLLawAmt# + LCalcAmt#)
            LawMatAmt# = Round((GCalcAmt# * EPct#) * .01)
            TotLawMatAmt# = Round(TotLawMatAmt# + LawMatAmt#)
          ELSE
            EPct# = GPct#
            TVGenAmt# = Round(TVGenAmt# + VCalcAmt#)
            TLGenAmt# = Round(TLGenAmt# + LCalcAmt#)
            GenMatAmt# = Round((GCalcAmt# * EPct#) * .01)
            TotGenMatAmt# = Round(TotGenMatAmt# + GenMatAmt#)
          END IF
          GOSUB PrintThisOne
        END IF
        EXIT DO
      ELSE
        TransRecNum& = CLNG(TransHRec(1).PrevTransRec)
      END IF
      
    LOOP
    
SkipEm:
    ShowPctComp RecNo, NumOfRecs
  NEXT
  
  IF EPrinted > 1 THEN
    SortT K401Rec(1), EPrinted, 0, LEN(K401Rec(1)), 0, 1
    Ret$ = K401Rec(1).RetType
    FOR Cnt = 2 TO EPrinted
      IF K401Rec(Cnt).RetType <> Ret$ THEN
        OffSet = Cnt - 1
        EXIT FOR
      END IF
    NEXT
    SortT K401Rec(1), OffSet, 0, LEN(K401Rec(1)), 1, 10
    SortT K401Rec(OffSet + 1), EPrinted - OffSet, 0, LEN(K401Rec(1)), 1, 10
  END IF
  
  
  GOSUB K401Header
  PRINT #1, "General"
  PRINT #1,
  FOR Cnt = 1 TO OffSet
    PRINT #1, K401Rec(Cnt).EmpName; K401Rec(Cnt).SSN; "   "; K401Rec(Cnt).VAmt; " "; K401Rec(Cnt).LAmt;
    PRINT #1, "  "; K401Rec(Cnt).MAmt; "    "; K401Rec(Cnt).HDate; " "; K401Rec(Cnt).RetType;
    IF VassFlag THEN
      PRINT #1, "  "; K401Rec(Cnt).Gross;
    END IF
    PRINT #1,

  NEXT
  PRINT #1,
  PRINT #1, "Law Enforcement"
  PRINT #1,
  FOR Cnt = OffSet + 1 TO EPrinted
    'PRINT #1, K401Rec(Cnt).EmpNum; "  "; K401Rec(Cnt).EmpName; K401Rec(Cnt).SSN; "   "; K401Rec(Cnt).VAmt; " "; K401Rec(Cnt).LAmt;
    PRINT #1, K401Rec(Cnt).EmpName; K401Rec(Cnt).SSN; "   "; K401Rec(Cnt).VAmt; " "; K401Rec(Cnt).LAmt;
    PRINT #1, "  "; K401Rec(Cnt).MAmt; "    "; K401Rec(Cnt).HDate; " "; K401Rec(Cnt).RetType;
    IF VassFlag THEN
      PRINT #1, "  "; K401Rec(Cnt).Gross;
    END IF
    PRINT #1,
  NEXT
  
  GOSUB PrintGrandTots
  
  FClose DHandle
  FClose THandle
  PRINT #1, PrnDef$(1);
  CLOSE
  
  ERASE TransHRec, EMP2Rec, K401Rec, IdxBuff, DedCodes
  
  PrintRptFile RptTitle$, RptName$, 1, RetCode, 0
  
  EXIT SUB
  
PrintThisOne:
  'EPrinted = EPrinted + 1
  IF EPct# > 0 OR VCalcAmt# > 0 OR LCalcAmt# > 0 THEN
    EPrinted = EPrinted + 1
    
    REDIM PRESERVE K401Rec(1 TO EPrinted)        AS K401RptType
    
    TotalVAmt# = Round(TotalVAmt# + VCalcAmt#)
    TotalLAmt# = Round(TotalLAmt# + LCalcAmt#)
    TMatchAmt# = Round((GCalcAmt# * EPct#) * .01)
    TotalGross# = Round(TotalGross# + GCalcAmt#)
    TotalMatchAmt# = Round(TotalMatchAmt# + TMatchAmt#)
    
    RSET K401Rec(EPrinted).EmpNum = EMP2Rec(1).EmpNo
    K401Rec(EPrinted).EmpName = LTRIM$(RTRIM$(EMP2Rec(1).EMPFNAME)) + " " + LTRIM$(RTRIM$(EMP2Rec(1).EMPLNAME))
    K401Rec(EPrinted).SSN = LEFT$(EMP2Rec(1).EMPSSN, 3) + "-" + MID$(EMP2Rec(1).EMPSSN, 4, 2) + "-" + MID$(EMP2Rec(1).EMPSSN, 6, 4)
    RSET K401Rec(EPrinted).VAmt = LTRIM$(RTRIM$(FUsing(STR$(VCalcAmt#), Image1$)))
    IF VCalcAmt# = 0 THEN
      RSET K401Rec(EPrinted).VAmt = ".00"
    END IF
    
    RSET K401Rec(EPrinted).LAmt = LTRIM$(RTRIM$(FUsing(STR$(LCalcAmt#), Image1$)))
    IF LCalcAmt# = 0 THEN
      RSET K401Rec(EPrinted).LAmt = ".00"
    END IF
    
    RSET K401Rec(EPrinted).MAmt = LTRIM$(RTRIM$(FUsing(STR$(TMatchAmt#), Image1$)))
    IF TMatchAmt# = 0 THEN
      RSET K401Rec(EPrinted).MAmt = ".00"
    END IF
    
    RSET K401Rec(EPrinted).Batch = QPTrim$(Unit(1).BBTBATCH)
    K401Rec(EPrinted).RetType = EmpRType$
    
    TempDate$ = Num2Date(HighDate)
    K401Rec(EPrinted).HDate = LEFT$(TempDate$, 6) + RIGHT$(TempDate$, 2)
    K401Rec(EPrinted).Gross = FUsing(STR$(GCalcAmt#), Image1$)
    
  END IF
RETURN
  
K401Header:
  TempDate$ = DATE$
  TempDate$ = LEFT$(TempDate$, 6) + RIGHT$(TempDate$, 2)
  Text$(1) = "Supplemental Retirement Income Plan of NC"
  Text$(2) = "BB&T Administrator"
  Text$(3) = "Payroll Center Name: " + QPTrim$(Unit(1).UFEMPR)
  Text$(4) = "Payroll Center Number: " + QPTrim$(Unit(1).BBTCNTNO)
  
  Text1Len = LEN(Text$(1))
  Text2Len = LEN(Text$(2))
  Text3Len = LEN(Text$(3))
  Text4Len = LEN(Text$(4))
  
  PRINT #1, SPACE$(48 - (Text1Len \ 2)); Text$(1); TAB(81); "Date: "; TempDate$
  PRINT #1, SPACE$(48 - (Text2Len \ 2)); Text$(2); TAB(81); "Page: 1 "
  PRINT #1, SPACE$(48 - (Text3Len \ 2)); Text$(3)
  PRINT #1, SPACE$(48 - (Text4Len \ 2)); Text$(4)
  PRINT #1,
  PRINT #1, "                                                 Pre-Tax    Post-Tax     Employer"
  '  PRINT #1, "Employee Name                       SS#          Vol Cont   Loan Pmts    Contrib'n      Batch     Date"
  IF VassFlag THEN
    PRINT #1, "Employee Name                       SS#          Vol Cont   Loan Pmts    Contrib'n      Date   RetType   Gross"
  ELSE
    PRINT #1, "Employee Name                       SS#          Vol Cont   Loan Pmts    Contrib'n      Date"
  END IF
  PRINT #1, ""
  RETURN
  
PrintGrandTots:
  Totals$ = SPACE$(95)
  
  RSET K401Totals(1).VAmt = LTRIM$(RTRIM$(FUsing(STR$(TotalVAmt#), Image1$)))
  RSET K401Totals(1).LAmt = LTRIM$(RTRIM$(FUsing(STR$(TotalLAmt#), Image1$)))
  RSET K401Totals(1).MAmt = LTRIM$(RTRIM$(FUsing(STR$(TotalMatchAmt#), Image2$)))
  
  IF TotalVAmt# = 0 THEN
    RSET K401Totals(1).VAmt = ".00"
  END IF
  IF TotalLAmt# = 0 THEN
    RSET K401Totals(1).LAmt = ".00"
  END IF
  IF TotalMatchAmt# = 0 THEN
    RSET K401Totals(1).MAmt = ".00"
  END IF
  
  LSET Totals$ = "Totals"
  MID$(Totals$, 22) = FUsing(STR$(TotalGross#), Image2$)
  MID$(Totals$, 37) = STR$(EPrinted - 1)
  MID$(Totals$, 47) = K401Totals(1).VAmt
  MID$(Totals$, 59) = K401Totals(1).LAmt
  MID$(Totals$, 72) = K401Totals(1).MAmt
  MID$(Totals$, 87) = LTRIM$(RTRIM$(FUsing(STR$(Round#(TotalVAmt# + TotalLAmt# + TotalMatchAmt#)), Image2$)))
  
  PRINT #1,
  PRINT #1, Totals$  '; "     "; FUsing(STR$(TotalGross#), Image2$)
  PRINT #1,
  PRINT #1, "Subtotals"
  PRINT #1,
  PRINT #1, "  Total Vol Law "; USING "#####.##"; TVLawAmt#
  PRINT #1, "  Total Vol Gen "; USING "#####.##"; TVGenAmt#
  PRINT #1,
  PRINT #1, " Total Loan Law "; USING "#####.##"; TLLawAmt#
  PRINT #1, " Total Loan Gen "; USING "#####.##"; TLGenAmt#
  PRINT #1,
  PRINT #1, "Total Match Law "; USING "#####.##"; TotLawMatAmt#
  PRINT #1, "Total Match Gen "; USING "#####.##"; TotGenMatAmt#
  
  PRINT #1, FF$
  
  RETURN
  
END SUB

SUB YTDFundReport
  
  ExitFlag = False
  FirstTime = True
  OKFlag = False
  
  REDIM Frm(1) AS FormInfo
  
  FormName$ = "YTDFUND"
  NumFlds = LibNumberOfFields(MiscQLib, FormName$)
  REDIM Form$(NumFlds, 2)       'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo               'DIM the field information array
  StartEl = 0   'Load first form at array start
  LibGetFldDef MiscQLib, FormName$, StartEl, Fld(), Form$(), ErrCode
  
  
  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1
  
  '----- Setup TYPE for setting and reading form editing information.
  
  Frm(1).FldNo = 1              'Start editing on field #1
  Frm(1).InsStat = False        'Set insert state (True = Insert on)
  Frm(1).StartEl = 0            'Set form starting element to 0 and
  
  
  BlockClear
  DisplayMiscScrn FormName$
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      Form$(1, 0) = RIGHT$(DATE$, 4)
      'lset Form$(1, 0) =
      Action = 2
    END IF
    
    SELECT CASE Frm(1).KeyCode
    CASE F0Key
      Year$ = Form$(1, 0)
      LowDate = Date2Num("01-01-" + Year$)
      HiDate = Date2Num("12-31-" + Year$)
      OKFlag = True
    CASE EscKey
      ExitFlag = True
    END SELECT
    
  LOOP UNTIL OKFlag OR ExitFlag 'proper key not set
  
  CursorOff
  
  ERASE Form$, Fld, Frm
  
  IF ExitFlag THEN EXIT SUB
  
  RptName$ = "PRRPTS\YTDWAGE.RPT"
  RptTitle$ = "YTD Wage Distribution Report."
  
  REDIM THRec(1)          AS TransRecType
  REDIM Unit(1)           AS UnitFileRecType
  REDIM Pg(1)             AS STRING * 3
  REDIM Dash(1)           AS STRING * 80
  REDIM FundYTD(1 TO 1)  AS YTDFundRptType
  REDIM ToDisk(1) AS STRING * 78
  REDIM SysRec(1)     AS RegDSysFileRecType
  
  FGetAH SysFileName, SysRec(1), LEN(SysRec(1)), 1
  FGetAH UnitFileName, Unit(1), LEN(Unit(1)), 1
  
  'FundNumLen = 5
  
  FundNumLen = SysRec(1).AcctCnt
  FundYTDLen = LEN(FundYTD(1))
  FundCnt = 0
  FirstFund = True
  
  MaxLines = 59
  
  Dash(1) = STRING$(80, "-")
  Image1$ = "#####.##"
  Image2$ = "######.##"
  Image3$ = "#######.##"
  
  TRecSize = LEN(THRec(1))
  NumOfRecs& = FileSize(TransHistFileName) \ TRecSize
  REDIM FundList(1 TO NumOfRecs&) AS LONG
  
  
  FCreate RptName$
  FOpenS RptName$, RHandle
  RPTSetupPRN 9, RHandle
  FOpenS TransHistFileName, THandle
  
  TCol = 40 - (LEN(RptTitle$) \ 2) + 5
  PctRow = 11
  ShowProcessingScrn RptTitle$
  
  QPrintRC "Scanning Transactions", PctRow, TCol, 112
  
  FOR Cnt = 1 TO NumOfRecs&
    FGetRTA THandle, THRec(1), CLNG(Cnt), TRecSize
    IF (THRec(1).CheckDate >= LowDate) AND (THRec(1).CheckDate <= HiDate) THEN
      FOR Cnt2 = 1 TO 8         'eight possiable distributions
        TAcct$ = THRec(1).TDist(Cnt2).DAcct
        DO
          DashPos = INSTR(TAcct$, "-")
          IF DashPos THEN
            TAcct$ = LEFT$(TAcct$, DashPos - 1) + MID$(TAcct$, DashPos + 1)
          END IF
        LOOP WHILE DashPos
        IF LEN(QPTrim$(TAcct$)) > 0 THEN
          TFundNum& = QPValL(LEFT$(TAcct$, FundNumLen))
          GOSUB Parse2Fund
        END IF
      NEXT
    END IF
    ShowPctComp Cnt, NumOfRecs&
  NEXT
  
  FClose THandle
  
  SortT FundYTD(1), FundCnt, 0, FundYTDLen, 0, -2
  'sort by fund number
  QPrintRC "  Summarizing Funds  ", PctRow, TCol, 112
  GOSUB PrintRptHeader
  
  FOR Cnt = 1 TO FundCnt
    GOSUB PrintFundInfo
    ShowPctComp Cnt, FundCnt
  NEXT
  FPut RHandle, FF$
  TotalsFlag = True
  GOSUB PrintRptHeader
  
  FOR Cnt = 1 TO FundCnt
    
    RegTot# = 0
    OTTot# = 0
    RegWageTot# = 0
    OTWageTot# = 0
    
    FOR MonthCnt = 1 TO 12
      RegTot# = Round(RegTot# + FundYTD(Cnt).Mths(MonthCnt).RegHrs)
      RegWageTot# = Round(RegWageTot# + FundYTD(Cnt).Mths(MonthCnt).RegWage)
      OTTot# = Round(OTTot# + FundYTD(Cnt).Mths(MonthCnt).OTHrs)
      OTWageTot# = Round(OTWageTot# + FundYTD(Cnt).Mths(MonthCnt).OTWage)
    NEXT
    LSET ToDisk(1) = "  " + QPStrL(FundYTD(Cnt).FundNum)
    MID$(ToDisk(1), 19) = FUsing(STR$(RegTot#), Image2$)
    MID$(ToDisk(1), 34) = FUsing(STR$(OTTot#), Image2$)
    MID$(ToDisk(1), 48) = FUsing(STR$(RegWageTot#), Image3$)
    MID$(ToDisk(1), 63) = FUsing(STR$(OTWageTot#), Image3$)
    
    FPut RHandle, ToDisk(1) + CrLf$
    LineCnt = LineCnt + 1
    IF LineCnt > MaxLines THEN
      FPut RHandle, FF$
      GOSUB PrintRptHeader
    END IF
  NEXT
  FPut RHandle, FF$
  RPTSetupPRN 0, RHandle
  FClose RHandle
  
  ERASE THRec, Unit, Pg, Dash, FundYTD, FundList
  
  PrintRptFile RptTitle$, RptName$, 1, RetCode, 0
  
  EXIT SUB
  
  
Parse2Fund:
  
  FundPointer = 0
  NewFund = True                'assume this is a new fund
  
  IF FirstFund THEN             'if this is the first fund processed
    FirstFund = False           'skip search for fund part
  ELSE
    FOR FCnt = 1 TO FundCnt     'look through the fund list and see
      IF FundList(FCnt) = TFundNum& THEN        'if this fund is already in the list
        FundPointer = FCnt      'it is point to this fund in array
        NewFund = False         'Not a new fund
        EXIT FOR                '
      END IF
    NEXT
  END IF
  
  IF NewFund THEN               'if this fund wasn't found in the list
    FundCnt = FundCnt + 1       'total funds count
    FundPointer = FundCnt       'set the array pointer to the new entry
    REDIM PRESERVE FundYTD(1 TO FundCnt)  AS YTDFundRptType
    'resize the YTD Fund array
    FundList(FundCnt) = TFundNum&               'add this fund to the fund list
    FundYTD(FundCnt).FundNum = TFundNum&        'set the new fund into the array
  END IF
  
  MonthNum = QPValI(LEFT$(Num2Date(THRec(1).CheckDate), 2))
  'get the month of this transaction
  
  FundYTD(FundPointer).Mths(MonthNum).RegHrs = Round(FundYTD(FundPointer).Mths(MonthNum).RegHrs + THRec(1).TDist(Cnt2).DRHrs)
  FundYTD(FundPointer).Mths(MonthNum).OTHrs = Round(FundYTD(FundPointer).Mths(MonthNum).OTHrs + THRec(1).TDist(Cnt2).DOHrs)
  FundYTD(FundPointer).Mths(MonthNum).RegWage = Round(FundYTD(FundPointer).Mths(MonthNum).RegWage + THRec(1).TDist(Cnt2).DRWage)
  FundYTD(FundPointer).Mths(MonthNum).OTWage = Round(FundYTD(FundPointer).Mths(MonthNum).OTWage + THRec(1).TDist(Cnt2).DOWage)
  'set trans data into the YTD Fund Array for the month
  
  '  FundYTD(FundPointer).Mths(MonthNum).VACHRS  =Round(FundYTD(FCnt).Mths(MonthNum).VACHRS  +
  '  FundYTD(FundPointer).Mths(MonthNum).SICKHRS =Round(FundYTD(FCnt).Mths(MonthNum).SICKHRS +
  '  FundYTD(FundPointer).Mths(MonthNum).HOLHRS  =Round(FundYTD(FCnt).Mths(MonthNum).HOLHRS  +
  '  FundYTD(FundPointer).Mths(MonthNum).CTHrs   =Round(FundYTD(FCnt).Mths(MonthNum).CTHrs   +
  '  FundYTD(FundPointer).Mths(MonthNum).CWHrs   =Round(FundYTD(FCnt).Mths(MonthNum).CWHrs   +
  
  
  RETURN
  
  
PrintRptHeader:
  Page = Page + 1
  RSET Pg(1) = STR$(Page)
  UTemp$ = SPACE$(71)
  LSET UTemp$ = Unit(1).UFEMPR
  MID$(UTemp$, 62) = "Page:" + Pg(1) + CrLf$
  FPut RHandle, UTemp$
  FPut RHandle, "YTD Wage Distribution Report for Year: " + Year$ + CrLf$
  FPut RHandle, "Report Date: " + DATE$ + CrLf$ + CrLf$
  IF TotalsFlag THEN
    FPut RHandle, "Totals" + CrLf$
    FPut RHandle, "Fund No.            Reg Hrs         OT Hrs      Reg Wages       OT Wages" + CrLf$
    LineCnt = 6
  ELSE
    FPut RHandle, "Fund No." + CrLf$
    FPut RHandle, "Month               Reg Hrs         OT Hrs      Reg Wages       OT Wages" + CrLf$
  END IF
  FPut RHandle, Dash(1) + CrLf$
  RETURN
  
PrintFundInfo:
  
  NumPrinted = NumPrinted + 1
  IF NumPrinted = 4 THEN
    NumPrinted = 1
    FPut RHandle, FF$
    GOSUB PrintRptHeader
  END IF
  FPut RHandle, "Fund: " + QPStrL(FundYTD(Cnt).FundNum) + CrLf$
  '    RegTot# = 0
  '    OTTot# = 0
  FOR MonthCnt = 1 TO 12
    '     RegTot# = Round(RegTot# + FundYTD(Cnt).Mths(MonthCnt).RegHrs)
    '     OTTot# = Round(OTTot# + FundYTD(Cnt).Mths(MonthCnt).OTHrs)
    
    LSET ToDisk(1) = "  " + MID$("JanFebMarAprMayJunJulAugSepOctNovDec", (((MonthCnt - 1) * 3)) + 1, 3)
    MID$(ToDisk(1), 20) = FUsing(STR$(FundYTD(Cnt).Mths(MonthCnt).RegHrs), Image1$)
    MID$(ToDisk(1), 35) = FUsing(STR$(FundYTD(Cnt).Mths(MonthCnt).OTHrs), Image1$)
    MID$(ToDisk(1), 49) = FUsing(STR$(FundYTD(Cnt).Mths(MonthCnt).RegWage), Image2$)
    MID$(ToDisk(1), 64) = FUsing(STR$(FundYTD(Cnt).Mths(MonthCnt).OTWage), Image2$)
    
    FPut RHandle, QPRTrim$(ToDisk(1)) + CrLf$
  NEXT
  FPut RHandle, CrLf$
  '  LSET ToDisk(1) = "Totals:"
  '  MID$(ToDisk(1), 19) = FUsing(STR$(RegTot#), "######.##")
  '  MID$(ToDisk(1), 39) = FUsing(STR$(OTTot#), "######.##")
  '  FPut RHandle, ToDisk(1) + CrLf$
  '  FPut RHandle, CrLf$
  
  RETURN
  
  
END SUB

