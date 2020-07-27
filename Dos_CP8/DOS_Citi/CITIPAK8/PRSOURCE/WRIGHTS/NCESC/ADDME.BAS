DEFINT A-Z
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)

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
'$INCLUDE: 'PRRpts.BI'

CONST False = 0
CONST True = NOT False


SUB ESC2Disk

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
      Form$(2, 0) = RIGHT$(DATE$, 2)
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

  'RPTSetupPRN 7, RHandle

  FOpenS TransHistFileName, THandle
  FOpenS EmpData2Name, DHandle   'open employee data file

  GOSUB PrintESCHeader

  FOR RecNo = 1 TO NumOfRecs
    FGetRTA DHandle, Emp2Rec(1), CLNG(IdxBuff(RecNo)), EmpRecSize
    IF Emp2Rec(1).LastTransRec <= 0 THEN
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
  'RPTSetupPRN 0, RHandle
  
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
      EQtrWage(Cnt2) = Round(EQtrWage(Cnt2) + TransHRec(1).GrossPay)
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

