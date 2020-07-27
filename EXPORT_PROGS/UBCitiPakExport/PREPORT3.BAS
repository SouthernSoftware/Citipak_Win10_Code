DEFINT A-Z

DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB MPaintBox (ULRow, ULCol, LRRow, LRCol, Colr)
DECLARE SUB VertMenuT (Items() AS ANY, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)

DECLARE SUB ReadFileX (FSpec$, DirSize&, SEG Element AS ANY)
  
DECLARE FUNCTION ASCII% (Strng$)
DECLARE FUNCTION DOSError% ()
DECLARE FUNCTION FCount% (FSpec$)

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
'$INCLUDE: 'DedCodes.Bi'
'$INCLUDE: 'ErnCodes.Bi'
'$INCLUDE: 'PRUNIT.Bi'
'$INCLUDE: 'PRSYSCTR.Bi'
'$INCLUDE: 'PRPPDEF.Bi'
'$INCLUDE: 'PRTRANS.Bi'
'$INCLUDE: 'PRRpts.BI'

  
  CONST False = 0
  CONST True = NOT False
  
  TYPE FullInfo
    BaseName AS STRING * 8
    ExtName  AS STRING * 3
    FileSize AS LONG
    FileDate AS STRING * 8
    FileTime AS STRING * 6
    Attrib   AS STRING * 1
  END TYPE
  
  TYPE RPTLstType
    BaseName AS STRING * 8
    RPTDesc  AS STRING * 26
  END TYPE
  
  TYPE FLen
    V AS STRING * 43
  END TYPE

  TYPE CompSortType
    CompCode AS STRING * 12
    RecNo    AS LONG
  END TYPE

SUB CheckIssuedRpt

  REDIM TempScrn(1)

  ExitFlag = False
  FirstTime = True

  OKFlag = False

  REDIM Frm(1) AS FormInfo

  FormName$ = "CHKISSUE"
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
  Image2$ = "###,###.##"

  REDIM Dash(1) AS STRING * 80
  REDIM TransHRec(1)   AS TransRecType
  REDIM EMP2Rec(1)     AS EmpData2Type
  REDIM Unit(1)        AS UnitFileRecType
  REDIM Pg(1)          AS STRING * 3

  ToPrint$ = SPACE$(78)
  ChkNum$ = SPACE$(12)
  
  MaxLines = 57
  LineCnt = 0
  Dash(1) = STRING$(78, "-") + CrLf$
  
  EmpRecSize = LEN(EMP2Rec(1))
  TRecSize = LEN(TransHRec(1))

  IdxRecLen = 2
  IdxFileSize& = FileSize(EmpIdxLName)
  NumOfRecs = IdxFileSize& \ IdxRecLen

  REDIM IdxBuff(1 TO NumOfRecs)
  FGetAH EmpIdxLName, IdxBuff(1), IdxRecLen, NumOfRecs

  FGetAH UnitFileName, Unit(1), LEN(Unit(1)), 1

  RptTitle$ = "Employee Checks Issued Report"
  ShowProcessingScrn RptTitle$

  RptName$ = "PRRPTS\CHKISSUE.RPT"

  FCreate RptName$
  FOpenS RptName$, RHandle

  RPTSetupPRN 10, RHandle

  FOpenS TransHistFileName, THandle
  FOpenS EmpData2Name, DHandle  'open employee data file

  GOSUB PrintIssuedHeader

  FOR RecNo = 1 TO NumOfRecs
    UsingThisOne = False
    FGetRTA DHandle, EMP2Rec(1), CLNG(IdxBuff(RecNo)), EmpRecSize
    IF EMP2Rec(1).LastTransRec <= 0 THEN
      GOTO SkipEm4
    END IF
    TransRecNum& = CLNG(EMP2Rec(1).LastTransRec)

    DO
      FGetRTA THandle, TransHRec(1), TransRecNum&, TRecSize

      SELECT CASE TransHRec(1).CheckDate

      CASE LowDate TO HiDate
        GOSUB PrintChkInfo
        UsingThisOne = True
      CASE ELSE
      END SELECT
      IF TransHRec(1).PrevTransRec <= 0 THEN    'if no more transactions
        IF UsingThisOne THEN    'if there were any at all
          FPut RHandle, CrLf$
          LineCnt = LineCnt + 1
          IF LineCnt > MaxLines THEN            'bottom of page?
            FPut RHandle, FF$   'yes; form feed
            GOSUB PrintIssuedHeader              'write title lines
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

  FPut RHandle, Dash(1)

  RPTSetupPRN 0, RHandle

  FClose DHandle
  FClose THandle
  FClose RHandle

  ERASE Dash, TransHRec, EMP2Rec, Unit, Pg, IdxBuff

  PrintRptFile RptTitle$, RptName$, 1, RetCode, 0
  
  EXIT SUB

PrintIssuedHeader:
  Page = Page + 1
  RSET Pg(1) = QPStrI$(Page)
  UTemp$ = SPACE$(80)
  LSET UTemp$ = Unit(1).UFEMPR
  MID$(UTemp$, 71) = "Page:" + Pg(1) + CrLf$
  FPut RHandle, UTemp$
  FPut RHandle, "Checks Issued Report" + CrLf$
  FPut RHandle, Num2Date$(LowDate) + " thru " + Num2Date$(HiDate) + CrLf$
  FPut RHandle, "Check Date    Check Number        Employee Name                       Net Pay" + CrLf$
  FPut RHandle, Dash(1)
  LineCnt = 5
  RETURN

PrintChkInfo:
  LSET ToPrint$ = Num2Date(TransHRec(1).CheckDate)
  RSET ChkNum$ = QPStrL(TransHRec(1).CheckNUM)
  MID$(ToPrint$, 15) = ChkNum$
  MID$(ToPrint$, 35) = QPTrim$(EMP2Rec(1).EMPFNAME) + " " + QPTrim$(EMP2Rec(1).EMPLNAME)
  MID$(ToPrint$, 70) = FUsing(STR$(TransHRec(1).NetPay), Image1$)
  FPut RHandle, ToPrint$ + CrLf$
  LineCnt = LineCnt + 1
RETURN

END SUB

SUB CompWageReport
  
  REDIM TempScrn(1)
  REDIM CompSort(1 TO 1)  AS CompSortType

  ExitFlag = False
  FirstTime = True
  
  OKFlag = False
  
  REDIM Frm(1) AS FormInfo
  
  FormName$ = "GWAGERPT"
  NumFlds = LibNumberOfFields(MiscQLib, FormName$)
  REDIM Form$(NumFlds, 2)           'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo   'DIM the field information array
  StartEl = 0                       'Load first form at array start
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
      QPrintRC "Annual Workers Comp Report", 7, 28, 126
      FirstTime = False
      'LSET Form$(1, 0) = "06-01-2000"
      LSET Form$(2, 0) = DATE$
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
  
  REDIM Dash(1) AS STRING * 80
  REDIM TransHRec(1)   AS TransRecType
  REDIM EMP2Rec(1)     AS EmpData2Type
  REDIM Unit(1)        AS UnitFileRecType
  REDIM Pg(1)          AS STRING * 3
  
  REDIM GrsRpt(1 TO 3) AS GrossWageRptType
  
  MaxLines = 57
  LineCnt = 0
  Dash(1) = STRING$(80, "-")
  
  EmpRecSize = LEN(EMP2Rec(1))
  TRecSize = LEN(TransHRec(1))
  
  IdxRecLen = 2
  IdxFileSize& = FileSize(EmpIdxNName)
  NumOfRecs = IdxFileSize& \ IdxRecLen
  
  REDIM IdxBuff(1 TO NumOfRecs)
  FGetAH EmpIdxNName, IdxBuff(1), IdxRecLen, NumOfRecs
  
  FGetAH UnitFileName, Unit(1), LEN(Unit(1)), 1
  
  RptTitle$ = "Employee Gross Wage Report"
  ShowProcessingScrn RptTitle$
  
  RptName$ = "PRRPTS\COMPWAGE.RPT"
  
  RptFile = FREEFILE
  OPEN RptName$ FOR OUTPUT AS RptFile

  'RPTSetupPRN 5, RHandle
  
  FOpenS TransHistFileName, THandle
  FOpenS EmpData2Name, DHandle  'open employee data file

  GOSUB MakeCompIndex

  GOSUB PrintGrossHeader
  
  FOR RecNo = 1 TO NumOfRecs
    UsingThisOne = False
    FGetRTA DHandle, EMP2Rec(1), CompSort(RecNo).RecNo, EmpRecSize
    IF EMP2Rec(1).LastTransRec <= 0 THEN
      GOTO SkipEmComp1
    END IF

    PayFreq$ = QPTrim$(EMP2Rec(1).EMPPFREQ)
    SELECT CASE PayFreq$
    CASE "Weekly"
      PayFreq = 52
      Hrs# = 40
    CASE "Bi-Weekly"
      PayFreq = 26
      Hrs# = 80
    CASE "Semi-Monthly"
      PayFreq = 24
      Hrs# = 86.66
    CASE "Monthly"
      PayFreq = 12
      Hrs# = 173.33
    CASE "Quarterly"
      PayFreq = 4
      Hrs# = 520
    CASE "Semi-Annually"
      PayFreq = 2
      Hrs# = 1040
    CASE "Annually"
      PayFreq = 1
      Hrs# = 2080
    END SELECT

    AnnualGross# = 0
    IF EMP2Rec(1).EMPPRATE < 0 THEN
      EMP2Rec(1).EMPPRATE = 0
    END IF

    PayType$ = UCASE$(QPTrim$(EMP2Rec(1).EMPPTYPE))
    SELECT CASE PayType$
    CASE "HOURLY"
       AnnualGross# = Round#(Hrs# * EMP2Rec(1).EMPPRATE)' * (0# + PayFreq))
    CASE "SALARIED"
       AnnualGross# = Round#(EMP2Rec(1).EMPPRATE * (0# + PayFreq))
    END SELECT

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
            PRINT #RptFile, FF$
            'FPut RHandle, FF$   'yes; form feed
            GOSUB PrintGrossHeader              'write title lines
          END IF
        END IF
        EXIT DO
      ELSE      'there are more trans to check
        TransRecNum& = CLNG(TransHRec(1).PrevTransRec)
      END IF
    LOOP
SkipEmComp1:
    ShowPctComp RecNo, NumOfRecs
  NEXT
  
  GOSUB PrintGrossGTotals

  'RPTSetupPRN 0, RHandle
  
  FClose DHandle
  FClose THandle
  CLOSE

  ERASE Dash, TransHRec, EMP2Rec, Unit, Pg, GrsRpt, IdxBuff
  
  PrintRptFile RptTitle$, RptName$, 1, RetCode, 0
  
  
  EXIT SUB
  
PrintGrossHeader:
  Page = Page + 1
  RSET Pg(1) = QPStrI$(Page)
  UTemp$ = SPACE$(80)
  LSET UTemp$ = Unit(1).UFEMPR
  MID$(UTemp$, 65) = "Page:" + Pg(1)
  PRINT #RptFile, UTemp$
  PRINT #RptFile, "Annual Workers Comp Report"
  PRINT #RptFile, Num2Date$(LowDate) + " thru " + Num2Date$(HiDate); "           Fiscal      Annual      Fiscal       Comp"
  '               "01-01-0000 thru 01-01-0000
  PRINT #RptFile, "Emp #      Name                       Gross       Gross    Overtime       Code"
  PRINT #RptFile, Dash(1)
  LineCnt = 5
  RETURN
  
SumEmpGross:
  EGross# = Round#(EGross# + TransHRec(1).GrossPay)
  EOTWage# = Round#(EOTWage# + TransHRec(1).TotOTWage)
RETURN
  
PrintEmpGrossTotals:
  EmpInfo$ = SPACE$(80)
  CompCode$ = SPACE$(10)
  LSET EmpInfo$ = QPTrim$(EMP2Rec(1).EmpNo)
  MID$(EmpInfo$, 11) = QPTrim$(EMP2Rec(1).EMPLNAME) + ", " + QPTrim$(EMP2Rec(1).EMPFNAME)
  MID$(EmpInfo$, 34) = FUsing(STR$(EGross#), Image2$)
  MID$(EmpInfo$, 46) = FUsing(STR$(AnnualGross#), Image2$)
  MID$(EmpInfo$, 58) = FUsing(STR$(EOTWage#), Image2$)
  RSET CompCode$ = QPTrim$(EMP2Rec(1).EMPWCCLS)
  MID$(EmpInfo$, 69) = CompCode$
  PRINT #RptFile, EmpInfo$
  
  LineCnt = LineCnt + 1

  TEGross# = Round#(TEGross# + EGross#)
  TAnnualGross# = Round#(TAnnualGross# + AnnualGross#)
  TEOTWage# = Round#(TEOTWage# + EOTWage#)

  EGross# = 0
  EOTWage# = 0
  AnnualGross# = 0
RETURN
  
PrintGrossGTotals:
  EmpInfo$ = SPACE$(80)
  LSET EmpInfo$ = "Totals:"
  MID$(EmpInfo$, 34) = FUsing(STR$(TEGross#), Image3$)
  MID$(EmpInfo$, 46) = FUsing(STR$(TAnnualGross#), Image3$)
  MID$(EmpInfo$, 58) = FUsing(STR$(TEOTWage#), Image3$)
  PRINT #RptFile, Dash(1)
  PRINT #RptFile, EmpInfo$
  PRINT #RptFile, FF$
RETURN

MakeCompIndex:
  REDIM CompSort(1 TO NumOfRecs)  AS CompSortType

  FOR RecNo = 1 TO NumOfRecs
    FGetRTA DHandle, EMP2Rec(1), CLNG(RecNo), EmpRecSize
    RSET CompSort(RecNo).CompCode = EMP2Rec(1).EMPWCCLS
    CompSort(RecNo).RecNo = RecNo
  NEXT
  SortT CompSort(1), NumOfRecs, 0, 16, 0, 12

RETURN

END SUB

SUB RePrintRpts

  CursorOff

  FSpec$ = "PRRPTS\*.RPT"       'default to all files
  Count = FCount%(FSpec$)       'see how many files match
  IF Count = 0 THEN
     DisplayMiscScrn "NOREPTR"
     WaitForAction
     EXIT SUB
  END IF
  
  REDIM Items(1 TO Count) AS FLen
  REDIM FileInfo(1 TO Count) AS FullInfo           'make an array to hold the full file info
  
  ReadFileX FSpec$, DirSize&, FileInfo(1)          'ReadFileX gets all the info in one shot
  
  ItemLen = LEN(Items(1).V)

  REDIM RPTList(1 TO 16) AS RPTLstType
  FGetAH ReportListFileName, RPTList(1), LEN(RPTList(1)), 16

  REDIM DedCodes(1 TO 12)   AS DedCodeRecType
  FGetAH DedCodeFileName, DedCodes(1), LEN(DedCodes(1)), 12
  
  FOR x = 1 TO Count            'print 'em to show it worked
    FOR Cnt2 = 1 TO 16
      IF INSTR(FileInfo(x).BaseName, QPTrim$(RPTList(Cnt2).BaseName)) THEN
        SELECT CASE RPTList(Cnt2).BaseName
'        CASE "QUARTER "
'          QStart = 8
'          GOSUB MakeQtrString
        CASE "ESCQTR  "
          QStart = 7
          GOSUB MakeQtrString
        CASE "DEDUCT  "
          DedNum = QPValI(MID$(FileInfo(x).BaseName, 7))
          TDesc$ = QPTrim$(RPTList(Cnt2).RPTDesc) + " " + DedCodes(DedNum).DCDESC1
          LSET Items(x).V = TDesc$
        CASE ELSE
          LSET Items(x).V = RPTList(Cnt2).RPTDesc
        END SELECT
        MID$(Items(x).V, 32) = FileInfo(x).FileDate
        MID$(Items(x).V, LEN(Items(x).V) - 1, 2) = QPStrI$(x)
        EXIT FOR
      END IF
    NEXT
    IF Cnt2 > 16 THEN
      LSET Items(x).V = FileInfo(x).BaseName
      MID$(Items(x).V, 32) = FileInfo(x).FileDate
      MID$(Items(x).V, LEN(Items(x).V) - 1, 2) = QPStrI$(x)
    END IF
  NEXT

  SortT Items(1), Count, 0, ItemLen, 0, ItemLen
  
  VAction = 1
  BoxBot = 18
  BoxTop = 9
  MaxLen = 40
  LeftCol = ((80 - MaxLen) \ 2) - 1
  
  DO
    HideCursor
    LOCATE BoxTop, LeftCol
    BlockClear
    WazzWind 3, LeftCol, 5, LeftCol + MaxLen + 3, Cnf.MenBox, 1, True
    QPrintRC "Reprint Reports Menu", 4, LeftCol + (MaxLen \ 2) - 8, Cnf.PulBar

    QPrintRC "  Report Description             File Date  ", BoxTop - 1, LeftCol, Cnf.PulBar
    MPaintBox BoxTop, LeftCol + MaxLen + 4, BoxTop, LeftCol + MaxLen + 5, 8
    WazzWind 21, LeftCol, 23, LeftCol + MaxLen + 3, Cnf.PulBar, 2, True
    QPrintRC " Use (" + CHR$(24) + "-" + CHR$(25) + ") to select", 22, LeftCol + (MaxLen \ 2) - 8, Cnf.PulBar
    
    ShowCursor
    
    DO
      
      VertMenuT Items(), Choice, MaxLen, BoxBot, Ky$, VAction, Cnf
      
      SELECT CASE ASCII(Ky$)
      CASE EscKey
        Choice = EscKey
        ExitFlag = True
        EXIT DO
      CASE 13
        CursorOff
        RptTitle$ = QPTrim$(LEFT$(Items(Choice).V, 26))
        FileName$ = "PRRPTS\" + QPTrim$(FileInfo(QPValI(RIGHT$(Items(Choice).V, 2))).BaseName) + ".RPT"
        PrintRptFile RptTitle$, FileName$, 1, RetCode%, 1
        EXIT DO
      CASE ELSE
      END SELECT
    LOOP
  LOOP UNTIL ExitFlag
  ERASE Items, FileInfo, RPTList, DedCodes
  
  CursorOff
  EXIT SUB
  
MakeQtrString:
  Qtr = QPValI(MID$(FileInfo(x).BaseName, QStart))
  SELECT CASE Qtr
  CASE 1
    Qtr$ = "1st"
  CASE 2
    Qtr$ = "2nd"
  CASE 3
    Qtr$ = "3rd"
  CASE 4
    Qtr$ = "4th"
  END SELECT
  LSET Items(x).V = RPTList(Cnt2).RPTDesc
  MID$(Items(x).V, 1) = Qtr$
RETURN

END SUB

SUB SCRetirementRpt

  SHARED PrnDef$()

  FirstTime = True

  REDIM TempScrn(1)
  REDIM Frm(1) AS FormInfo
  FormName$ = "ERETIRPT"
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

  RptName$ = "PRRPTS\RETIRE.RPT"

  REDIM TransHRec(1)   AS TransRecType
  REDIM EMP2Rec(1)     AS EmpData2Type
  REDIM Unit(1)        AS UnitFileRecType
  REDIM Pg(1)          AS STRING * 3
  REDIM Dash(1)        AS STRING * 78

  Dash(1) = STRING$(78, "-")

  EmpRecSize = LEN(EMP2Rec(1))
  TRecSize = LEN(TransHRec(1))

  IdxRecLen = 2
  IdxFileSize& = FileSize(EmpIdxLName)
  NumOfRecs = IdxFileSize& \ IdxRecLen

  REDIM IdxBuff(1 TO NumOfRecs)
'  REDIM RetList(1 TO 6) AS RetireRecType

  FGetAH EmpIdxLName, IdxBuff(1), IdxRecLen, NumOfRecs
  FGetAH UnitFileName, Unit(1), LEN(Unit(1)), 1
  UTemp$ = "Reporting Unit: " + UCASE$(Unit(1).UFEMPR)

'  FGetAH RetireFileName, RetList(1), LEN(RetList(1)), 6

  MaxLines = 55

'-=-=-=-=-=-=-=-=-=-=-=

  RptTitle$ = "Employee Retirement Report."
  ShowProcessingScrn RptTitle$

  TCol = 40 - (LEN(RptTitle$) \ 2) + 1
  PctRow = 11

'  QPrintRC "Pass 1 of 2", PctRow, TCol, 112

'  OPEN RptName$ FOR OUTPUT AS #1

  FCreate RptName$
  FOpenS RptName$, RHandle
  RPTSetupPRN 11, RHandle
  FClose RHandle

  OPEN RptName$ FOR APPEND AS #1

  FOpenS TransHistFileName, THandle
  FOpenS EmpData2Name, DHandle   'open employee data file

  Pass = 1

PassLoop:
  GOSUB SCRetRptHeader
  FOR RecNo = 1 TO NumOfRecs
    UsingThisOne = False
    RetAmt# = 0
    RetWage# = 0
    MatAmt# = 0

    'IF CLNG(IdxBuff(RecNo)) <= ThisEmp& THEN STOP

    ThisEmp& = CLNG(IdxBuff(RecNo))

    FGetRTA DHandle, EMP2Rec(1), ThisEmp&, EmpRecSize
    IF Pass = 1 THEN
      IF LEFT$(EMP2Rec(1).EMPRETTP, 1) <> "G" THEN
        GOTO SCSkipEm
      END IF
    ELSE
      IF LEFT$(EMP2Rec(1).EMPRETTP, 1) <> "L" THEN
        GOTO SCSkipEm
      END IF
    END IF
        
    IF EMP2Rec(1).LastTransRec <= 0 THEN
      GOTO SCSkipEm
    END IF

    TransRecNum& = CLNG(EMP2Rec(1).LastTransRec)

    DO
      FGetRTA THandle, TransHRec(1), TransRecNum&, TRecSize
      SELECT CASE TransHRec(1).CheckDate
      CASE LowDate TO HighDate
        IF TransHRec(1).RetireAmt <> 0 THEN
          RetAmt# = Round(RetAmt# + TransHRec(1).RetireAmt)
          RetWage# = Round(RetWage# + TransHRec(1).RetGrossPay)
          MatAmt# = Round(MatAmt# + TransHRec(1).MatchRetAmt)
          UsingThisOne = True
        END IF
      CASE ELSE
      END SELECT

      IF TransHRec(1).PrevTransRec <= 0 THEN
        IF UsingThisOne THEN
          'GOSUB GetRetType
          GOSUB SCPrintEmpRetLine
          IF LineCnt >= MaxLines THEN
            PRINT #1, FF$
            GOSUB SCRetRptHeader
          END IF
        END IF
        EXIT DO
      ELSE
        TransRecNum& = CLNG(TransHRec(1).PrevTransRec)
      END IF

    LOOP

SCSkipEm:
    ShowPctComp RecNo, NumOfRecs
  NEXT

  IF Pass = 1 THEN
    GOSUB SCRetTotals
    GRTotal# = RTotal#
    GWTotal# = WTotal#
    GETotal# = ETotal#

    RTotal# = 0
    WTotal# = 0
    ETotal# = 0
    
    Pass = Pass + 1
    GOTO PassLoop
  END IF

  GOSUB SCRetGTotals
'***********************************************************
  GOSUB SCNORetRptHeader

  ThisEmp& = 0
  FOR RecNo = 1 TO NumOfRecs
    UsingThisOne = False
    RetAmt# = 0
    RetWage# = 0
    MatAmt# = 0

    'IF CLNG(IdxBuff(RecNo)) <= ThisEmp& THEN STOP
    ThisEmp& = CLNG(IdxBuff(RecNo))
    
    FGetRTA DHandle, EMP2Rec(1), ThisEmp&, EmpRecSize

    IF EMP2Rec(1).LastTransRec <= 0 THEN
      GOTO SCSkipNOEm
    END IF

    TransRecNum& = CLNG(EMP2Rec(1).LastTransRec)

    DO
      FGetRTA THandle, TransHRec(1), TransRecNum&, TRecSize
      SELECT CASE TransHRec(1).CheckDate
      CASE LowDate TO HighDate
        RetAmt# = Round(RetAmt# + TransHRec(1).RetireAmt)
        RetWage# = Round(RetWage# + TransHRec(1).RetGrossPay)
        MatAmt# = Round(MatAmt# + TransHRec(1).MatchRetAmt)
      CASE ELSE
      END SELECT

      IF TransHRec(1).PrevTransRec <= 0 THEN
        IF (RetAmt# = 0 AND MatAmt# = 0) AND RetWage# > 0 THEN
          GOSUB SCPrintEmpNORetLine
          IF LineCnt >= MaxLines THEN
            PRINT #1, FF$
            GOSUB SCNORetRptHeader
          END IF
        END IF
        EXIT DO
      ELSE
        TransRecNum& = CLNG(TransHRec(1).PrevTransRec)
      END IF
    LOOP

SCSkipNOEm:
    ShowPctComp RecNo, NumOfRecs
  NEXT

  GOSUB SCNORetTotals

'****************************************************************

  FClose THandle
  FClose DHandle   'open employee data file
  PRINT #1, PrnDef$(1);
  CLOSE

  ERASE TransHRec, EMP2Rec, Unit, Pg, Dash, IdxBuff
  PrintRptFile RptTitle$, RptName$, 1, RetCode, 0

EXIT SUB


SCPrintEmpRetLine:
  PRINT #1, LEFT$(EMP2Rec(1).EMPSSN, 3); "-"; MID$(EMP2Rec(1).EMPSSN, 4, 2); "-"; MID$(EMP2Rec(1).EMPSSN, 6, 4);
  PRINT #1, TAB(16); QPTrim$(EMP2Rec(1).EMPLNAME); ", "; QPTrim$(EMP2Rec(1).EMPFNAME);
  PRINT #1, TAB(42); USING "#####,.##"; RetWage#;
  PRINT #1, TAB(56); USING "#####,.##"; RetAmt#;
  PRINT #1, TAB(68); USING "#####,.##"; MatAmt#
  LineCnt = LineCnt + 1     'employeesprinted = employeesprinted + 1
  RTotal# = Round(RTotal# + RetAmt#)
  WTotal# = Round(WTotal# + RetWage#)
  ETotal# = Round(ETotal# + MatAmt#)

RETURN

SCPrintEmpNORetLine:
  PRINT #1, LEFT$(EMP2Rec(1).EMPSSN, 3); "-"; MID$(EMP2Rec(1).EMPSSN, 4, 2); "-"; MID$(EMP2Rec(1).EMPSSN, 6, 4);
  PRINT #1, TAB(16); QPTrim$(EMP2Rec(1).EMPLNAME); ", "; QPTrim$(EMP2Rec(1).EMPFNAME);
  PRINT #1, TAB(47); USING "#####,.##"; RetWage#
  LineCnt = LineCnt + 1     'employeesprinted = employeesprinted + 1
  NWTotal# = Round(NWTotal# + RetWage#)

RETURN

SCRetTotals:
  PRINT #1, Dash(1)
  PRINT #1, TAB(28); "Totals:";
  PRINT #1, TAB(40); USING "#######,.##"; WTotal#;
  PRINT #1, TAB(54); USING "#######,.##"; RTotal#;
  PRINT #1, TAB(68); USING "#####,.##"; ETotal#
  PRINT #1, FF$
RETURN

SCRetGTotals:
  PRINT #1, Dash(1)
  PRINT #1, TAB(28); "Totals:";
  PRINT #1, TAB(40); USING "#######,.##"; WTotal#;
  PRINT #1, TAB(54); USING "#######,.##"; RTotal#;
  PRINT #1, TAB(68); USING "#####,.##"; ETotal#
  PRINT #1, TAB(22); "Grand Totals:";
  PRINT #1, TAB(40); USING "#######,.##"; Round#(GWTotal# + WTotal#);
  PRINT #1, TAB(54); USING "#######,.##"; Round#(GRTotal# + RTotal#);
  PRINT #1, TAB(68); USING "#####,.##"; Round#(GETotal# + ETotal#)
  PRINT #1, FF$
RETURN

SCNORetTotals:
  PRINT #1, Dash(1)
  PRINT #1, TAB(28); "Totals:";
  PRINT #1, TAB(45); USING "#######,.##"; NWTotal#
  'PRINT #1, TAB(54); USING "#######,.##"; RTotal#;
  'PRINT #1, TAB(68); USING "#####,.##"; ETotal#
  PRINT #1, FF$
RETURN

SCRetRptHeader:
  Page = Page + 1
  PRINT #1, "S.C. Retirement System Report"; TAB(68); "Page:"; Page
  PRINT #1, "Monthly Report of Subject Wages and Retirement Contributions."
  PRINT #1, "Report Date:  "; Num2Date(LowDate); " to "; Num2Date(HighDate)
  PRINT #1, UTemp$
  IF Pass = 1 THEN
    PRINT #1, "General";
  ELSE
    PRINT #1, "Law Enforcement";
  END IF
  PRINT #1, TAB(37); "Wages Subject    Retirement    Employer"
  PRINT #1, "Soc Sec #      Employee Name         to Retirement    Deductions    Matching"
  PRINT #1, Dash(1)
  LineCnt = 7
RETURN

SCNORetRptHeader:
  Page = Page + 1
  PRINT #1, "S.C. Retirement System Report"; TAB(68); "Page:"; Page
  PRINT #1, "Monthly Report of Subject Wages and Retirement Contributions."
  PRINT #1, "Report Date:  "; Num2Date(LowDate); " to "; Num2Date(HighDate)
  PRINT #1, UTemp$
  PRINT #1,
  PRINT #1, "Soc Sec #      Employee Name         NON-Retirement Wages"
  PRINT #1, Dash(1)
  LineCnt = 6
RETURN

END SUB

SUB VARetirementRpt

  SHARED PrnDef$()

  FirstTime = True

  REDIM TempScrn(1)
  REDIM Frm(1) AS FormInfo
  FormName$ = "ERETIRPT"
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

  RptName$ = "PRRPTS\RETIRE.RPT"

  REDIM TransHRec(1)   AS TransRecType
  REDIM EMP2Rec(1)     AS EmpData2Type
  REDIM Unit(1)        AS UnitFileRecType
  REDIM Pg(1)          AS STRING * 3
  REDIM Dash(1)        AS STRING * 78

  Dash(1) = STRING$(78, "-")

  EmpRecSize = LEN(EMP2Rec(1))
  TRecSize = LEN(TransHRec(1))

  IdxRecLen = 2
  IdxFileSize& = FileSize(EmpIdxLName)
  NumOfRecs = IdxFileSize& \ IdxRecLen

  REDIM IdxBuff(1 TO NumOfRecs)
'  REDIM RetList(1 TO 6) AS RetireRecType

  FGetAH EmpIdxLName, IdxBuff(1), IdxRecLen, NumOfRecs
  FGetAH UnitFileName, Unit(1), LEN(Unit(1)), 1

'  FGetAH RetireFileName, RetList(1), LEN(RetList(1)), 6

  MaxLines = 55

'-=-=-=-=-=-=-=-=-=-=-=

  RptTitle$ = "Employee Retirement Report."
  ShowProcessingScrn RptTitle$

  TCol = 40 - (LEN(RptTitle$) \ 2) + 1
  PctRow = 11

'  QPrintRC "Pass 1 of 2", PctRow, TCol, 112

'  OPEN RptName$ FOR OUTPUT AS #1

  FCreate RptName$
  FOpenS RptName$, RHandle
  RPTSetupPRN 11, RHandle
  FClose RHandle

  OPEN RptName$ FOR APPEND AS #1

  FOpenS TransHistFileName, THandle
  FOpenS EmpData2Name, DHandle   'open employee data file

  GOSUB oLRetRptHeader

  FOR RecNo = 1 TO NumOfRecs
    UsingThisOne = False
    RetAmt# = 0
    RetWage# = 0
    MatAmt# = 0
    FGetRTA DHandle, EMP2Rec(1), CLNG(IdxBuff(RecNo)), EmpRecSize

    IF EMP2Rec(1).LastTransRec <= 0 THEN
      GOTO oSkipEm6
    END IF

    TransRecNum& = CLNG(EMP2Rec(1).LastTransRec)

    DO
      FGetRTA THandle, TransHRec(1), TransRecNum&, TRecSize
      SELECT CASE TransHRec(1).CheckDate
      CASE LowDate TO HighDate
        IF TransHRec(1).RetireAmt <> 0 THEN
          RetAmt# = Round(RetAmt# + TransHRec(1).RetireAmt)
          RetWage# = Round(RetWage# + TransHRec(1).RetGrossPay)
          MatAmt# = Round(MatAmt# + TransHRec(1).MatchRetAmt)
          UsingThisOne = True
        END IF
      CASE ELSE
      END SELECT

      IF TransHRec(1).PrevTransRec <= 0 THEN
        IF UsingThisOne THEN
          'GOSUB GetRetType
          GOSUB oPrintEmpRetLine
          IF LineCnt >= MaxLines THEN
            PRINT #1, FF$
            GOSUB oLRetRptHeader
          END IF
        END IF
        EXIT DO
      ELSE
        TransRecNum& = CLNG(TransHRec(1).PrevTransRec)
      END IF
    LOOP

oSkipEm6:
    ShowPctComp RecNo, NumOfRecs
  NEXT
  GOSUB oRetLTotals

  FClose THandle
  FClose DHandle   'open employee data file
  PRINT #1, PrnDef$(1);
  CLOSE

  ERASE TransHRec, EMP2Rec, Unit, Pg, Dash, IdxBuff

  PrintRptFile RptTitle$, RptName$, 1, RetCode, 0

EXIT SUB


oPrintEmpRetLine:
  PRINT #1, LEFT$(EMP2Rec(1).EMPSSN, 3); "-"; MID$(EMP2Rec(1).EMPSSN, 4, 2); "-"; MID$(EMP2Rec(1).EMPSSN, 6, 4);
  PRINT #1, TAB(16); QPTrim$(EMP2Rec(1).EMPLNAME); ", "; QPTrim$(EMP2Rec(1).EMPFNAME);
  PRINT #1, TAB(42); USING "#####,.##"; RetWage#;
  PRINT #1, TAB(56); USING "#####,.##"; RetAmt#;
  PRINT #1, TAB(68); USING "#####,.##"; MatAmt#
  LineCnt = LineCnt + 1     'employeesprinted = employeesprinted + 1
  RTotal# = Round(RTotal# + RetAmt#)
  WTotal# = Round(WTotal# + RetWage#)
  ETotal# = Round(ETotal# + MatAmt#)

RETURN

oRetLTotals:
  PRINT #1, Dash(1)
  PRINT #1, TAB(28); "Totals:";
  PRINT #1, TAB(40); USING "#######,.##"; WTotal#;
  PRINT #1, TAB(54); USING "#######,.##"; RTotal#;
  PRINT #1, TAB(68); USING "#####,.##"; ETotal#
  PRINT #1, FF$
RETURN

oLRetRptHeader:
  PRINT #1, "Retirement Deduction Report"
  PRINT #1, ""
  PRINT #1, "Report Date:  "; Num2Date(LowDate); " to "; Num2Date(HighDate)
  PRINT #1,
  PRINT #1, "                                     Wages Subject    Retirement    Employer"
  PRINT #1, "Soc Sec #      Employee Name         to Retirement    Deductions    Matching"
  PRINT #1, Dash(1)
  LineCnt = 7
RETURN

END SUB

