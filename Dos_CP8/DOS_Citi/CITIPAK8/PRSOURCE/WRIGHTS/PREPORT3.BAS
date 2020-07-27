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
'$INCLUDE: 'PREmpRec.biw'
'$INCLUDE: 'DedCodes.Bi'
'$INCLUDE: 'ErnCodes.Bi'
'$INCLUDE: 'PRUNIT.Bi'
'$INCLUDE: 'PRSYSCTR.Bi'
'$INCLUDE: 'PRPPDEF.Bi'
'$INCLUDE: 'PRTRANS.Biw'
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

  TYPE ChkInfoType
    ChkDate AS INTEGER
    ChkNum  AS LONG
    RecNo   AS LONG
    ERecNo  AS LONG
  END TYPE

SUB CheckIssuedRpt

  REDIM DftChks(0)  AS ChkInfoType
  REDIM ChkChks(0)  AS ChkInfoType
  ChkInfoLen = LEN(ChkChks(0))

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
  Image2$ = ",#########.##"

  REDIM Dash(1) AS STRING * 80
  REDIM TransHRec(1)   AS TransRecType
  REDIM Emp2Rec(1)     AS EmpData2Type
  REDIM Unit(1)        AS UnitFileRecType
  REDIM Pg(1)          AS STRING * 3

  ToPrint$ = SPACE$(78)
  ChkNum$ = SPACE$(12)
  
  MaxLines = 57
  LineCnt = 0
  Dash(1) = STRING$(78, "-") + CrLf$
  
  EmpRecSize = LEN(Emp2Rec(1))
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
    FGetRTA DHandle, Emp2Rec(1), CLNG(IdxBuff(RecNo)), EmpRecSize
    IF Emp2Rec(1).LastTransRec <= 0 THEN
      GOTO SkipEm4
    END IF
    TransRecNum& = CLNG(Emp2Rec(1).LastTransRec)

    DO
      FGetRTA THandle, TransHRec(1), TransRecNum&, TRecSize
      SELECT CASE TransHRec(1).CheckDate
      CASE LowDate TO HiDate
        SELECT CASE TransHRec(1).DraftFlag
        CASE True
          DCnt = DCnt + 1
          REDIM PRESERVE DftChks(0 TO DCnt) AS ChkInfoType
          DftChks(DCnt).RecNo = TransRecNum&
          DftChks(DCnt).ChkNum = TransHRec(1).CheckNum
          DftChks(DCnt).ChkDate = TransHRec(1).CheckDate
          DftChks(DCnt).ERecNo = IdxBuff(RecNo)
        CASE False
          CCnt = CCnt + 1
          REDIM PRESERVE ChkChks(0 TO CCnt) AS ChkInfoType
          ChkChks(CCnt).RecNo = TransRecNum&
          ChkChks(CCnt).ChkNum = TransHRec(1).CheckNum
          ChkChks(CCnt).ChkDate = TransHRec(1).CheckDate
          ChkChks(CCnt).ERecNo = IdxBuff(RecNo)
        END SELECT
      END SELECT
      IF TransHRec(1).PrevTransRec <= 0 THEN    'if no more transactions
        EXIT DO
      ELSE      'there are more trans to check
        TransRecNum& = CLNG(TransHRec(1).PrevTransRec)
      END IF
    LOOP
SkipEm4:
    ShowPctComp RecNo, NumOfRecs
  NEXT

'sort Drafts by numbers within each date
  IF DCnt > 0 THEN
    SortT DftChks(1), DCnt, 0, ChkInfoLen, 0, -1
    Start = 1
    FOR zz = 1 TO DCnt
      IF zz = 1 THEN
        ThisDate = DftChks(1).ChkDate
      ELSE
        IF DftChks(zz).ChkDate <> ThisDate THEN
          Last = zz - 1
          SortT DftChks(Start), (Last - Start) + 1, 0, ChkInfoLen, 2, -2
          ThisDate = DftChks(zz).ChkDate
          Start = zz
        END IF
      END IF
    NEXT
    Last = zz - 1
    SortT DftChks(Start), (Last - Start) + 1, 0, ChkInfoLen, 2, -2
  END IF

'next sort Checks by numbers within each date
'first sortem by check/draft date
  IF CCnt > 0 THEN
    SortT ChkChks(1), CCnt, 0, ChkInfoLen, 0, -1
    Start = 1
    FOR zz = 1 TO CCnt
      IF zz = 1 THEN
        ThisDate = ChkChks(1).ChkDate
      ELSE
        IF ChkChks(zz).ChkDate <> ThisDate THEN
          Last = zz - 1
          SortT ChkChks(Start), (Last - Start) + 1, 0, ChkInfoLen, 2, -2
          ThisDate = ChkChks(zz).ChkDate
          Start = zz
        END IF
      END IF
    NEXT
    Last = zz - 1
    SortT ChkChks(Start), (Last - Start) + 1, 0, ChkInfoLen, 2, -2
  END IF

  ChkTot# = 0
  FOR Cnt = 1 TO CCnt
    FGetRTA THandle, TransHRec(1), ChkChks(Cnt).RecNo, TRecSize
    FGetRTA DHandle, Emp2Rec(1), ChkChks(Cnt).ERecNo, EmpRecSize
    ChkTot# = Round#(ChkTot# + TransHRec(1).NetPay)
    GOSUB PrintChkInfo
    LineCnt = LineCnt + 1
    IF LineCnt > MaxLines THEN            'bottom of page?
      FPut RHandle, FF$   'yes; form feed
      GOSUB PrintIssuedHeader              'write title lines
    END IF
    ShowPctComp Cnt, CCnt
  NEXT

  FPut RHandle, Dash(1)
  LSET ToPrint$ = "Total of Checks: " + STR$(CCnt)
  MID$(ToPrint$, 65) = FUsing(STR$(ChkTot#), Image2$)
  FPut RHandle, ToPrint$ + CrLf$
  FPut RHandle, FF$   'yes; form feed

  GOSUB PrintIssuedHeader              'write title lines

  DftTot# = 0
  DftFlag = True
  FOR Cnt = 1 TO DCnt
    FGetRTA THandle, TransHRec(1), DftChks(Cnt).RecNo, TRecSize
    FGetRTA DHandle, Emp2Rec(1), DftChks(Cnt).ERecNo, EmpRecSize
    DftTot# = Round#(DftTot# + TransHRec(1).NetPay)
    GOSUB PrintChkInfo
    LineCnt = LineCnt + 1
    IF LineCnt > MaxLines THEN            'bottom of page?
      FPut RHandle, FF$   'yes; form feed
      GOSUB PrintIssuedHeader              'write title lines
    END IF
    ShowPctComp Cnt, DCnt
  NEXT

  FPut RHandle, Dash(1)
  LSET ToPrint$ = "Total of Drafts: " + STR$(DCnt)
  MID$(ToPrint$, 65) = FUsing(STR$(DftTot#), Image2$)
  FPut RHandle, ToPrint$ + CrLf$
  FPut RHandle, FF$   'yes; form feed
  RPTSetupPRN 0, RHandle

  FClose DHandle
  FClose THandle
  FClose RHandle

  ERASE Dash, TransHRec, Emp2Rec, Unit, Pg, IdxBuff

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
  FPut RHandle, "Check Date    Check/Draft No.     Employee Name                       Net Pay" + CrLf$
  FPut RHandle, Dash(1)
  LineCnt = 5
RETURN

PrintChkInfo:
  LSET ToPrint$ = Num2Date(TransHRec(1).CheckDate)
  RSET ChkNum$ = QPStrL(TransHRec(1).CheckNum)
  MID$(ToPrint$, 15) = ChkNum$
  IF DftFlag THEN
    MID$(ToPrint$, 15) = "Dft"
  ELSE
    MID$(ToPrint$, 15) = "Chk"
  END IF
  MID$(ToPrint$, 35) = QPTrim$(Emp2Rec(1).EMPFNAME) + " " + QPTrim$(Emp2Rec(1).EMPLNAME)
  MID$(ToPrint$, 70) = FUsing(STR$(TransHRec(1).NetPay), Image1$)
  FPut RHandle, ToPrint$ + CrLf$
  LineCnt = LineCnt + 1
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
  REDIM Emp2Rec(1)     AS EmpData2Type
  REDIM Unit(1)        AS UnitFileRecType
  REDIM Pg(1)          AS STRING * 3
  REDIM Dash(1)        AS STRING * 78

  Dash(1) = STRING$(78, "-")

  EmpRecSize = LEN(Emp2Rec(1))
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
    RETAMT# = 0
    RetWage# = 0
    MatAmt# = 0
    FGetRTA DHandle, Emp2Rec(1), CLNG(IdxBuff(RecNo)), EmpRecSize

    IF Emp2Rec(1).LastTransRec <= 0 THEN
      GOTO oSkipEm6
    END IF

    TransRecNum& = CLNG(Emp2Rec(1).LastTransRec)

    DO
      FGetRTA THandle, TransHRec(1), TransRecNum&, TRecSize
      SELECT CASE TransHRec(1).CheckDate
      CASE LowDate TO HighDate
        IF TransHRec(1).RetireAmt <> 0 THEN
          RETAMT# = Round(RETAMT# + TransHRec(1).RetireAmt)
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

  ERASE TransHRec, Emp2Rec, Unit, Pg, Dash, IdxBuff

  PrintRptFile RptTitle$, RptName$, 1, RetCode, 0

EXIT SUB


oPrintEmpRetLine:
  PRINT #1, LEFT$(Emp2Rec(1).EMPSSN, 3); "-"; MID$(Emp2Rec(1).EMPSSN, 4, 2); "-"; MID$(Emp2Rec(1).EMPSSN, 6, 4);
  PRINT #1, TAB(16); QPTrim$(Emp2Rec(1).EMPLNAME); ", "; QPTrim$(Emp2Rec(1).EMPFNAME);
  PRINT #1, TAB(42); USING "#####,.##"; RetWage#;
  PRINT #1, TAB(56); USING "#####,.##"; RETAMT#;
  PRINT #1, TAB(68); USING "#####,.##"; MatAmt#
  LineCnt = LineCnt + 1     'employeesprinted = employeesprinted + 1
  RTotal# = Round(RTotal# + RETAMT#)
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

