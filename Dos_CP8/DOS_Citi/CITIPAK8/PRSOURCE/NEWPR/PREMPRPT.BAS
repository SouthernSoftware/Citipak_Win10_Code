DEFINT A-Z
DECLARE SUB RPTSetupPRN (RPTNum%, FHandle%)

DECLARE SUB GetEmpNumbs (LowEmp&, HiEmp&)
DECLARE FUNCTION EmpAsk2Delete% ()
DECLARE SUB EMPrintTermEmpList ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)

DECLARE SUB EMPShowTransHist (RecNo%)

DECLARE FUNCTION CheckFldEmpty% (FldData$)
DECLARE SUB EMPLoadRFldData (ReqFld%(), ReqNumOfFlds%)
DECLARE SUB EMPShowYTDTotals (RecNo%)

DECLARE SUB EMPrintEmpData ()
DECLARE SUB EMPrintEmpList ()
DECLARE FUNCTION EMPNumNull% (Form$())
DECLARE SUB EmpLookUpEdit ()
DECLARE SUB EMPEditRec (RecNo%)
DECLARE SUB InitEmpMulti (Fld() AS ANY)
DECLARE SUB EmpSaveNewRec (RecNo%, Form$())
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB WaitForAction ()
DECLARE FUNCTION EmpNumOfRecs% ()

DECLARE SUB FCreate (FileName$)
DECLARE SUB FOpenS (FileName$, Handle%)
'DECLARE SUB FOpenAll (FileName$, AccessMode%, ShareMode%, Handle%)
'DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FClose (Handle%)

DECLARE FUNCTION DOSError% ()
DECLARE FUNCTION FLInput$ (Handle%, Buff$)
DECLARE FUNCTION WhichError% ()

DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FPutRT (Handle%, Source AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutRTA (Handle%, SEG Source AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutT (Handle%, Source AS ANY, NumBytes%)
DECLARE SUB FPut (Handle%, ToDisk$)

DECLARE SUB FGetA (Handle%, SEG Element AS ANY, NumOfBytes&)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetRT (Handle%, Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FGetT (Handle%, Dest AS ANY, NumBytes%)

DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB FSeek (Handle%, Location&)

DECLARE SUB FFlush (Handle%)

DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FEof% (Handle%)
DECLARE FUNCTION Round# (DoubleNum#)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION FLoc& (Handle%)
DECLARE FUNCTION FLof& (Handle%)
DECLARE SUB InsertT (SEG StartElement AS ANY, ElSize%, NumEls%)
DECLARE SUB DeleteT (SEG StartElement AS ANY, ElSize%, NumEls%)

DECLARE SUB MScrnSave (ULRow, ULCol, LRRow, LRCol, SEG Address)
DECLARE SUB MScrnRest (ULRow, ULCol, LRRow, LRCol, SEG Address)
DECLARE SUB StuffBuf (Ky$)
DECLARE SUB GetCursor (X, y, Button)
DECLARE SUB MQPrint (X$, Colr)
DECLARE SUB MPaintBox (ULRow, ULCol, LRRow, LRCol, Colr)
DECLARE SUB ButtonPress (ButNo%, Down%, Presses%, X%, y%)
DECLARE SUB HideCursor ()
DECLARE SUB TextCursor (Fg, Bg)
DECLARE SUB ShowCursor ()
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)

DECLARE FUNCTION LibNumberOfFields (LibName$, FrmName$)
DECLARE FUNCTION LibScrName$ (Scr, ScrnLib())
DECLARE FUNCTION ScrnLibSize (LibName$)
DECLARE SUB EditForm (Form$(), Fld() AS ANY, Frm AS ANY, Cnf AS ANY, Action)
DECLARE SUB Lib2Scrn (NameInLib$, ScrnLib(), MonoCode, Attribute, ErrorCode)
DECLARE SUB LibGetFldDef (LibName$, FrmName$, StartEl, Fld() AS ANY, Form$(), ErrCode)
DECLARE SUB LoadScrnLib (LibName$, ScrnLib(), ErrorCode)
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB PrintArray (FirstFld, LastFld, Form$(), Fld() AS ANY)
DECLARE SUB UnPackBuffer (FirstFld, LastFld, Form$(), Fld() AS ANY)
DECLARE FUNCTION FldNum (FldName$, Fld() AS ANY)
DECLARE FUNCTION EndOfForms (Fld() AS ANY)
DECLARE FUNCTION PDQTimer& ()
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE SUB SearchT2 (SEG Array AS ANY, NumEls%, Match, Found%, Dir%, Code%, StructSize%, MemberOff%, MemberSize%)
'DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPRTrim$ (Text$)
DECLARE FUNCTION PromptSaveData% ()
DECLARE SUB PayMenu (BYVAL MenuNum%, Choice%, NumOfItems%)

DECLARE SUB VertMenu (Item$(), Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB DisplayMiscScrn (ScrnName$)
DECLARE SUB CursorOff ()
DECLARE SUB RestScrn (ScrnArray%())
DECLARE SUB SaveScrn (ScrnArray%())
DECLARE FUNCTION ASCII% (Strng$)
DECLARE FUNCTION WEnvTest% ()
DECLARE FUNCTION Monitor% ()
DECLARE SUB Pause (Ticks)
DECLARE FUNCTION Compare3% (BYVAL Seg1, BYVAL Adr1, BYVAL Seg2, BYVAL Adr2, NumBytes)
DECLARE SUB ClearScrn ()
DECLARE SUB SearchDataFile (LookUpType%, Form$(), MatchCnt%)
DECLARE SUB EMPChooseList (MatchCnt%, LookUpType%)
DECLARE SUB Sort (Array$(), Direction%)
DECLARE FUNCTION EmpNumOfRecs% ()
DECLARE SUB EmpSaveOldRec (RecNo%, Form$())
DECLARE SUB BlockClear ()
DECLARE FUNCTION EMPCheckEMPNumber% (EMPNumFld$, Frm AS ANY)
DECLARE SUB EMPAddShowName (First$, Last$)
DECLARE SUB SaveField (FldNo%, Form$(), Fld() AS ANY, BadFld%)

DECLARE FUNCTION QPStrI$ (Num%)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPStrL$ (LongNum&)
DECLARE FUNCTION QPValI (NumString$)
DECLARE FUNCTION QPValL& (NumString$)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION Num2Date$ (DateNum%)


'$INCLUDE: 'DefCnf.bi'                  'Defines TYPE for monitor/color info.
'$INCLUDE: 'FieldInf.bi'                '        "        field information
'$INCLUDE: 'FormEdit.bi'                '        "        form editing info.
'$INCLUDE: 'PRFiles.bi'
'$INCLUDE: 'PREmpRec.bi'
'$INCLUDE: 'PageInfo.bi'    '        "          Page information
'$INCLUDE: 'Retire.Bi'
'$INCLUDE: 'DedCodes.Bi'
'$INCLUDE: 'ERNCodes.Bi'
'$INCLUDE: 'PRTRANS.Bi'
'$INCLUDE: 'PRCHECK.Bi'
'$INCLUDE: 'PRUNIT.Bi'


CONST False = 0
CONST True = NOT False

SUB EMPrintEmpData

  CursorOff

  RptName$ = "PRRPTS\EMPDATA.RPT"

  FldFlag$ = CHR$(27)

  DescFlag$ = CHR$(25)

  IdxRecLen = 2
  IdxFileSize& = FileSize(EmpIdxLName)
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen

 ' DIM SHARED EmpData1Rec   AS EmpData1Type
 ' DIM SHARED EmpData3Rec   AS EmpData3Type

  REDIM Emp2Data(1)    AS EmpData2Type
'-Emp2Data(1)

  Emp2RecLen = LEN(Emp2Data(1))
  DataFileSize& = FileSize(EmpData2Name)
  DataNumOfRecs& = DataFileSize& \ Emp2RecLen

  REDIM TempScrn(0)

  REDIM Frm(1)    AS FormInfo

  REDIM Desc$(1 TO 15)
  REDIM DedCodeRec(1 TO 12)  AS DedCodeRecType
  REDIM ERNCodeRec(1 TO 3)   AS ErnCodeRecType
  
  FGetAH DedCodeFileName, DedCodeRec(1), LEN(DedCodeRec(1)), 12
  FGetAH ErnCodeFileName, ERNCodeRec(1), LEN(ERNCodeRec(1)), 3

  FOR Cnt = 1 TO 12
    Desc$(Cnt) = DedCodeRec(Cnt).DCDESC1
    IF LEN(Desc$(Cnt)) = 0 THEN
      Desc$(Cnt) = " "
    END IF
  NEXT
  FOR Cnt = 13 TO 15
    Desc$(Cnt) = ERNCodeRec(Cnt - 12).ERNCODE1
    IF LEN(Desc$(Cnt)) = 0 THEN
      Desc$(Cnt) = " "
    END IF
  NEXT

'**********************************************************************

  NumOfScrns = 3

  REDIM FormName$(1 TO NumOfScrns)
  FOR Cnt = 1 TO NumOfScrns
    FormName$(Cnt) = "EMPMA" + LTRIM$(STR$(Cnt))
  NEXT
  NumFlds = -1

  FOR Cnt = 1 TO NumOfScrns           '-- Get total number of fields
    NumFlds = NumFlds + LibNumberOfFields(EmployeeQLib, FormName$(Cnt)) + 1
  NEXT

  REDIM Form$(NumFlds, 2)           'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo   'DIM the field information array

  StartEl = 0                         'Load first form at array start
  FOR Cnt = 1 TO NumOfScrns           'Get field defs
    LibGetFldDef EmployeeQLib, FormName$(Cnt), StartEl, Fld(), Form$(), ErrCode
    StartEl = StartEl + Fld(StartEl).Fields + 1 'Inc StartEl to next
  NEXT

  FOR n = 1 TO NumFlds             'For all fields on both pages
    LSET Form$(n, 0) = ""        'Clear each field
  NEXT

  Frm(1).StartEl = 0               'Set form starting element to 0 and
'--------------------------------------------------------------------

  REDIM PrintBuff$(1 TO 100)

  BufSize = 82                            'enough to read up to 80 chars + CRLF
  Buffer$ = SPACE$(BufSize)               '  (or use whatever is appropriate)

  LineCnt = 1
  FOpenS EmpDataFileMask, Handle           'open for QuickPak Professional Binary

  DO                                      'the main file read loop
    Temp$ = FLInput$(Handle, Buffer$)    'get a line of input
    WHILE WhichError% = 83               'a "Buffer too small" error means
      Temp$ = Temp$ + FLInput$(Handle, Buffer$)     'we have to go back
    WEND                                             'and get some more
    IF DOSError% THEN EXIT DO            'exit if error (probably end of file)
    PrintBuff$(LineCnt) = Temp$
    LineCnt = LineCnt + 1
  LOOP

  FClose Handle                           'always remember to close the file

  REDIM PRESERVE PrintBuff$(1 TO LineCnt)

  LinesPerRec = UBOUND(PrintBuff$)

  FirstDesc = 1
  FOR Cnt = 1 TO LinesPerRec
    DescPos = INSTR(PrintBuff$(Cnt), DescFlag$)
    IF DescPos THEN
      MID$(PrintBuff$(Cnt), DescPos) = Desc$(FirstDesc)
      FirstDesc = FirstDesc + 1
      IF FirstDesc > 15 THEN EXIT FOR
    END IF
  NEXT

  REDIM IdxBuff(1 TO IdxNumOfRecs) AS EmployeeIndexType  'load index file
  FGetAH EmpIdxLName, IdxBuff(1), IdxRecLen, IdxNumOfRecs

  Form$(0, 0) = SPACE$(Emp2RecLen)

  RptTitle$ = "Employee Information Listing"
  ShowProcessingScrn RptTitle$

  FCreate RptName$
  FOpenS RptName$, RptHandel

  RPTSetupPRN 9, RptHandel

  FOpenS EmpData2Name, D2Handle

  ToPrint$ = SPACE$(80)

  FOR Cnt = 1 TO IdxNumOfRecs
    FGetRTA D2Handle, Emp2Data(1), CLNG(IdxBuff(Cnt).DataRecNum), Emp2RecLen
    IF NOT Emp2Data(1).Deleted THEN
      'LSET Form$(0, 0) = ""              'clear old data and refresh
      BCopy VARSEG(Emp2Data(1)), VARPTR(Emp2Data(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), Emp2RecLen, 0
      UnPackBuffer 0, 0, Form$(), Fld()
      GOSUB PrintEmpData
    END IF
    ShowPctComp Cnt, IdxNumOfRecs
  NEXT

  RPTSetupPRN 0, RptHandel
  FClose D2Handle
  FClose RptHandel

  ERASE TempScrn, Desc$, DedCodeRec, ERNCodeRec
  ERASE Form$, Fld, Frm, PrintBuff$, Emp2Data

  PrintRptFile RptTitle$, RptName$, 1, RetCode, 0

EXIT SUB

PrintEmpData:


  FldOffSet = 1

  FOR Cnt2 = 1 TO LinesPerRec
    LSET ToPrint$ = PrintBuff$(Cnt2)
    DO
      FldPos = INSTR(ToPrint$, FldFlag$)
      IF FldPos THEN
        RptFldNo = QPValI(MID$(ToPrint$, FldPos + 1))
        MID$(ToPrint$, FldPos) = "    "
        MID$(ToPrint$, FldPos) = Form$(RptFldNo, 0)
      END IF
    LOOP WHILE FldPos
    FPut RptHandel, QPRTrim$(ToPrint$) + CrLf$
  NEXT

  FPut RptHandel, FF$

RETURN

END SUB

SUB EMPrintEmpList

  REDIM TempScrn(1)
  
  ExitFlag = False
  FirstTime = True
  
  OKFlag = False
  
  REDIM Frm(1) AS FormInfo
  
  FormName$ = "ELISTRPT"
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
    
    SELECT CASE Frm(1).KeyCode
    CASE F0Key
      AN$ = QPTrim$(Form$(1, 0))
      IF LEN(AN$) = 0 THEN 'Year$ = Form$(2, 0)
        CursorOff
        SaveScrn TempScrn()
        DisplayMiscScrn ReportParmError
        WaitForAction
        RestScrn TempScrn()
        Frm(1).FldNo = 1
        Action = 2
        OKFlag = False
      ELSE
        OKFlag = True
      END IF

    CASE EscKey
      OKFlag = True
      ExitFlag = True
    END SELECT
    
  LOOP UNTIL OKFlag             'proper key not set
  
  CursorOff
  
  IF ExitFlag THEN
    GOTO ExitListRpt
  ELSE
    OKFlag = False
  END IF

  'IF ExitFlag THEN EXIT SUB

'EXIT SUB

  IF AN$ = "A" THEN
    GOTO AlphaList
  ELSE
    NumFlag = True
  END IF

  GetEmpNumbs LowEmp&, HiEmp&

  FormName$ = "ELISTNUM"
  NumFlds = LibNumberOfFields(MiscQLib, FormName$)
  REDIM Form$(NumFlds, 2)       'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo               'DIM the field information array
  StartEl = 0   'Load first form at array start
  LibGetFldDef MiscQLib, FormName$, StartEl, Fld(), Form$(), ErrCode


  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1

  '----- Setup TYPE for setting and reading form editing information.
  FirstTime = True
  Frm(1).FldNo = 1              'Start editing on field #1
  Frm(1).InsStat = False        'Set insert state (True = Insert on)
  Frm(1).StartEl = 0            'Set form starting element to 0 and

  BlockClear
  DisplayMiscScrn FormName$

  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      LSET Form$(1, 0) = QPStrL(LowEmp&)
      LSET Form$(2, 0) = QPStrL(HiEmp&)
      Action = 2
    END IF

    SELECT CASE Frm(1).KeyCode
    CASE F0Key

'HERE

      LowNum& = QPValL(Form$(1, 0))
      HiNum& = QPValL(Form$(2, 0))

      IF HiNum& < LowNum& THEN
        CursorOff
        SaveScrn TempScrn()
        DisplayMiscScrn ReportParmError
        WaitForAction
        RestScrn TempScrn()
        Frm(1).FldNo = 1
        Action = 2
        OKFlag = False
      ELSE
        IF LowNum& = 0 AND HiNum& = 0 THEN AllFlag = True
        OKFlag = True
      END IF

    CASE EscKey
      OKFlag = True
      ExitFlag = True
    END SELECT

  LOOP UNTIL OKFlag             'proper key not set


ExitListRpt:
  ERASE Form$, Fld, TempScrn, Frm
  IF ExitFlag THEN EXIT SUB

'******

AlphaList:
  Image1$ = ",####.##"

  REDIM Dash(1) AS STRING * 78
  REDIM Emp2Rec(1)     AS EmpData2Type
  REDIM Unit(1)        AS UnitFileRecType
  REDIM Pg(1)          AS STRING * 3

  REDIM ENumb(1) AS STRING * 11
  REDIM EName(1) AS STRING * 23
  REDIM ETitle(1) AS STRING * 12
  REDIM BRate(1) AS STRING * 9

  REDIM SH(1) AS STRING * 3

  REDIM BDate(1) AS STRING * 10
  REDIM HDate(1) AS STRING * 10
  
  MaxLines = 55
  LineCnt = 0
  Dash(1) = STRING$(78, "-")
  EmpRecSize = LEN(Emp2Rec(1))

  IdxRecLen = 2
  IdxFileSize& = FileSize(EmpIdxLName)
  NumOfRecs = IdxFileSize& \ IdxRecLen

  REDIM IdxBuff(1 TO NumOfRecs)

  IF NumFlag THEN
    FGetAH EmpIdxNName, IdxBuff(1), IdxRecLen, NumOfRecs
  ELSE
    FGetAH EmpIdxLName, IdxBuff(1), IdxRecLen, NumOfRecs
  END IF

  FGetAH UnitFileName, Unit(1), LEN(Unit(1)), 1

  RptTitle$ = "Employee Listing"
  ShowProcessingScrn RptTitle$

  RptName$ = "PRRPTS\EMPLIST.RPT"

  FCreate RptName$
  FOpenS RptName$, RptHandel
  RPTSetupPRN 10, RptHandel
  FOpenS EmpData2Name, EHandle   'open employee data file

  GOSUB PrintEmpListHeader

  FOR RecNo = 1 TO NumOfRecs
    FGetRTA EHandle, Emp2Rec(1), CLNG(IdxBuff(RecNo)), EmpRecSize
    IF NOT Emp2Rec(1).Deleted THEN
      IF NumFlag AND NOT AllFlag THEN
        EmpNo& = QPValL(Emp2Rec(1).EmpNo)
        IF NOT (EmpNo& >= LowNum& AND EmpNo& <= HiNum&) THEN
           GOTO SkipThisOne
        END IF
      END IF
      IF Emp2Rec(1).EmpTDate > 0 THEN
         GOTO SkipThisOne
      END IF

      EmpCnt = EmpCnt + 1
      GOSUB PrintEmpListData
      IF LineCnt >= MaxLines THEN
        FPut RptHandel, FF$
        GOSUB PrintEmpListHeader
      END IF
    END IF

SkipThisOne:
    ShowPctComp RecNo, NumOfRecs
  NEXT

  GOSUB PrintEmpListTotals
  RPTSetupPRN 0, RptHandel
  FClose EHandle
  FClose RptHandel

  ERASE Dash, Emp2Rec, Unit, Pg, ENumb, EName, ETitle, BRate, SH, IdxBuff
  ERASE BDate, HDate

  PrintRptFile RptTitle$, RptName$, 1, RetCode, 0

EXIT SUB


PrintEmpListHeader:
  Page = Page + 1
  FPut RptHandel, Unit(1).UFEMPR + CrLf$
  FPut RptHandel, "Employee Listing" + CrLf$
  FPut RptHandel, "Report Date: " + DATE$ + "                                            Page: " + QPStrI$(Page) + CrLf$
  FPut RptHandel, CrLf$
  FPut RptHandel, "Number     Name                     SSN              Rate     B-Date    H-Date" + CrLf$
  FPut RptHandel, Dash$(1) + CrLf$
  LineCnt = 6
RETURN

PrintEmpListData:
   ENumb(1) = QPTrim$(Emp2Rec(1).EmpNo)
   EName(1) = QPTrim$(Emp2Rec(1).EMPLNAME) + ", " + QPTrim$(Emp2Rec(1).EMPFNAME)
   'ETitle(1) = QPTrim$(Emp2Rec(1).EMPJOB)
   ETitle(1) = LEFT$(Emp2Rec(1).EmpSSN, 3) + "-" + MID$(Emp2Rec(1).EmpSSN, 4, 2) + "-" + MID$(Emp2Rec(1).EmpSSN, 6, 4)

   RSET BRate(1) = QPTrim$(FUsing$(STR$(Emp2Rec(1).EMPPRATE), Image1$))
   MID$(SH(1), 2, 1) = LEFT$(QPTrim$(Emp2Rec(1).EMPPTYPE), 1)

   BDate$ = Num2Date(Emp2Rec(1).EMPBDAY)
   RSET BDate(1) = LEFT$(BDate$, 6) + RIGHT$(BDate$, 2)
   HDate$ = Num2Date(Emp2Rec(1).EMPHDATE)
   RSET HDate(1) = LEFT$(HDate$, 6) + RIGHT$(HDate$, 2)

   FPut RptHandel, ENumb(1) + EName(1) + ETitle(1) + BRate(1) + SH(1) + BDate(1) + HDate(1) + CrLf$
   LineCnt = LineCnt + 1
RETURN

PrintEmpListTotals:
  FPut RptHandel, Dash$(1) + CrLf$
  FPut RptHandel, "Total Employees: " + QPStrI(EmpCnt) + CrLf$
  FPut RptHandel, FF$
RETURN

END SUB

SUB EMPrintTermEmpList

  REDIM Dash(1) AS STRING * 75
  REDIM Emp2Rec(1)     AS EmpData2Type
  REDIM Unit(1)        AS UnitFileRecType
  REDIM Pg(1)          AS STRING * 3

  REDIM ENumb(1) AS STRING * 11
  REDIM EName(1) AS STRING * 26
  REDIM ETitle(1) AS STRING * 23
  REDIM TDate(1) AS STRING * 12

  MaxLines = 55
  LineCnt = 0
  Dash(1) = STRING$(75, "-")
  EmpRecSize = LEN(Emp2Rec(1))

  IdxRecLen = 2
  IdxFileSize& = FileSize(EmpIdxLName)
  NumOfRecs = IdxFileSize& \ IdxRecLen

  REDIM IdxBuff(1 TO NumOfRecs)

  FGetAH EmpIdxLName, IdxBuff(1), IdxRecLen, NumOfRecs
  FGetAH UnitFileName, Unit(1), LEN(Unit(1)), 1

  RptTitle$ = "Terminted Employee Listing"
  ShowProcessingScrn RptTitle$

  RptName$ = "PRRPTS\TERMLIST.RPT"

  FCreate RptName$
  FOpenS RptName$, RptHandel
  RPTSetupPRN 11, RptHandel
  FOpenS EmpData2Name, EHandle   'open employee data file

  GOSUB PrintTEmpListHeader

  FOR RecNo = 1 TO NumOfRecs
    FGetRTA EHandle, Emp2Rec(1), CLNG(IdxBuff(RecNo)), EmpRecSize
    IF Emp2Rec(1).EmpTDate > 0 THEN
      EmpCnt = EmpCnt + 1
      GOSUB PrintTEmpListData
      IF LineCnt >= MaxLines THEN
        FPut RptHandel, FF$
        GOSUB PrintTEmpListHeader
      END IF
    END IF
    ShowPctComp RecNo, NumOfRecs
  NEXT

  GOSUB PrintTEmpListTotals
  RPTSetupPRN 0, RptHandel
  FClose EHandle
  FClose RptHandel

  ERASE Dash, Emp2Rec, Unit, Pg, ENumb, EName, ETitle, TDate, IdxBuff

  PrintRptFile RptTitle$, RptName$, 1, RetCode, 0

EXIT SUB

PrintTEmpListHeader:
  Page = Page + 1
  FPut RptHandel, Unit(1).UFEMPR + CrLf$
  FPut RptHandel, "Terminted Employee Listing" + CrLf$
  FPut RptHandel, "Report Date: " + DATE$ + "                                            Page: " + QPStrI$(Page) + CrLf$
  FPut RptHandel, CrLf$
  FPut RptHandel, "Number     Name                      Title                    Termination" + CrLf$
  FPut RptHandel, Dash$(1) + CrLf$
  LineCnt = 6
RETURN

PrintTEmpListData:
   ENumb(1) = QPTrim$(Emp2Rec(1).EmpNo)
   EName(1) = QPTrim$(Emp2Rec(1).EMPLNAME) + ", " + QPTrim$(Emp2Rec(1).EMPFNAME)
   ETitle(1) = QPTrim$(Emp2Rec(1).EMPJOB)
   RSET TDate(1) = Num2Date(Emp2Rec(1).EmpTDate)
   FPut RptHandel, ENumb(1) + EName(1) + ETitle(1) + TDate(1) + CrLf$
   LineCnt = LineCnt + 1
RETURN

PrintTEmpListTotals:
  FPut RptHandel, Dash$(1) + CrLf$
  FPut RptHandel, "Total Employees: " + QPStrI(EmpCnt) + CrLf$
  FPut RptHandel, FF$
RETURN


END SUB

SUB GetEmpNumbs (LowEmp&, HiEmp&)
  NumOfRecs = FileSize(EMPNumFileName) \ 10
  REDIM EmpNumbers(1 TO NumOfRecs)  AS EmpNumType
  FGetAH EMPNumFileName, EmpNumbers(1), 10, NumOfRecs
  LowEmp& = 1000000
  HiEmp& = 0
  FOR Cnt = 1 TO NumOfRecs
    EmpNum& = QPValL(EmpNumbers(Cnt).EmpNum)
    IF EmpNum& < LowEmp& THEN LowEmp& = EmpNum&
    IF EmpNum& > HiEmp& THEN HiEmp& = EmpNum&
  NEXT
  ERASE EmpNumbers
END SUB

