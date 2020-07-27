DEFINT A-Z
DECLARE FUNCTION GetPassword$ (Row%, Col%, MaxLen%, Abort%)
DECLARE SUB PRW2Menu ()
DECLARE FUNCTION W2GetEmpName$ (EdRecNo%)
DECLARE SUB W2PrintForms ()
DECLARE SUB W2Report ()
DECLARE SUB W2EditRec (EdRecNo%)
DECLARE SUB W2ChooseEditW2 ()
DECLARE SUB W2LoadEmpList (GoodRecs$())
DECLARE SUB EMPEditRec (RecNo%)
DECLARE SUB ExtractW2Info (W2Type%)
  
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB WaitForAction ()
  
DECLARE SUB FCreate (FileName$)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FClose (Handle%)
  
DECLARE FUNCTION DOSError% ()
DECLARE FUNCTION WhichError% ()
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FClose (Handle%)
  
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
  
DECLARE SUB MScrnSave (ULRow, ULCol, LRRow, LRCol, SEG Address)
DECLARE SUB MScrnRest (ULRow, ULCol, LRRow, LRCol, SEG Address)
DECLARE SUB StuffBuf (Ky$)
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
DECLARE FUNCTION QPRTrim$ (Text$)
DECLARE FUNCTION PromptSaveData% ()
DECLARE SUB PayMenu (BYVAL MenuNum%, Choice%, NumOfItems%)
DECLARE SUB VertMenu (Item$(), Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB DisplayMiscScrn (ScrnName$)
DECLARE SUB CursorOff ()
DECLARE SUB RestScrn (ScrnArray%())
DECLARE SUB SaveScrn (ScrnArray%())
DECLARE FUNCTION ASCII% (Strng$)
DECLARE SUB Pause (Ticks)
DECLARE SUB ClearScrn ()
DECLARE SUB Sort (Array$(), Direction%)
DECLARE SUB BlockClear ()
DECLARE SUB SaveField (FldNo%, Form$(), Fld() AS ANY, BadFld%)
  
DECLARE SUB InitW2Multi (Fld() AS ANY)
  
DECLARE FUNCTION QPStrI$ (Num%)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPStrL$ (LongNum&)
DECLARE FUNCTION QPValI (NumString$)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num (DateIn$)
DECLARE FUNCTION Value# (e$, ErCode%)
  
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB DisplayMiscScrn (ScrnName$)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LptPort%, RETCODE%, EntryPoint%)
DECLARE FUNCTION Monitor ()
DECLARE SUB MakeBKGround ()
DECLARE SUB PRW2Menu ()
DECLARE SUB WaitForKey ()
DECLARE SUB WazzWind2 (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)


'$INCLUDE: 'DefCnf.bi'      'Defines a TYPE for monitor and color info..
'$INCLUDE: 'FieldInf.bi'    '        "          field information
'$INCLUDE: 'FormEdit.bi'    '        "          form editing information
'$INCLUDE: 'SetCnf.bi'
'$INCLUDE: 'PRFiles.bi'
'$INCLUDE: 'PRW2.bi'
''$INCLUDE: 'oPRtypes.bi'
'$INCLUDE: 'PREmpRec.bi'
'$INCLUDE: 'DedCodes.Bi'
'$INCLUDE: 'PRUNIT.Bi'
'$INCLUDE: 'PRTRANS.Bi'
'$INCLUDE: 'setCnf.bi'      'Defines a TYPE for monitor and color info..

CONST False = 0
CONST True = NOT False

  CrLf$ = CHR$(13) + CHR$(10)

'  PRW2Menu
'
'  IF Exist("CPMENU.EXE") THEN
''    RUN "cpmenu"
'  ELSE
'    HideCursor
'    ClearScrn
'  END IF
'  END
'
'DEFINT A-Z
LOCATE 1, 1, 0

PassCol = 44
OkFlag = -1
GOTO here

GetPW:
  DO
    MakeBKGround
    WazzWind2 7, 15, 17, 65, 3, 2, 1
    QPrintRC "Southern Software Systems, Inc.", 9, 24, 15
    QPrintRC "W-2 Processing v8.20 (C)1993-99", 11, 24, 15
    QPrintRC "Enter PASSWORD:", 14, PassCol - 16, 14
    QPrintRC "               ", 14, PassCol, 15

    IF Try > 0 THEN
      IF Try <> 3 THEN
        QPrintRC "Attempt:" + STR$(Try), 16, 35, 7
      ELSE
        QPrintRC "LAST CHANCE", 16, 34, 143
      END IF
    END IF

    COLOR 15

    PW$ = GetPassword(14, PassCol, 10, Abort)

    IF Abort THEN RUN "CPMENU"

    IF PW$ = "PRW2TIME" THEN
      OkFlag = -1
      EXIT DO
    END IF

    Try = Try + 1

    'QPrintRC "Attempt:" + STR$(Try), 12, 1, 7

    IF Try = 4 THEN
      ExitFlag = -1
      MakeBKGround
      WazzWind2 8, 15, 16, 65, 15, 4, 1
      QPrintRC "Get the password from the system operator.", 11, 20, 14
      QPrintRC "Press any key to return to the Main Menu.", 13, 20, 14
      WaitForKey
      EXIT DO
    END IF

    QPrintRC "Invalid Password. Press any key to try again.", 14, 18, 31

    IF K = 27 THEN
      GOTO GetOut
    END IF
  LOOP

here:
  IF ExitFlag THEN
    RUN "CPMenu"
  ELSEIF OkFlag THEN
    PRW2Menu
    'MakeBKGround
    'RUN "PR"
  ELSE
    GOTO GetPW
  END IF

GetOut:
CLS
END












  IF Exist("CPMENU.EXE") THEN
'    RUN "cpmenu"
  ELSE
    HideCursor
    ClearScrn
  END IF
  END

SUB ExtractW2Info (W2Type%)
  
  SHARED Choice$()
  
  REDIM sp2(1)   AS STRING * 2
  REDIM W2SetUpRec(1) AS W2SetUpType
  
  REDIM Deds(0 TO 12) AS DOUBLE
  
  W2SetUpRecLen = LEN(W2SetUpRec(1))
  
  ExitFlag = False
  OkFlag = False
  
  StartEl = 0   'Load first form at array start

  FormName$ = "W2SETUP2"
  
  NumFlds = LibNumberOfFields(W2QLib, FormName$)
  
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)       'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo               'DIM the field information array
  
  LibGetFldDef W2QLib, FormName$, StartEl, Fld(), Form$(), ErrCode
  
  InitW2Multi Fld()
  
  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1
  
  '----- Setup TYPE for setting and reading form editing information.
  
  Frm(1).FldNo = 1              'Start editing on field #1
  Frm(1).InsStat = False        'Set insert state (True = Insert on)
  Frm(1).StartEl = 0            'Set form starting element to 0 and
  
  'if not first time in then get setup file
  IF FileSize(W2SetupFile) > 0 THEN
    Form$(0, 0) = SPACE$(W2SetUpRecLen)
    FGetAH W2SetupFile, W2SetUpRec(1), W2SetUpRecLen, 1
    BCopy VARSEG(W2SetUpRec(1)), VARPTR(W2SetUpRec(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), W2SetUpRecLen, 0
    UnPackBuffer 0, 0, Form$(), Fld()
  END IF
  
  BlockClear
  
  LibFile2Scrn W2QLib, FormName$, MonoCode, 0, 0
  
  FOR Cnt = 1 TO 12
    IF LEN(QPTrim$(Choice$(Cnt, 1))) > 0 THEN
      RSET sp2(1) = QPStrI(Cnt)
      QPrintRC sp2(1), Cnt + 9, 7, -1
      QPrintRC Choice$(Cnt, 1), Cnt + 9, 10, Cnf.PulBar
    ELSE
      EXIT FOR
    END IF
  NEXT
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF Frm(1).Presses AND Frm(1).MRow = 25 THEN
      
      SELECT CASE Frm(1).MCol
      CASE 3 TO 11              'F1
        PressButton F1Key, 25, 3, 11
      CASE 35 TO 44             'EscKey
        PressButton EscKey, 25, 35, 44
      CASE 69 TO 78             'F10
        PressButton F0Key, 25, 69, 78
      END SELECT
    END IF
    
    SELECT CASE Frm(1).KeyCode
    CASE F0Key  '
      OkFlag = True
    CASE EscKey 'no
      OkFlag = True
      ExitFlag = True
    END SELECT
    
  LOOP UNTIL OkFlag             'proper key not set
  
  CursorOff
  
  IF ExitFlag THEN
    GOTO ExitW2SetUp
  END IF
  
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(W2SetUpRec(1)), VARPTR(W2SetUpRec(1)), W2SetUpRecLen, 0
  FPutAH W2SetupFile, W2SetUpRec(1), W2SetUpRecLen, 1
  
  W2Year = QPValI(Form$(1, 0))
  StrDate = Date2Num("01-01-" + Form$(1, 0))
  EndDate = Date2Num("12-31-" + Form$(1, 0))
  
  REDIM Emp2Rec(1)    AS EmpData2Type
  REDIM TranRec(1)    AS TransRecType
  REDIM UnitRec(1)    AS UnitFileRecType
  REDIM TPntr(0 TO 200)
  REDIM W2FormRec(1) AS W2FormType
  REDIM BW2FormRec(1) AS W2FormType
  
  UnitRecLen = LEN(UnitRec(1))
  W2FormRecLen = LEN(W2FormRec(1))
  Emp2RecLen = LEN(Emp2Rec(1))
  TranRecLen = LEN(TranRec(1))
  
  ENumOfRec = FileSize(EmpData2Name) \ Emp2RecLen
  TNumOfRec = FileSize(TransHistFileName) \ TranRecLen
  
  FGetAH UnitFileName, UnitRec(1), LEN(UnitRec(1)), 1
  BlockClear
  RptTitle$ = "Extracting W-2 Information"
  ShowProcessingScrn RptTitle$
  
  QPrintRC "Pass 1 of 2", 11, 35, 112
  
  'get trans action history pins
  REDIM TPins(1 TO TNumOfRec)
  FOpenS TransHistFileName, THandle             'FOR RANDOM AS #1 LEN = TranRecLen
  FOR Cnt = 1 TO TNumOfRec
    FGetRTA THandle, TranRec(1), CLNG(Cnt), TranRecLen
    TPins(Cnt) = TranRec(1).EmpPin
    ShowPctComp Cnt, TNumOfRec
  NEXT
  
  FOpenS EmpData2Name, EHandle  'FOR RANDOM AS #1 LEN = Emp2RecLen
  FOpenS W2InfoFile, WHandle    'FOR RANDOM AS #2 LEN = W2FormRecLen
  
  QPrintRC "Pass 2 of 2", 11, 35, 112
  FOR ECnt = 1 TO ENumOfRec
    FGetRTA EHandle, Emp2Rec(1), CLNG(ECnt), Emp2RecLen
    GOSUB GetEmpTranRecNums
    SELECT CASE W2Type%
    CASE 1
      IF Emp2Rec(1).EMPSOCX <> "Y" THEN
        GOSUB CheckEmp
      ELSE
        W2FormRec(1) = BW2FormRec(1)
      END IF
    CASE 2
      IF Emp2Rec(1).EMPSOCX = "Y" THEN
        GOSUB CheckEmp
      ELSE
        W2FormRec(1) = BW2FormRec(1)
      END IF
    END SELECT
    FPutRTA WHandle, W2FormRec(1), CLNG(ECnt), W2FormRecLen
    ShowPctComp ECnt, ENumOfRec
  NEXT
  
  FClose EHandle
  FClose WHandle
  FClose THandle
  
  '********ERASE arrays here
  ERASE Emp2Rec, TranRec, UnitRec
  ERASE TPntr, W2FormRec, BW2FormRec
  ERASE sp2, W2SetUpRec
  ERASE Frm, Form$, Fld
  
  BlockClear
  DisplayMiscScrn UpdatedOK
  WaitForAction
  
  GOTO ExitW2SetUp
  
GetEmpTranRecNums:
  REDIM TPntr(0 TO 1600)
  TotalTransRecs = 0
  FOR TCnt = 1 TO TNumOfRec
    IF TPins(TCnt) = Emp2Rec(1).EmpPin THEN
      TotalTransRecs = TotalTransRecs + 1
      TPntr(TotalTransRecs) = TCnt
    END IF
    TPntr(0) = TotalTransRecs
  NEXT
  RETURN
  
SumW2Info:
  W2FormRec(1) = BW2FormRec(1)
  FOR CntZZ = 0 TO 12
    Deds(CntZZ) = 0
  NEXT

  FOR Cnt = 1 TO TPntr(0)
    FGetRTA THandle, TranRec(1), CLNG(TPntr(Cnt)), TranRecLen
    IF TranRec(1).CheckDate >= StrDate AND TranRec(1).CheckDate <= EndDate THEN

      W2FormRec(1).FEDWAGE = Round(W2FormRec(1).FEDWAGE + TranRec(1).FedGrossPay)
      W2FormRec(1).FEDTAXWH = Round(W2FormRec(1).FEDTAXWH + TranRec(1).FedTaxAmt)
      W2FormRec(1).SOCWAGE = Round(W2FormRec(1).SOCWAGE + TranRec(1).SocGrossPay)
      W2FormRec(1).SOCTAXWH = Round(W2FormRec(1).SOCTAXWH + TranRec(1).SocTaxAmt)
      W2FormRec(1).MEDWAGES = Round(W2FormRec(1).MEDWAGES + TranRec(1).MedGrossPay)
      W2FormRec(1).MEDTAXWH = Round(W2FormRec(1).MEDTAXWH + TranRec(1).MedTaxAmt)
      W2FormRec(1).ADVEIC = Round(W2FormRec(1).ADVEIC + TranRec(1).EICAmt)
      W2FormRec(1).BENFBOX1 = Round(W2FormRec(1).BENFBOX1 + TranRec(1).TaxFring)
      W2FormRec(1).STATE = UnitRec(1).UFSTATE
      W2FormRec(1).STAWAGE = Round(W2FormRec(1).STAWAGE + TranRec(1).StaGrossPay)
      W2FormRec(1).STATAXWH = Round(W2FormRec(1).STATAXWH + TranRec(1).StaTaxAmt)
      
      Deds(0) = Round(Deds(0) + TranRec(1).RetireAmt)
      FOR CntZZ = 1 TO 12
        Deds(CntZZ) = Round(Deds(CntZZ) + TranRec(1).DAMT(CntZZ))
      NEXT
    END IF
  NEXT
  RETURN
  
ApplyForm:
'*******************************
  IF W2FormRec(1).BENFBOX1 > 0 THEN
    W2FormRec(1).FEDWAGE = Round(W2FormRec(1).FEDWAGE + W2FormRec(1).BENFBOX1)
    W2FormRec(1).SOCWAGE = Round(W2FormRec(1).SOCWAGE + W2FormRec(1).BENFBOX1)
    W2FormRec(1).MEDWAGES = Round(W2FormRec(1).MEDWAGES + W2FormRec(1).BENFBOX1)
'    W2FormRec(1).STAWAGE = Round(W2FormRec(1).STAWAGE + W2FormRec(1).BENFBOX1)
  END IF
'*******************************
  LSET W2FormRec(1).BOX15A = ""
  LSET W2FormRec(1).BOX15B = ""
  LSET W2FormRec(1).BOX15C = ""
  LSET W2FormRec(1).BOX15D = ""
  LSET W2FormRec(1).BOX15E = ""
  LSET W2FormRec(1).BOX15F = ""
  LSET W2FormRec(1).BOX15G = ""

  FOR CntZZ = 0 TO 12
    IF LEN(QPTrim$(W2SetUpRec(1).Deds(CntZZ).CHKDED)) AND Deds(CntZZ) > 0 THEN
      SELECT CASE LEFT$(W2SetUpRec(1).Deds(CntZZ).CHKDED, 1)
      CASE "P"
        W2FormRec(1).BOX15C = "X"
      CASE "D"
        W2FormRec(1).BOX15G = "X"
      END SELECT
    END IF
  NEXT
  
  FOR CntZZ = 0 TO 12
    SELECT CASE W2SetUpRec(1).Deds(CntZZ).AMTBOX
    CASE "13a"
      W2FormRec(1).BOX13AMT = Round#(W2FormRec(1).BOX13AMT + Deds(CntZZ))
      IF Deds(CntZZ) > 0 THEN
        W2FormRec(1).BOX13TXT = W2SetUpRec(1).Deds(CntZZ).DEDCODE
      END IF
    CASE "13b"
      W2FormRec(1).BOX13AM1 = Round#(W2FormRec(1).BOX13AM1 + Deds(CntZZ))
      IF Deds(CntZZ) > 0 THEN
        W2FormRec(1).BOX13TX1 = W2SetUpRec(1).Deds(CntZZ).DEDCODE
      END IF
    CASE "14a"
      W2FormRec(1).BOX14AMT = Round#(W2FormRec(1).BOX14AMT + Deds(CntZZ))
      IF Deds(CntZZ) > 0 THEN
        W2FormRec(1).BOX14TXT = W2SetUpRec(1).Deds(CntZZ).DEDCODE
      END IF
    CASE "14b"
      W2FormRec(1).BOX14AM1 = Round#(W2FormRec(1).BOX14AM1 + Deds(CntZZ))
      IF Deds(CntZZ) > 0 THEN
        W2FormRec(1).BOX14TX1 = W2SetUpRec(1).Deds(CntZZ).DEDCODE
      END IF
    END SELECT
  NEXT
  
  RETURN

CheckEmp:
  IF TPntr(0) THEN            'if this emp has any transactions
    GOSUB SumW2Info           'sum emp w2 info
    GOSUB ApplyForm
  ELSE
    W2FormRec(1) = BW2FormRec(1)
  END IF
RETURN

ExitW2SetUp:
  'erase arrays
END SUB

FUNCTION GetPassword$ (Row%, Col%, MaxLen%, Abort%)
  
Start:
  
  CurPos = Col
  
  DO
    
    LOCATE Row, CurPos, 1
    P$ = INPUT$(1)
    
    SELECT CASE P$
      
    CASE CHR$(8)                'Backspace
      CurPos = CurPos - 1
      
      IF CurPos <= Col THEN     'Check for backspace key at
        LOCATE Row, Col         ' start of input
        PRINT " "
        PW$ = ""
        GOTO Start
      END IF
      
      LOCATE Row, CurPos
      PRINT " "             'Clear the last *
      
    CASE CHR$(13)               'Enter
      EXIT DO
      
    CASE CHR$(27)               'Esc
      Abort = -1
      EXIT FUNCTION
      
    CASE CHR$(32) TO CHR$(126)  'ignore all other keys
      LOCATE Row, CurPos        ' except keyboard chars
      PRINT "*"
      PW$ = PW$ + P$
      CurPos = CurPos + 1
      
    END SELECT
    
  LOOP UNTIL LEN(PW$) = MaxLen
  
  GetPassword$ = UCASE$(PW$)
  LOCATE , , 0
END FUNCTION

SUB InitW2Multi (Fld() AS FieldInfo)
  SHARED Choice$()
  
  REDIM DedCodeRec(1 TO 12)  AS DedCodeRecType
  
  REDIM Choice$(14, 2)          '
  
  FOR Cnt = 0 TO 12
    FldName$ = "CHKDED" + QPTrim$(STR$(Cnt))
    Choice$(0, 0) = Choice$(0, 0) + STR$(FldNum(FldName$, Fld())) + ","
  NEXT
  Choice$(1, 0) = "Pension Plan"
  Choice$(2, 0) = "Deferred Compensation"
  Choice$(0, 2) = STR$(FldNum("AMTBOXR", Fld())) + ","
  FOR Cnt = 1 TO 12
    FldName$ = "AMTBOX" + QPTrim$(STR$(Cnt))
    Choice$(0, 2) = Choice$(0, 2) + STR$(FldNum(FldName$, Fld())) + ","
  NEXT
  Choice$(1, 2) = "13a"
  Choice$(2, 2) = "13b"
  Choice$(3, 2) = "14a"
  Choice$(4, 2) = "14b"
  
  FGetAH DedCodeFileName, DedCodeRec(1), LEN(DedCodeRec(1)), 12
  
  Choice$(0, 1) = ""
  FOR Cnt = 1 TO 12
    Choice$(Cnt, 1) = QPTrim$(DedCodeRec(Cnt).DCDESC1)
    IF Choice$(Cnt, 1) = "" THEN
      FldOffSet = FldNum("CHKDED" + QPTrim$(STR$(Cnt)), Fld())
      FOR zz = FldOffSet TO FldOffSet + 2
        Fld(zz).Protected = True
      NEXT
    END IF
  NEXT
  
  ERASE DedCodeRec
  
END SUB

SUB MakeBKGround

BK$ = STRING$(80, 177)
FOR RowCnt = 1 TO 25
  QPrintRC BK$, RowCnt, 1, 3
NEXT

END SUB

SUB PRW2Menu
  DO
    
    CursorOff
    
    Choice = 0
    
    PayMenu YTDMenu, Choice, 5
    
    SELECT CASE Choice
    CASE 1
      ExtractW2Info 1            'unrem
    CASE 2
      ExtractW2Info 2            'unrem
    CASE 3
      W2ChooseEditW2            'unrem
    CASE 4
      W2Report
    CASE 5
      W2PrintForms
      
    END SELECT
    
  LOOP UNTIL Choice = EscKey
  
END SUB

SUB W2ChooseEditW2
  
  VAction = 1
  BoxBot = 17
  BoxTop = 6
  MaxLen = 40
  LeftCol = ((80 - MaxLen) \ 2) - 1
  
  REDIM GoodRecs$(1 TO 1)
  
  W2LoadEmpList GoodRecs$()
  
  DO
    HideCursor
    LOCATE BoxTop, LeftCol
    BlockClear
    QPrintRC "  Last Name        First Name     Number    ", BoxTop - 1, LeftCol, Cnf.HiLite
    MPaintBox BoxTop, LeftCol + MaxLen + 4, BoxTop, LeftCol + MaxLen + 5, 8
    WazzWind 21, LeftCol, 23, LeftCol + MaxLen + 3, Cnf.MenBox, 2, True
    QPrintRC " Use (" + CHR$(24) + "-" + CHR$(25) + ") to select", 22, LeftCol + (MaxLen \ 2) - 8, Cnf.ActivCh
    ShowCursor
    
    DO
      VertMenu GoodRecs$(), Choice, MaxLen, BoxBot, Ky$, VAction, Cnf
      SELECT CASE ASCII(Ky$)
      CASE EscKey
        Choice = EscKey
        ExitFlag = True
        EXIT DO
      CASE 13
        EdRecNo = CVI(RIGHT$(GoodRecs$(Choice), 2))
        W2EditRec EdRecNo
        EXIT DO
      CASE ELSE
      END SELECT
    LOOP
  LOOP UNTIL ExitFlag
  
  ERASE GoodRecs$
  
  CursorOff
  
END SUB

SUB W2EditRec (EdRecNo)
  
  REDIM W2FormRec(1) AS W2FormType
  
  W2FormRecLen = LEN(W2FormRec(1))
  
  ExitFlag = False
  OkFlag = False
  
  StartEl = 0   'Load first form at array start
  FormName$ = W2Info
  NumFlds = LibNumberOfFields(W2QLib, FormName$)
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)       'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo               'DIM the field information array
  LibGetFldDef W2QLib, FormName$, StartEl, Fld(), Form$(), ErrCode
  
  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1
  '----- Setup TYPE for setting and reading form editing information.
  Frm(1).FldNo = 1              'Start editing on field #1
  Frm(1).InsStat = False        'Set insert state (True = Insert on)
  Frm(1).StartEl = 0            'Set form starting element to 0 and
  
  Form$(0, 0) = SPACE$(W2FormRecLen)
  FOpenS W2InfoFile, WHandle
  FGetRTA WHandle, W2FormRec(1), CLNG(EdRecNo), W2FormRecLen
  FClose WHandle
  
  BCopy VARSEG(W2FormRec(1)), VARPTR(W2FormRec(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), W2FormRecLen, 0
  UnPackBuffer 0, 0, Form$(), Fld()
  
  BlockClear
  
  LibFile2Scrn W2QLib, FormName$, MonoCode, 0, 0
  QPrintRC W2GetEmpName$(EdRecNo), 1, 15, 112
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF Frm(1).Presses AND Frm(1).MRow = 25 THEN
      
      SELECT CASE Frm(1).MCol
      CASE 3 TO 11              'F1
        PressButton F1Key, 25, 3, 11
      CASE 35 TO 44             'EscKey
        PressButton EscKey, 25, 35, 44
      CASE 69 TO 78             'F10
        PressButton F0Key, 25, 69, 78
      END SELECT
    END IF
    
    SELECT CASE Frm(1).KeyCode
    CASE F0Key  '
      OkFlag = True
    CASE EscKey 'no
      IF Frm(1).Edited THEN
        OkFlag = PromptSaveData
      END IF
      
      IF OkFlag = 1 THEN
        OkFlag = 0
      ELSE
        ExitFlag = True
      END IF
      
    END SELECT
    
  LOOP UNTIL OkFlag OR ExitFlag 'proper key not set
  
  CursorOff
  
  IF OkFlag THEN
    BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(W2FormRec(1)), VARPTR(W2FormRec(1)), W2FormRecLen, 0
    FOpenS W2InfoFile, WHandle
    FPutRTA WHandle, W2FormRec(1), CLNG(EdRecNo), W2FormRecLen
    FClose WHandle
    
    DisplayMiscScrn UpdatedOK
    WaitForAction
  END IF
  
  ERASE W2FormRec, Frm, Form$, Fld
  
END SUB

FUNCTION W2GetEmpName$ (EdRecNo%)
  REDIM Emp1Rec(1)    AS EmpData1Type
  FOpenS EmpData1Name, Handle
  FGetRTA Handle, Emp1Rec(1), CLNG(EdRecNo), LEN(Emp1Rec(1))
  FClose Handle
  W2GetEmpName$ = QPTrim$(Emp1Rec(1).EMPFNAME) + " " + QPTrim$(Emp1Rec(1).EMPLNAME)
  ERASE Emp1Rec
  
END FUNCTION

SUB W2LoadEmpList (GoodRecs$())
  
  REDIM E1Rec(1) AS EmpData1Type
  REDIM TempData(1) AS STRING * 43
  
  E1RecLen = LEN(E1Rec(1))
  
  EmpNumOfRecs = FileSize(EmpData1Name) \ E1RecLen
  
  REDIM GoodRecs$(1 TO EmpNumOfRecs)
  
  FOpenS EmpData1Name, E1Handle
  FOR Cnt = 1 TO EmpNumOfRecs
    FGetRTA E1Handle, E1Rec(1), CLNG(Cnt), E1RecLen
    IF NOT E1Rec(1).Deleted THEN
      ECnt = ECnt + 1
      GoodRecs$(ECnt) = TempData(1)
      LSET GoodRecs$(ECnt) = E1Rec(1).EMPLNAME
      MID$(GoodRecs$(ECnt), 18) = QPTrim$(E1Rec(1).EMPFNAME)
      MID$(GoodRecs$(ECnt), 30) = RIGHT$(E1Rec(1).EmpNo, 8)
      MID$(GoodRecs$(ECnt), 42) = MKI$(E1Rec(1).Data1RecNum)
    END IF
  NEXT
  FClose E1Handle
  
  REDIM PRESERVE GoodRecs$(1 TO ECnt)
  Sort GoodRecs$(), 0
  
  ERASE E1Rec, TempData
  
END SUB

SUB W2PrintForms
  
  FirstTime = True

  REDIM PEMPCity(1)  AS STRING * 20
  REDIM PEmpSSN(1)  AS STRING * 15
  REDIM BTxt14(1)   AS STRING * 5
  
  REDIM Unit(1)        AS UnitFileRecType
  REDIM W2InfoRec(1)   AS W2FormType
  REDIM Emp2Rec(1)     AS EmpData2Type
  'W2InfoRec(1)
  'Dash(1) = STRING$(78, "-") + CrLf$
  
  Image1$ = "######.##"
  Image2$ = "#####.##"
  Image$ = "######"
  
  UnitRecLen = LEN(Unit(1))
  W2InfoRecLen = LEN(W2InfoRec(1))
  Emp2RecLen = LEN(Emp2Rec(1))
  
  FGetAH UnitFileName, Unit(1), UnitRecLen, 1
  
  ExitFlag = False
  OkFlag = False
  
  StartEl = 0   'Load first form at array start
  FormName$ = "W2FRMPRN"
  
  NumFlds = LibNumberOfFields(W2QLib, FormName$)
  
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)       'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo               'DIM the field information array
  
  LibGetFldDef W2QLib, FormName$, StartEl, Fld(), Form$(), ErrCode
  
  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1
  
  '----- Setup TYPE for setting and reading form editing information.
  
  Frm(1).FldNo = 1              'Start editing on field #1
  Frm(1).InsStat = False        'Set insert state (True = Insert on)
  Frm(1).StartEl = 0            'Set form starting element to 0 and
  
  GOSUB W2PrintFormScrn
  

  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF Frm(1).Presses AND Frm(1).MRow = 25 THEN
      SELECT CASE Frm(1).MCol
      CASE 2 TO 12              'F1
        PressButton F1Key, 25, 2, 12
      CASE 28 TO 47             'F5
        PressButton F5KEY, 25, 28, 47
      CASE 48 TO 67             'F10
        PressButton F0Key, 25, 48, 67
      CASE 68 TO 79             'Escape
        PressButton EscKey, 25, 68, 79
      END SELECT
    END IF
    
    SELECT CASE Frm(1).KeyCode
    CASE F5KEY
      CursorOff
      BlockClear
      LptPort = QPValI(Form$(3, 0))
      IF LptPort < 1 OR LptPort > 2 THEN LptPort = 1
      SELECT CASE QPValI(Form$(1, 0))
      CASE 0, 1
         PrintRptFile "", W2PrintMask1, LptPort, RETCODE, 5
      CASE 2
         PrintRptFile "", W2PrintMask2, LptPort, RETCODE, 5
      CASE ELSE
      END SELECT
      GOSUB W2PrintFormScrn
    CASE F0Key  '
      OkFlag = True
    CASE EscKey 'no
      OkFlag = True
      ExitFlag = True
    END SELECT
  LOOP UNTIL ExitFlag OR OkFlag
  
  CursorOff
  
  IF ExitFlag THEN
    GOTO ExitW2PrintForm
  END IF
  
  '*** Start of W2 Printing
  
  TwoWide = QPValI(Form$(1, 0))
  LptPort = QPValI(Form$(3, 0))
  IF LptPort < 1 OR LptPort > 2 THEN LptPort = 1

  SELECT CASE TwoWide
    CASE 0, 1
      TwoWide = False
    CASE 2
      TwoWide = True
  END SELECT

  CtrlNumb = QPValI(Form$(2, 0))
  
  RptName$ = W2PrintFile
  RptTitle$ = "W-2 Forms Printing"
  
  PrnCnt = 0
  MaxPrn = 41
  
  'Page = 0
  IdxRecLen = 2
  IdxFileSize& = FileSize(EmpIdxLName)
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  REDIM IdxBuff(1 TO IdxNumOfRecs) AS EmployeeIndexType         'load index file
  FGetAH EmpIdxLName, IdxBuff(1), IdxRecLen, IdxNumOfRecs
  
  BlockClear
  
  ShowProcessingScrn RptTitle$
  
  FCreate RptName$
  
  '  FOpenS RptName$, RHandle            'open
  
  OPEN RptName$ FOR OUTPUT AS #1
  WIDTH #1, 255

  FOpenS W2InfoFile, WHandle    'those
  FOpenS EmpData2Name, EHandle  'files.
  
  FOR Cnt = 1 TO IdxNumOfRecs
    FGetRTA WHandle, W2InfoRec(1), CLNG(IdxBuff(Cnt).DataRecNum), W2InfoRecLen
    FGetRTA EHandle, Emp2Rec(1), CLNG(IdxBuff(Cnt).DataRecNum), Emp2RecLen
    GOSUB PrintW2Form
    'IF PrnCnt = MaxPrn THEN
    '  GOSUB PrintW2SubTotal
    'END IF
'    IF DidOne > 0 THEN EXIT FOR
    ShowPctComp Cnt, IdxNumOfRecs
  NEXT
  
  CLOSE
  FClose WHandle
  FClose EHandle
  
  PrintRptFile RptTitle$, RptName$, LptPort, RETCODE, 1

ExitW2PrintForm:
  ERASE Frm, Form$, Fld, IdxBuff
  ERASE PEMPCity, PEmpSSN, BTxt14, Unit, W2InfoRec, Emp2Rec
  
  EXIT SUB
  
PrintW2Form:
  
  IF W2InfoRec(1).FEDWAGE = 0 AND W2InfoRec(1).FEDTAXWH = 0 AND W2InfoRec(1).SOCWAGE = 0 THEN
    IF W2InfoRec(1).SOCTAXWH = 0 AND W2InfoRec(1).MEDWAGES = 0 AND W2InfoRec(1).MEDTAXWH = 0 THEN
      IF W2InfoRec(1).SOCTIPS = 0 AND W2InfoRec(1).ALLOCTIP = 0 AND W2InfoRec(1).ADVEIC = 0 THEN
        IF W2InfoRec(1).DEPNDCAR = 0 AND W2InfoRec(1).NQPLAN = 0 AND W2InfoRec(1).BOX13AMT = 0 THEN
          IF W2InfoRec(1).BOX13AM1 = 0 THEN
            GOTO DontPrintEm
          END IF
        END IF
      END IF
    END IF
  END IF

  DidOne = DidOne + 1

  'GOSUB SumW2SubTotal
  
  LSET PEMPCity(1) = Emp2Rec(1).EMPCITY
  LSET PEmpSSN(1) = LEFT$(Emp2Rec(1).EMPSSN, 3) + "-" + MID$(Emp2Rec(1).EMPSSN, 4, 2) + "-" + RIGHT$(QPTrim$(Emp2Rec(1).EMPSSN), 4)

'start of w2 forms printing
  PRINT #1,
  PRINT #1,
  PRINT #1,

  PRINT #1, TAB(6); FUsing(STR$(CtrlNumb), Image$);
  IF TwoWide THEN
    PRINT #1, TAB(75); FUsing(STR$(CtrlNumb), Image$)
  ELSE
    PRINT #1,
  END IF

  PRINT #1,

  PRINT #1, TAB(5); Unit(1).UFFEDID; TAB(42); FUsing(STR$(W2InfoRec(1).FEDWAGE), Image1$); TAB(59); FUsing(STR$(W2InfoRec(1).FEDTAXWH), Image1$);
  IF TwoWide THEN
    PRINT #1, TAB(74); Unit(1).UFFEDID; TAB(111); FUsing(STR$(W2InfoRec(1).FEDWAGE), Image1$); TAB(128); FUsing(STR$(W2InfoRec(1).FEDTAXWH), Image1$)
  ELSE
    PRINT #1,
  END IF

  PRINT #1,

  PRINT #1, TAB(42); FUsing(STR$(W2InfoRec(1).SOCWAGE), Image1$); TAB(59); FUsing(STR$(W2InfoRec(1).SOCTAXWH), Image1$);
  IF TwoWide THEN
    PRINT #1, TAB(111); FUsing(STR$(W2InfoRec(1).SOCWAGE), Image1$); TAB(128); FUsing(STR$(W2InfoRec(1).SOCTAXWH), Image1$)
  ELSE
    PRINT #1,
  END IF

  PRINT #1, TAB(5); Unit(1).UFEMPR;
  IF TwoWide THEN
    PRINT #1, TAB(74); Unit(1).UFEMPR
  ELSE
    PRINT #1,
  END IF

  PRINT #1, TAB(5); Unit(1).UFADDR1; TAB(42); FUsing(STR$(W2InfoRec(1).MEDWAGES), Image1$); TAB(59); FUsing(STR$(W2InfoRec(1).MEDTAXWH), Image1$);
  IF TwoWide THEN
    PRINT #1, TAB(74); Unit(1).UFADDR1; TAB(111); FUsing(STR$(W2InfoRec(1).MEDWAGES), Image1$); TAB(128); FUsing(STR$(W2InfoRec(1).MEDTAXWH), Image1$)
  ELSE
    PRINT #1,
  END IF

  PRINT #1, TAB(5); Unit(1).UFADDR2;
  IF TwoWide THEN
    PRINT #1, TAB(74); Unit(1).UFADDR2
  ELSE
    PRINT #1,
  END IF

  IF W2InfoRec(1).SOCTIPS > 0 THEN
    SocTip$ = FUsing(STR$(W2InfoRec(1).SOCTIPS), Image1$)
  ELSE
    SocTip$ = SPACE$(LEN(Image1$))
  END IF

  IF W2InfoRec(1).ALLOCTIP > 0 THEN
    AlocTip$ = FUsing(STR$(W2InfoRec(1).ALLOCTIP), Image1$)
  ELSE
    AlocTip$ = SPACE$(LEN(Image1$))
  END IF

  PRINT #1, TAB(5); LEFT$(QPTrim$(Unit(1).UFCITY), 15) + " " + Unit(1).UFSTATE + " " + Unit(1).UFZIP; TAB(42); SocTip$; TAB(59); AlocTip$;
  IF TwoWide THEN
    PRINT #1, TAB(74); LEFT$(QPTrim$(Unit(1).UFCITY), 15) + " " + Unit(1).UFSTATE + " " + Unit(1).UFZIP; TAB(111); SocTip$; TAB(128); AlocTip$
  ELSE
    PRINT #1,
  END IF
  
  PRINT #1,     ' line 13

  IF W2InfoRec(1).ADVEIC > 0 THEN
    AdvEicP$ = FUsing(STR$(W2InfoRec(1).ADVEIC), Image1$)
  ELSE
    AdvEicP$ = SPACE$(LEN(Image1$))
  END IF

  IF W2InfoRec(1).DEPNDCAR > 0 THEN
    DepCare$ = FUsing(STR$(W2InfoRec(1).DEPNDCAR), Image1$)
  ELSE
    DepCare$ = SPACE$(LEN(Image1$))
  END IF

  PRINT #1, TAB(5); PEmpSSN(1); TAB(42); AdvEicP$; TAB(59); DepCare$;
  IF TwoWide THEN
    PRINT #1, TAB(74); PEmpSSN(1); TAB(111); AdvEicP$; TAB(128); DepCare$
  ELSE
    PRINT #1,
  END IF

  PRINT #1,

  IF W2InfoRec(1).NQPLAN > 0 THEN
    NQP$ = FUsing(STR$(W2InfoRec(1).NQPLAN), Image1$)
  ELSE
    NQP$ = SPACE$(LEN(Image1$))
  END IF

  IF W2InfoRec(1).BENFBOX1 > 0 THEN
    BenBox$ = FUsing(STR$(W2InfoRec(1).BENFBOX1), Image1$)
  ELSE
    BenBox$ = SPACE$(LEN(Image1$))
  END IF

  PRINT #1, TAB(5); QPTrim$(Emp2Rec(1).EMPFNAME); " "; QPTrim$(Emp2Rec(1).EMPLNAME); TAB(42); NQP$; TAB(59); BenBox$;
  IF TwoWide THEN
    PRINT #1, TAB(74); QPTrim$(Emp2Rec(1).EMPFNAME); " "; QPTrim$(Emp2Rec(1).EMPLNAME); TAB(111); NQP$; TAB(128); BenBox$
  ELSE
    PRINT #1,
  END IF

  PRINT #1,

  BTxt14(1) = QPTrim$(W2InfoRec(1).BOX14TXT) + SPACE$(5)        'line 14 a
  IF W2InfoRec(1).BOX13AMT > 0 THEN
    Box13Amt1$ = FUsing(STR$(W2InfoRec(1).BOX13AMT), Image1$)
  ELSE
    Box13Amt1$ = SPACE$(LEN(Image1$))
  END IF
  IF W2InfoRec(1).BOX14AMT > 0 THEN
    Box14Amt1$ = FUsing(STR$(W2InfoRec(1).BOX14AMT), Image1$)
  ELSE
    Box14Amt1$ = SPACE$(LEN(Image1$))
  END IF

  PRINT #1, TAB(5); QPTrim$(Emp2Rec(1).EMPADDR1); TAB(40); LEFT$(QPTrim$(W2InfoRec(1).BOX13TXT) + SPACE$(1), 1); " "; Box13Amt1$; TAB(54); BTxt14(1); Box14Amt1$;
  IF TwoWide THEN
    PRINT #1, TAB(74); QPTrim$(Emp2Rec(1).EMPADDR1); TAB(109); LEFT$(QPTrim$(W2InfoRec(1).BOX13TXT) + SPACE$(1), 1); " "; Box13Amt1$; TAB(123); BTxt14(1); Box14Amt1$
  ELSE
    PRINT #1,
  END IF

  BTxt14(1) = QPTrim$(W2InfoRec(1).BOX14TX1) + SPACE$(5)        'line 14 b
  IF W2InfoRec(1).BOX13AM1 > 0 THEN
    Box13Amt1$ = FUsing(STR$(W2InfoRec(1).BOX13AM1), Image1$)
  ELSE
    Box13Amt1$ = SPACE$(LEN(Image1$))
  END IF
  IF W2InfoRec(1).BOX14AM1 > 0 THEN
    Box14Amt1$ = FUsing(STR$(W2InfoRec(1).BOX14AM1), Image1$)
  ELSE
    Box14Amt1$ = SPACE$(LEN(Image1$))
  END IF

  PRINT #1, TAB(5); QPTrim$(Emp2Rec(1).EMPADDR2); TAB(40); LEFT$(QPTrim$(W2InfoRec(1).BOX13TX1) + SPACE$(1), 1); " "; Box13Amt1$; TAB(54); BTxt14(1); Box14Amt1$;
  IF TwoWide THEN
    PRINT #1, TAB(74); QPTrim$(Emp2Rec(1).EMPADDR2); TAB(109); LEFT$(QPTrim$(W2InfoRec(1).BOX13TX1) + SPACE$(1), 1); " "; Box13Amt1$; TAB(123); BTxt14(1); Box14Amt1$
  ELSE
    PRINT #1,
  END IF

  PRINT #1, TAB(5); QPTrim$(PEMPCity(1)) + " "; Emp2Rec(1).EMPSTATE; " "; QPTrim$(Emp2Rec(1).EMPZIP);
  IF TwoWide THEN
    PRINT #1, TAB(74); QPTrim$(PEMPCity(1)) + " "; Emp2Rec(1).EMPSTATE; " "; QPTrim$(Emp2Rec(1).EMPZIP)
  ELSE
    PRINT #1,
  END IF

  PRINT #1,
  PRINT #1,

  'PRINT #1, TAB(39); W2InfoRec(1).BOX15A; TAB(45); W2InfoRec(1).BOX15B; TAB(51); W2InfoRec(1).BOX15C; TAB(56); W2InfoRec(1).BOX15D; TAB(55); W2InfoRec(1).BOX15E; TAB(59); W2InfoRec(1).BOX15F; TAB(63); W2InfoRec(1).BOX15G;

  'LSET W2InfoRec(1).BOX15A = W2InfoRec(1).BOX15A
  'LSET W2InfoRec(1).BOX15B = W2InfoRec(1).BOX15B
  'LSET W2InfoRec(1).BOX15C = W2InfoRec(1).BOX15C
  'LSET W2InfoRec(1).BOX15D = W2InfoRec(1).BOX15D
  'LSET W2InfoRec(1).BOX15G = W2InfoRec(1).BOX15G

  PRINT #1, TAB(39); W2InfoRec(1).BOX15A; TAB(45); W2InfoRec(1).BOX15B; TAB(51); W2InfoRec(1).BOX15C; TAB(57); W2InfoRec(1).BOX15D; TAB(63); W2InfoRec(1).BOX15G;
  IF TwoWide THEN
    'PRINT #1, TAB(107); W2InfoRec(1).BOX15A; TAB(111); W2InfoRec(1).BOX15B; TAB(115); W2InfoRec(1).BOX15C; TAB(119); W2InfoRec(1).BOX15D; TAB(123); W2InfoRec(1).BOX15E; TAB(127); W2InfoRec(1).BOX15F; TAB(131); W2InfoRec(1).BOX15G
    PRINT #1, TAB(109); W2InfoRec(1).BOX15A; TAB(114); W2InfoRec(1).BOX15B; TAB(120); W2InfoRec(1).BOX15C; TAB(126); W2InfoRec(1).BOX15D; TAB(132); W2InfoRec(1).BOX15G
  ELSE
    PRINT #1,
  END IF

  PRINT #1,

  IF W2InfoRec(1).STAWAGE > 0 THEN
    StateWage$ = FUsing(STR$(W2InfoRec(1).STAWAGE), Image1$)
  ELSE
    StateWage$ = SPACE$(LEN(Image1$))
  END IF

  IF W2InfoRec(1).STATAXWH > 0 THEN
    StateTax$ = FUsing(STR$(W2InfoRec(1).STATAXWH), Image1$)
  ELSE
    StateTax$ = SPACE$(LEN(Image1$))
  END IF

  PRINT #1, TAB(5); W2InfoRec(1).STATE; TAB(9); QPTrim$(Unit(1).UFSTAID); TAB(23); StateWage$; TAB(33); StateTax$;
  IF TwoWide THEN
    PRINT #1, TAB(74); W2InfoRec(1).STATE; TAB(78); QPTrim$(Unit(1).UFSTAID); TAB(91); StateWage$; TAB(101); StateTax$
  ELSE
    PRINT #1,
  END IF

  PRINT #1,
  PRINT #1,
  PRINT #1,
  PRINT #1,
  PRINT #1,
  PRINT #1,
  PRINT #1,
  PRINT #1,

  PrnCnt = PrnCnt + 1
  CtrlNumb = CtrlNumb + 1
  
DontPrintEm:
  RETURN
  
'PrintW2SubTotal:
'  PRINT #1,
'  PRINT #1,
'  PRINT #1,
'
'  PRINT #1, TAB(6); FUsing(STR$(CtrlNumb), Image$);
'  IF TwoWide THEN
'    PRINT #1, TAB(75); FUsing(STR$(CtrlNumb), Image$)
'  ELSE
'    PRINT #1,
'  END IF
'
'  PRINT #1,
'
'  PRINT #1, TAB(5); Unit(1).UFFEDID; TAB(42); FUsing(STR$(FEDWAGE#), Image1$); TAB(59); FUsing(STR$(FEDTAXWH#), Image1$);
'  IF TwoWide THEN
'    PRINT #1, TAB(74); Unit(1).UFFEDID; TAB(111); FUsing(STR$(FEDWAGE#), Image1$); TAB(128); FUsing(STR$(FEDTAXWH#), Image1$)
'  ELSE
'    PRINT #1,
'  END IF
'
'  PRINT #1,
'
'  PRINT #1, TAB(42); FUsing(STR$(SOCWAGE#), Image1$); TAB(59); FUsing(STR$(SOCTAXWH#), Image1$);
'  IF TwoWide THEN
'    PRINT #1, TAB(111); FUsing(STR$(SOCWAGE#), Image1$); TAB(128); FUsing(STR$(SOCTAXWH#), Image1$)
'  ELSE
'    PRINT #1,
'  END IF
'
'  PRINT #1, TAB(5); Unit(1).UFEMPR;
'  IF TwoWide THEN
'    PRINT #1, TAB(74); Unit(1).UFEMPR
'  ELSE
'    PRINT #1,
'  END IF
'
'  PRINT #1, TAB(5); Unit(1).UFADDR1; TAB(42); FUsing(STR$(MEDWAGES#), Image1$); TAB(59); FUsing(STR$(MEDTAXWH#), Image1$);
'  IF TwoWide THEN
'    PRINT #1, TAB(74); Unit(1).UFADDR1; TAB(111); FUsing(STR$(MEDWAGES#), Image1$); TAB(128); FUsing(STR$(MEDTAXWH#), Image1$)
'  ELSE
'    PRINT #1,
'  END IF
'
'  PRINT #1, TAB(5); Unit(1).UFADDR2;
'  IF TwoWide THEN
'    PRINT #1, TAB(74); Unit(1).UFADDR2
'  ELSE
'    PRINT #1,
'  END IF
'
'  'PRINT #1, TAB(5); Unit(1).UFCITY; "  "; Unit(1).UFSTATE; "  "; Unit(1).UFZIP; TAB(42); FUsing(STR$(SOCTIPS#), Image1$); TAB(59); FUsing(STR$(ALLOCTIP#), Image1$)
'
'  PRINT #1, TAB(5); LEFT$(QPTrim$(Unit(1).UFCITY), 15) + " " + Unit(1).UFSTATE + " " + Unit(1).UFZIP; TAB(42); FUsing(STR$(SOCTIPS#), Image1$); TAB(59); FUsing(STR$(ALLOCTIP#), Image1$);
'  IF TwoWide THEN
'    PRINT #1, TAB(74); LEFT$(QPTrim$(Unit(1).UFCITY), 15) + " " + Unit(1).UFSTATE + " " + Unit(1).UFZIP; TAB(111); FUsing(STR$(SOCTIPS#), Image1$); TAB(128); FUsing(STR$(ALLOCTIP#), Image1$)
'  ELSE
'    PRINT #1,
'  END IF
'
'  PRINT #1,     ' line 13
'
'  PRINT #1, TAB(42); FUsing(STR$(ADVEIC#), Image1$); TAB(59); FUsing(STR$(DEPNDCAR#), Image1$);
'  IF TwoWide THEN
'    PRINT #1, TAB(111); FUsing(STR$(ADVEIC#), Image1$); TAB(128); FUsing(STR$(DEPNDCAR#), Image1$)
'  ELSE
'    PRINT #1,
'  END IF
'
'  PRINT #1,
'
'  PRINT #1, TAB(42); FUsing(STR$(NQPLAN#), Image1$); TAB(59); FUsing(STR$(BENFBOX1#), Image1$);
'  IF TwoWide THEN
'    PRINT #1, TAB(111); FUsing(STR$(NQPLAN#), Image1$); TAB(128); FUsing(STR$(BENFBOX1#), Image1$)
'  ELSE
'    PRINT #1,
'  END IF
'
'  PRINT #1,
'
'  PRINT #1, TAB(40); "  "; FUsing(STR$(BOX13AMT#), Image1$);
'  IF TwoWide THEN
'    PRINT #1, TAB(109); "  "; FUsing(STR$(BOX13AMT#), Image1$)
'  ELSE
'    PRINT #1,
'  END IF
'
'
'  PRINT #1, TAB(40); "  "; FUsing(STR$(BOX13AM1#), Image1$);
'  IF TwoWide THEN
'    PRINT #1, TAB(109); "  "; FUsing(STR$(BOX13AM1#), Image1$)
'  ELSE
'    PRINT #1,
'  END IF
'
'  PRINT #1,
'  PRINT #1,
'  PRINT #1,
'
'  PRINT #1, TAB(59); "X";
'  IF TwoWide THEN
'    PRINT #1, TAB(128); "X"
'  ELSE
'    PRINT #1,
'  END IF
'
'  PRINT #1,
'
'  PRINT #1, TAB(23); FUsing(STR$(STAWAGE#), Image1$); TAB(33); FUsing(STR$(STATAXWH#), Image1$);
'  IF TwoWide THEN
'    PRINT #1, TAB(91); FUsing(STR$(STAWAGE#), Image1$); TAB(101); FUsing(STR$(STATAXWH#), Image1$)
'  ELSE
'    PRINT #1,
'  END IF
'
'  PRINT #1,
'  PRINT #1,
'  PRINT #1,
'  PRINT #1,
'  PRINT #1,
'  PRINT #1,
'  PRINT #1,
'  PRINT #1,
'
'  BeenInSubTotal = True
  
'  FEDWAGE# = 0
'  FEDTAXWH# = 0
'  SOCWAGE# = 0
'  SOCTAXWH# = 0
'  MEDWAGES# = 0
'  MEDTAXWH# = 0
'  SOCTIPS# = 0
'  ALLOCTIP# = 0
'  ADVEIC# = 0
'  DEPNDCAR# = 0
'  NQPLAN# = 0
'  BENFBOX1# = 0
'  BOX13AMT# = 0
'  BOX13AM1# = 0
'
'  STATAXWH# = 0
'  STAWAGE# = 0
'
'  PrnCnt = 0
'  CtrlNumb = CtrlNumb + 1
'  RETURN

W2PrintFormScrn:
  BlockClear
  LibFile2Scrn W2QLib, "W2BUTTON", MonoCode, 0, 0
  LibFile2Scrn W2QLib, FormName$, MonoCode, 0, 0
  IF FirstTime THEN
    FirstTime = NOT FirstTime
    Form$(1, 0) = "1"
    'Form$(2, 0) = "1"
    Form$(3, 0) = "1"
    Fld(1).Protected = False
    Action = 1
  ELSE
    PrintArray 0, 0, Form$(), Fld()
  END IF
RETURN
  
SumW2SubTotal:
  FEDWAGE# = Round(FEDWAGE# + W2InfoRec(1).FEDWAGE)
  FEDTAXWH# = Round(FEDTAXWH# + W2InfoRec(1).FEDTAXWH)
  SOCWAGE# = Round(SOCWAGE# + W2InfoRec(1).SOCWAGE)
  SOCTAXWH# = Round(SOCTAXWH# + W2InfoRec(1).SOCTAXWH)
  MEDWAGES# = Round(MEDWAGES# + W2InfoRec(1).MEDWAGES)
  MEDTAXWH# = Round(MEDTAXWH# + W2InfoRec(1).MEDTAXWH)
  SOCTIPS# = Round(SOCTIPS# + W2InfoRec(1).SOCTIPS)
  ALLOCTIP# = Round(ALLOCTIP# + W2InfoRec(1).ALLOCTIP)
  ADVEIC# = Round(ADVEIC# + W2InfoRec(1).ADVEIC)
  DEPNDCAR# = Round(DEPNDCAR# + W2InfoRec(1).DEPNDCAR)
  NQPLAN# = Round(NQPLAN# + W2InfoRec(1).NQPLAN)
  '11
  BENFBOX1# = Round(BENFBOX1# + W2InfoRec(1).BENFBOX1)
  BOX13AMT# = Round(BOX13AMT# + W2InfoRec(1).BOX13AMT)
  BOX13AM1# = Round(BOX13AM1# + W2InfoRec(1).BOX13AM1)
  STATAXWH# = Round(STATAXWH# + W2InfoRec(1).STATAXWH)
  STAWAGE# = Round(STAWAGE# + W2InfoRec(1).STAWAGE)
RETURN
  
  
END SUB

SUB W2Report
  
  CursorOff

  REDIM EmpCnt(1) AS STRING * 26
  REDIM FEDGROSS(1)   AS STRING * 13
  REDIM STAGROSS(1)   AS STRING * 13
  REDIM SOCGROSS(1)   AS STRING * 13
  REDIM MEDGROSS(1)   AS STRING * 13
  REDIM Box13a(1)     AS STRING * 13
  REDIM Box13b(1)     AS STRING * 13
  REDIM FEDTAX(1)     AS STRING * 13
  REDIM STATAX(1)     AS STRING * 13
  REDIM SOCTAX(1)     AS STRING * 13
  REDIM MEDTAX(1)     AS STRING * 13
  REDIM Box14a(1)     AS STRING * 13
  REDIM Box14b(1)     AS STRING * 13
  REDIM Box12(1)      AS STRING * 13
  REDIM ADVEIC(1)     AS STRING * 13
  REDIM Box10(1)      AS STRING * 13
  REDIM Box11(1)      AS STRING * 13
  
  REDIM Pg(1)          AS STRING * 3
  REDIM Dash(1)        AS STRING * 80
  REDIM Unit(1)        AS UnitFileRecType
  REDIM W2InfoRec(1)   AS W2FormType
  REDIM Emp2Rec(1)     AS EmpData2Type
  REDIM Line1(1)       AS STRING * 82
  
  Dash(1) = STRING$(78, "-") + CrLf$
  
  Image1$ = ",########.##"
  'Image2$ = ",######.##"
  
  UnitRecLen = LEN(Unit(1))
  W2InfoRecLen = LEN(W2InfoRec(1))
  Emp2RecLen = LEN(Emp2Rec(1))
  
  FGetAH UnitFileName, Unit(1), UnitRecLen, 1
  
  RptName$ = "PRRPTS\W2REPORT.RPT"
  
  LineCnt = 0
  MaxLines = 50
  Page = 0
  IdxRecLen = 2
  IdxFileSize& = FileSize(EmpIdxLName)
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  REDIM IdxBuff(1 TO IdxNumOfRecs) AS EmployeeIndexType         'load index file
  FGetAH EmpIdxLName, IdxBuff(1), IdxRecLen, IdxNumOfRecs
  
  RptTitle$ = "W-2 Information Report"
  
  BlockClear
  
  ShowProcessingScrn RptTitle$
  
  FCreate RptName$
  
  FOpenS RptName$, RHandle
  FOpenS W2InfoFile, WHandle
  FOpenS EmpData2Name, EHandle
  
  GOSUB PrintW2Header
  
  FOR Cnt = 1 TO IdxNumOfRecs
    FGetRTA WHandle, W2InfoRec(1), CLNG(IdxBuff(Cnt).DataRecNum), W2InfoRecLen
    FGetRTA EHandle, Emp2Rec(1), CLNG(IdxBuff(Cnt).DataRecNum), Emp2RecLen
    IF NOT Emp2Rec(1).Deleted THEN
      GOSUB PrintW2Data
      IF LineCnt >= MaxLines THEN
        FPut RHandle, FF$
        GOSUB PrintW2Header
      END IF
    END IF
    ShowPctComp Cnt, IdxNumOfRecs
  NEXT
  
  GOSUB PrintW2Summary
  
  FClose WHandle
  FClose RHandle
  FClose EHandle
  
  ERASE EmpCnt, FEDGROSS, STAGROSS, SOCGROSS, MEDGROSS, Box13a
  ERASE Box13b, FEDTAX, STATAX, SOCTAX, MEDTAX, Box14a
  ERASE Box14b, Box12, ADVEIC, Box10, Box11, Pg, Dash
  ERASE Unit, W2InfoRec, Emp2Rec, Line1, IdxBuff
  
  PrintRptFile RptTitle$, RptName$, 1, RETCODE, 0
  
  EXIT SUB
  
PrintW2Header:
  Page = Page + 1
  RSET Pg(1) = STR$(Page)
  UTemp$ = SPACE$(80)
  LSET UTemp$ = Unit(1).UFEMPR
  MID$(UTemp$, 71) = "Page:" + Pg(1) + CrLf$
  
  FPut RHandle, UTemp$
  FPut RHandle, "W-2 Information Report" + CrLf$
  FPut RHandle, "Report Date: " + DATE$ + CrLf$
  FPut RHandle, CrLf$
  FPut RHandle, "Employee Name                   Adv EIC       Box 10       Box 11       Box 12" + CrLf$
  FPut RHandle, "    Fed Gross  State Gross    Soc Gross    Med Gross       Box 13       Box 13" + CrLf$
  FPut RHandle, "      Fed Tax    State Tax      Soc Tax      Med Tax       Box 14       Box 14" + CrLf$
  FPut RHandle, Dash(1)
  LineCnt = 8
  UTemp$ = ""
  RETURN
  
  
PrintW2Data:

  IF W2InfoRec(1).FEDWAGE = 0 AND W2InfoRec(1).FEDTAXWH = 0 AND W2InfoRec(1).SOCWAGE = 0 THEN
    IF W2InfoRec(1).SOCTAXWH = 0 AND W2InfoRec(1).MEDWAGES = 0 AND W2InfoRec(1).MEDTAXWH = 0 THEN
      IF W2InfoRec(1).SOCTIPS = 0 AND W2InfoRec(1).ALLOCTIP = 0 AND W2InfoRec(1).ADVEIC = 0 THEN
        IF W2InfoRec(1).DEPNDCAR = 0 AND W2InfoRec(1).NQPLAN = 0 AND W2InfoRec(1).BOX13AMT = 0 THEN
          IF W2InfoRec(1).BOX13AM1 = 0 THEN
            GOTO SkipDontPrintEm
          END IF
        END IF
      END IF
    END IF
  END IF

  ECnt = ECnt + 1
  TFedGross# = Round#(TFedGross# + W2InfoRec(1).FEDWAGE)
  TStaGross# = Round#(TStaGross# + W2InfoRec(1).STAWAGE)
  TSocGross# = Round#(TSocGross# + W2InfoRec(1).SOCWAGE)
  TMedGross# = Round#(TMedGross# + W2InfoRec(1).MEDWAGES)
  TBox13a# = Round#(TBox13a# + W2InfoRec(1).BOX13AMT)
  TBox13b# = Round#(TBox13b# + W2InfoRec(1).BOX13AM1)
  TFedTax# = Round#(TFedTax# + W2InfoRec(1).FEDTAXWH)
  TStaTax# = Round#(TStaTax# + W2InfoRec(1).STATAXWH)
  TSocTax# = Round#(TSocTax# + W2InfoRec(1).SOCTAXWH)
  TMedTax# = Round#(TMedTax# + W2InfoRec(1).MEDTAXWH)
  TBox14a# = Round#(TBox14a# + W2InfoRec(1).BOX14AMT)
  TBox14b# = Round#(TBox14b# + W2InfoRec(1).BOX14AM1)
  TBox12# = Round#(TBox12# + W2InfoRec(1).BENFBOX1)
  TAdvEIC# = Round#(TAdvEIC# + W2InfoRec(1).ADVEIC)
  TBox10# = Round#(TBox10# + W2InfoRec(1).DEPNDCAR)
  TBox11# = Round#(TBox11# + W2InfoRec(1).NQPLAN)
  
  RSET FEDGROSS(1) = FUsing$(STR$(W2InfoRec(1).FEDWAGE), Image1$)
  RSET STAGROSS(1) = FUsing$(STR$(W2InfoRec(1).STAWAGE), Image1$)
  RSET SOCGROSS(1) = FUsing$(STR$(W2InfoRec(1).SOCWAGE), Image1$)
  RSET MEDGROSS(1) = FUsing$(STR$(W2InfoRec(1).MEDWAGES), Image1$)
  
  RSET Box13a(1) = FUsing$(STR$(W2InfoRec(1).BOX13AMT), Image1$)
  RSET Box13b(1) = FUsing$(STR$(W2InfoRec(1).BOX13AM1), Image1$)
  RSET FEDTAX(1) = FUsing$(STR$(W2InfoRec(1).FEDTAXWH), Image1$)
  RSET STATAX(1) = FUsing$(STR$(W2InfoRec(1).STATAXWH), Image1$)
  RSET SOCTAX(1) = FUsing$(STR$(W2InfoRec(1).SOCTAXWH), Image1$)
  RSET MEDTAX(1) = FUsing$(STR$(W2InfoRec(1).MEDTAXWH), Image1$)
  RSET Box14a(1) = FUsing$(STR$(W2InfoRec(1).BOX14AMT), Image1$)
  RSET Box14b(1) = FUsing$(STR$(W2InfoRec(1).BOX14AM1), Image1$)
  RSET ADVEIC(1) = FUsing$(STR$(W2InfoRec(1).ADVEIC), Image1$)
  RSET Box10(1) = FUsing$(STR$(W2InfoRec(1).DEPNDCAR), Image1$)
  RSET Box11(1) = FUsing$(STR$(W2InfoRec(1).NQPLAN), Image1$)
  RSET Box12(1) = FUsing$(STR$(W2InfoRec(1).BENFBOX1), Image1$)
  
  LSET Line1(1) = UCASE$(QPTrim$(Emp2Rec(1).EMPLNAME) + ", " + QPTrim$(Emp2Rec(1).EMPFNAME))
  MID$(Line1(1), 27) = ADVEIC(1)
  MID$(Line1(1), 40) = Box10(1)
  MID$(Line1(1), 53) = Box11(1)
  MID$(Line1(1), 66) = Box12(1)

'FPut RHandle, Emp2Rec(1).EmpNo + CrLf$

  FPut RHandle, Line1(1) + CrLf$
  FPut RHandle, FEDGROSS(1) + STAGROSS(1) + SOCGROSS(1) + MEDGROSS(1) + Box13a(1) + Box13b(1) + CrLf$
  FPut RHandle, FEDTAX(1) + STATAX(1) + SOCTAX(1) + MEDTAX(1) + Box14a(1) + Box14b(1) + CrLf$
  FPut RHandle, CrLf$
  LineCnt = LineCnt + 4
SkipDontPrintEm:
RETURN
  
PrintW2Summary:
  EmpOk = False
  RSET EmpCnt(1) = FUsing$(STR$(ECnt), "####")
  RSET FEDGROSS(1) = FUsing$(STR$(TFedGross#), Image1$)
  RSET STAGROSS(1) = FUsing$(STR$(TStaGross#), Image1$)
  RSET SOCGROSS(1) = FUsing$(STR$(TSocGross#), Image1$)
  RSET MEDGROSS(1) = FUsing$(STR$(TMedGross#), Image1$)
  RSET Box13a(1) = FUsing$(STR$(TBox13a#), Image1$)
  RSET Box13b(1) = FUsing$(STR$(TBox13b#), Image1$)
  RSET FEDTAX(1) = FUsing$(STR$(TFedTax#), Image1$)
  RSET STATAX(1) = FUsing$(STR$(TStaTax#), Image1$)
  RSET SOCTAX(1) = FUsing$(STR$(TSocTax#), Image1$)
  RSET MEDTAX(1) = FUsing$(STR$(TMedTax#), Image1$)
  
  RSET ADVEIC(1) = FUsing$(STR$(TAdvEIC#), Image1$)
  RSET Box12(1) = FUsing$(STR$(TBox12#), Image1$)
  RSET Box10(1) = FUsing$(STR$(TBox10#), Image1$)
  RSET Box11(1) = FUsing$(STR$(TBox11#), Image1$)
  
  RSET Box14a(1) = FUsing$(STR$(TBox14a#), Image1$)
  RSET Box14b(1) = FUsing$(STR$(TBox14b#), Image1$)
  
PrintEm:
  FPut RHandle, FF$ + CrLf$
  FPut RHandle, Dash(1)
  FPut RHandle, "Report Totals" + CrLf$ + CrLf$
  FPut RHandle, "                  W2 Forms      Adv EIC       Box 10       Box 11       Box 12" + CrLf$
  FPut RHandle, EmpCnt(1) + ADVEIC(1) + Box11(1) + Box10(1) + Box12(1) + CrLf$ + CrLf$
  FPut RHandle, "    Fed Gross  State Gross    Soc Gross    Med Gross       Box 13       Box 13" + CrLf$
  FPut RHandle, FEDGROSS(1) + STAGROSS(1) + SOCGROSS(1) + MEDGROSS(1) + Box13a(1) + Box13b(1) + CrLf$ + CrLf$
  FPut RHandle, "      Fed Tax    State Tax      Soc Tax      Med Tax       Box 14       Box 14" + CrLf$
  FPut RHandle, FEDTAX(1) + STATAX(1) + SOCTAX(1) + MEDTAX(1) + Box14a(1) + Box14b(1) + CrLf$
  FPut RHandle, Dash(1)
  FPut RHandle, FF$

  RETURN
  
  
END SUB

SUB WaitForKey
LabelForKeyTrap:
DO
  a$ = INKEY$
LOOP UNTIL LEN(a$) > 0
END SUB

SUB WazzWind2 (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
  REDIM Frame$(1 TO 6)

  IF FrameType < 1 OR FrameType > 5 THEN FrameType = 1

  Style$ = MID$("⁄ƒø≥¿Ÿ’Õ∏≥‘æ÷ƒ∑∫”Ω…Õª∫»º      ", ((FrameType - 1) * 6) + 1, 6)

  FOR a = 1 TO 6
    Frame$(a) = MID$(Style$, a, 1)
  NEXT

  LDCol = LeftCol + ((RghtCol - LeftCol) \ 2)
  RDCol = RghtCol - ((RghtCol - LeftCol) \ 2)
  TDRow = TopRow + ((BotRow - TopRow) \ 2)
  BDRow = BotRow - ((BotRow - TopRow) \ 2)

  CALL HideCursor
  DO
    LDCol = LDCol - 3
    IF LDCol < LeftCol THEN LDCol = LeftCol
    RDCol = RDCol + 3
    IF RDCol > RghtCol THEN RDCol = RghtCol

    IF TDRow > TopRow THEN TDRow = TDRow - 1
    IF BDRow < BotRow THEN BDRow = BDRow + 1

    Cnt = (RDCol - LDCol) - 1

    T1$ = STRING$(Cnt, Frame$(2))
    Middle$ = Frame$(4) + STRING$(Cnt, 32) + Frame$(4)
    QPrintRC Frame$(1) + T1$ + Frame$(3), TDRow, LDCol, FrameColor
    FOR i = (TDRow + 1) TO (BDRow - 1)
      QPrintRC Middle$, i, LDCol, FrameColor
    NEXT
    QPrintRC Frame$(5) + T1$ + Frame$(6), BDRow, LDCol, FrameColor
    IF Shadow THEN
      MPaintBox BDRow + 1, LDCol + 2, BDRow + 1, LDCol + Cnt + 3, 8
      MPaintBox TDRow + 1, RDCol + 1, BDRow, RDCol + 2, 8
    END IF

  LOOP UNTIL LDCol = LeftCol AND BDRow = BotRow

  CALL ShowCursor
  ERASE Frame$

END SUB

