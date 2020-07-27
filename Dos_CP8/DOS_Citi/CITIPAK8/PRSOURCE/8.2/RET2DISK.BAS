DEFINT A-Z
DECLARE FUNCTION ASCII% (Strng$)
DECLARE SUB KillFile (FileName$)
DECLARE FUNCTION NotFromPRMain% ()
DECLARE SUB SaveField (FldNo%, Form$(), Fld() AS ANY, BadFld%)
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB WaitForAction ()
DECLARE SUB FCreate (FileName$)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FClose (Handle%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FPutRT (Handle%, Source AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutRTA (Handle%, SEG Source AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutT (Handle%, Source AS ANY, NumBytes%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetRT (Handle%, Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FGetT (Handle%, Dest AS ANY, NumBytes%)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB FSeek (Handle%, Location&)
DECLARE SUB FFlush (Handle%)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FEof% (Handle%)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION FLoc& (Handle%)
DECLARE FUNCTION FLof& (Handle%)
DECLARE SUB InsertT (SEG StartElement AS ANY, ElSize%, NumEls%)
DECLARE FUNCTION Monitor ()
DECLARE SUB MScrnSave (UlRow, UlCol, LRRow, LRCol, SEG Address)
DECLARE SUB MScrnRest (UlRow, UlCol, LRRow, LRCol, SEG Address)
DECLARE SUB StuffBuf (Ky$)
DECLARE SUB GetCursor (x, Y, Button)
DECLARE SUB MQPrint (x$, Colr)
DECLARE SUB MPaintBox (UlRow, UlCol, LRRow, LRCol, Colr)
DECLARE SUB ButtonPress (ButNo%, Down%, Presses%, x%, Y%)
DECLARE SUB HideCursor ()
DECLARE SUB TextCursor (Fg, Bg)
DECLARE SUB ShowCursor ()
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
DECLARE FUNCTION QPStrI$ (Num%)
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
DECLARE FUNCTION QPRTrim$ (Text$)
DECLARE FUNCTION PromptSaveData% ()
DECLARE SUB PayMenu (BYVAL MenuNum%, Choice%, NumOfItems%)
DECLARE SUB VertMenu (Item$(), Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB DisplayMiscScrn (ScrnName$)
DECLARE SUB CursorOff ()
DECLARE SUB RestScrn (ScrnArray%())
DECLARE SUB SaveScrn (ScrnArray%())
DECLARE SUB Pause (Ticks)
DECLARE SUB BlockClear ()
DECLARE FUNCTION WEnvTest% ()
DECLARE FUNCTION QPValI% (Number$)
DECLARE SUB DisplayMiscScrn (ScrnName$)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE FUNCTION Date2Num (Dt$)
DECLARE FUNCTION Num2Date$ (N)
DECLARE FUNCTION QPTrim$ (String2Trim$)
DECLARE FUNCTION Round# (N#)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE SUB PrintRptFile (RptTitle$, RptFileName$, LPTPort, RetCode, EntryPoint)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)

'procedures in this module
DECLARE FUNCTION PadNum$ (TotNumLen%, Num$)
DECLARE SUB ALRet2Disk ()
  
'$INCLUDE: 'DefCnf.bi'                  'Defines TYPE for monitor/color info.
'$INCLUDE: 'FieldInf.bi'                '        "        field information
'$INCLUDE: 'FormEdit.bi'                '        "        form editing info.
'$INCLUDE: 'PRFiles.bi'
'$INCLUDE: 'PREmpRec.bi'
'$INCLUDE: 'PRTRANS.bi'
'$INCLUDE: 'PageInfo.bi'    '        "          Page information
'$INCLUDE: 'Retire.Bi'

TYPE EmpRetInfoRecType
  EmpName AS STRING * 30
  EmpSSN AS STRING * 11
  MoGross AS DOUBLE
  MoRet AS DOUBLE
END TYPE

'--original from bob jacquay
'TYPE ContCardRecType
'  Blank AS STRING * 6
'  EName AS STRING * 19
'  EAmt AS STRING * 11
'  ESSN AS STRING * 9
'  fill AS STRING * 26
'  RecCode AS STRING * 3
'END TYPE

'--new from ret system
TYPE ContCardRecType
  FullTimeSvc AS STRING * 5       '1-5
  Blank1 AS STRING * 1            '6
  EName AS STRING * 19            '7-25
  Blank2 AS STRING * 5            '26-30
  EAmt AS STRING * 6              '31-36
  ESSN AS STRING * 9              '37-45
  Blank3 AS STRING * 9            '46-54
  RecCode AS STRING * 3           '55-57
  RSA1 AS STRING * 7              '58-64
  PartSvc AS STRING * 5           '65-69
  RSADedMo AS STRING * 2          '70-71
  UnitCode AS STRING * 3          '72-74
  Blank4 AS STRING * 4            '75-78
  DeptCode AS STRING * 1          '79
  Blank5 AS STRING * 1            '80
END TYPE


STACK 5000
'$INCLUDE: 'SetCnf.bi'
CONST False = 0
CONST True = NOT False

'CLS
'REDIM ContCard(1) AS ContCardRecType: ContCardRecLen = LEN(ContCard(1))
'PRINT ContCardRecLen 'needs to be 80
'END


ALRet2Disk

SUB ALRet2Disk
  
  'ON LOCAL ERROR GOTO FixErr
  BlockClear

  LibName$ = "PRQL\PRMISCRN"
  FormName$ = "ALRET"
  REDIM TempScrn(0)
  REDIM Frm(1) AS FormInfo
  LibSize = ScrnLibSize(LibName$)
  REDIM ScrnLib(LibSize)
  LoadScrnLib LibName$, ScrnLib(), ErrorCode
  NumFlds = LibNumberOfFields(LibName$, FormName$)
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  StartEl = 0
  LibGetFldDef LibName$, FormName$, StartEl, Fld(), Form$(), ErrCode
  
  Frm(1).FldNo = 1               'Start editing on field #1
  Frm(1).InsStat = False         'Set the insert state (-1 = Insert on)
  Frm(1).StartEl = 0             'Set form starting element to 0 and
  
  Action = 1
  Lib2Scrn FormName$, ScrnLib(), MonoCode, -1, ErrCode
    
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F0Key, EscKey
      SELECT CASE Frm(1).KeyCode
        CASE EscKey
          CancelFlag = True
          ExitFlag = True
        CASE F0Key
          ExitFlag = True
      END SELECT
    END SELECT

  LOOP UNTIL ExitFlag = True
  IF CancelFlag THEN EXIT SUB
  
  CursorOff
  
  Drive$ = Form$(1, 0)
  DoDiskFile = LEN(QPTrim$(Drive$))
  BegDate = Date2Num(LEFT$(Form$(2, 0), 2) + "01" + RIGHT$(Form$(2, 0), 5))
  EndDate = Date2Num(Form$(2, 0))

  DIM Emp2 AS EmpData2Type
  ERecLen = LEN(Emp2)
  OPEN EmpData2Name FOR RANDOM AS #1 LEN = ERecLen
  NumOFERec = LOF(1) \ ERecLen

  DIM Tr1 AS TransRecType
  TRecLen = LEN(Tr1)
  OPEN TransHistFileName FOR RANDOM AS #2 LEN = TRecLen
  NumOfTRec = LOF(2) \ TRecLen

  REDIM EmpRetInfo(1 TO NumOFERec)  AS EmpRetInfoRecType
  RetInfoArrayLen = LEN(EmpRetInfo(1))

  RptFileName$ = "PRRPTS\CONTCARD.PRN"
  OPEN RptFileName$ FOR OUTPUT AS #3 LEN = 80
  
  IF DoDiskFile THEN
    DiskFileName$ = Drive$ + ":\CONTCARD"
    IF Exist(DiskFileName$) THEN
      KILL DiskFileName$
    END IF

    REDIM ContCard(1) AS ContCardRecType
    ContCardRecLen = LEN(ContCard(1))
    OPEN DiskFileName$ FOR RANDOM AS #4 LEN = ContCardRecLen
  END IF

  DisplayMiscScrn "PROCBAK"
  
  'ToPrint$ = SPACE$(80)
  E$ = SPACE$(27)
  Date1 = Date2Num(SearchDate$)
  RptTitle$ = "CONTCARD REPORT"
  FF$ = CHR$(12)
  MaxLines = 55
  Fmt$ = "######.##"
  GTMoGross# = 0
  GTMoRet# = 0
  NLen = 6   '--length of number
  LPTPort = 1
  P = 0  '--Num employees printed
  UnitCode$ = "BRE"

  GOSUB PageHeader

  FOR Cnt = 1 TO NumOFERec

    GET #1, Cnt, Emp2
    LSET E$ = QPTrim$(Emp2.EMPLNAME) + ", " + QPTrim$(Emp2.EMPFNAME)
    QPrintRC E$, 13, 28, -1

    IF QPTrim$(Emp2.EMPRETTP) = "MEMBER" THEN
      TrNum = Emp2.LastTransRec
      Found = 0

      '--Get employee's monthly gross pay and ret ded
      DO UNTIL TrNum = 0
        GET #2, TrNum, Tr1
        IF Tr1.CheckDate >= BegDate AND Tr1.CheckDate <= EndDate THEN
          EMoGross# = EMoGross# + Tr1.RetGrossPay
          EMoRet# = EMoRet# + Tr1.RetireAmt
        END IF
        TrNum = Tr1.PrevTransRec
      LOOP

      '--If employee has ret amts
      IF EMoGross# <> 0 OR EMoRet# <> 0 THEN
        P = P + 1
        EmpRetInfo(P).EmpName = QPTrim$(Emp2.EMPLNAME) + ", " + QPTrim$(Emp2.EMPFNAME)
        EmpRetInfo(P).EmpSSN = Emp2.EmpSSN
        EmpRetInfo(P).MoGross = EMoGross#
        EmpRetInfo(P).MoRet = EMoRet#
        EMoGross# = 0
        EMoRet# = 0
      END IF

    END IF
  NEXT

  CLOSE #1
  CLOSE #2
  
  REDIM PRESERVE EmpRetInfo(1 TO P) AS EmpRetInfoRecType

  'SortT SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%
  'TYPE EmpRetInfoRecType
  '  EmpName AS STRING * 30
  '  EmpSSN AS STRING * 11
  '  MoGross AS DOUBLE
  '  MoRet AS DOUBLE
  'END TYPE
  
  SortT EmpRetInfo(1), P, 0, 57, 0, 30

  '--put it to disk here after sorting if necessary
  IF DoDiskFile THEN
    GOSUB PrintReport2Disk
  END IF
  
  GOSUB PrintReport
  CLOSE #3
  PrintRptFile RptTitle$, RptFileName$, LPTPort, RetCode, EntryPoint
  COLOR 7, 0
  CLS

EXIT SUB


PageHeader:
  PRINT #3, "                              CONTCARD REPORT"
  PRINT #3, "                          MONTH ENDING: " + Form$(2, 0)
  PRINT #3,
  PRINT #3, "EMPLOYEE'S NAME              SOC-SEC-NUM     MONTH'S GROSS      MONTH'S RETIRE"
  PRINT #3, "=============================================================================="
  LineCnt = 4
  RETURN


PrintReport:
  FOR R = 1 TO P
    LineCnt = LineCnt + 1
    IF LineCnt >= MaxLines THEN
      PRINT #3, FF$
      GOSUB PageHeader
    END IF
    PRINT #3, EmpRetInfo(R).EmpName;
    PRINT #3, TAB(32); EmpRetInfo(R).EmpSSN;
    PRINT #3, TAB(50); USING Fmt$; EmpRetInfo(R).MoGross;
    PRINT #3, TAB(70); USING Fmt$; EmpRetInfo(R).MoRet
    GTMoGross# = GTMoGross# + EmpRetInfo(R).MoGross
    GTMoRet# = GTMoRet# + EmpRetInfo(R).MoRet
  NEXT

  PRINT #3,
  PRINT #3, "REPORT TOTALS";
  PRINT #3, TAB(50); USING Fmt$; GTMoGross#;
  PRINT #3, TAB(70); USING Fmt$; GTMoRet#
  PRINT #3,
  PRINT #3, USING "RECORDS PROCESSED  #####"; P
  PRINT #3, FF$

  RETURN


PrintReport2Disk:
  LastRec = P + 1
  '--print the detail records
  FOR R = 1 TO P
    RetAmt$ = PadNum$(NLen, FUsing$(STR$(EmpRetInfo(R).MoRet), Fmt$))

    '--Bob's record
    'ContCard(1).Blank = ""
    'ContCard(1).EName = EmpRetInfo(R).EmpName
    'ContCard(1).EAmt = RetAmt$
    'ContCard(1).ESSN = EmpRetInfo(R).EmpSSN
    'ContCard(1).fill = ""
    'ContCard(1).RecCode = "BRE"

    '--new from ret system
    LSET ContCard(1).FullTimeSvc = ""
    LSET ContCard(1).Blank1 = ""
    LSET ContCard(1).EName = EmpRetInfo(R).EmpName
    LSET ContCard(1).Blank2 = ""
    LSET ContCard(1).EAmt = RetAmt$
    LSET ContCard(1).ESSN = EmpRetInfo(R).EmpSSN
    LSET ContCard(1).Blank3 = ""
    LSET ContCard(1).RecCode = ""
    LSET ContCard(1).RSA1 = ""
    LSET ContCard(1).PartSvc = ""
    LSET ContCard(1).RSADedMo = ""
    LSET ContCard(1).UnitCode = UnitCode$
    LSET ContCard(1).Blank4 = ""
    LSET ContCard(1).DeptCode = ""
    LSET ContCard(1).Blank5 = ""
   
    PUT #4, R, ContCard(1)
  NEXT
  
  '--print the summary record

  '--orginal
  'ContCard(1).Blank = "END$"
  'ContCard(1).EName = ""
  'ContCard(1).EAmt = ""
  'ContCard(1).ESSN = ""
  'ContCard(1).fill = ""
  'ContCard(1).RecCode = ""

  LSET ContCard(1).FullTimeSvc = "END$"
  LSET ContCard(1).Blank1 = ""
  LSET ContCard(1).EName = ""
  LSET ContCard(1).Blank2 = ""
  LSET ContCard(1).EAmt = ""
  LSET ContCard(1).ESSN = ""
  LSET ContCard(1).Blank3 = ""
  LSET ContCard(1).RecCode = ""
  LSET ContCard(1).RSA1 = ""
  LSET ContCard(1).PartSvc = ""
  LSET ContCard(1).RSADedMo = ""
  LSET ContCard(1).UnitCode = ""
  LSET ContCard(1).Blank4 = ""
  LSET ContCard(1).DeptCode = ""
  LSET ContCard(1).Blank5 = ""
  
  PUT #4, LastRec, ContCard(1)

  CLOSE #4
  RETURN

  
FixErr:
  COLOR 15, 4: CLS : BEEP
  SELECT CASE ERR
    CASE 71 'Disk not ready
      QPrintRC "Disk Not Ready. Insert disk in drive " + Drive$ + ":", 10, 2, -1
      QPrintRC "Press Esc to cancel or any other key to continue.", 11, 2, -1
      K$ = INPUT$(1)
      IF K$ = CHR$(27) THEN
        CLOSE : CLS : END
      ELSE
        RESUME
      END IF
    CASE ELSE
      QPrintRC "System Error: " + STR$(ERR), 10, 2, -1
      QPrintRC "Press key to exit to DOS.", 11, 2, -1
      K$ = INPUT$(1)
      CLOSE : COLOR 0, 7: CLS : END
  END SELECT
  RETURN
  
END SUB

FUNCTION PadNum$ (TotNumLen%, NumIn$)

  Num$ = QPTrim(NumIn$)
  NumLen = LEN(Num$)
  StripNum$ = LEFT$(Num$, NumLen - 3) + RIGHT$(Num$, 2)
  PadNum$ = STRING$(TotNumLen - LEN(StripNum$), "0") + StripNum$

END FUNCTION

