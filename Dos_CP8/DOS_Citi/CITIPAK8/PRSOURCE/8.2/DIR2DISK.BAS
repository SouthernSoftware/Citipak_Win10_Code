DEFINT A-Z
DECLARE SUB DIRESCMagMediaRpt ()
DECLARE FUNCTION MakeDIRWages$ (EMPGrossAmt#)
DECLARE SUB ALESC2Disk ()
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
DECLARE FUNCTION Value# (Number$, ECode)

'procedures in this module
DECLARE FUNCTION PadNum$ (TotNumLen%, Num$)
DECLARE SUB ALRet2Disk ()
  
'$INCLUDE: 'DefCnf.bi'                  'Defines TYPE for monitor/color info.
'$INCLUDE: 'FieldInf.bi'                '        "        field information
'$INCLUDE: 'FormEdit.bi'                '        "        form editing info.
'$INCLUDE: 'PRFiles.bi'
'$INCLUDE: 'PRUNIT.Bi'
'$INCLUDE: 'PRSYSCTR.Bi'
'$INCLUDE: 'PREmpRec.bi'
'$INCLUDE: 'PRTRANS.bi'
'$INCLUDE: 'oPRtypes.bi'
'$INCLUDE: 'PageInfo.bi'    '        "          Page information
'$INCLUDE: 'Retire.Bi'
'$INCLUDE: 'PRRpts.BI'
'$INCLUDE: 'alquartr.BI'

TYPE EmpRetInfoRecType
  EmpName AS STRING * 30
  EMPSSN AS STRING * 11
  MoGross AS DOUBLE
  MoRet AS DOUBLE
END TYPE

TYPE ContCardRecType
  Blank AS STRING * 6
  EName AS STRING * 19
  EAmt AS STRING * 11
  ESSN AS STRING * 9
  fill AS STRING * 26
  RecCode AS STRING * 3
END TYPE

STACK 5000
'$INCLUDE: 'SetCnf.bi'
CONST False = 0
CONST True = NOT False

  DIRESCMagMediaRpt
  BlockClear
  DisplayMiscScrn "UPDATEOK"
  WaitForAction
CLS

SUB DIRESCMagMediaRpt
  
  REDIM TempScrn(1)
  
  ExitFlag = False
  FirstTime = True
  
  OKFlag = False
  
  REDIM Frm(1) AS FormInfo
  
  FormName$ = "ALESC"
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
  LOCATE 1, 1
  CursorOff
  BlockClear
  DisplayMiscScrn FormName$
  
'new
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF FirstTime THEN
      FirstTime = False
      LSET Form$(3, 0) = RIGHT$(DATE$, 4)
      Action = 1
    END IF

    SELECT CASE Frm(1).KeyCode
    CASE F0Key, EscKey
      SELECT CASE Frm(1).KeyCode
        CASE EscKey
          CancelFlag = True
          ExitFlag = True
        CASE F0Key
          GOSUB ChkInput
          IF InputOK THEN
            ExitFlag = True
          END IF
          Action = 1
      END SELECT
    END SELECT

  LOOP UNTIL ExitFlag = True

  IF CancelFlag THEN EXIT SUB
  
  CursorOff
  

  Disk$ = Form$(1, 0)
  DiskFileName$ = Disk$ + ":\ALAWAGES"
  Year$ = Form$(3, 0)

  IF Exist(DiskFileName$) THEN
    KILL DiskFileName$
  END IF

  ERASE Form$, Fld, TempScrn, Frm
  
  REDIM TransHRec(1)   AS TransRecType
  REDIM Emp2Rec(1)     AS EmpData2Type
  
  MaxPerBlock = 41
  TEmpCnt = 0
  
  EmpRecSize = LEN(Emp2Rec(1))
  TRecSize = LEN(TransHRec(1))
  
  IdxRecLen = 2
  IdxFileSize& = FileSize(EmpIdxLName)
  NumEmpRecs = IdxFileSize& \ IdxRecLen
  
  REDIM IdxBuff(1 TO NumEmpRecs)
  FGetAH EmpIdxLName, IdxBuff(1), IdxRecLen, NumEmpRecs
  
  RptTitle$ = "DIR Magnetic Media Quarterly Report"
  ShowProcessingScrn RptTitle$
  
  RHandle = FREEFILE
  OPEN DiskFileName$ FOR RANDOM AS RHandle LEN = 80

  THandle = FREEFILE
  OPEN TransHistFileName FOR RANDOM AS THandle LEN = TRecSize

  DHandle = FREEFILE
  OPEN EmpData2Name FOR RANDOM AS DHandle LEN = EmpRecSize'open employee data file
  
  GOSUB Make.ABE.Recs
  
  FOR RecNo = 1 TO NumEmpRecs
    UsingThisOne = False
    EMPGrossAmt# = 0
    GET DHandle, IdxBuff(RecNo), Emp2Rec(1)
    IF Emp2Rec(1).LastTransRec <= 0 THEN
      GOTO SkipEm4
    END IF
    TranRecNum = Emp2Rec(1).LastTransRec

    DO
      GET THandle, TranRecNum, TransHRec(1)

      SELECT CASE TransHRec(1).CheckDate
      CASE LoDate TO HiDate
        GOSUB SumEmpESCGross
        UsingThisOne = True
      CASE ELSE
      END SELECT

      IF TransHRec(1).PrevTransRec <= 0 THEN    'if no more transactions
        IF UsingThisOne THEN               'if there were any at all
          IF TEmpCnt = MaxPerBlock THEN    'bottom of page?
            GOSUB Make.I.Rec               'write I Rec
          END IF
          GOSUB Make.EmpS.Rec              'then write emp totals
        END IF
        EXIT DO
      ELSE      'there are more trans to check
        TranRecNum = TransHRec(1).PrevTransRec
      END IF
    LOOP

SkipEm4:
    ShowPctComp RecNo, NumEmpRecs
  NEXT
  GOSUB Make.I.Rec
  GOSUB Make.TF.Recs
  
  CLOSE
  
  ERASE TransHRec, Emp2Rec, IdxBuff
  
EXIT SUB
  
Make.ABE.Recs:
  REDIM ATRec(1) AS ATransmitterType
  REDIM BBRec(1) AS BBasAuthRecType
  REDIM EERec(1) AS EEmployerRecType

  ATRec(1).RecID = "A"                   '1   Must equal "A"
  ATRec(1).RptPeriod = RptPer$           '4   Must equal 0397,0697,0997,1297
  ATRec(1).FedEIN = "636001208"          '9   636001208
  LSET ATRec(1).TranName = "CITY OF BREWTON"  '35  LSET,FILLED SPACES
  LSET ATRec(1).Filler = ""              '31  filled with spaces

  BBRec(1).RecID = "B"                   '1   Must equal "B"
  BBRec(1).RptPeriod = RptPer$           '4   Must equal 0397,0697,0997,1297
  BBRec(1).FedEIN = "636001208"          '9   636001208
  LSET BBRec(1).Computer = "IBM"         '8   "IBM" lset filled with blanks
  LSET BBRec(1).RecCode = "ASCII"        '6   "ASCII" lset filled with blanks
  LSET BBRec(1).Filler = ""              '52  filled with spaces

  EERec(1).RecID = "E"                   '1   Must equal "E"
  EERec(1).RptPeriod = RptPer$           '4   Must equal 0397,0697,0997,1297
  EERec(1).FedEIN = "636001208"          '9   636001208
  LSET EERec(1).Employer = "CITY OF BREWTON"  '24  LSET,FILLED SPACES
  EERec(1).DIRAcctNo = "0213057600"      '10  DIR Acount No.
  EERec(1).NameCode = "S"                ' 1  "S" Surname
  EERec(1).StateCode = "01"              ' 2  "01"
  EERec(1).UnitNum = "00000"             ' 5  "00000"
  LSET EERec(1).Filler = ""              '24

  PUT RHandle, , ATRec(1)
  PUT RHandle, , BBRec(1)
  PUT RHandle, , EERec(1)

RETURN
  
SumEmpESCGross:
  EMPGrossAmt# = Round#(EMPGrossAmt# + TransHRec(1).GrossPay)
RETURN
  
Make.EmpS.Rec:

  REDIM SERec(1) AS SEmpWageRecType
  SERec(1).RecID = "S"                   '1   Must equal "S"
  LSET SERec(1).EmpSOC = QPTrim$(Emp2Rec(1).EMPSSN)
  LSET SERec(1).EmpName = (QPTrim$(Emp2Rec(1).EMPLNAME) + ", " + QPTrim$(Emp2Rec(1).EMPFNAME))
  SERec(1).StateCode = "01"              ' 2  "01"
  SERec(1).RptPeriod = RptPer$           '4   Must equal 0397,0697,0997,1297
  SERec(1).Wages = MakeDIRWages$(EMPGrossAmt#)
  SERec(1).DirNumber = "0213057600"      '10  DIR Acount No.
  LSET SERec(1).Filler = ""
  PUT RHandle, , SERec(1)

  TEmpCnt = TEmpCnt + 1
  EmpCnt = EmpCnt + 1
  TotGrossAmt# = Round#(TotGrossAmt# + EMPGrossAmt#)
  TmpGrossAmt# = Round#(TmpGrossAmt# + EMPGrossAmt#)

  'sum to quarter grand totals
  
RETURN

Make.I.Rec:
  REDIM IIRec(1) AS IInterMedRecType
  IIRec(1).RecID = "I"                   '1   Must equal "I"
  IIRec(1).TotWages = RIGHT$(("000000" + MakeDIRWages$(TmpGrossAmt#)), 10)
  LSET IIRec(1).Filler = ""
  PUT RHandle, , IIRec(1)
  TEmpCnt = 0
  TmpGrossAmt# = 0
RETURN

Make.TF.Recs:
  REDIM TTRec(1) AS TTotalRecType
  REDIM TFRec(1) AS TFinalRecType

  TTRec(1).RecID = "T"                   '1   Must equal "T"
  TTRec(1).TotalSRecs = RIGHT$(("0000000" + QPTrim$(STR$(EmpCnt))), 7)
  TTRec(1).TotAllWage = RIGHT$(("0000000000000" + MakeDIRWages$(TotGrossAmt#)), 10)
  LSET TTRec(1).Filler = ""              '31  filled with spaces

  TFRec(1).RecID = "F"                   '1   Must equal "F"
  TFRec(1).TotalSRecs = TTRec(1).TotalSRecs
  LSET TFRec(1).Filler = ""              '31  filled with spaces

  PUT RHandle, , TTRec(1)
  PUT RHandle, , TFRec(1)

RETURN

ChkInput:

  Year$ = Form$(3, 0)
  Year = Value(Form$(3, 0), ECode)
  Quarter = Value(Form$(2, 0), ECode)

  IF Year < 1998 OR Year > 2010 THEN
    InputOK = False
    GOTO NoGood
  END IF

  IF (Quarter >= 1 OR Quarter <= 4) THEN
    InputOK = True
    RRYear$ = RIGHT$(Year$, 2)
    SELECT CASE Quarter
    CASE 1
      LoDate = Date2Num("01-01-" + Year$)
      HiDate = Date2Num("03-31-" + Year$)
      RptPer$ = "03" + RRYear$
    CASE 2
      LoDate = Date2Num("04-01-" + Year$)
      HiDate = Date2Num("06-30-" + Year$)
      RptPer$ = "06" + RRYear$
    CASE 3
      LoDate = Date2Num("06-31-" + Year$)
      HiDate = Date2Num("09-30-" + Year$)
      RptPer$ = "09" + RRYear$
    CASE 4
      LoDate = Date2Num("10-01-" + Year$)
      HiDate = Date2Num("12-31-" + Year$)
      RptPer$ = "12" + RRYear$
    END SELECT
  ELSE
    InputOK = False
    GOTO NoGood
  END IF

  Disk$ = UCASE$(Form$(1, 0))
  IF INSTR(COMMAND$, "TEST") > 0 THEN
    InputOK = True
    GOTO NoGood:
  END IF
  IF Disk$ = "A" OR Disk$ = "B" THEN
    InputOK = True
  ELSE
    InputOK = False
  END IF

NoGood:
RETURN
  
END SUB

FUNCTION MakeDIRWages$ (EMPGrossAmt#)

'  IF EMPGrossAmt# = 600 THEN STOP

  Zero$ = STRING$(9, "0")      'RSET, ZERO FILL, NO NON-NUMERIC

  Gross$ = QPTrim$(STR$(Round#(EMPGrossAmt#)))

  DecPoint = INSTR(Gross$, ".")

  IF DecPoint > 0 THEN
    Cent$ = QPTrim$(MID$(Gross$, DecPoint + 1))
  END IF

  CentLen = LEN(Cent$)
  SELECT CASE CentLen
  CASE 0
    Cent$ = "00"
  CASE 1
    Cent$ = Cent$ + "0"
  END SELECT

  IF DecPoint > 0 THEN
    Dollars# = Value#(LEFT$(Gross$, DecPoint - 1), ECode)
  ELSE
    Dollars# = Value#(Gross$, ECode)
  END IF

  Dollars$ = QPTrim$(STR$(Dollars#))

  MakeDIRWages$ = RIGHT$((Zero$ + Dollars$ + Cent$), 9)

END FUNCTION

