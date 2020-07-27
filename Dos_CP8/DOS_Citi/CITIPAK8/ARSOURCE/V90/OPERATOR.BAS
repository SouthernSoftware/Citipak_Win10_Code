DEFINT A-Z
DECLARE SUB DisplayARScrn (ScrnName$)
DECLARE FUNCTION IsDeleted% (AcctNo&)
DECLARE SUB AddEditDeposit (OPERNUM%, PostDate$)
DECLARE SUB OperatorEntry (OperOKFlag%, OPERNUM%, PostDate$)
DECLARE SUB AddEditOperator ()
DECLARE SUB AuthorizeOperator (oktoadd%)
DECLARE SUB DisplayOperEntry ()
DECLARE SUB ValidateOperator (OPERNUM%, OperPassword$, OperOKFlag%)
DECLARE FUNCTION GetOperator% ()
DECLARE SUB StuffBuf (Ky$)
DECLARE FUNCTION PromptSaveData% ()
DECLARE SUB WaitForAction ()
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB AddEditPayment (OPERNUM, PostDate$)
DECLARE SUB PostTransEntries (OPERNUM, PostDate$)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB BlockClear ()
DECLARE SUB CursorOff ()
DECLARE SUB MPaintBox (UlRow%, UlCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB MScrnRest (UlRow, UlCol, LRRow, LRCol, SEG Address)
DECLARE SUB MScrnSave (UlRow, UlCol, LRRow, LRCol, SEG Address)
DECLARE SUB Pause3 (MillaSecs%, ProcessorSpeed&)
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
DECLARE SUB SaveScrn (Array())
DECLARE SUB RestScrn (Array())
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValL& (Numb$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
'DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB ClearScrn ()
DECLARE SUB HideCursor ()
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION WEnvTest% ()
DECLARE FUNCTION Date2Num% (WhatDate$)
DECLARE FUNCTION Num2Date$ (WhatDate%)
DECLARE FUNCTION Exist% (FileName$)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE FUNCTION FLof& (Handle%)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Round# (DoubleNum#)
DECLARE SUB FPutRTA (Handle%, SEG Source AS ANY, RecNo&, RecSize%)
DECLARE SUB FClose (FileHandle)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB UBLog (Text$)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPORT%, RetCode%, EntryPoint%)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB KillFile (FileName$)
DECLARE SUB ButtonPress (WhatBtn, N, MooseButton, MCol, MRow)   ' ----- Check Mouse activity
DECLARE SUB GetCursor (a, b, NewStatus)
DECLARE SUB OPENDrawer (RecpPort)
DECLARE SUB QPSound (Freq, Dur)

  '---- ASM subs

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.bi'
  '$INCLUDE: 'SetCnf.BI'
  ''$INCLUDE: 'newCUST.bi'
  ''$INCLUDE: 'UBtrans.bi'
  ''$INCLUDE: 'ubsetup.BI'
  '$INCLUDE: 'cmfiles.BI'

  ''$INCLUDE: 'init.BI'

  STACK 8000

  TYPE FLen2
    V AS STRING * 64
  END TYPE

  CONST False = 0, True = NOT False

  DIM SHARED OPERNUM AS INTEGER, OperOKFlag AS INTEGER

'DEFINT A-Z
'SUB DisplayOPScrn (ScrnName$)
'  LibFile2Scrn "UB", ScrnName$, MonoCode%, Attribute%, ErrorCode%
'  IF ErrorCode% <> 0 THEN
'    PRINT "Screen Error: "; ScrnName$
'    END
'  END IF
'
'END SUB

SUB AddEditOperator
  
  AuthorizeOperator oktoadd
  
  IF oktoadd = 0 THEN EXIT SUB
  
  LibName$ = "UB"
  ScrnName$ = "UBNEWOP"
  BlockClear
  DisplayARScrn ScrnName$
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  Frm(1).StayOnField = True
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  Action = 1
  REDIM CMOperRec(1) AS CMOperRecType
  CMOperRecLen = LEN(CMOperRec(1))
  CMFile = FREEFILE
  OPEN "CMOPER.DAT" FOR RANDOM SHARED AS CMFile LEN = CMOperRecLen
  NumRecs = LOF(CMFile) \ CMOperRecLen
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF Frm(1).PrevFld = 2 AND edit <> True THEN
      FOR Cnt = 1 TO NumRecs
        GET CMFile, Cnt, CMOperRec(1)
        IF VAL(Form$(1, 0)) = CMOperRec(1).OperatorNumber AND VAL(Form$(1, 0)) > 0 THEN
          Form$(2, 0) = CMOperRec(1).OperatorName
          Form$(3, 0) = CMOperRec(1).OperatorPassword
          Fld(1).Protected = True
          RecNo = Cnt
          edit = True
          Action = 1
          EXIT FOR
        END IF
      NEXT Cnt
    END IF
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      
      CMOperRec(1).OperatorNumber = VAL(Form$(1, 0))
      CMOperRec(1).OperatorName = Form$(2, 0)
      CMOperRec(1).OperatorPassword = Form$(3, 0)
      IF RecNo = 0 THEN
        PUT CMFile, LOF(CMFile) \ CMOperRecLen + 1, CMOperRec(1)
      ELSE
        PUT CMFile, RecNo, CMOperRec(1)
      END IF
      
      CLOSE CMFile
      Done = True
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  
END SUB

SUB AuthorizeOperator (oktoadd)
  
  LibName$ = "UB"
  ScrnName$ = "OPERAUTH"
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  Frm(1).StayOnField = True
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  BlockClear
  DisplayARScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      IF Form$(1, 0) = "CASHMGMT" THEN
        oktoadd = 1
      ELSE
        oktoadd = 0
      END IF
      Done = True
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  
END SUB

SUB DisplayOperEntry
  BlockClear
  DisplayARScrn "UBCHOSOP"
END SUB

SUB OperatorEntry (OperOKFlag, OPERNUM, PostDate$)
  
  Action = 1
  added = 0
  
  LibName$ = "UB.QSL"
  ScrnName$ = "UBCHOSOP"
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  Frm(1).StayOnField = True
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  DisplayOperEntry
  
  FirstTime = True
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      LSET Form$(3, 0) = DATE$
      Action = 1
    END IF
    
    SELECT CASE Frm(1).KeyCode
    CASE F7KEY
      AddEditOperator
      Action = 1
      DisplayOperEntry
    CASE F10Key
      OPERNUM = VAL(Form$(1, 0))
      OperPassword$ = Form$(2, 0)
      PostDate$ = Form$(3, 0)
      GOSUB CheckPostDate
      IF BadDate THEN
        QPrintRC "<-- DATE NOT SET", 12, 49, 4
        BEEP
        SLEEP 1
        DisplayOperEntry
      END IF
      
      IF OPERNUM > 0 THEN
        ValidateOperator OPERNUM, OperPassword$, OperOKFlag
        IF OperOKFlag > 0 THEN
          Done = True
        ELSE
          DisplayOperEntry
          Action = 1
          Frm(1).FldNo = 1
        END IF
      ELSE
        OperOKFlag = 0
        Done = True
      END IF
    CASE ESC
      OperOKFlag = -1
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  
  EXIT SUB
  
CheckPostDate:
  IF VAL(LEFT$(PostDate$, 2)) < 1 OR VAL(LEFT$(PostDate$, 2)) > 12 THEN
    BadDate = True
  ELSEIF VAL(MID$(PostDate$, 4, 2)) < 1 OR VAL(MID$(PostDate$, 4, 2)) > 31 THEN
    BadDate = True
  ELSE
    BadDate = False
  END IF
  RETURN
  
END SUB

SUB ValidateOperator (OPERNUM, OperPassword$, OperOKFlag)
  
  REDIM CMOperRec(1) AS CMOperRecType
  OperFileLength = LEN(CMOperRec(1))
  
  FileHandle = FREEFILE
  OPEN "CMOPER.DAT" FOR RANDOM SHARED AS #FileHandle LEN = OperFileLength
  
  NumOperRecs = LOF(FileHandle) \ OperFileLength
  
  IF NumOperRecs = 0 THEN
    OperOKFlag = 0
    GOTO ExitValidate
  END IF
  
  FOR Cnt = 1 TO NumOperRecs
    GET FileHandle, Cnt, CMOperRec(1)
    IF OPERNUM = CMOperRec(1).OperatorNumber AND OperPassword$ = CMOperRec(1).OperatorPassword THEN
      OperOKFlag = Cnt
      EXIT FOR
    END IF
  NEXT
  
ExitValidate:
  CLOSE
  IF OperOKFlag = 0 THEN
    BEEP
    OK = MsgBox("UB", "BADOPER")
  END IF
  
END SUB

