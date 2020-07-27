DEFINT A-Z    'CmMENU.BAS

DECLARE SUB CMLog (Text$)
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FUsing$ (Number$, Mask$)
DECLARE FUNCTION FindAcct% (AcctNum$)
DECLARE FUNCTION GetDEFPort% ()
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION QPTrim$ (X$)
DECLARE FUNCTION QPValI% (Number$)
DECLARE FUNCTION Round# (ThisAmount#)
DECLARE FUNCTION WEnvTest ()
DECLARE SUB AddEditDeposit (OPERNUM%, PostDate$)
DECLARE SUB AddEditOperator ()
DECLARE SUB AddEditPayment (OPERNUM%, PostDate$)
DECLARE SUB AddMiscCode ()
DECLARE SUB ArPaymen (OPERNUM%, PostDate$)
DECLARE SUB AuthorizeOperator (oktoadd)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB BlockClear ()
DECLARE SUB CMJourMenu (OperRecNumber%)
DECLARE SUB DisplayCMScrn (ScrnName$)
DECLARE SUB DisplayOperEntry ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB EnterPayment (OperRecNum%, PostDate$)
DECLARE SUB EnterTaxPayments (OPERNUM%, PostDate$)
DECLARE SUB GetMiscCodeRecord (MCFile, RecNo, MiscCodeValue$)
DECLARE SUB HideCursor ()
DECLARE SUB MPaintBox (ULRow%, ULCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB MiscPayEntry (OperRecNumber, PostDate$)
DECLARE SUB OPENDrawer ()
DECLARE SUB OpenAcctIdx (AcctIdxFileNum, NumIdxRecs)
DECLARE SUB OpenMiscCodeFile (NumOfMiscRecs, MCFile)
DECLARE SUB OpenSetupFile (SetUpFileNum%)
DECLARE SUB OperatorEntry (Results, OPERNUM, PostDate$)
DECLARE SUB PostAndPrint (Posted)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintJournal (OperNumb)
DECLARE SUB PrintMiscCodeList ()
'DECLARE SUB PrintReceipt ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPORT%, RetCode%, EntryPoint%)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB QPrintRC (t$, r%, c%, clr%)
DECLARE SUB RestScrn (ScrnArray())
DECLARE SUB SaveScrn (ScrnArray())
DECLARE SUB ShowCursor ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB ValidateOperator (OPERNUM, OperPassword$, Results)
DECLARE SUB WaitForAction ()

  TYPE Struct
    who AS STRING * 14
    RecNum AS INTEGER
  END TYPE
  
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'Qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'txpaymnt.BI'

  '*** Citipak Include Files Needed ***
  ''$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'CMFILES.BI'
  '$INCLUDE: 'gl.bi'

  STACK 6000
  
  CONST False = 0, True = NOT False

'GOTO here:

  REDIM Scrn1(0) AS INTEGER

  LOCATE 1, 1
  BlockClear

  IF INSTR(COMMAND$, "TEST") > 0 THEN
    OperOKFlag = 2
    OPERNUM = 2
    PostDate$ = DATE$
    GOTO SkipToHere
  END IF
  
  DO
    OperatorEntry OperOKFlag, OPERNUM, PostDate$
  LOOP WHILE OperOKFlag = 0

SkipToHere:
  IF OperOKFlag < 0 THEN GOTO ExitCMMenuProg
  OperRecNumber = OperOKFlag    'Set Operator Rec Num
  
Top:
  REDIM CMOperRec(1) AS CMOperRecType
  CMOperRecLen = LEN(CMOperRec(1))
  
  CMFile = FREEFILE
  OPEN "CMOPER.DAT" FOR RANDOM SHARED AS CMFile LEN = CMOperRecLen
  GET CMFile, OperRecNumber, CMOperRec(1)
  CLOSE CMFile
  
  OperName$ = LEFT$(CMOperRec(1).OperatorName, 19)
  OperNumb = CMOperRec(1).OperatorNumber
  
  LibName$ = "CM"
  ScrnName$ = "CMSOURC"
  
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  Help$ = "Cash Management System V8.0"
  
  GOSUB InitMulti
  DisplayCMScrn ScrnName$
  QPrintRC OperName$, 3, 55, 15
  
  SaveScrn Scrn1()
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE AltO
'unrem
      OPENDrawer
    CASE F7KEY
      ERASE CMOperRec, Frm, Form$, Fld
'unrem
      CMJourMenu OperRecNumber

'      PrintJournal OperRecNumber

      Action = 1
      WhatSource = 0
      GOTO Top
    CASE ESC
      Done = True

'    CASE IS <> 0
'      STOP
    END SELECT
    
    IF Frm(1).KeyCode > 0 THEN
      Source$ = QPTrim$(Form$(1, 0))
      IF LEN(Source$) > 0 THEN
        FOR Cnt = 1 TO 5
          IF INSTR(Choice$(Cnt, 0), Source$) THEN
            WhatSource = Cnt
          END IF
        NEXT
      END IF
    END IF
    
    IF WhatSource THEN
      SaveScrn Scrn1()
      SELECT CASE WhatSource
'      CASE 1
'        AddEditPayment OPERNUM, PostDate$
'      CASE 2
'        AddEditDeposit OPERNUM, PostDate$
'      CASE 3
'        EnterTaxPayments OPERNUM, PostDate$
      CASE 1
        'MiscPayEntry OperRecNumber, PostDate$

'      CASE 5
'        EnterPayment OPERNUM, PostDate$
      END SELECT
      LSET Form$(1, 0) = ""
      Frm(1).FldNo = 1
      Action = 1
      WhatSource = 0
      GOSUB InitMulti
      RestScrn Scrn1()
    END IF
  LOOP UNTIL Done
  
ExitCMMenuProg:
  
  IF Exist("CPMENU.EXE") THEN
    RUN "CPMENU"
  END IF
  
  HideCursor
  CLS
  END
  
InitMulti:
  REDIM Choice$(0 TO 5, 0 TO 0)
  Choice$(0, 0) = "1"
'  Choice$(1, 0) = "Utility Billing Payment"
'  Choice$(2, 0) = "Utility Deposit Payment"
'  Choice$(3, 0) = "Property Tax Payment"
  Choice$(1, 0) = "Miscellaneous Payment"
'  Choice$(5, 0) = "Business License Payment"

RETURN

'here:
'  REDIM PayList(1 TO 1)      AS PayListType
'  RecLen = LEN(PayList(1))
'  OPEN "taxlop2.dat" FOR RANDOM AS #1 LEN = RecLen
'  GET #1, , PayList(1)
'  CLOSE
'  PRINT PayList(1).BillRec

'here:
'
'  DIM CMTRRec(1) AS CMTransRecType
'  CMTrRecLen = LEN(CMTRRec(1))
'  NAME "CMTRANS.DAT" AS "OCMTRANS.DAT"
'
'  CHandle = FREEFILE
'  OPEN "CMTRANS.DAT" FOR RANDOM SHARED AS CHandle LEN = CMTrRecLen
'
'  OHandle = FREEFILE
'  OPEN "OCMTRANS.DAT" FOR RANDOM SHARED AS OHandle LEN = CMTrRecLen
'    RecCnt& = LOF(OHandle) \ CMTrRecLen
'    FOR zz& = 1 TO RecCnt&
'      GET OHandle, zz&, CMTRRec(1)
'      SELECT CASE zz&
'      CASE 15677, 15678, 15680
'      CASE ELSE
'        PUT CHandle, , CMTRRec(1)
'      END SELECT
'    NEXT
'
'  CLOSE

SUB CMLog (Text$) STATIC

  Today$ = DATE$
  Today$ = LEFT$(Today$, 2) + MID$(Today$, 4, 2) + RIGHT$(Today$, 2)

  TheTime$ = TIME$
  IF LEFT$(TheTime$, 1) = "0" THEN
    Hour = VAL(MID$(TheTime$, 2, 1))
  ELSE
    Hour = VAL(MID$(TheTime$, 1, 2))
  END IF

  SELECT CASE Hour
  CASE IS > 11
    Hour = Hour - 12
    IF Hour = 0 THEN Hour = 12
    AmPm$ = "pm"
  CASE 1 TO 12
    AmPm$ = "am"
  CASE 0
    Hour = 12
    AmPm$ = "am"
  END SELECT
  SELECT CASE Hour
    CASE 1 TO 9
      Hour$ = "0" + QPTrim$(STR$(Hour))
    CASE ELSE
      Hour$ = QPTrim$(STR$(Hour))
  END SELECT
  TheTime$ = Hour$ + ":" + MID$(TheTime$, 4) + AmPm$

  LogFile = FREEFILE
  OPEN "C:\CMLOG.DAT" FOR APPEND SHARED AS LogFile LEN = 255
  PRINT #LogFile, "CM: "; Today$; " @"; TheTime$
  PRINT #LogFile, "    "; Text$
  CLOSE #LogFile
  Text$ = "": TheTime$ = ""

END SUB

SUB PrintHelp (Help$)
  t$ = SPACE$(40)
  LSET t$ = Help$
  QPrintRC t$, 25, 2, -1
  t$ = ""
END SUB

SUB PrintTitle (Title$)
  t$ = SPACE$(40)
  LSET t$ = Title$
  QPrintRC t$, 1, 2, -1
  QPrintRC DATE$, 1, 70, -1
  t$ = ""
END SUB

