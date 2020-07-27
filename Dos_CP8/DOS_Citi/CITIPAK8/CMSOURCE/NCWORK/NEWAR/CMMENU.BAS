DEFINT A-Z    'CmMENU.BAS
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
DECLARE SUB PrintReceipt ()
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

  STACK 8000
  
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
      OPENDrawer
    CASE F7KEY
      ERASE CMOperRec, Frm, Form$, Fld
      'CMJourMenu OperRecNumber
      'PrintJournal OperRecNumber

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
      CASE 1
'        AddEditPayment OPERNUM, PostDate$
      CASE 2
'        AddEditDeposit OPERNUM, PostDate$
      CASE 3
'        EnterTaxPayments OPERNUM, PostDate$
      CASE 4
'        MiscPayEntry OperRecNumber, PostDate$
      CASE 5
        EnterPayment OPERNUM, PostDate$
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
  Choice$(1, 0) = "Utility Billing Payment"
  Choice$(2, 0) = "Utility Deposit Payment"
  Choice$(3, 0) = "Property Tax Payment"
  Choice$(4, 0) = "Miscellaneous Payment"
  Choice$(5, 0) = "Business License Payment"
RETURN

'here:
'  REDIM PayList(1 TO 1)      AS PayListType
'  RecLen = LEN(PayList(1))
'  OPEN "taxlop2.dat" FOR RANDOM AS #1 LEN = RecLen
'  GET #1, , PayList(1)
'  CLOSE
'  PRINT PayList(1).BillRec

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

