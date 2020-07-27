DEFINT A-Z    'CmMENU.BAS
DECLARE SUB EnterTaxPayment (OperRecNum%, PostDate$)
DECLARE SUB OpenARCustFile (NumOfArRecs%, ARFile%)
DECLARE SUB EnterPayment (OperRecNum%, PostDate$)
DECLARE SUB AddEditDeposit (OPERNUM%, PostDate$)
DECLARE SUB AddEditPayment (OPERNUM%, PostDate$)
DECLARE SUB ArPaymen (OPERNUM%, PostDate$)
DECLARE SUB AddEditOperator ()
DECLARE SUB AddMiscCode ()
DECLARE SUB AuthorizeOperator (oktoadd)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB BlockClear ()
DECLARE SUB DisplayOperEntry ()
DECLARE SUB GetMiscCodeRecord (MCFile, RecNo, MiscCodeValue$)
DECLARE SUB HideCursor ()
DECLARE SUB InvalidCodeNumber ()
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE SUB MPaintBox (ULRow%, ULCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB MiscPayEntry (OperRecNumber, PostDate$)
DECLARE SUB OpenAcctIdx (AcctIdxFileNum, NumIdxRecs)
DECLARE SUB OpenMiscCodeFile (NumOfMiscRecs, MCFile)
DECLARE SUB OpenSetupFile (SetUpFileNum%)
DECLARE SUB OperatorEntry (Results, OPERNUM, PostDate$)
DECLARE SUB PostAndPrint (Posted)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintJournal (OperNumb)
DECLARE SUB PrintReceipt (RECPTPort)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPORT%, RetCode%, EntryPoint%)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB QPrintRC (T$, r%, c%, clr%)
DECLARE SUB RestScrn (ScrnArray())
DECLARE SUB SaveScrn (ScrnArray())
DECLARE SUB ShowCursor ()
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE SUB ValidateOperator (OPERNUM, OperPassword$, Results)
DECLARE SUB WaitForAction ()
DECLARE FUNCTION QPValI% (Number$)
DECLARE FUNCTION GetDEFPort% ()
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FUsing$ (Number$, Mask$)
DECLARE FUNCTION FindAcct% (AcctNum$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION QPTrim$ (X$)
DECLARE FUNCTION Round# (ThisAmount#)
DECLARE FUNCTION WEnvTest ()
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice, MaxLen%, BoxBot, Ky$, Action, Cnf AS ANY)

  TYPE FLen2
    V AS STRING * 64
  END TYPE

  TYPE Struct
    who AS STRING * 14
    RecNum AS INTEGER
  END TYPE
  
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'Qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  
  '*** Citipak Include Files Needed ***
  '$INCLUDE: 'CMFILES.BI'
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'ubsetup.BI'
  '$INCLUDE: 'gl.bi'
  '$INCLUDE: 'lcTAX.BI'
  
  CONST False = 0, True = NOT False
  
  DIM SHARED RECPTPort AS INTEGER
  DIM SHARED GLSETUP AS GLSetupRecType
  DIM SHARED Acct AS GLAcctRecType
  DIM SHARED AcctIdx AS GLAcctIndexType

  
  REDIM Scrn1(0) AS INTEGER
  
  LOCATE 1, 1, 0
  BlockClear
  ShowCursor
  
  DO
     OperatorEntry OperOKFlag, OPERNUM, PostDate$
  LOOP WHILE OperOKFlag = 0
  
 '  OperOKFlag = 1
 ' OPERNUM = 1

  'PostDate$ = DATE$
  '  TestFlag = False
  '
  '  IF WEnvTest% THEN
  '    FileExt$ = ".BAS"
  '  ELSE
  '    FileExt$ = ".EXE"
  '  END IF
  '
  
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
  
  help$ = "Cash Management System V8.0"
  
  GOSUB InitMulti
  
  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%
  QPrintRC OperName$, 3, 55, 15
  
  SaveScrn Scrn1()
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F5KEY
      RUN "CMSENDGL"
    CASE F7KEY
      ERASE CMOperRec, Frm, Form$, Fld
      PrintJournal (OperRecNumber)
      Action = 1
      WhatSource = 0
      GOTO Top
    CASE ESC
      Done = True
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
        AddEditPayment OPERNUM, PostDate$
      CASE 2
        AddEditDeposit OPERNUM, PostDate$
      CASE 3
        MiscPayEntry OperRecNumber, PostDate$
      CASE 4
        EnterTaxPayment OperRecNumber, PostDate$
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
  Choice$(2, 0) = "Utility Deposit Entry  "
  Choice$(3, 0) = "Misc Payment           "
  Choice$(4, 0) = "Tax Payment"
  Choice$(5, 0) = "Business License Payment"
  RETURN

SUB AddMiscCode
  SHARED MCHoice$
  
  
  Size = 500
  Start = 1     'start at array element 1
  Dir = 0       'sort direction - use anything else for descending
  SSize = 16    'total size of each TYPE element
  MOff = 0      'offset into the TYPE for the key element
  MSize = 7     'size of the key element - coded as follows:
  '   -1 = integer
  '   -2 = long integer
  '   -3 = single precision
  '   -4 = double precision
  '   +N = TYPE array/fixed-length string of length N
  
  REDIM array(1 TO Size)  AS Struct
  
  LibName$ = "CM"
  ScrnName$ = "CMNWCODE"
  help$ = "Cash Management Operator Entry"
  LOCATE 1, 1, 0
  
  help$ = "Add Misc Codes V8.0"
  ShowCursor
  LibFile2Scrn "CM.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  PrintHelp help$
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  Action = 1
  NoProcess = False
  
  
  REDIM MiscCodeRec(1) AS MiscCodeRecType
  MiscCodeRecLen = LEN(MiscCodeRec(1))
  MCFile = FREEFILE
  OPEN "CMMISCCD.DAT" FOR RANDOM AS MCFile LEN = MiscCodeRecLen
  NumOfMiscRecs = LOF(MCFile) \ MiscCodeRecLen
  
  IF NumOfMiscRecs > 0 THEN
    
    GOSUB SortMiscCodes
    
    REDIM MCHoice$(1 TO NumOfMiscRecs)
    FOR Cnt = 1 TO NumOfMiscRecs
      GET MCFile, array(Cnt).RecNum, MiscCodeRec(1)
      MCHoice$(Cnt) = SPACE$(50)
      LSET MCHoice$(Cnt) = MiscCodeRec(1).MiscCode
      MID$(MCHoice$(Cnt), 9) = MiscCodeRec(1).Description
    NEXT Cnt
    
    MaxLen = 50 'Set menu width to zero
    BoxBot = 17 'limit the box length to go no lower than line 20
    Action = 0  '0 means stay in the menu until they select something
    Choice = 1  'Pre-load choice to highlight
    
    TText$ = SPACE$(MaxLen + 4)
    LSET TText$ = "  Code    Description"
    
    '--Center Menu within Screen
    Row = 8
    Col = 15
    
    DO
      
      '--Set upper left corner of menu, turn off the cursor
      LOCATE Row, Col, 0
      LibFile2Scrn "CM.QSL", "MENUBAK", MonoCode, -1, ErrorCode
      ShowCursor
      QPrintRC TText$, Row - 1, Col, 112
      VertMenu MCHoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        RecNo = 0
        ExitFlag = True
      ELSE
        RecNo = array(Choice).RecNum
        ExitFlag = True
      END IF
      
    LOOP UNTIL ExitFlag
  END IF
  
  
  IF RecNo > 0 THEN
    GET MCFile, RecNo, MiscCodeRec(1)
    Form$(1, 0) = MiscCodeRec(1).MiscCode
    Form$(2, 0) = MiscCodeRec(1).Description
    Form$(3, 0) = MiscCodeRec(1).GlAcctNumb
  END IF
  
  ShowCursor
  LibFile2Scrn "CM.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  PrintHelp help$
  
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      
      IF RecNo <= 0 THEN
        ' check for duplicate account number
        GOSUB CheckCodes
        IF validcode = False THEN
          LOCATE 14, 18: PRINT "DUPLICATE CODE ENTERED : PLEASE ENTER NEW CODE";
          PRINT CHR$(7);
          SLEEP 1
          Frm(1).FldNo = 1
          Action = 1
          LOCATE 14, 18: PRINT "                                               "
          Done = False
          NoProcess = True
        END IF
        
        ' check for valid glacctnumber
        GOSUB CheckGLAcctNumber
        
        IF AcctRecNo = False THEN
          LOCATE 14, 18: PRINT "INVALID ACCOUNT NUMBER : PLEASE CHECK YOUR ACCT";
          PRINT CHR$(7);
          WaitForAction
          Frm(1).FldNo = 3
          Action = 1
          LOCATE 14, 18: PRINT "                                               "
          Done = False
          NoProcess = True
        END IF
      END IF
      IF NoProcess = False THEN
        MiscCodeRec(1).MiscCode = Form$(1, 0)
        MiscCodeRec(1).Description = Form$(2, 0)
        MiscCodeRec(1).GlAcctNumb = Form$(3, 0)
        
        IF RecNo > 0 THEN
          PUT MCFile, RecNo, MiscCodeRec(1)
        ELSE
          PUT MCFile, LOF(MCFile) \ MiscCodeRecLen + 1, MiscCodeRec(1)
        END IF
        CLOSE MCFile
        Done = True
      END IF
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  
  EXIT SUB
  
  
CheckGLAcctNumber:
  OpenSetupFile SetUpFileNum
  IF LOF(SetUpFileNum) = 0 THEN AcctRecNo = True: CLOSE SetUpFileNum: RETURN
  GET SetUpFileNum, 1, GLSETUP
  Fund = GLSETUP.FundLen
  ACCNT = GLSETUP.AcctLen
  DET = GLSETUP.DetLen
  GLNumber$ = LEFT$(Form$(3, 0), Fund) + "-" + MID$(Form$(3, 0), Fund + 1, ACCNT) + "-" + MID$(Form$(3, 0), Fund + ACCNT + 1, DET)
  
  AcctRecNo = FindAcct(GLNumber$)
  
  CLOSE SetUpFileNum
  RETURN
  
CheckCodes:
  FOR Cnt = 1 TO NumOfMiscRecs
    GET MCFile, Cnt, MiscCodeRec(1)
    IF Form$(1, 0) = MiscCodeRec(1).MiscCode THEN validcode = False: RETURN
  NEXT Cnt
  validcode = True
  RETURN
  
SortMiscCodes:
  FOR Cnt = 1 TO NumOfMiscRecs
    GET MCFile, Cnt, MiscCodeRec(1)
    array(Cnt).who = MiscCodeRec(1).MiscCode + STRING$(7, " ")
    array(Cnt).RecNum = Cnt
  NEXT Cnt
  SortT array(Start), NumOfMiscRecs, Dir, SSize, MOff, MSize
  RETURN
  
END SUB

SUB EnterTaxPayment (OperRecNum%, PostDate$)

  SHARED Choice$()
  
  REDIM TempScrn(0)
  
  'Tax Files Here
  REDIM TBLog(1)  AS TBLog
  REDIM TBTrans(1) AS TBTrans
  REDIM TBRECust(1) AS TBRECust
  REDIM TBRECustBal(1) AS TBREBal
  REDIM TBPPCust(1) AS TBPPCust
  REDIM TBPPCustBal(1) AS TBPPBal
  
  
taxmainbody:
  CLOSE
  REDIM CMOperRec(1) AS CMOperRecType
  REDIM UBSetUpRec(1) AS UBSetupRecType
  
  CustomerGrabed = 0
  ARIdxFile = 0
  ARFile = 0
  
  LibName$ = "CM"
  ScrnName$ = "TBPAY"
  help$ = "Enter Payments"
  LOCATE 1, 1, 0
  
  GOSUB SetTaxOperatorName
  
  ShowCursor
  LibFile2Scrn "CM.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  PrintHelp help$
  
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  'Rem Set Choice
  REDIM Choice$(3, 0)
  Choice$(0, 0) = "8"
  Choice$(1, 0) = "Cash"
  Choice$(2, 0) = "Check"
  Choice$(3, 0) = "Cash & Check"
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  Form$(1, 0) = OperName$
  Form$(2, 0) = PostDate$
  Fld(1).Protected = True
  Fld(2).Protected = True
  Fld(5).Protected = True
  
  Fld(11).Protected = True
  Fld(12).Protected = True
  
  Frm(1).StayOnField = -1
  Action = 1
  
  'Set Field Defaults
  CustAcctFld = 4
  OwedAmtFld = 7
  
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    'This Gets the Customer
    IF Frm(1).FldNo > CustAcctFld AND Frm(1).PrevFld = CustAcctFld THEN
      LSET Form$(7, 0) = "0.00"
      LSET Form$(8, 0) = ""
      LSET Form$(9, 0) = "0.00"
      LSET Form$(10, 0) = "0.00"
      LSET Form$(11, 0) = "0.00"
      LSET Form$(12, 0) = "0.00"

      GOSUB GetTaxCustomer
      IF AccountRecord = 0 THEN
        CLOSE : Action = 1: GOTO taxmainbody
      END IF
      Action = 1
    END IF
    
    'This Makes Sure that the Amount Owed Field is Not Equal than Zero
    IF Frm(1).FldNo > OwedAmtFld AND Frm(1).PrevFld = OwedAmtFld THEN
      IF Value(Form$(7, 0), ecode) = 0 THEN Frm(1).FldNo = 7
      Action = 1
    END IF
    
    
    'Now Get Money Tendered
    
    IF LEFT$(Form$(8, 0), 6) = "Cash  " THEN
      Fld(10).Protected = True
      LSET Form$(10, 0) = "0.00"
      Fld(9).Protected = False
      IF PolledIt = 0 THEN
        Action = 1
        PolledIt = 1
        PayHow$ = LEFT$(Form$(8, 0), 2)
      END IF
    END IF
    
    IF LEFT$(Form$(8, 0), 2) = "Ch" THEN
      Fld(9).Protected = True
      LSET Form$(9, 0) = "0.00"
      Fld(10).Protected = False
      IF PolledIt = 0 THEN
        Action = 1
        PolledIt = 1
        PayHow$ = LEFT$(Form$(8, 0), 2)
      END IF
    END IF
    
    IF LEFT$(Form$(8, 0), 6) = "Cash &" THEN
      Fld(9).Protected = False
      Fld(10).Protected = False
      IF PolledIt = 0 THEN
        Action = 1: PolledIt = 1: PayHow$ = LEFT$(Form$(8, 0), 2)
      END IF
    END IF
    
    IF PolledIt = 1 AND LEFT$(Form$(8, 0), 2) <> PayHow$ THEN
      PolledIt = 0
      Action = 1
    END IF
    
    IF Frm(1).PrevFld = 8 AND Frm(1).FldEdited = -1 THEN
      LSET Form$(9, 0) = "0.00"
      LSET Form$(10, 0) = "0.00"
      LSET Form$(11, 0) = "0.00"
      LSET Form$(12, 0) = "0.00"
      Action = 1
    END IF
    
    'Calculate and Display the Change Due
    IF (Frm(1).FldNo <> 7) AND Frm(1).PrevFld = 7 THEN
      TotalRec# = Value(Form$(9, 0), ecode) + Value(Form$(10, 0), ecode)
      LSET Form$(11, 0) = STR$(TotalRec#)
      ChangeDue# = TotalRec# - Value(Form$(7, 0), ecode)
      IF ChangeDue# < 0 THEN ChangeDue# = 0
      LSET Form$(12, 0) = STR$(ChangeDue#)
      Action = 1
      CalcCode = 1
    END IF
    IF (Frm(1).FldNo <> 10) AND Frm(1).PrevFld = 10 THEN
      TotalRec# = Value(Form$(9, 0), ecode) + Value(Form$(10, 0), ecode)
      LSET Form$(11, 0) = STR$(TotalRec#)
      ChangeDue# = TotalRec# - Value(Form$(7, 0), ecode)
      IF ChangeDue# < 0 THEN ChangeDue# = 0
      LSET Form$(12, 0) = STR$(ChangeDue#)
      Action = 1
      CalcCode = 1
    END IF
    
    IF Value(Form$(7, 0), ecode) <> AmountOwed# THEN
      CalcCode = 0
      AmountOwed# = Value(Form$(7, 0), ecode)
    END IF
    
    IF Value(Form$(9, 0), ecode) <> CashIn# THEN
      CalcCode = 0
      CashIn# = Value(Form$(9, 0), ecode)
    END IF
    
    IF Value(Form$(10, 0), ecode) <> ChkIn# THEN
      CalcCode = 0
      ChkIn# = Value(Form$(10, 0), ecode)
    END IF
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      IF Value(Form$(7, 0), ecode) <> 0 THEN
        CustAcct$ = Form$(4, 0)
        CustName$ = Form$(5, 0)
        Desc$ = Form$(6, 0)
        TotalOwed# = Value(Form$(7, 0), ecode%)
        TotalCash# = Value(Form$(9, 0), ecode%)
        TotalCheck# = Value(Form$(10, 0), ecode%)
        TotalPaid# = Value(Form$(11, 0), ecode%)
        TotalChange# = Value(Form$(12, 0), ecode%)
        GOSUB PrintTaxReceipt
        GOSUB PostTaxTransaction
      END IF
      CLOSE
      Done = True
      GOTO taxmainbody
    CASE EscKey
      EXIT SUB
    END SELECT
  LOOP
  
  
  '*****************
GetTaxCustomer:
  CustomerGrabed = 0
  AccountRecord = VAL(Form$(4, 0))
  TaxType$ = Form$(3, 0)
  CLOSE
  GOSUB taxfile
  IF NOT (TaxType$ = "R" OR TaxType$ = "P") THEN RETURN

  IF TaxType$ = "R" THEN
    NumOfRecs = LOF(TBFile) / LEN(TBRECust(1))
    NumOfIdxRecs = LOF(TBIdxFile) / 24
  ELSE
    NumOfRecs = LOF(TBFile) / LEN(TBPPCust(1))
    NumOfIdxRecs = LOF(TBIdxFile) / 24
  END IF
  
  IF AccountRecord = 0 THEN
    
    MaxLen = 50 'Set menu width to zero
    BoxBot = 17 'limit the box length to go no lower than line 20
    Action = 0  '0 means stay in the menu until they select something
    Choice = 1  'Pre-load choice to highlight
    
    TText$ = SPACE$(MaxLen + 4)
    LSET TText$ = "Customer Name             Acct #"
    
    '--Center Menu within Screen
    Row = 8
    Col = 15
    
    REDIM MCHoice(1 TO NumOfIdxRecs) AS FLen2
    
    ChoiceCounter = 0
    sAcct$ = SPACE$(6)
    FOR Cnt = 1 TO NumOfIdxRecs
      GET TBIdxFile, Cnt
      ChoiceCounter = ChoiceCounter + 1
      'Mchoice$(ChoiceCounter) = SPACE$(16)
      LSET MCHoice(ChoiceCounter).V = LEFT$(LastNam$, 20)
      RSET sAcct$ = LTRIM$(STR$(CVS(Mrec$)))
      MID$(MCHoice(ChoiceCounter).V, 30) = sAcct$
      MID$(MCHoice(ChoiceCounter).V, 61) = LTRIM$(STR$(CVS(Mrec$)))
    NEXT Cnt
    
    DO
'here
      '--Set upper left corner of menu, turn off the cursor
      LOCATE Row, Col, 0
      LibFile2Scrn "CM.QSL", "MENUBAK", MonoCode, -1, ErrorCode
      ShowCursor
      QPrintRC TText$, Row - 1, Col, 112
      'VertMenu Mchoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      VertMenuT2 MCHoice(), Choice, MaxLen%, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        AccountRecord = 0
        ExitFlag = True
      ELSE
        'fix me
        AccountRecord = VAL(RIGHT$(MCHoice(Choice).V, 5))
        ERASE MCHoice

        LSET Form$(4, 0) = LTRIM$(STR$(AccountRecord))
        ExitFlag = True
      END IF
      
    LOOP UNTIL ExitFlag
    
    LibName$ = "CM"
    ScrnName$ = "TBPAY"
    LOCATE 1, 1, 0
    
    ShowCursor
    LibFile2Scrn "CM.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
    
  END IF
  
  '************************************************************************
  IF AccountRecord > 0 AND AccountRecord <= NumOfRecs THEN
    IF TaxType$ = "R" THEN
      GET TBFile, AccountRecord, TBRECust(1)
      LSET Form$(5, 0) = TBRECust(1).CustName
    ELSE
      GET TBFile, AccountRecord, TBPPCust(1)
      LSET Form$(5, 0) = TBPPCust(1).CustName
    END IF

    GOSUB FigureBalance

    IF CurBal# = 0 AND PastDue# = 0 THEN

      FOR zzz = 4 TO NumFlds
        LSET Form$(zzz, 0) = ""
        Frm(1).FldNo = 4
        CustomerGrabed = 0
        Action = 1
        'COLOR 15
      NEXT

    ELSE
      LSET Form$(7, 0) = LTRIM$(STR$((CurBal# + PastDue#)))
      Frm(1).FldNo = 6
      CustomerGrabed = 1
      Action = 1
      COLOR 15
    END IF
  END IF
  RETURN
  
SetTaxOperatorName:
  Action = 1
  OperRecNumber = OperRecNum%
  CMOperRecLen = LEN(CMOperRec(1))
  CMFile = FREEFILE
  OPEN "CMOPER.DAT" FOR RANDOM AS CMFile LEN = CMOperRecLen
  GET CMFile, OperRecNumber, CMOperRec(1)
  OperName$ = LEFT$(CMOperRec(1).OperatorName, 18)
  Operator = CMOperRec(1).OperatorNumber
  Operator$ = STR$(Operator)
  Operator$ = RIGHT$(Operator$, LEN(Operator$) - 1)
  CLOSE CMFile
RETURN

PrintTaxReceipt:
  help$ = "Citipak Cash Management System v8.0"
  ShowCursor
  LibName$ = "CM.QSL"
  ScrnName$ = "CMPOST"
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  LibFile2Scrn "CM.QSL", "CMPOST", MonoCode%, Attribute%, ErrorCode%
  PrintTitle User$
  PrintHelp help$
  
  Form$(1, 0) = "Y"
  Action = 1
TaxPrtReceipt:
  REDIM TmpScrn(0)
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      IF Form$(1, 0) = "Y" THEN
        SaveScrn TmpScrn()
        GOSUB PrintTaxReceiptFunction
        RestScrn TmpScrn()
        Form$(1, 0) = "N"
        Frm(1).FldNo = 1
        Action = 1
        GOTO TaxPrtReceipt
      END IF
      Posted = True
      Done = True
    CASE ESC
      Posted = False
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
RETURN
  
PrintTaxReceiptFunction:
  REDIM CMTRRec(1) AS CMTransRecType
  CMTrRecLen = LEN(CMTRRec(1))
  CHandle = FREEFILE
  OPEN "CMTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS CHandle LEN = CMTrRecLen
  RecNumber = (LOF(CHandle) \ CMTrRecLen) + 1
  CLOSE CHandle
  
  
  UBSetupLen = LEN(UBSetUpRec(1))
  FGetAH "UBSETUP.DAT", UBSetUpRec(1), UBSetupLen, 1            'load it
  TOWNNAME$ = UBSetUpRec(1).UTILNAME
  'Set Printer Port for Receipts Printer if none set to 1
  
  IF INSTR(UBSetUpRec(1).UTILNAME, "SPENCER") THEN
    SpencerFlag = True
  END IF
  Work$ = SPACE$(19)
  
  RMFile = FREEFILE
  REDIM RMRec(1) AS RMReceiptRecType
  RMRecLen = LEN(RMRec(1))
  
  OPEN "RECPT.PRN" FOR OUTPUT SHARED AS #15
  WIDTH #15, 255
  
  PRINT #15, ""
  PRINT #15, TOWNNAME$
  IF TaxType$ = "R" THEN
    PRINT #15, "      REAL ESTATE"
    PRINT #15, "      TAX PAYMENT"
  ELSE
    PRINT #15, "    PERSONAL PROPERTY"
    PRINT #15, "       TAX PAYMENT"
  END IF
  
  PRINT #15, "Date: "; PostDate$
  PRINT #15,
  PRINT #15, "Account #"; CustAcct$
  PRINT #15, CustName$
  PRINT #15, Address$
  PRINT #15, Desc$
  PRINT #15,
  PRINT #15, "Total Owed: "; USING "$$####,#.##"; TotalOwed#
  PRINT #15, ""
  PRINT #15, "  Cash Amt: "; USING "$$####,#.##"; TotalCash#
  PRINT #15, " Check Amt: "; USING "$$####,#.##"; TotalCheck#
  PRINT #15, "             -----------"
  PRINT #15, "Total Paid: "; USING "$$####,#.##"; TotalPaid#
  PRINT #15, ""
  PRINT #15, "    Change: "; USING "$$####,#.##"; TotalChange#
  PRINT #15,
  'put spencer trap here
  IF SpencerFlag THEN
    PRINT #15, CHR$(27); "="; CHR$(2)           'only to display
    LSET Work$ = "THANK YOU!"
    PRINT #15, Work$
    LSET Work$ = "Change:" + FUsing$(STR$(TotalChange#), "$$####.##")
    PRINT #15, Work$; CHR$(10)
    PRINT #15, CHR$(27); "="; CHR$(1)           'only to printer
  END IF
  
  PRINT #15,
  PRINT #15, "Operator: "; OperRecNum%
  PRINT #15, "Receipt#: "; USING "######"; RecNumber
  PRINT #15,
  PRINT #15, TAB(7); "T H A N K   Y O U !"
  PRINT #15,
  PRINT #15,
  PRINT #15,
  PRINT #15,
  PRINT #15,
  PRINT #15,
  PRINT #15, CHR$(27); "p"; CHR$(0); CHR$(25); CHR$(250)
  CLOSE
  PrintRptFile Header$, "RECPT.PRN", GetDEFPort%, RetCode%, 5
  RETURN
  
PostTaxTransaction:
  IF Posted <> True THEN RETURN
  
  GOSUB PostTAXDetail
  IF TaxType$ = "R" THEN
    GOSUB taxfile
    PUT TBFile, AccountRecord, TBRECust(1)
    PUT TBBalFile, AccountRecord, TBRECustBal(1)
    GOSUB PostRealHistory
  ELSE
    GOSUB taxfile
    PUT TBFile, AccountRecord, TBPPCust(1)
    PUT TBBalFile, AccountRecord, TBPPCustBal(1)
    GOSUB PostPersHistory
  END IF
  
  PostDate = Date2Num%(PostDate$)
  REDIM CMTRRec(1) AS CMTransRecType
  CMTrRecLen = LEN(CMTRRec(1))
  CHandle = FREEFILE
  OPEN "CMTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS CHandle LEN = CMTrRecLen
  CMTRRec(1).TransDate = PostDate
  CMTRRec(1).TransAmount = TotalPaid#
  CMTRRec(1).TransCash = TotalCash#
  CMTRRec(1).TransCheck = TotalCheck#
  CMTRRec(1).TransAmtOwed = TotalOwed#
  CMTRRec(1).TransDesc = Desc$
  CMTRRec(1).TransSource = 31
  CMTRRec(1).TransName = CustName$
  CMTRRec(1).TransAcctNum = AcctRecord
  CMTRRec(1).TransDetNum = DetailRecord!
  CMTRRec(1).TransOperNum = Operator
  CMTRRec(1).TransPad = ""
  FldFactor = 0
  
  CMTRRec(1).TransRevAmt(1) = (TotalPaid# - TotalChange#) - (Interest# + Penalty#)
  CMTRRec(1).TransRevAmt(2) = Interest#
  CMTRRec(1).TransRevAmt(3) = Penalty#
  
  FOR Cnt = 4 TO 15
    CMTRRec(1).TransRevAmt(Cnt) = 0
  NEXT Cnt
  
  PUT CHandle, (LOF(CHandle) / CMTrRecLen) + 1, CMTRRec(1)
  CLOSE CHandle
  
  RETURN
  
taxfile:
  IF TaxType$ = "R" THEN
    TBFile = FREEFILE
    OPEN TBPath$ + "TBRECust.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBRECust(1))
    TBIdxFile = FREEFILE
    OPEN TBPath$ + "TBRENAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIdxFile LEN = 24
    FIELD TBIdxFile, 20 AS LastNam$, 4 AS Mrec$
    
  ELSE
    TBFile = FREEFILE
    OPEN TBPath$ + "TBPPCust.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBPPCust(1))

    TBIdxFile = FREEFILE
    OPEN TBPath$ + "TBPPNAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIdxFile LEN = 24
    FIELD TBIdxFile, 20 AS LastNam$, 4 AS Mrec$
    
  END IF
  
  
  TBBalFile = FREEFILE
  IF TaxType$ = "R" THEN
    OPEN TBPath$ + "TBREBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBRECustBal(1))
  ELSE
    OPEN TBPath$ + "TBPPBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBPPCustBal(1))
  END IF
  
  
  TBTranFile = FREEFILE
  OPEN TBPath$ + "TBTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBTranFile LEN = LEN(TBTrans(1))
  RETURN
  
  
  
FigureBalance:
  CurBal# = 0: PastDue# = 0
  IF TaxType$ = "R" THEN
    CurBal# = TBRECust(1).CurYrTaxAmount + TBRECust(1).CurYrIntAmount + TBRECust(1).CurYrPenAmount + TBRECust(1).CurYrOthAmount
    CurBal# = INT((CurBal# * 100) + .5) / 100
    FOR ll = 1 TO 10
      PastDue# = PastDue# + TBRECustBal(1).TaxAmt(ll) + TBRECustBal(1).IntAmt(ll) + TBRECustBal(1).PenAmt(ll) + TBRECustBal(1).OthAmt(ll) + TBRECustBal(1).LateAmt(ll)
      PastDue# = INT((PastDue# * 100) + .5) / 100
    NEXT ll
  ELSE
    CurBal# = TBPPCust(1).CurYrPersPropTax + TBPPCust(1).CurYrFarmEquipTax + TBPPCust(1).CurYrMachToolsTax + TBPPCust(1).CurYrMerchCapTax + TBPPCust(1).CurYrMobileHomeTax + TBPPCust(1).CurYrLateFeeTax + TBPPCust(1).CurYrIntAmount + TBPPCust(1). _
CurYrPenAmount
    CurBal# = INT((CurBal# * 100) + .5) / 100
    FOR ll = 1 TO 10
      PastDue# = PastDue# + TBPPCustBal(1).PerPropTaxAmt(ll) + TBPPCustBal(1).FarmEquipAmt(ll) + TBPPCustBal(1).MachToolsAmt(ll) + TBPPCustBal(1).MerchCapAmt(ll) + TBPPCustBal(1).MobileHomeAmt(ll) + TBPPCustBal(1).LateFeeAmt(ll) + TBPPCustBal(1). _
IntAmt(ll) + TBPPCustBal(1).PenAmt(ll)
      PastDue# = INT((PastDue# * 100) + .5) / 100
    NEXT ll
  END IF

RETURN
  
PostRealHistory:
  GET TBFile, AccountRecord, TBRECust(1)
  GET TBBalFile, AccountRecord, TBRECustBal(1)
  GOSUB FigureBalance
  AcctBalance# = CurBal# + PastDue#
  TBTrans(1).TransDate = Date2Num%(PostDate$)
  TBTrans(1).TransAmt = TotalPaid# - TotalChange#
  TBTrans(1).TrTaxYear = ""
  TBTrans(1).TrDesc = Desc$
  TBTrans(1).TrType = 2
  TBTrans(1).TrTax = (TotalPaid# - TotalChange#) - (Interest# + Penalty#)
  TBTrans(1).TrInterest = Interest#
  TBTrans(1).TrPenalty = Penalty#
  TBTrans(1).TrLate = 0
  TBTrans(1).TrMachTools = 0
  TBTrans(1).TrMobileHome = 0
  TBTrans(1).TrMerchCap = 0
  TBTrans(1).TrFarmEquip = 0
  TBTrans(1).TrCustomerRecord = AcctRecord
  TBTrans(1).TrNextTrans = 0
  TBTrans(1).TrRunningBal = AcctBalance#
  TBTrans(1).TrSource = "R"
  TBTrans(1).TrPad = ""
  PUT TBTranFile, LOF(TBTranFile) / LEN(TBTrans(1)) + 1, TBTrans(1)
  TransRecord! = LOF(TBTranFile) / LEN(TBTrans(1))
  PrevTrans! = TBRECust(1).LastTrans
  TBRECust(1).LastTrans = TransRecord!
  PUT TBFile, AccountRecord, TBRECust(1)
  GET TBTranFile, PrevTrans!, TBTrans(1)
  TBTrans(1).TrNextTrans = TransRecord!
  PUT TBTranFile, PrevTrans!, TBTrans(1)
  CLOSE TBTranFile
  RETURN
  
PostPersHistory:
  GET TBFile, AccountRecord, TBPPCust(1)
  GET TBBalFile, AccountRecord, TBPPCustBal(1)
  GOSUB FigureBalance
  AcctBalance# = CurBal# + PastDue#
  TBTrans(1).TransDate = Date2Num%(PostDate$)
  TBTrans(1).TransAmt = TotalPaid# - TotalChange#
  TBTrans(1).TrTaxYear = ""
  TBTrans(1).TrDesc = Desc$
  TBTrans(1).TrType = 2
  TBTrans(1).TrTax = (TotalPaid# - TotalChange#) - (Interest# + Penalty#)
  TBTrans(1).TrInterest = Interest#
  TBTrans(1).TrPenalty = Penalty#
  TBTrans(1).TrLate = 0
  TBTrans(1).TrMachTools = 0
  TBTrans(1).TrMobileHome = 0
  TBTrans(1).TrMerchCap = 0
  TBTrans(1).TrFarmEquip = 0
  TBTrans(1).TrCustomerRecord = AcctRecord
  TBTrans(1).TrNextTrans = 0
  TBTrans(1).TrRunningBal = AcctBalance#
  TBTrans(1).TrSource = "P"
  TBTrans(1).TrPad = ""
  PUT TBTranFile, LOF(TBTranFile) / LEN(TBTrans(1)) + 1, TBTrans(1)
  TransRecord! = LOF(TBTranFile) / LEN(TBTrans(1))
  PrevTrans! = TBPPCust(1).LastTrans
  TBPPCust(1).LastTrans = TransRecord!
  PUT TBFile, AccountRecord, TBPPCust(1)
  GET TBTranFile, PrevTrans!, TBTrans(1)
  TBTrans(1).TrNextTrans = TransRecord!
  PUT TBTranFile, PrevTrans!, TBTrans(1)
  CLOSE TBTranFile
  RETURN
  
PostTAXDetail:
  IF TaxType$ = "R" THEN
    REM process real estate
    TotalCredit# = TotalPaid# - TotalChange#
    IF PastDue# > 0 THEN
      FOR Pyr = 10 TO 1 STEP -1
        IF TBRECustBal(1).OthAmt(Pyr) > TotalCredit# THEN
          TBRECustBal(1).OthAmt(Pyr) = TBRECustBal(1).OthAmt(Pyr) - TotalCredit#
          RETURN
        END IF
        IF TBRECustBal(1).OthAmt(Pyr) >= 0 THEN
          TotalCredit# = TotalCredit# - TBRECustBal(1).OthAmt(Pyr)
          TBRECustBal(1).OthAmt(Pyr) = 0
          TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
        END IF
        IF TotalCredit# = 0 THEN RETURN
        IF TBRECustBal(1).PenAmt(Pyr) > TotalCredit# THEN
          TBRECustBal(1).PenAmt(Pyr) = TBRECustBal(1).PenAmt(Pyr) - TotalCredit#
          Penalty# = Penalty# + TotalCredit#
          RETURN
        END IF
        IF TBRECustBal(1).PenAmt(Pyr) >= 0 THEN
          TotalCredit# = TotalCredit# - TBRECustBal(1).PenAmt(Pyr)
          Penalty# = Penalty# + TBRECustBal(1).PenAmt(Pyr)
          TBRECustBal(1).PenAmt(Pyr) = 0
          TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
        END IF
        IF TotalCredit# = 0 THEN RETURN
        IF TBRECustBal(1).IntAmt(Pyr) > TotalCredit# THEN
          TBRECustBal(1).IntAmt(Pyr) = TBRECustBal(1).IntAmt(Pyr) - TotalCredit#
          Interest# = Interest# + TotalCredit#
          RETURN
        END IF
        IF TBRECustBal(1).IntAmt(Pyr) >= 0 THEN
          TotalCredit# = TotalCredit# - TBRECustBal(1).IntAmt(Pyr)
          Interest# = Interest# + TBRECustBal(1).IntAmt(Pyr)
          TBRECustBal(1).IntAmt(Pyr) = 0
          TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
        END IF
        IF TotalCredit# = 0 THEN RETURN
        IF TBRECustBal(1).TaxAmt(Pyr) > TotalCredit# THEN
          TBRECustBal(1).TaxAmt(Pyr) = TBRECustBal(1).TaxAmt(Pyr) - TotalCredit#
          RETURN
        END IF
        IF TBRECustBal(1).TaxAmt(Pyr) >= 0 THEN
          TotalCredit# = TotalCredit# - TBRECustBal(1).TaxAmt(Pyr)
          TBRECustBal(1).TaxAmt(Pyr) = 0
          TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
        END IF
        IF TotalCredit# = 0 THEN RETURN
      NEXT Pyr
    END IF
    REM put current here
    IF TBRECust(1).CurYrOthAmount > TotalCredit# THEN
      TBRECust(1).CurYrOthAmount = TBRECust(1).CurYrOthAmount - TotalCredit#
      RETURN
    END IF
    IF TBRECust(1).CurYrOthAmount >= 0 THEN
      TotalCredit# = TotalCredit# - TBRECust(1).CurYrOthAmount
      TBRECust(1).CurYrOthAmount = 0
      TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    IF TBRECust(1).CurYrPenAmount >= TotalCredit# THEN
      TBRECust(1).CurYrPenAmount = TBRECust(1).CurYrPenAmount - TotalCredit#
      Penalty# = Penalty# + TotalCredit#
      RETURN
    END IF
    IF TBRECust(1).CurYrPenAmount >= 0 THEN
      TotalCredit# = TotalCredit# - TBRECust(1).CurYrPenAmount
      Penalty# = Penalty# + TBRECust(1).CurYrPenAmount
      TBRECust(1).CurYrPenAmount = 0
      TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    IF TBRECust(1).CurYrIntAmount > TotalCredit# THEN
      TBRECust(1).CurYrIntAmount = TBRECust(1).CurYrIntAmount - TotalCredit#
      Interest# = Interest# + TotalCredit#
      RETURN
    END IF
    IF TBRECust(1).CurYrIntAmount >= 0 THEN
      TotalCredit# = TotalCredit# - TBRECust(1).CurYrIntAmount
      Interest# = Interest# + TBRECust(1).CurYrIntAmount
      TBRECust(1).CurYrIntAmount = 0
      TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    TBRECust(1).CurYrTaxAmount = TBRECust(1).CurYrTaxAmount - TotalCredit#
    RETURN
  END IF
  
  IF TaxType$ = "P" THEN
    REM auto process here
    TotalCredit# = TotalPaid# - TotalChange#
    IF PastDue# > 0 THEN
      FOR Pyr = 10 TO 1 STEP -1
        IF TBPPCustBal(1).PenAmt(Pyr) > TotalCredit# THEN
          TBPPCustBal(1).PenAmt(Pyr) = TBPPCustBal(1).PenAmt(Pyr) - TotalCredit#
          Penalty# = Penalty# + TotalCredit#
          RETURN
        END IF
        IF TBPPCustBal(1).PenAmt(Pyr) >= 0 THEN
          TotalCredit# = TotalCredit# - TBPPCustBal(1).PenAmt(Pyr)
          Penalty# = Penalty# + TBPPCustBal(1).PenAmt(Pyr)
          TBPPCustBal(1).PenAmt(Pyr) = 0
          TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
        END IF
        IF TotalCredit# = 0 THEN RETURN
        IF TBPPCustBal(1).IntAmt(Pyr) > TotalCredit# THEN
          TBPPCustBal(1).IntAmt(Pyr) = TBPPCustBal(1).IntAmt(Pyr) - TotalCredit#
          Interest# = Interest# + TotalCredit#
          RETURN
        END IF
        IF TBPPCustBal(1).IntAmt(Pyr) >= 0 THEN
          TotalCredit# = TotalCredit# - TBPPCustBal(1).IntAmt(Pyr)
          Interest# = Interest# + TBPPCustBal(1).IntAmt(Pyr)
          TBPPCustBal(1).IntAmt(Pyr) = 0
          TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
        END IF
        IF TotalCredit# = 0 THEN RETURN
        IF TBPPCustBal(1).LateFeeAmt(Pyr) > TotalCredit# THEN
          TBPPCustBal(1).LateFeeAmt(Pyr) = TBPPCustBal(1).LateFeeAmt(Pyr) - TotalCredit#
          RETURN
        END IF
        IF TBPPCustBal(1).LateFeeAmt(Pyr) >= 0 THEN
          TotalCredit# = TotalCredit# - TBPPCustBal(1).LateFeeAmt(Pyr)
          TBPPCustBal(1).LateFeeAmt(Pyr) = 0
          TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
        END IF
        IF TotalCredit# = 0 THEN RETURN
        IF TBPPCustBal(1).MobileHomeAmt(Pyr) > TotalCredit# THEN
          TBPPCustBal(1).MobileHomeAmt(Pyr) = TBPPCustBal(1).MobileHomeAmt(Pyr) - TotalCredit#
          RETURN
        END IF
        IF TBPPCustBal(1).MobileHomeAmt(Pyr) >= 0 THEN
          TotalCredit# = TotalCredit# - TBPPCustBal(1).MobileHomeAmt(Pyr)
          TBPPCustBal(1).MobileHomeAmt(Pyr) = 0
          TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
        END IF
        IF TotalCredit# = 0 THEN RETURN
        IF TBPPCustBal(1).MerchCapAmt(Pyr) > TotalCredit# THEN
          TBPPCustBal(1).MerchCapAmt(Pyr) = TBPPCustBal(1).MerchCapAmt(Pyr) - TotalCredit#
          RETURN
        END IF
        IF TBPPCustBal(1).MerchCapAmt(Pyr) >= 0 THEN
          TotalCredit# = TotalCredit# - TBPPCustBal(1).MerchCapAmt(Pyr)
          TBPPCustBal(1).MerchCapAmt(Pyr) = 0
          TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
        END IF
        IF TotalCredit# = 0 THEN RETURN
        IF TBPPCustBal(1).MachToolsAmt(Pyr) > TotalCredit# THEN
          TBPPCustBal(1).MachToolsAmt(Pyr) = TBPPCustBal(1).MachToolsAmt(Pyr) - TotalCredit#
          RETURN
        END IF
        IF TBPPCustBal(1).MachToolsAmt(Pyr) >= 0 THEN
          TotalCredit# = TotalCredit# - TBPPCustBal(1).MachToolsAmt(Pyr)
          TBPPCustBal(1).MachToolsAmt(Pyr) = 0
          TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
        END IF
        IF TotalCredit# = 0 THEN RETURN
        IF TBPPCustBal(1).FarmEquipAmt(Pyr) > TotalCredit# THEN
          TBPPCustBal(1).FarmEquipAmt(Pyr) = TBPPCustBal(1).FarmEquipAmt(Pyr) - TotalCredit#
          RETURN
        END IF
        IF TBPPCustBal(1).FarmEquipAmt(Pyr) >= 0 THEN
          TotalCredit# = TotalCredit# - TBPPCustBal(1).FarmEquipAmt(Pyr)
          TBPPCustBal(1).FarmEquipAmt(Pyr) = 0
          TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
        END IF
        IF TotalCredit# = 0 THEN RETURN
        IF TBPPCustBal(1).PerPropTaxAmt(Pyr) > TotalCredit# THEN
          TBPPCustBal(1).PerPropTaxAmt(Pyr) = TBPPCustBal(1).PerPropTaxAmt(Pyr) - TotalCredit#
          RETURN
        END IF
        IF TBPPCustBal(1).PerPropTaxAmt(Pyr) >= 0 THEN
          TotalCredit# = TotalCredit# - TBPPCustBal(1).PerPropTaxAmt(Pyr)
          TBPPCustBal(1).PerPropTaxAmt(Pyr) = 0
          TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
        END IF
        IF TotalCredit# = 0 THEN RETURN
      NEXT Pyr
      
    END IF
    REM current here
    IF TBPPCust(1).CurYrPenAmount > TotalCredit# THEN
      TBPPCust(1).CurYrPenAmount = TBPPCust(1).CurYrPenAmount - TotalCredit#
      Penalty# = Penalty# + TotalCredit#
      RETURN
    END IF
    IF TBPPCust(1).CurYrPenAmount >= 0 THEN
      TotalCredit# = TotalCredit# - TBPPCust(1).CurYrPenAmount
      Penalty# = Penalty# + TBPPCust(1).CurYrPenAmount
      TBPPCust(1).CurYrPenAmount = 0
      TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    IF TBPPCust(1).CurYrIntAmount > TotalCredit# THEN
      TBPPCust(1).CurYrIntAmount = TBPPCust(1).CurYrIntAmount - TotalCredit#
      Interest# = Interest# + TotalCredit#
      RETURN
    END IF
    IF TBPPCust(1).CurYrIntAmount >= 0 THEN
      TotalCredit# = TotalCredit# - TBPPCust(1).CurYrIntAmount
      Interest# = Interest# + TBPPCust(1).CurYrIntAmount
      TBPPCust(1).CurYrIntAmount = 0
      TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    IF TBPPCust(1).CurYrLateFeeTax > TotalCredit# THEN
      TBPPCust(1).CurYrLateFeeTax = TBPPCust(1).CurYrLateFeeTax - TotalCredit#
      RETURN
    END IF
    IF TBPPCust(1).CurYrLateFeeTax >= 0 THEN
      TotalCredit# = TotalCredit# - TBPPCust(1).CurYrLateFeeTax
      TBPPCust(1).CurYrLateFeeTax = 0
      TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    IF TBPPCust(1).CurYrMobileHomeTax > TotalCredit# THEN
      TBPPCust(1).CurYrMobileHomeTax = TBPPCust(1).CurYrMobileHomeTax - TotalCredit#
      RETURN
    END IF
    IF TBPPCust(1).CurYrMobileHomeTax >= 0 THEN
      TotalCredit# = TotalCredit# - TBPPCust(1).CurYrMobileHomeTax
      TBPPCust(1).CurYrMobileHomeTax = 0
      TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    IF TBPPCust(1).CurYrMerchCapTax > TotalCredit# THEN
      TBPPCust(1).CurYrMerchCapTax = TBPPCust(1).CurYrMerchCapTax - TotalCredit#
      RETURN
    END IF
    IF TBPPCust(1).CurYrMerchCapTax >= 0 THEN
      TotalCredit# = TotalCredit# - TBPPCust(1).CurYrMerchCapTax
      TBPPCust(1).CurYrMerchCapTax = 0
      TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    IF TBPPCust(1).CurYrMachToolsTax > TotalCredit# THEN
      TBPPCust(1).CurYrMachToolsTax = TBPPCust(1).CurYrMachToolsTax - TotalCredit#
      RETURN
    END IF
    IF TBPPCust(1).CurYrMachToolsTax >= 0 THEN
      TotalCredit# = TotalCredit# - TBPPCust(1).CurYrMachToolsTax
      TBPPCust(1).CurYrMachToolsTax = 0
      TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    IF TBPPCust(1).CurYrFarmEquipTax > TotalCredit# THEN
      TBPPCust(1).CurYrFarmEquipTax = TBPPCust(1).CurYrFarmEquipTax - TotalCredit#
      RETURN
    END IF
    IF TBPPCust(1).CurYrFarmEquipTax >= 0 THEN
      TotalCredit# = TotalCredit# - TBPPCust(1).CurYrFarmEquipTax
      TBPPCust(1).CurYrFarmEquipTax = 0
      TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    TBPPCust(1).CurYrPersPropTax = TBPPCust(1).CurYrPersPropTax - TotalCredit#
    RETURN
  END IF
  RETURN
  
END SUB

  '
  '
  '*****************************************************************************
  'Searches the acct index for a matching account number and returns the record
  'number of the account
  '
  '    Input: AcctNum$ as a formatted G/L account number string
  '  Returns: Record number of the account
  '*****************************************************************************
  '
FUNCTION FindAcct (AcctNum$)
  '
  OpenAcctIdx AcctIdxFileNum, NumIdxRecs
  '
  IF NumIdxRecs = 0 THEN
    RecordNum = 0
    CLOSE AcctIdxFileNum
    EXIT FUNCTION
  END IF
  '
  Match = False
  FirstRec = 1
  LastRec = NumIdxRecs
  Lookfor$ = LTRIM$(RTRIM$(AcctNum$))
  '
  DO UNTIL LastRec < FirstRec
    '
    MiddleRec = (LastRec + FirstRec) \ 2
    '
    GET AcctIdxFileNum, MiddleRec, AcctIdx
    '
    'IF Lookfor$ = RTRIM$(AcctIdx.AcctNum) THEN
    Match = -1
    EXIT DO
    'ELSEIF Lookfor$ < AcctIdx.AcctNum THEN
    LastRec = MiddleRec - 1
    'ELSE
    FirstRec = MiddleRec + 1
    'END IF
    '
  LOOP
  '
  IF Match THEN
    RecNum = AcctIdx.RecNum
  ELSE
    RecNum = 0
  END IF
  '
  FindAcct = RecNum
  '
  CLOSE AcctIdxFileNum
  '
END FUNCTION

FUNCTION GetDEFPort%
  IF Exist("CMSETUP.DAT") THEN
    REDIM zz(1) AS STRING * 1
    CMFile = FREEFILE
    OPEN "CMSETUP.DAT" FOR RANDOM AS CMFile LEN = 1
    IF LOF(CMFile) > 0 THEN
      GET CMFile, 1, zz(1)
    ELSE
      zz(1) = "1"
    END IF
    CLOSE CMFile
  ELSE
    zz(1) = "1"
  END IF
  
  GetDEFPort% = QPValI(zz(1))
  
  ERASE zz
  
END FUNCTION

SUB GetMiscCodeRecord (MCFile, RecNo, MiscCodeValue$)
  
  OpenMiscCodeFile NumOfMiscRecs, MCFile
  REDIM MiscCodeRec(1) AS MiscCodeRecType
  
  Size = 250
  Start = 1     'start at array element 1
  Dir = 0       'sort direction - use anything else for descending
  SSize = 16    'total size of each TYPE element
  MOff = 0      'offset into the TYPE for the key element
  MSize = 7     'size of the key element - coded as follows:
  '   -1 = integer
  '   -2 = long integer
  '   -3 = single precision
  '   -4 = double precision
  '   +N = TYPE array/fixed-length string of length N
  
  REDIM array(1 TO Size)  AS Struct
  
  
  
  
  IF LEFT$(MiscCodeValue$, 1) = " " THEN
    Action = 1
    Row = 12
    Col = 15
    
    
    REDIM MCHoice$(1 TO NumOfMiscRecs)
    GOSUB SortMiscCode
    
    REDIM MCHoice$(1 TO NumOfMiscRecs)
    FOR Cnt = 1 TO NumOfMiscRecs
      GET MCFile, array(Cnt).RecNum, MiscCodeRec(1)
      MCHoice$(Cnt) = SPACE$(50)
      LSET MCHoice$(Cnt) = MiscCodeRec(1).MiscCode
      MID$(MCHoice$(Cnt), 9) = MiscCodeRec(1).Description
    NEXT Cnt
    
    
    MaxLen = 50 'Set menu width to zero
    BoxBot = 17 'limit the box length to go no lower than line 20
    Action = 0  '0 means stay in the menu until they select something
    Choice = 1  'Pre-load choice to highlight
    
    '--Center Menu within Screen
    
    DO
      '--Set upper left corner of menu, turn off the cursor
      LOCATE Row, Col, 0
      'LibFile2Scrn "CM.QSL", "MENUBAK", MonoCode, -1, ErrorCode
      'ShowCursor
      QPrintRC TText$, Row - 1, Col, 112
      VertMenu MCHoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      
      IF Ky$ = CHR$(27) THEN
        RecNo = 0
        ExitFlag = True
      ELSE
        RecNo = array(Choice).RecNum
        GET MCFile, RecNo, MiscCodeRec(1)
        help$ = "Code Desc: " + MiscCodeRec(1).Description
        PrintHelp help$
        
        ExitFlag = True
      END IF
      CODE = True
    LOOP UNTIL ExitFlag
    
    
  ELSE
    FOR Cnt = 1 TO NumOfMiscRecs
      GET MCFile, Cnt, MiscCodeRec(1)
      IF MiscCodeValue$ = MiscCodeRec(1).MiscCode THEN
        RecNo = Cnt
        help$ = "Code Desc: " + MiscCodeRec(1).Description
        PrintHelp help$
        EXIT FOR
      END IF
      RecNo = 0
    NEXT Cnt
  END IF
  
  CLOSE MCFile
  EXIT SUB
  
  
SortMiscCode:
  FOR Cnt = 1 TO NumOfMiscRecs
    GET MCFile, Cnt, MiscCodeRec(1)
    array(Cnt).who = MiscCodeRec(1).MiscCode + STRING$(7, " ")
    array(Cnt).RecNum = Cnt
  NEXT Cnt
  SortT array(Start), NumOfMiscRecs, Dir, SSize, MOff, MSize
  RETURN
  
END SUB

SUB InvalidCodeNumber
  LOCATE 19, 50
  COLOR 15
  PRINT "INVALID CODE : REDO"
  PRINT CHR$(7);
  SLEEP 1
  LOCATE 19, 50
  PRINT "                    "
  
END SUB

SUB MiscPayEntry (OperRecNumber, PostDate$)
  
  REDIM MiscRecNo(10)
  REDIM MiscCodeRec(1) AS MiscCodeRecType
  SHARED Choice$()
  
  LibName$ = "CM"
  ScrnName$ = "CMMISPAY"
  help$ = "Cash Management Operator Entry"
  LOCATE 1, 1, 0
  
  help$ = "Misc Payment Entry"
  ShowCursor
  LibFile2Scrn "CM.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  PrintHelp help$
  GOSUB SetOperatorName
  
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Rem Set Choice
  REDIM Choice$(3, 0)
  Choice$(0, 0) = "5"
  Choice$(1, 0) = "Cash"
  Choice$(2, 0) = "Check"
  Choice$(3, 0) = "Cash & Check"
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  ' Get operator name and display
  
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    
    'Force amount in field 4 cannot be 0 or less than 0
    IF Frm(1).FldNo > 4 AND Frm(1).PrevFld = 4 AND VAL(Form$(4, 0)) = 0 THEN
      Frm(1).FldNo = 4
    END IF
    
    IF Frm(1).FldNo >= 4 AND VAL(Form$(9, 0)) < 0 THEN
      IF ClearIt = 0 THEN
        LOCATE 18, 28: PRINT STRING$(15, 32)
        ClearIt = 1
      END IF
    END IF
    
    'rem set cash/check/both fields
    IF Frm(1).PrevFld = 5 AND Frm(1).FldEdited = -1 THEN
      Form$(6, 0) = "0.00": Form$(7, 0) = "0.00"
      Form$(8, 0) = "0.00": Form$(9, 0) = "0.00"
      Action = 1
    END IF
    IF LEFT$(Form$(5, 0), 6) = "Cash  " THEN
      Fld(7).Protected = True
      Form$(7, 0) = "0.00"
      Fld(6).Protected = False
      IF PolledIt = 0 THEN
        Action = 1: PolledIt = 1: PayHow$ = LEFT$(Form$(5, 0), 2)
      END IF
    END IF
    
    IF LEFT$(Form$(5, 0), 2) = "Ch" THEN
      Fld(6).Protected = True: Form$(6, 0) = "0.00"
      Fld(7).Protected = False
      IF PolledIt = 0 THEN
        Action = 1: PolledIt = 1: PayHow$ = LEFT$(Form$(5, 0), 2)
        '      CalcFields 0, 8, form$(), Fld()
      END IF
    END IF
    
    IF LEFT$(Form$(5, 0), 6) = "Cash &" THEN
      Fld(6).Protected = False
      Fld(7).Protected = False
      IF PolledIt = 0 THEN
        Action = 1: PolledIt = 1: PayHow$ = LEFT$(Form$(5, 0), 2)
        
      END IF
    END IF
    
    IF PolledIt = 1 AND LEFT$(Form$(5, 0), 2) <> PayHow$ THEN
      PolledIt = 0: Action = 1
      '   CalcFields 0, 8, form$(), Fld()
    END IF
    
    IF Frm(1).FldNo > 9 AND Frm(1).PrevFld <= 9 AND VAL(Form$(9, 0)) < 0 THEN
      Frm(1).FldNo = 5
    END IF
    
    ' ****************** Code Reconciliation Right Side of Screen ***********
    
    
    GOSUB PollMiscCodeEntry
    
    
    IF Frm(1).PrevFld = 21 AND Frm(1).FldNo = 20 THEN
      Frm(1).FldNo = 10
    END IF
    
    
    
    
    
    SELECT CASE Frm(1).KeyCode
    CASE F7KEY
      AddMiscCode
      help$ = "Misc Payment Entry"
      ShowCursor
      LibFile2Scrn "CM.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
      PrintHelp help$
      GOSUB SetOperatorName
      Action = 1
      Done = False
    CASE F10Key
      RecAmtOwed# = CVD(MID$(Form$(0, 0), Fld(4).Fields, 8))
      RecAmtRecd# = CVD(MID$(Form$(0, 0), Fld(8).Fields, 8))
      MiscAmtRecd# = CVD(MID$(Form$(0, 0), Fld(20).Fields, 8))
      IF RecAmtOwed# <= 0 AND RecAmtOwed# <= -.01 THEN
        Done = True
      ELSE
        
        IF RecAmtOwed# > RecAmtRecd# OR MiscAmtRecd# <> RecAmtOwed# THEN
          Frm(1).FldNo = 4
          Action = 1
        ELSE
          GOSUB StoreReceipt
          PostAndPrint Posted
        END IF
        IF Posted = True THEN
          GOSUB PostTransaction ' Normal Posting
          Done = True
        ELSE
          help$ = "Misc Payment Entry"
          ShowCursor
          LibFile2Scrn "CM.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
          PrintHelp help$
          GOSUB SetOperatorName
          Action = 1
          Done = False
        END IF
      END IF
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  EXIT SUB
  
SetOperatorName:
  Action = 1
  REDIM CMOperRec(1) AS CMOperRecType
  CMOperRecLen = LEN(CMOperRec(1))
  CMFile = FREEFILE
  OPEN "CMOPER.DAT" FOR RANDOM AS CMFile LEN = CMOperRecLen
  GET CMFile, OperRecNumber, CMOperRec(1)
  LOCATE 3, 55: COLOR 15: PRINT LEFT$(CMOperRec(1).OperatorName, 19)
  LOCATE 4, 44: COLOR 11: PRINT "Post Date: "; : COLOR 15: PRINT PostDate$
  Operator = CMOperRec(1).OperatorNumber
  Operator$ = STR$(Operator)
  Operator$ = RIGHT$(Operator$, LEN(Operator$) - 1)
  CLOSE CMFile
  RETURN
  
PollMiscCodeEntry:
  IF Frm(1).FldNo < 10 THEN RETURN
  
  IF (Frm(1).FldNo = 11 AND Frm(1).PrevFld = 10) OR (Frm(1).FldNo = 13 AND Frm(1).PrevFld = 12) OR (Frm(1).FldNo = 15 AND Frm(1).PrevFld = 14) OR (Frm(1).FldNo = 17 AND Frm(1).PrevFld = 16) OR (Frm(1).FldNo = 19 AND Frm(1).PrevFld = 18) THEN
    
    MiscCodeValue$ = Form$(Frm(1).PrevFld, 0)
    
    IF (Form$(4, 0) = Form$(20, 0)) AND (Frm(1).FldNo = 13 OR Frm(1).FldNo = 15 OR Frm(1).FldNo = 17 OR Frm(1).FldNo = 19) THEN Frm(1).FldNo = 21: Action = 1: RETURN
    
    GetMiscCodeRecord MCFile, RecNo, MiscCodeValue$
    
    IF RecNo = 0 THEN
      InvalidCodeNumber
      Action = 1
      Frm(1).FldNo = Frm(1).PrevFld
      RETURN
    END IF
    ' get record and continue
    OpenMiscCodeFile NumOfMiscRecs, MCFile
    GET MCFile, RecNo, MiscCodeRec(1)
    Form$(Frm(1).PrevFld, 0) = MiscCodeRec(1).MiscCode
    MiscRecNo((Frm(1).PrevFld) - 9) = RecNo
    Action = 1
    CLOSE MCFile
  END IF
  
  RETURN
  
StoreReceipt:
  REDIM CMTRRec(1) AS CMTransRecType
  CMTrRecLen = LEN(CMTRRec(1))
  CHandle = FREEFILE
  OPEN "CMTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS CHandle LEN = CMTrRecLen
  RecNumber = (LOF(CHandle) \ CMTrRecLen) + 1
  CLOSE CHandle
  
  RMFile = FREEFILE
  REDIM RMRec(1) AS RMReceiptRecType
  RMRecLen = LEN(RMRec(1))
  OPEN "CMRECPT.DAT" FOR RANDOM AS RMFile LEN = RMRecLen
  RMRec(1).RecName = Form$(1, 0)
  RMRec(1).RecAddress = Form$(2, 0)
  RMRec(1).RecDesc = Form$(3, 0)
  RMRec(1).RecAmtOwed = CVD(MID$(Form$(0, 0), Fld(4).Fields, 8))
  IF LEFT$(Form$(5, 0), 6) = "Cash  " THEN
    RMRec(1).RecPayType = 1
  END IF
  IF LEFT$(Form$(5, 0), 6) = "Check " THEN
    RMRec(1).RecPayType = 2
  END IF
  IF LEFT$(Form$(5, 0), 6) = "Cash &" THEN
    RMRec(1).RecPayType = 3
  END IF
  RMRec(1).RecCashAmt = CVD(MID$(Form$(0, 0), Fld(6).Fields, 8))
  RMRec(1).RecCheckAmt = CVD(MID$(Form$(0, 0), Fld(7).Fields, 8))
  RMRec(1).RecChangeDue = CVD(MID$(Form$(0, 0), Fld(9).Fields, 8))
  RMRec(1).RecDate = PostDate$
  RMRec(1).RecOperator = Operator$
  RMRec(1).RecptNumber = RecNumber
  PUT RMFile, 1, RMRec(1)
  CLOSE RMFile
  RETURN
  
PostTransaction:
  PostDate = Date2Num%(PostDate$)
  REDIM CMTRRec(1) AS CMTransRecType
  CMTrRecLen = LEN(CMTRRec(1))
  CHandle = FREEFILE
  OPEN "CMTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS CHandle LEN = CMTrRecLen
  CMTRRec(1).TransDate = PostDate
  CMTRRec(1).TransAmount = Value(Form$(8, 0), ecode)
  CMTRRec(1).TransCash = Value(Form$(6, 0), ecode)
  CMTRRec(1).TransCheck = Value(Form$(7, 0), ecode)
  CMTRRec(1).TransAmtOwed = Value(Form$(4, 0), ecode)
  CMTRRec(1).TransDesc = Form$(2, 0)
  CMTRRec(1).TransSource = 1
  CMTRRec(1).TransName = Form$(1, 0)
  CMTRRec(1).TransAcctNum = 99999
  CMTRRec(1).TransDetNum = DetailRecord!
  CMTRRec(1).TransOperNum = Operator
  CMTRRec(1).TransPad = ""
  
  FldFactor = 0
  FOR Cnt = 1 TO 5
    CMTRRec(1).TransRevAmt(Cnt) = Value(Form$(Cnt + 10 + FldFactor, 0), ecode)
    FldFactor = FldFactor + 1
  NEXT Cnt
  
  FldFactor = 0
  FOR Cnt = 1 TO 5              ' Store the Misc Code Record Number in Rev Amt 6-10
    CMTRRec(1).TransRevAmt(Cnt + 5) = MiscRecNo(Cnt + FldFactor)
    FldFactor = FldFactor + 1
  NEXT Cnt
  
  PUT CHandle, (LOF(CHandle) / CMTrRecLen) + 1, CMTRRec(1)
  CLOSE CHandle
  
  RETURN
  
  
  
END SUB

SUB OpenAcctIdx (AcctIdxFileNum, NumIdxRecs)
  
  AcctIdxLen = LEN(AcctIdx)
  AcctIdxFileNum = FREEFILE
  OPEN "GLACCT.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS AcctIdxFileNum LEN = AcctIdxLen
  NumIdxRecs = LOF(AcctIdxFileNum) \ AcctIdxLen
  
END SUB

SUB OpenMiscCodeFile (NumOfMiscRecs, MCFile)
  REDIM MiscCodeRec(1) AS MiscCodeRecType
  MiscCodeRecLen = LEN(MiscCodeRec(1))
  MCFile = FREEFILE
  OPEN "CMMISCCD.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS MCFile LEN = MiscCodeRecLen
  NumOfMiscRecs = LOF(MCFile) \ MiscCodeRecLen
END SUB

SUB OpenSetupFile (SetUpFileNum)
  GLSetupRecLen = LEN(GLSETUP)
  SetUpFileNum = FREEFILE
  OPEN "GLSETUP.DAT" FOR RANDOM AS SetUpFileNum LEN = GLSetupRecLen
END SUB

SUB PostAndPrint (Posted)
  
MainPrintPost:
  REDIM TmpScrn(0)
  help$ = "Citipak Cash Management System v8.0"
  ShowCursor
  LibName$ = "CM.QSL"
  ScrnName$ = "CMPOST"
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  LibFile2Scrn "CM.QSL", "CMPOST", MonoCode%, Attribute%, ErrorCode%
  PrintTitle User$
  PrintHelp help$
  
  Form$(1, 0) = "Y"
  
PrtReceipt:
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      IF Form$(1, 0) = "Y" THEN
        SaveScrn TmpScrn()
        PrintReceipt RECPTPort
        RestScrn TmpScrn()
        Form$(1, 0) = "N"
        Frm(1).FldNo = 1
        Action = 1
        GOTO PrtReceipt
      END IF
      Posted = True
      Done = True
    CASE ESC
      GOSUB AREYOUSURE
      Posted = False
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  EXIT SUB
  
AREYOUSURE:
  'Put ESC Confirmation Here
  
  RETURN
  
END SUB

