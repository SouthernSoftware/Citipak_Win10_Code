DECLARE SUB CMJourMenu (OperRecNumber%)
DEFINT A-Z    'CmMENU.BAS
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
DECLARE SUB PrintReceipt ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPORT%, RetCode%, EntryPoint%)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB QPrintRC (t$, r%, c%, clr%)
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
  
  TYPE Struct
    who AS STRING * 12
    RecNum AS LONG
  END TYPE

'$INCLUDE: 'DefCnf.BI'
'$INCLUDE: 'formedit.BI'
'$INCLUDE: 'fieldinf.BI'
'$INCLUDE: 'Qscr.BI'
'$INCLUDE: 'SetCnf.BI'

'*** Citipak Include Files Needed ***
'$INCLUDE: 'CMFILES.BI'
'$INCLUDE: 'ubtrans.BI'
'$INCLUDE: 'gl.bi'

  CONST False = 0, True = NOT False

  DIM SHARED RECPTPort AS INTEGER
  DIM SHARED GLSETUP AS GLSetupRecType
  DIM SHARED ACCT AS GLAcctRecType
  DIM SHARED AcctIdx AS GLAcctIndexType

  REDIM Scrn1(0) AS INTEGER

  LOCATE 1, 1, 0
  BlockClear
  ShowCursor

  DO
    OperatorEntry OperOKFlag, OPERNUM, PostDate$
  LOOP WHILE OperOKFlag = 0

'  OperOKFlag = 1
'  OPERNUM = 1


  'PostDate$
'  TestFlag = False
'
'  IF WEnvTest% THEN
'    FileExt$ = ".BAS"
'  ELSE
'    FileExt$ = ".EXE"
'  END IF
'

  IF OperOKFlag < 0 THEN GOTO ExitCMMenuProg
  
  OperRecNumber = OperOKFlag 'Set Operator Rec Num
  
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

  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%
  QPrintRC OperName$, 3, 55, 15

  SaveScrn Scrn1()
  
  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode
    CASE F5KEY
     'RUN "CMSENDGL"

    CASE F7KEY
      ERASE CMOperRec, Frm, Form$, Fld
      CMJourMenu (OperRecNumber)
      'PrintJournal (OperRecNumber)
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
          EnterPayment OPERNUM, PostDate$
        CASE 5
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

  RUN "CPRDMENU"

  HideCursor
  CLS
  END

InitMulti:
  REDIM Choice$(0 TO 5, 0 TO 0)
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Utility Billing Payment"
  Choice$(2, 0) = "Utility Deposit Entry  "
  Choice$(3, 0) = "Misc Payment           "
  Choice$(4, 0) = "Business License Payment"
  '  Choice$(4, 0) = "Tax Payment"
RETURN


  SUB AddMiscCode
  SHARED Mchoice$
  

  size = 500
  Start = 1               'start at array element 1
  Dir = 0                 'sort direction - use anything else for descending
  SSize = 16               'total size of each TYPE element
  MOff = 0                'offset into the TYPE for the key element
  MSize = 7              'size of the key element - coded as follows:
                        '   -1 = integer
                        '   -2 = long integer
                        '   -3 = single precision
                        '   -4 = double precision
                        '   +N = TYPE array/fixed-length string of length N

  REDIM array(1 TO size)  AS Struct

  LibName$ = "CM"
  ScrnName$ = "CMNWCODE"
  Help$ = "Cash Management Operator Entry"
  LOCATE 1, 1, 0

  Help$ = "Add Misc Codes V8.0"
  ShowCursor
  LibFile2Scrn "CM.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  PrintHelp Help$
  
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

  REDIM Mchoice$(1 TO NumOfMiscRecs)
  FOR Cnt = 1 TO NumOfMiscRecs
    GET MCFile, array(Cnt).RecNum, MiscCodeRec(1)
    Mchoice$(Cnt) = SPACE$(50)
    LSET Mchoice$(Cnt) = MiscCodeRec(1).MiscCode
    MID$(Mchoice$(Cnt), 9) = MiscCodeRec(1).Description
  NEXT Cnt
  
   MaxLen = 50     'Set menu width to zero
   BoxBot = 17    'limit the box length to go no lower than line 20
   Action = 0     '0 means stay in the menu until they select something
   Choice = 1     'Pre-load choice to highlight

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
      VertMenu Mchoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
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
  PrintHelp Help$


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

SUB DisplayCMScrn (ScrnName$)
  LibFile2Scrn "CM.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
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

  size = 250
  Start = 1               'start at array element 1
  Dir = 0                 'sort direction - use anything else for descending
  SSize = 16               'total size of each TYPE element
  MOff = 0                'offset into the TYPE for the key element
  MSize = 7              'size of the key element - coded as follows:
                        '   -1 = integer
                        '   -2 = long integer
                        '   -3 = single precision
                        '   -4 = double precision
                        '   +N = TYPE array/fixed-length string of length N

  REDIM array(1 TO size)  AS Struct




IF LEFT$(MiscCodeValue$, 1) = " " THEN
  Action = 1
  Row = 12
  Col = 15


  REDIM Mchoice$(1 TO NumOfMiscRecs)
  GOSUB SortMiscCode

  REDIM Mchoice$(1 TO NumOfMiscRecs)
  FOR Cnt = 1 TO NumOfMiscRecs
    GET MCFile, array(Cnt).RecNum, MiscCodeRec(1)
    Mchoice$(Cnt) = SPACE$(50)
    LSET Mchoice$(Cnt) = MiscCodeRec(1).MiscCode
    MID$(Mchoice$(Cnt), 9) = MiscCodeRec(1).Description
  NEXT Cnt


  MaxLen = 50     'Set menu width to zero
   BoxBot = 17    'limit the box length to go no lower than line 20
   Action = 0     '0 means stay in the menu until they select something
   Choice = 1     'Pre-load choice to highlight

   '--Center Menu within Screen

   DO
      '--Set upper left corner of menu, turn off the cursor
      LOCATE Row, Col, 0
      'LibFile2Scrn "CM.QSL", "MENUBAK", MonoCode, -1, ErrorCode
      'ShowCursor
      QPrintRC TText$, Row - 1, Col, 112
      VertMenu Mchoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf

      IF Ky$ = CHR$(27) THEN
        RecNo = 0
        ExitFlag = True
      ELSE
        RecNo = array(Choice).RecNum
    GET MCFile, RecNo, MiscCodeRec(1)
    Help$ = "Code Desc: " + MiscCodeRec(1).Description
    PrintHelp Help$

        ExitFlag = True
      END IF
      CODE = True
   LOOP UNTIL ExitFlag


ELSE
  FOR Cnt = 1 TO NumOfMiscRecs
  GET MCFile, Cnt, MiscCodeRec(1)
   IF MiscCodeValue$ = MiscCodeRec(1).MiscCode THEN
      RecNo = Cnt
      Help$ = "Code Desc: " + MiscCodeRec(1).Description
      PrintHelp Help$
      EXIT FOR
   END IF
  RecNo = 0
  NEXT Cnt
 END IF

CLOSE
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
  Help$ = "Cash Management Operator Entry"
  LOCATE 1, 1, 0
  
  Help$ = "Misc Payment Entry"
  ShowCursor
  LibFile2Scrn "CM.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  PrintHelp Help$
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
      Help$ = "Misc Payment Entry"
      ShowCursor
      LibFile2Scrn "CM.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
      PrintHelp Help$
      GOSUB SetOperatorName
      Action = 1
      Done = False
    CASE F10Key
     RecAmtOwed# = CVD(MID$(Form$(0, 0), Fld(4).Fields, 8))
     RecAmtRecd# = CVD(MID$(Form$(0, 0), Fld(8).Fields, 8))
     MiscAmtRecd# = CVD(MID$(Form$(0, 0), Fld(20).Fields, 8))
     IF RecAmtOwed# <= 0 AND RecAmtOwed# <= -.01 AND LEFT$(Form$(1, 0), 10) <> "CASH SHORT" THEN
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
        GOSUB PostTransaction         ' Normal Posting
       Done = True
       ELSE
      Help$ = "Misc Payment Entry"
      ShowCursor
      LibFile2Scrn "CM.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
      PrintHelp Help$
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
  RecNumber! = (LOF(CHandle) \ CMTrRecLen) + 1
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
    RMRec(1).RecptNumber = RecNumber!
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
  CMTRRec(1).TransDetNum = RecNumber!
  CMTRRec(1).TransOperNum = Operator
  CMTRRec(1).TransPad = ""

  FldFactor = 0
  FOR Cnt = 1 TO 5
    CMTRRec(1).TransRevAmt(Cnt) = Value(Form$(Cnt + 10 + FldFactor, 0), ecode)
    FldFactor = FldFactor + 1
  NEXT Cnt

 FldFactor = 0
 FOR Cnt = 1 TO 5               ' Store the Misc Code Record Number in Rev Amt 6-10
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
  Help$ = "Citipak Cash Management System v8.0"
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
  PrintHelp Help$

  Form$(1, 0) = "Y"

PrtReceipt:

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      IF Form$(1, 0) = "Y" THEN
        SaveScrn TmpScrn()
        PrintReceipt
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

