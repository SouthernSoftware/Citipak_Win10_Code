 'CmMENU.BAS
DEFINT A-Z
DECLARE SUB OldUtilPayEntry (OperRecNumber%, PostDate$)
DECLARE SUB OpenSetupFile (SetUpFileNum%)
DECLARE SUB OpenAcctIdx (AcctIdxFileNum, NumIdxRecs)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB PrintJournal (OPERNUMB)
DECLARE SUB PrintReceipt ()
DECLARE SUB PostAndPrint (Posted)
DECLARE SUB InvalidCodeNumber ()
DECLARE SUB OpenMiscCodeFile (NumOfMiscRecs, MCFile)
DECLARE SUB GetMiscCodeRecord (MCFile, RecNo, MiscCodeValue$)
DECLARE SUB AddMiscCode ()
DECLARE SUB MiscPayEntry (OperRecNumber, PostDate$)
DECLARE SUB UtilPayEntry (OperRecNumber, PostDate$)
DECLARE SUB SourceOfReceipt (OperRecNumber, source$)
DECLARE SUB AuthorizeOperator (oktoadd)
DECLARE SUB AddEditOperator ()
DECLARE SUB ValidateOperator (OperNum, OperPassword$, Results)
DECLARE SUB OperatorEntry (Results, OperNum, PostDate$)
DECLARE SUB DisplayOperEntry ()
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB BLPrint (LPTPort%, Text$, ErrCode%)
DECLARE SUB FileView (FileName$, Ky, action, FVI AS ANY, SEG array)
DECLARE SUB HideCursor ()
DECLARE SUB MScrnSave (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MScrnRest (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MPaintBox (ULRow%, ULCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE FUNCTION FindAcct% (AcctNum$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE FUNCTION Round# (ThisAmount#)
DECLARE FUNCTION MGetKey% (Row%, Col%)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION FUsing$ (Number$, Mask$)
DECLARE FUNCTION Num2Date$ (Dat%)
'DECLARE FUNCTION Peek1% (Segment%, Address%)
'DECLARE FUNCTION WEnvTest ()
DECLARE SUB QPrint (X$, Colr%, page%)
DECLARE SUB QPrintRC (T$, r%, c%, clr%)
DECLARE FUNCTION QPTrim$ (X$)
DECLARE SUB ShowCursor ()
' DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE FUNCTION WaitKey% ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PrintLoading ()
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
''$INCLUDE: 'gl.bi'

  CONST False = 0, True = NOT False

'DIM SHARED GLSETUP AS GLSetupRecType
'DIM SHARED ACCT AS GLAcctRecType
'DIM SHARED AcctIdx AS GLAcctIndexType

  HideCursor
  OperatorEntry Results, OperNum, PostDate$
  
  IF Results < 0 THEN RUN "cpmenu"
  IF Results = 0 THEN RUN
  
  ' Next Find Out Source of Receipt 1=Misc : 2=Util : 3=Util Dep : 4=Tax : 5=Lic
  ' Only 1 & 2 Are Valid Now

  source = 0    ' Set to 0
  
  OperRecNumber = Results       ' Set Operator Record Number to be pulled when needed
  
  DO

    SourceOfReceipt OperRecNumber, source$
    
    IF source$ = "INVALID" THEN RUN

    IF source$ = "M" THEN
      MiscPayEntry OperRecNumber, PostDate$
    END IF

   ' IF source$ = "U" THEN
   '   OldUtilPayEntry OperRecNumber, PostDate$
   ' END IF

  LOOP UNTIL source$ = "INVALID"

  END

SUB AddEditOperator
  
  AuthorizeOperator oktoadd
  
  IF oktoadd = 0 THEN EXIT SUB
  
  LibName$ = "CM"
  ScrnName$ = "CMNEWOP"
  help$ = "Cash Management Operator Entry"
  LOCATE 1, 1, 0
  
  help$ = "Add Operator v8.0"
  ShowCursor
  LibFile2Scrn "CM.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  PrintHelp help$
  
  
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  action = 1
  REDIM CMOperRec(1) AS CMOperRecType
  CMOperRecLen = LEN(CMOperRec(1))
  CMFile = FREEFILE
  OPEN "CMOPER.DAT" FOR RANDOM AS CMFile LEN = CMOperRecLen
  NumRecs = LOF(CMFile) / CMOperRecLen
  
  DO
    EditForm Form$(), Fld(), frm(1), Cnf, action
   IF frm(1).PrevFld = 2 AND edit <> True THEN
   FOR Cnt = 1 TO NumRecs
    GET CMFile, Cnt, CMOperRec(1)
    IF VAL(Form$(1, 0)) = CMOperRec(1).OperatorNumber AND VAL(Form$(1, 0)) > 0 THEN
      Form$(2, 0) = CMOperRec(1).operatorname
      Form$(3, 0) = CMOperRec(1).OperatorPassword
      Fld(1).Protected = True
      RecNo = Cnt
      edit = True
      action = 1
      EXIT FOR
    END IF
   NEXT Cnt
   END IF
   
   

    
    SELECT CASE frm(1).KeyCode
      
    CASE F10Key
 
      CMOperRec(1).OperatorNumber = VAL(Form$(1, 0))
      CMOperRec(1).operatorname = Form$(2, 0)
      CMOperRec(1).OperatorPassword = Form$(3, 0)
      IF RecNo = 0 THEN
       PUT CMFile, LOF(CMFile) / CMOperRecLen + 1, CMOperRec(1)
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
  help$ = "Cash Management Operator Entry"
  LOCATE 1, 1, 0

  help$ = "Add Misc Codes V8.0"
  ShowCursor
  LibFile2Scrn "CM.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  PrintHelp help$
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)

  ' Define Quick Screen Form Editing Arrays
  REDIM frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode

  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F

  action = 1
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
   action = 0     '0 means stay in the menu until they select something
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
      VertMenu Mchoice$(), Choice, MaxLen, BoxBot, Ky$, action, Cnf
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
    EditForm Form$(), Fld(), frm(1), Cnf, action



    SELECT CASE frm(1).KeyCode

    CASE F10Key
      
      IF RecNo <= 0 THEN
      ' check for duplicate account number
      GOSUB CheckCodes
      IF validcode = False THEN
       LOCATE 14, 18: PRINT "DUPLICATE CODE ENTERED : PLEASE ENTER NEW CODE";
       PRINT CHR$(7);
       SLEEP 1
       frm(1).FldNo = 1
       action = 1
       LOCATE 14, 18: PRINT "                                               "
       Done = False
       NoProcess = True
      END IF

      ' check for valid glacctnumber
      GOSUB CheckGLAcctNumber
      
      IF AcctRecNo = False THEN
       LOCATE 14, 18: PRINT "INVALID ACCOUNT NUMBER : PLEASE CHECK YOUR ACCT";
       PRINT CHR$(7);
       SLEEP 1
       frm(1).FldNo = 3
       action = 1
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
IF LOF(SetUpFileNum) = 0 THEN AcctRecNo = False: CLOSE SetUpFileNum: RETURN
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

SUB AuthorizeOperator (oktoadd)
  LibName$ = "CM"
  ScrnName$ = "OPERAUTH"
  help$ = "Cash Management Operator Entry"
  LOCATE 1, 1, 0
  
  help$ = "Authorization Required"
  ShowCursor
  LibFile2Scrn "CM.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  PrintHelp help$
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  DO
    EditForm Form$(), Fld(), frm(1), Cnf, action
    
    SELECT CASE frm(1).KeyCode
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
  LOCATE 1, 1, 0
  help$ = "Citipak Cash Management System v8.0"
  ShowCursor
  LibFile2Scrn "CM.QSL", "CMCHOSOP", MonoCode%, Attribute%, ErrorCode%
  PrintTitle User$
  PrintHelp help$
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
  action = 1
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
   action = 0     '0 means stay in the menu until they select something
   Choice = 1     'Pre-load choice to highlight

   '--Center Menu within Screen

   DO
      '--Set upper left corner of menu, turn off the cursor
      LOCATE Row, Col, 0
      'LibFile2Scrn "CM.QSL", "MENUBAK", MonoCode, -1, ErrorCode
      'ShowCursor
      QPrintRC TText$, Row - 1, Col, 112
      VertMenu Mchoice$(), Choice, MaxLen, BoxBot, Ky$, action, Cnf

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
  REDIM frm(1) AS FormInfo
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
    EditForm Form$(), Fld(), frm(1), Cnf, action


    'Force amount in field 4 cannot be 0 or less than 0
    IF frm(1).FldNo > 4 AND frm(1).PrevFld = 4 AND VAL(Form$(4, 0)) <= 0 THEN
     frm(1).FldNo = 4
    END IF

    IF frm(1).FldNo >= 4 AND VAL(Form$(9, 0)) < 0 THEN
     IF ClearIt = 0 THEN
      LOCATE 18, 28: PRINT STRING$(15, 32)
      ClearIt = 1
     END IF
    END IF

   'rem set cash/check/both fields
    IF frm(1).PrevFld = 5 AND frm(1).FldEdited = -1 THEN
      Form$(6, 0) = "0.00": Form$(7, 0) = "0.00"
      Form$(8, 0) = "0.00": Form$(9, 0) = "0.00"
      action = 1
    END IF
    IF LEFT$(Form$(5, 0), 6) = "Cash  " THEN
     Fld(7).Protected = True
     Form$(7, 0) = "0.00"
     Fld(6).Protected = False
     IF PolledIt = 0 THEN
      action = 1: PolledIt = 1: PayHow$ = LEFT$(Form$(5, 0), 2)
     END IF
    END IF

    IF LEFT$(Form$(5, 0), 2) = "Ch" THEN
     Fld(6).Protected = True: Form$(6, 0) = "0.00"
     Fld(7).Protected = False
     IF PolledIt = 0 THEN
      action = 1: PolledIt = 1: PayHow$ = LEFT$(Form$(5, 0), 2)
'      CalcFields 0, 8, form$(), Fld()
     END IF
    END IF

    IF LEFT$(Form$(5, 0), 6) = "Cash &" THEN
     Fld(6).Protected = False
     Fld(7).Protected = False
     IF PolledIt = 0 THEN
      action = 1: PolledIt = 1: PayHow$ = LEFT$(Form$(5, 0), 2)
 
     END IF
    END IF

    IF PolledIt = 1 AND LEFT$(Form$(5, 0), 2) <> PayHow$ THEN
     PolledIt = 0: action = 1
  '   CalcFields 0, 8, form$(), Fld()
    END IF

    IF frm(1).FldNo > 9 AND frm(1).PrevFld <= 9 AND VAL(Form$(9, 0)) < 0 THEN
     frm(1).FldNo = 5
    END IF

   ' ****************** Code Reconciliation Right Side of Screen ***********


   GOSUB PollMiscCodeEntry

 
   IF frm(1).PrevFld = 21 AND frm(1).FldNo = 20 THEN
    frm(1).FldNo = 10
   END IF

   



    SELECT CASE frm(1).KeyCode
    CASE F7KEY
      AddMiscCode
      help$ = "Misc Payment Entry"
      ShowCursor
      LibFile2Scrn "CM.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
      PrintHelp help$
      GOSUB SetOperatorName
      action = 1
      Done = False
    CASE F10Key
     RecAmtOwed# = CVD(MID$(Form$(0, 0), Fld(4).Fields, 8))
     RecAmtRecd# = CVD(MID$(Form$(0, 0), Fld(8).Fields, 8))
     MiscAmtRecd# = CVD(MID$(Form$(0, 0), Fld(20).Fields, 8))
     IF RecAmtOwed# <= 0 AND RecAmtOwed# <= -.01 THEN
      Done = True
      ELSE

     IF RecAmtOwed# > RecAmtRecd# OR MiscAmtRecd# <> RecAmtOwed# THEN
      frm(1).FldNo = 4
      action = 1
     ELSE
      GOSUB StoreReceipt
      PostAndPrint Posted
     END IF
      IF Posted = True THEN
        GOSUB PostTransaction         ' Normal Posting
        GOSUB PostOldTransaction     ' Hohenwald Tenn Only
        GOSUB PostCMIF                ' G/L I/F File
       Done = True
       ELSE
      help$ = "Misc Payment Entry"
      ShowCursor
      LibFile2Scrn "CM.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
      PrintHelp help$
      GOSUB SetOperatorName
      action = 1
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
  action = 1
  REDIM CMOperRec(1) AS CMOperRecType
  CMOperRecLen = LEN(CMOperRec(1))
  CMFile = FREEFILE
  OPEN "CMOPER.DAT" FOR RANDOM AS CMFile LEN = CMOperRecLen
  GET CMFile, OperRecNumber, CMOperRec(1)
  LOCATE 3, 55: COLOR 15: PRINT LEFT$(CMOperRec(1).operatorname, 19)
  LOCATE 4, 44: COLOR 11: PRINT "Post Date: "; : COLOR 15: PRINT PostDate$
  Operator = CMOperRec(1).OperatorNumber
  Operator$ = STR$(Operator)
  Operator$ = RIGHT$(Operator$, LEN(Operator$) - 1)
  CLOSE CMFile
  RETURN

PollMiscCodeEntry:
    IF frm(1).FldNo < 10 THEN RETURN

    IF (frm(1).FldNo = 11 AND frm(1).PrevFld = 10) OR (frm(1).FldNo = 13 AND frm(1).PrevFld = 12) OR (frm(1).FldNo = 15 AND frm(1).PrevFld = 14) OR (frm(1).FldNo = 17 AND frm(1).PrevFld = 16) OR (frm(1).FldNo = 19 AND frm(1).PrevFld = 18) THEN

    MiscCodeValue$ = Form$(frm(1).PrevFld, 0)
    
    IF (Form$(4, 0) = Form$(20, 0)) AND (frm(1).FldNo = 13 OR frm(1).FldNo = 15 OR frm(1).FldNo = 17 OR frm(1).FldNo = 19) THEN frm(1).FldNo = 21: action = 1: RETURN

    GetMiscCodeRecord MCFile, RecNo, MiscCodeValue$
    
     IF RecNo = 0 THEN
       InvalidCodeNumber
       action = 1
       frm(1).FldNo = frm(1).PrevFld
       RETURN
     END IF
     ' get record and continue
     OpenMiscCodeFile NumOfMiscRecs, MCFile
     GET MCFile, RecNo, MiscCodeRec(1)
     Form$(frm(1).PrevFld, 0) = MiscCodeRec(1).MiscCode
     MiscRecNo((frm(1).PrevFld) - 9) = RecNo
     action = 1
     CLOSE MCFile
     END IF

     RETURN

StoreReceipt:
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
  PUT RMFile, 1, RMRec(1)
  CLOSE RMFile
  RETURN

PostTransaction:
  PostDate = Date2Num(PostDate$)
  REDIM UBTRRec(1) AS UBTransRecType
  UBTRRecLen = LEN(UBTRRec(1))
  Thandle = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM LOCK READ WRITE AS Thandle LEN = UBTRRecLen
  UBTRRec(1).TransDate = PostDate
  UBTRRec(1).TransType = 99
  UBTRRec(1).TransDesc = Form$(1, 0)
  UBTRRec(1).TransAmt = CVD(MID$(Form$(0, 0), Fld(4).Fields, 8))
  UBTRRec(1).RunBalance = 0
  FldFactor = 0
 FOR Cnt = 1 TO 5
  UBTRRec(1).RevAmt(Cnt) = CVD(MID$(Form$(0, 0), Fld(Cnt + 10 + FldFactor).Fields, 8))
  FldFactor = FldFactor + 1
 NEXT Cnt
 FOR Cnt = 6 TO 15
  UBTRRec(1).RevAmt(Cnt) = 0
 NEXT Cnt
 FldFactor = 0
 FOR Cnt = 1 TO 5               ' Store the Misc Code Record Number in Reading
    UBTRRec(1).CurRead(Cnt) = MiscRecNo(Cnt + FldFactor)
    FldFactor = FldFactor + 1
 NEXT Cnt
 FOR Cnt = 6 TO 7
    UBTRRec(1).CurRead(Cnt) = 0
 NEXT Cnt
 FOR Cnt = 1 TO 7
  UBTRRec(1).PrevRead(Cnt) = 0
  'UBTRRec(1).Consumption(Cnt) = 0
  UBTRRec(1).EstRead(Cnt) = ""
 NEXT Cnt
'UBTRRec(1).CustLocation = ""
UBTRRec(1).CustStatus = ""
UBTRRec(1).CustCode = ""
UBTRRec(1).CustCode = ""
UBTRRec(1).billcycle = ""
UBTRRec(1).PayTypeCode = RMRec(1).RecPayType
UBTRRec(1).OperatorNumber = OperRecNumber
UBTRRec(1).CustAcctNo = 0
UBTRRec(1).PrevTrans = 0
UBTRRec(1).NextTrans = 0
UBTRRec(1).CashAmount = CVD(MID$(Form$(0, 0), Fld(6).Fields, 8))
UBTRRec(1).CheckAmount = CVD(MID$(Form$(0, 0), Fld(7).Fields, 8))
PUT Thandle, LOF(Thandle) \ UBTRRecLen + 1, UBTRRec(1)
CLOSE Thandle
RETURN

PostCMIF:

RETURN


PostOldTransaction:
   REM FOR HOHENWALD TENN. and SAWMILLS NC
   Thandle = FREEFILE
   OPEN "wbartran.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #Thandle LEN = 128
   FIELD Thandle, 8 AS tdate$, 8 AS tamt$, 2 AS ttype$, 15 AS tdesc$, 4 AS tmrec$, 4 AS tnexttr$, 4 AS tcuramt$
   FIELD Thandle, 45 AS tu1$, 4 AS wbrev1$, 4 AS wbrev2$, 4 AS wbrev3$, 4 AS wbrev4$, 4 AS wbrev5$, 4 AS wbrev6$, 4 AS wbrev7$, 4 AS wbrev8$
   FIELD Thandle, 77 AS tu2$, 1 AS testimate$, 8 AS tcurread$, 8 AS tpastread$
   FIELD Thandle, 94 AS TU3$, 8 AS tconsump$, 2 AS tbook$, 2 AS tcusttype$, 1 AS tcustloc$: REM tcustloc$ = 1 for cash and 2 for check in payment transaction
   FIELD Thandle, 107 AS TU4$, 1 AS tstatus$, 2 AS OPERNUMB$, 4 AS POSTTIME$, 1 AS tglpcode$, 13 AS wbblank$

  
    LSET tdate$ = LEFT$(PostDate$, 6) + RIGHT$(PostDate$, 2)
    LSET tamt$ = MKD$(CVD(MID$(Form$(0, 0), Fld(6).Fields, 8)))
    IF LEFT$(Form$(5, 0), 6) = "Cash  " THEN
     LSET ttype$ = MKI$(1)
    END IF
    IF LEFT$(Form$(5, 0), 6) = "Check " THEN
     LSET ttype$ = MKI$(2)
    END IF
    IF LEFT$(Form$(5, 0), 6) = "Cash &" THEN
     LSET ttype$ = MKI$(3)
    END IF
    LSET tdesc$ = Form$(3, 0)
    LSET tmrec$ = MKS$(999999)
    LSET tnexttr$ = MKS$(0)
    LSET tcuramt$ = MKS$(CVD(MID$(Form$(0, 0), Fld(4).Fields, 8)))
    LSET wbrev1$ = MKS$(MiscRecNo(1))
    LSET wbrev2$ = MKS$(MiscRecNo(3))
    LSET wbrev3$ = MKS$(MiscRecNo(5))
    LSET wbrev4$ = MKS$(MiscRecNo(7))
    LSET wbrev5$ = MKS$(MiscRecNo(9))
    MiscRev1# = CVD(MID$(Form$(0, 0), Fld(11).Fields, 8))
    MiscRev1# = Round(MiscRev1#)
    LSET wbrev6$ = MKS$(MiscRev1#)
    IF MiscRecNo(3) > 0 THEN
     MiscRev2# = CVD(MID$(Form$(0, 0), Fld(13).Fields, 8))
     MiscRev2# = Round(MiscRev2#)
     ELSE
     MiscRev2# = 0
    END IF
    LSET wbrev7$ = MKS$(MiscRev2#)
    IF MiscRecNo(5) > 0 THEN
     MiscRev3# = CVD(MID$(Form$(0, 0), Fld(15).Fields, 8))
     MiscRev3# = Round(MiscRev3#)
     ELSE
     MiscRev3# = 0
    END IF
    LSET wbrev8$ = MKS$(MiscRev3#)
    LSET testimate$ = "N"
    IF MiscRecNo(7) > 0 THEN
     MiscRev4# = CVD(MID$(Form$(0, 0), Fld(17).Fields, 8))
     MiscRev4# = Round(MiscRev4#)
     ELSE
     MiscRev4# = 0
    END IF
    LSET tcurread$ = MKD$(MiscRev4#)
    IF MiscRecNo(9) > 0 THEN
     MiscRev5# = CVD(MID$(Form$(0, 0), Fld(19).Fields, 8))
     MiscRev5# = Round(MiscRev5#)
     ELSE
     MiscRev5# = 0
    END IF
    LSET tpastread$ = MKD$(MiscRev5#)
    LSET tconsump$ = MKD$(CVD(MID$(Form$(0, 0), Fld(7).Fields, 8)))
    LSET tbook$ = MKI$(0)
    LSET tcusttype$ = MKI$(0)
    IF LEFT$(Form$(5, 0), 6) = "Cash  " THEN
     LSET tcustloc$ = "1"
    END IF
    IF LEFT$(Form$(5, 0), 6) = "Check " THEN
     LSET tcustloc$ = "2"
    END IF
    IF LEFT$(Form$(5, 0), 6) = "Cash &" THEN
     LSET tcustloc$ = "3"
    END IF
  LSET tstatus$ = ""
  LSET OPERNUMB$ = Operator$
  LSET POSTTIME$ = LEFT$(TIME$, 2) + MID$(TIME$, 4, 2)
  LSET tglpcode$ = "N"
  LSET wbblank$ = Form$(1, 0)
  PUT Thandle, LOF(Thandle) / 128 + 1
  CLOSE Thandle
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
  OPEN "CMMISCCD.DAT" FOR RANDOM AS MCFile LEN = MiscCodeRecLen
  NumOfMiscRecs = LOF(MCFile) \ MiscCodeRecLen
  END SUB

SUB OpenSetupFile (SetUpFileNum)
   GLSetupRecLen = LEN(GLSETUP)
   SetUpFileNum = FREEFILE
   OPEN "GLSETUP.DAT" FOR RANDOM AS SetUpFileNum LEN = GLSetupRecLen
END SUB

SUB OperatorEntry (Results, OperNum, PostDate$)
  
mainsub:
  ' show main screen
  action = 1
  added = 0
  DisplayOperEntry
  LibName$ = "CM.QSL"
  ScrnName$ = "CMCHOSOP"
  help$ = "Cash Management Operator Entry"
  
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
    Form$(3, 0) = DATE$
  DO
    
    EditForm Form$(), Fld(), frm(1), Cnf, action
    
    SELECT CASE frm(1).KeyCode
    CASE F7KEY
      AddEditOperator
      added = 1: EXIT DO
      Done = False
    CASE F10Key
      OperNum = VAL(Form$(1, 0))
      OperPassword$ = Form$(2, 0)
      PostDate$ = Form$(3, 0)
      GOSUB CheckPostDate
      IF BadDate = True THEN
       LOCATE 12, 49: COLOR 4: PRINT "<-- DATE NOT SET"; CHR$(7); : SLEEP 1: GOTO mainsub
      END IF
      IF OperNum > 0 THEN
       ValidateOperator OperNum, OperPassword$, Results
       Done = True
      ELSE
       Results = 0
       Done = True
      END IF
    CASE ESC
      Results = -1
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  
  IF added = 1 THEN GOTO mainsub
  EXIT SUB


CheckPostDate:
    IF VAL(LEFT$(PostDate$, 2)) < 1 OR VAL(LEFT$(PostDate$, 2)) > 12 THEN BadDate = True: RETURN
    IF VAL(MID$(PostDate$, 4, 2)) < 1 OR VAL(MID$(PostDate$, 4, 2)) > 31 THEN BadDate = True: RETURN
    BadDate = False
    RETURN

  
END SUB

SUB PostAndPrint (Posted)
  help$ = "Citipak Cash Management System v8.0"
  ShowCursor
  LibName$ = "CM.QSL"
  ScrnName$ = "CMPOST"

  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)

  ' Define Quick Screen Form Editing Arrays
  REDIM frm(1) AS FormInfo
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

    EditForm Form$(), Fld(), frm(1), Cnf, action

    SELECT CASE frm(1).KeyCode
    CASE F10Key
      IF Form$(1, 0) = "Y" THEN
        PrintReceipt
        Form$(1, 0) = "N"
        frm(1).FldNo = 1
        action = 1
        GOTO PrtReceipt
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
END SUB

  SUB PrintJournal (OperRecNumb)

  SHARED User$
  LibName$ = "CM.QSL"
  ScrnName$ = "PRNTJOUR"         'form to edit
  DevFld = 4    'Set to output device field number

  SHARED Choice$()

  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)

  ' Define Quick Screen Form Editing Arrays
  REDIM frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode

  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F


  REDIM Choice$(0 TO 3, 0 TO 1)
  

  Choice$(0, 0) = "1"  'Choices for Output Order
  Choice$(1, 0) = "Entry Order"
  Choice$(2, 0) = "Name"

  Choice$(0, 1) = STR$(DevFld)  'Choices for Output Device
  Choice$(1, 1) = "Screen"
  Choice$(2, 1) = "Prn#1"
  Choice$(3, 1) = "Prn#2"

  LibFile2Scrn "CM.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  
  action = 1

  DO

    EditForm Form$(), Fld(), frm(1), Cnf, action

    SELECT CASE frm(1).KeyCode  'Check the user's key presses
    
    CASE F10Key 'SaveButton
    
      BegDate = Date2Num(Form$(2, 0))
      BegDate$ = Form$(2, 0)
      EndDate = Date2Num(Form$(3, 0))
      EndDate$ = Form$(3, 0)

      IF LEN(LTRIM$(RTRIM$(Form$(DevFld, 0)))) = 0 THEN
        dev$ = "S"
        LPTNo = 1
      ELSE
        dev$ = LEFT$(Form$(DevFld, 0), 1)
        LPTNo = VAL(RIGHT$(RTRIM$(Form$(DevFld, 0)), 1))
      END IF

      IF BegDate > EndDate THEN
        Ok2Go = 0
        frm(1).FldNo = 5
        action = 1
      ELSE
        Ok2Go = 1
      END IF

    CASE ESC          'ESC Key
      EXIT SUB

    END SELECT

  LOOP UNTIL Ok2Go

  LOCATE , , 0



  
  IF LEFT$(Form$(1, 0), 2) = "En" THEN
   SortOrder$ = "Entry Order"
   ELSE
   SortOrder$ = "Name Order"
  END IF

  ERASE Form$, Fld, Choice$, frm




  'End of Input
  '=====================================================
  'Start Report Processing

  ReportFile$ = "CMJOURNL.PRN"  'Report File Name
  CommaFmt$ = "########,.##"    'format takes 13 chars
  TotalFmt$ = "#########,.##"   'format takes 14 chars
  SumLine$ = STRING$(13, "-")   'column summary line
  DivLine$ = STRING$(77, "-")   'dashed line
  DivLine2$ = STRING$(77, "=")  'Double Line
  FF$ = CHR$(12)
  MaxLines = 53
  LineCnt = 0
  TotDr# = 0
  TotCr# = 0
  size = 2500
  Start = 1               'start at array element 1
  Dir = 0                 'sort direction - use anything else for descending
  SSize = 16               'total size of each TYPE element
  MOff = 0                'offset into the TYPE for the key element
  MSize = 16              'size of the key element - coded as follows:
                        '   -1 = integer
                        '   -2 = long integer
                        '   -3 = single precision
                        '   -4 = double precision
                        '   +N = TYPE array/fixed-length string of length N

  REDIM array(1 TO size) AS Struct
  REDIM UBTRRec(1) AS UBTransRecType     ' open transaction file
  UBTRRecLen = LEN(UBTRRec(1))
  TrHandle = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS TrHandle LEN = UBTRRecLen
  TrNumRecs = LOF(TrHandle) \ UBTRRecLen
  

  GOSUB GetReportInformation

  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle

  REDIM OperRec(1) AS CMOperRecType          ' opens operatorfile
  OperRecLen = LEN(OperRec(1))
  OperFile = FREEFILE
  OPEN "CMOPER.DAT" FOR RANDOM AS OperFile LEN = OperRecLen
  NumOperRecs = LOF(OperFile) / OperRecLen
  GET OperFile, OperRecNumb, OperRec(1)

  OpenMiscCodeFile NumOfMiscRecs, MCFile      ' opens misc code file
  REDIM MiscCodeRec(1) AS MiscCodeRecType

  PRINT #RptHandle, CHR$(27); CHR$(58); ' oki 320 12 cpi

  GOSUB PrintRptHeader

  FOR Cnt = 1 TO Count
   GET TrHandle, array(Cnt).RecNum, UBTRRec(1)
   IF LineCnt >= MaxLines THEN
    PRINT #RptHandle, FF$
    GOSUB PrintRptHeader
   END IF
   PRINT #RptHandle, Num2Date(UBTRRec(1).TransDate); TAB(12); UBTRRec(1).TransDesc;
    PRINT #RptHandle, TAB(40); USING "#####,#.##"; UBTRRec(1).CashAmount;
    PRINT #RptHandle, TAB(55); USING "#####,#.##"; UBTRRec(1).CheckAmount;
    PRINT #RptHandle, TAB(70); USING "#####,#.##"; UBTRRec(1).TransAmt;
    PRINT #RptHandle, TAB(84); USING "#####,#.##"; (UBTRRec(1).CashAmount + UBTRRec(1).CheckAmount) - UBTRRec(1).TransAmt
    TotalCash# = TotalCash# + UBTRRec(1).CashAmount
    TotalCash# = Round(TotalCash#)
    TotalCheck# = TotalCheck# + UBTRRec(1).CheckAmount
    TotalCheck# = Round(TotalCheck#)
    TotalAmount# = TotalAmount# + UBTRRec(1).TransAmt
    TotalAmount# = Round(TotalAmount#)
    TotalReceipts = TotalReceipts + 1
    LineCnt = LineCnt + 1
   'Second Line of Print is Misc Code Breakdown Dist.
    PrintMiscFlag = 0
    FOR MCnt = 1 TO 5
    MiscRevAmt# = (UBTRRec(1).RevAmt(MCnt))
    MiscRevAmt# = Round#(MiscRevAmt#)
     IF MiscRevAmt# > 0 THEN
      GET MCFile, UBTRRec(1).CurRead(MCnt), MiscCodeRec(1)
       PRINT #RptHandle, "Code BrkDwn:";
       PRINT #RptHandle, TAB(14); MiscCodeRec(1).MiscCode;
       PRINT #RptHandle, TAB(25); MiscCodeRec(1).Description;
       PRINT #RptHandle, TAB(55); USING "#####,#.##"; MiscRevAmt#
       PrintMiscFlag = 1
       LineCnt = LineCnt + 1
     END IF
  NEXT MCnt
     IF PrintMiscFlag = 1 THEN PRINT #RptHandle, STRING$(96, "-"): LineCnt = LineCnt + 1
  NEXT Cnt
    GOSUB PrintRptEnding
    PRINT #RptHandle, CHR$(18); ' oki 320 12 cpi
    CLOSE                       'Close all open files now

  SELECT CASE dev$
    CASE "S"
      EntryPoint = 2
    CASE "P"
      EntryPoint = 5
  END SELECT
  
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  KILL ReportFile$
  EXIT SUB



PrintRptHeader:
page = page + 1
PRINT #RptHandle, TAB(27); "Cash Receipts Journal : Cash Management System"
PRINT #RptHandle, "Beginning Date: "; BegDate$
PRINT #RptHandle, "   Ending Date: "; EndDate$
PRINT #RptHandle, "      Operator: "; OperRec(1).operatorname$; TAB(83); "Page #"; page
PRINT #RptHandle, ""
PRINT #RptHandle, "   Date"; TAB(12); "Description"; TAB(40); "Cash Amount"; TAB(55); "Check Amount"; TAB(70); "Amount Owed"; TAB(84); "Change Given"
PRINT #RptHandle, STRING$(96, "=")
LineCnt = 5
RETURN

PrintRptEnding:
  PRINT #RptHandle, STRING$(96, "-")
    PRINT #RptHandle, "Operator # "; OperRec(1).OperatorNumber; " Totals .. ";
    PRINT #RptHandle, TAB(40); USING "#####,#.##"; TotalCash#;
    PRINT #RptHandle, TAB(55); USING "#####,#.##"; TotalCheck#;
    PRINT #RptHandle, TAB(70); USING "#####,#.##"; TotalAmount#;
    PRINT #RptHandle, TAB(84); USING "#####,#.##"; (TotalCash# + TotalCheck#) - TotalAmount#
    PRINT #RptHandle, "Number of Receipts .. "; USING "####,#"; TotalReceipts



  PRINT #RptHandle, FF$
  RETURN

GetReportInformation:
  BegRecNumb = TrNumRecs - 7500          ' Move back 7500 records to begin
  IF BegRecNumb < 1 THEN BegRecNumb = 1  ' Don't Allow Less Than 1


  FOR Cnt! = BegRecNumb TO TrNumRecs
   GET TrHandle, Cnt!, UBTRRec(1)
   TransDate = UBTRRec(1).TransDate

    IF UBTRRec(1).CustAcctNo = 0 AND UBTRRec(1).OperatorNumber = OperRecNumb AND UBTRRec(1).TransType = 99 AND (TransDate >= BegDate AND TransDate <= EndDate) THEN
      Count = Count + 1
    IF SortOrder$ = "Entry Order" THEN
      array(Count).who = STR$(Cnt!)
    ELSE
      array(Count).who = LEFT$(UBTRRec(1).TransDesc, 14)
    END IF
     array(Count).RecNum = Cnt!
   END IF
  NEXT Cnt!

SortT array(Start), Count, Dir, SSize, MOff, MSize
RETURN


END SUB

SUB PrintReceipt
  RMFile = FREEFILE
  REDIM RMRec(1) AS RMReceiptRecType
  RMRecLen = LEN(RMRec(1))
  OPEN "CMRECPT.DAT" FOR RANDOM AS RMFile LEN = RMRecLen
  LOCK #RMFile, 1
  GET RMFile, 1, RMRec(1)
  GOSUB GetReceiptNumber

  LPRINT
  LPRINT "CITY OF HOHENWALD"
  LPRINT "  MISC PAYMENT"
  LPRINT "Date: "; RMRec(1).RecDate
  LPRINT
  LPRINT "CUSTOMER NAME & DESC. OF PAYMENT"
  LPRINT RMRec(1).RecName
  LPRINT RMRec(1).RecAddress
  LPRINT RMRec(1).RecDesc
  LPRINT
  LPRINT "Total Owed: "; USING "$$####,#.##"; RMRec(1).RecAmtOwed
  LPRINT "Total Paid: "; USING "$$####,#.##"; RMRec(1).RecCashAmt + RMRec(1).RecCheckAmt
  LPRINT "Change Due: "; USING "$$####,#.##"; RMRec(1).RecChangeDue
  LPRINT
  LPRINT
  LPRINT "Operator: "; RMRec(1).RecOperator
  LPRINT "Receipt#: "; USING "######"; receiptnumber!
  LPRINT
  LPRINT TAB(7); "T H A N K   Y O U !"
  LPRINT
  LPRINT
  CLOSE RMFile
  EXIT SUB

GetReceiptNumber:
  Thandle = FREEFILE
  OPEN "wbartran.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #Thandle LEN = 128
  receiptnumber! = LOF(Thandle) / 128
  CLOSE Thandle
  RETURN
END SUB

FUNCTION Round# (ThisAmount#)
  ThisAmount# = INT((ThisAmount# * 100) + .5) / 100
  Round# = ThisAmount#
END FUNCTION

SUB ShowPctComp (BYVAL RecNo, BYVAL NumOfRecs) STATIC
END SUB

'
SUB SourceOfReceipt (OperRecNumber, source$)
startmainsub:


  SHARED Choice$()
  
  LibName$ = "CM"
  ScrnName$ = "CMSOURCE"
  help$ = "Cash Management System V8.0"
  LOCATE 1, 1, 0
  
  
  ShowCursor
  LibFile2Scrn "CM.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  PrintHelp help$

  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Rem Set Choice
  REDIM Choice$(2, 0)
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Utility Payment"
  Choice$(2, 0) = "Misc Payment"
  'Choice$(3, 0) = "Utility Deposit"
  'Choice$(4, 0) = "Tax Payment"
  'Choice$(5, 0) = "License Payment"
  
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  
  
  REDIM CMOperRec(1) AS CMOperRecType
  CMOperRecLen = LEN(CMOperRec(1))
  CMFile = FREEFILE
  OPEN "CMOPER.DAT" FOR RANDOM AS CMFile LEN = CMOperRecLen
  GET CMFile, OperRecNumber, CMOperRec(1)
  LOCATE 3, 55: COLOR 15: PRINT LEFT$(CMOperRec(1).operatorname, 19)
  OPERNUMB = CMOperRec(1).OperatorNumber
  CLOSE CMFile
  
  DO
    EditForm Form$(), Fld(), frm(1), Cnf, action
    
    source$ = LEFT$(Form$(1, 0), 1)
    IF ASC(LEFT$(source$, 1)) > 32 THEN EXIT SUB
    
    SELECT CASE frm(1).KeyCode
    CASE F7KEY
      PrintJournal (OperRecNumber)
      action = 1
      GOTO startmainsub
    CASE ESC
      source$ = "INVALID"
      Done = True
    CASE ELSE
      Done = False
    END SELECT

  LOOP UNTIL Done
  
END SUB

SUB UtilPayEntry (OperRecNumber, PostDate$)
  REDIM MiscRecNo(10)
  REDIM MiscCodeRec(1) AS MiscCodeRecType
  SHARED Choice$()
  LibName$ = "CM"
  ScrnName$ = "UBCMUTIL"
  help$ = "Cash Management Operator Entry"
  LOCATE 1, 1, 0

  help$ = "Utility Payment Entry"
  ShowCursor
  LibFile2Scrn "CM.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  PrintHelp help$

  GOSUB UtilOperatorName


  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)

  ' Define Quick Screen Form Editing Arrays
  REDIM frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  'Rem Set Choice
  REDIM Choice$(3, 0)
  'Choice$(0, 0) = "5"
  'Choice$(1, 0) = "Cash"
  'Choice$(2, 0) = "Check"
  'Choice$(3, 0) = "Cash & Check"
   
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode

  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
 
   
  DO
    EditForm Form$(), Fld(), frm(1), Cnf, action
  

    SELECT CASE frm(1).KeyCode
    CASE F7KEY
      Done = False
    CASE F10Key
      Done = False
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  EXIT SUB

UtilOperatorName:
  action = 1
  REDIM CMOperRec(1) AS CMOperRecType
  CMOperRecLen = LEN(CMOperRec(1))
  CMFile = FREEFILE
  OPEN "CMOPER.DAT" FOR RANDOM AS CMFile LEN = CMOperRecLen
  GET CMFile, OperRecNumber, CMOperRec(1)
  LOCATE 3, 55: COLOR 15: PRINT LEFT$(CMOperRec(1).operatorname, 19)
  LOCATE 4, 44: COLOR 11: PRINT "Post Date: "; : COLOR 15: PRINT PostDate$
  CLOSE CMFile
  RETURN

UtilStoreReceipt:
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
  PUT RMFile, 1, RMRec(1)
  CLOSE RMFile
  RETURN

UtilPostTransaction:
  PostDate = Date2Num(PostDate$)
  REDIM UBTRRec(1) AS UBTransRecType
  UBTRRecLen = LEN(UBTRRec(1))
  Thandle = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM LOCK READ WRITE AS Thandle LEN = UBTRRecLen
  UBTRRec(1).TransDate = PostDate
  UBTRRec(1).TransType = 2
  UBTRRec(1).TransDesc = Form$(1, 0)
  UBTRRec(1).TransAmt = CVD(MID$(Form$(0, 0), Fld(4).Fields, 8))
  UBTRRec(1).RunBalance = 0
  FldFactor = 0
 FOR Cnt = 1 TO 5
  UBTRRec(1).RevAmt(Cnt) = CVD(MID$(Form$(0, 0), Fld(Cnt + 10 + FldFactor).Fields, 8))
  FldFactor = FldFactor + 1
 NEXT Cnt
 FOR Cnt = 6 TO 10
  UBTRRec(1).RevAmt(Cnt) = 0
 NEXT Cnt

    FldFactor = 0
 FOR Cnt = 1 TO 5
    UBTRRec(1).CurRead(Cnt) = MiscRecNo(Cnt + FldFactor)
    FldFactor = FldFactor + 1
 NEXT Cnt
 FOR Cnt = 6 TO 7
    UBTRRec(1).CurRead(Cnt) = 0
 NEXT Cnt
 FOR Cnt = 1 TO 7
  UBTRRec(1).PrevRead(Cnt) = 0
  'UBTRRec(1).Consumption(Cnt) = 0
  UBTRRec(1).EstRead(Cnt) = ""
 NEXT Cnt
'UBTRRec(1).CustLocation = ""
UBTRRec(1).CustStatus = ""
UBTRRec(1).CustCode = ""
UBTRRec(1).CustCode = ""
UBTRRec(1).billcycle = ""
UBTRRec(1).PayTypeCode = RMRec(1).RecPayType
UBTRRec(1).OperatorNumber = OperRecNumb
UBTRRec(1).CustAcctNo = 0
UBTRRec(1).PrevTrans = 0
UBTRRec(1).NextTrans = 0
UBTRRec(1).CashAmount = CVD(MID$(Form$(0, 0), Fld(6).Fields, 8))
UBTRRec(1).CheckAmount = CVD(MID$(Form$(0, 0), Fld(7).Fields, 8))
PUT Thandle, LOF(Thandle) \ UBTRRecLen + 1, UBTRRec(1)
CLOSE Thandle
RETURN

END SUB

SUB ValidateOperator (OperNum, OperPassword$, Results)
  
  ' First Check for Operator File '
  
  DIM CMOperRec AS CMOperRecType
  OperFileLength = LEN(CMOperRec)
  filehandle = FREEFILE
  OPEN "CMOPER.DAT" FOR RANDOM ACCESS READ WRITE AS #filehandle LEN = OperFileLength
  IF LOF(filehandle) = 0 THEN
    Results = 0
  ELSE
    FOR Cnt = 1 TO LOF(filehandle) / OperFileLength
      GET filehandle, Cnt, CMOperRec
      
      IF OperNum = CMOperRec.OperatorNumber AND OperPassword$ = CMOperRec.OperatorPassword THEN
        Results = Cnt
        EXIT SUB
      END IF
    NEXT Cnt
    Results = 0
  END IF
  
  IF Results = 0 THEN
    PRINT CHR$(7);
    LibName$ = "CM"
    ScrnName$ = "BADOPER"
    LibFile2Scrn "CM.QSL", "BADOPER", MonoCode%, Attribute%, ErrorCode%
    help$ = "Cash Management Operator Entry"
    
    ' Define Fields
    NumFlds = LibNumberOfFields(LibName$, ScrnName$)
    
    ' Define Quick Screen Form Editing Arrays
    REDIM frm(1) AS FormInfo
    REDIM Form$(NumFlds, 2)
    REDIM Fld(NumFlds) AS FieldInfo
    
    ' Get 1st & Last Fields
    StartEl = 0
    LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
    
    ' Clear Fields
    FOR F = 1 TO NumFlds
      LSET Form$(F, 0) = ""
    NEXT F
    
    DO
      EditForm Form$(), Fld(), frm(1), Cnf, action
      
      
      SELECT CASE frm(1).KeyCode
        
      CASE ESC
        Done = True
      CASE ELSE
        Done = False
      END SELECT
    LOOP UNTIL Done
    
  END IF
  
END SUB

