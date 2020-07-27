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
DECLARE FUNCTION FindAcct% (AcctNum$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE FUNCTION Round# (ThisAmount#)
DECLARE SUB FileView (FileName$, Ky, Action, FVI AS ANY, SEG array)
DECLARE FUNCTION FUsing$ (Number$, Mask$)
DECLARE SUB HideCursor ()
DECLARE FUNCTION MGetKey% (Row%, Col%)
DECLARE FUNCTION Monitor% ()
DECLARE SUB MScrnSave (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MScrnRest (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MPaintBox (ULRow%, ULCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
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
'$INCLUDE: 'gl.bi'

  CONST False = 0, True = NOT False

DIM SHARED GLSETUP AS GLSetupRecType
DIM SHARED ACCT AS GLAcctRecType
DIM SHARED AcctIdx AS GLAcctIndexType

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

    IF source$ = "U" THEN
      OldUtilPayEntry OperRecNumber, PostDate$
    END IF

  LOOP UNTIL source$ = "INVALID"

  END

SUB AddEditOperator
  
  AuthorizeOperator oktoadd
  
  IF oktoadd = 0 THEN EXIT SUB
  
  LibName$ = "UB"
  ScrnName$ = "CMNEWOP"
  help$ = "Cash Management Operator Entry"
  LOCATE 1, 1, 0
  
  help$ = "Add Operator v8.0"
  ShowCursor
  LibFile2Scrn "UB.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  PrintHelp help$
  
  
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET form$(F, 0) = ""
  NEXT F
  
  Action = 1
  REDIM CMOperRec(1) AS CMOperRecType
  CMOperRecLen = LEN(CMOperRec(1))
  CMFile = FREEFILE
  OPEN "CMOPER.DAT" FOR RANDOM AS CMFile LEN = CMOperRecLen
  NumRecs = LOF(CMFile) / CMOperRecLen
  
  DO
    EditForm form$(), Fld(), Frm(1), Cnf, Action
   IF Frm(1).PrevFld = 2 AND edit <> True THEN
   FOR Cnt = 1 TO NumRecs
    GET CMFile, Cnt, CMOperRec(1)
    IF VAL(form$(1, 0)) = CMOperRec(1).OperatorNumber AND VAL(form$(1, 0)) > 0 THEN
      form$(2, 0) = CMOperRec(1).operatorname
      form$(3, 0) = CMOperRec(1).OperatorPassword
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
 
      CMOperRec(1).OperatorNumber = VAL(form$(1, 0))
      CMOperRec(1).operatorname = form$(2, 0)
      CMOperRec(1).OperatorPassword = form$(3, 0)
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

  LibName$ = "UB"
  ScrnName$ = "CMNWCODE"
  help$ = "Cash Management Operator Entry"
  LOCATE 1, 1, 0

  help$ = "Add Misc Codes V8.0"
  ShowCursor
  LibFile2Scrn "UB.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  PrintHelp help$
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)

  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), form$(), ErrCode

  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET form$(F, 0) = ""
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
      LibFile2Scrn "UB.QSL", "MENUBAK", MonoCode, -1, ErrorCode
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
     form$(1, 0) = MiscCodeRec(1).MiscCode
     form$(2, 0) = MiscCodeRec(1).Description
     form$(3, 0) = MiscCodeRec(1).GlAcctNumb
  END IF

  ShowCursor
  LibFile2Scrn "UB.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  PrintHelp help$


  DO
    EditForm form$(), Fld(), Frm(1), Cnf, Action



    SELECT CASE Frm(1).KeyCode

    CASE F10Key
      
      
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
       SLEEP 1
       Frm(1).FldNo = 3
       Action = 1
       LOCATE 14, 18: PRINT "                                               "
       Done = False
       NoProcess = True
      END IF

       IF NoProcess = False THEN
      MiscCodeRec(1).MiscCode = form$(1, 0)
      MiscCodeRec(1).Description = form$(2, 0)
      MiscCodeRec(1).GlAcctNumb = form$(3, 0)

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
GLNumber$ = LEFT$(form$(3, 0), Fund) + "-" + MID$(form$(3, 0), Fund + 1, ACCNT) + "-" + MID$(form$(3, 0), Fund + ACCNT + 1, DET)
  
AcctRecNo = FindAcct(GLNumber$)

CLOSE SetUpFileNum
RETURN

CheckCodes:
 FOR Cnt = 1 TO NumOfMiscRecs
 GET MCFile, Cnt, MiscCodeRec(1)
 IF form$(1, 0) = MiscCodeRec(1).MiscCode THEN validcode = False: RETURN
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
  LibName$ = "UB"
  ScrnName$ = "OPERAUTH"
  help$ = "Cash Management Operator Entry"
  LOCATE 1, 1, 0
  
  help$ = "Authorization Required"
  ShowCursor
  LibFile2Scrn "UB.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  PrintHelp help$
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET form$(F, 0) = ""
  NEXT F
  
  DO
    EditForm form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      IF form$(1, 0) = "CASHMGMT" THEN
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
  LibFile2Scrn "UB.QSL", "CMCHOSOP", MonoCode%, Attribute%, ErrorCode%
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
      IF Lookfor$ = RTRIM$(AcctIdx.AcctNum) THEN
         Match = -1
         EXIT DO
      ELSEIF Lookfor$ < AcctIdx.AcctNum THEN
         LastRec = MiddleRec - 1
      ELSE
         FirstRec = MiddleRec + 1
      END IF
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
      'LibFile2Scrn "UB.QSL", "MENUBAK", MonoCode, -1, ErrorCode
      'ShowCursor
      QPrintRC TText$, Row - 1, Col, 112
      VertMenu Mchoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf

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

  LibName$ = "UB"
  ScrnName$ = "UBCMMISC"
  help$ = "Cash Management Operator Entry"
  LOCATE 1, 1, 0
  
  help$ = "Misc Payment Entry"
  ShowCursor
  LibFile2Scrn "UB.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  PrintHelp help$
  GOSUB SetOperatorName
  
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  ' Rem Set Choice
  REDIM Choice$(3, 0)
  Choice$(0, 0) = "5"
  Choice$(1, 0) = "Cash"
  Choice$(2, 0) = "Check"
  Choice$(3, 0) = "Cash & Check"
  


  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET form$(F, 0) = ""
  NEXT F
 ' Get operator name and display
  

  DO
    EditForm form$(), Fld(), Frm(1), Cnf, Action




    'Force amount in field 4 cannot be 0 or less than 0
    IF Frm(1).FldNo > 4 AND Frm(1).PrevFld = 4 AND VAL(form$(4, 0)) <= 0 THEN
     Frm(1).FldNo = 4
    END IF

    IF Frm(1).FldNo >= 4 AND VAL(form$(9, 0)) < 0 THEN
     IF ClearIt = 0 THEN
      LOCATE 18, 28: PRINT STRING$(15, 32)
      ClearIt = 1
     END IF
    END IF

   'rem set cash/check/both fields
    IF Frm(1).PrevFld = 5 AND Frm(1).FldEdited = -1 THEN
      form$(6, 0) = "0.00": form$(7, 0) = "0.00"
      form$(8, 0) = "0.00": form$(9, 0) = "0.00"
      Action = 1
    END IF
    IF LEFT$(form$(5, 0), 6) = "Cash  " THEN
     Fld(7).Protected = True
     form$(7, 0) = "0.00"
     Fld(6).Protected = False
     IF PolledIt = 0 THEN
      Action = 1: PolledIt = 1: PayHow$ = LEFT$(form$(5, 0), 2)
     END IF
    END IF

    IF LEFT$(form$(5, 0), 2) = "Ch" THEN
     Fld(6).Protected = True: form$(6, 0) = "0.00"
     Fld(7).Protected = False
     IF PolledIt = 0 THEN
      Action = 1: PolledIt = 1: PayHow$ = LEFT$(form$(5, 0), 2)
'      CalcFields 0, 8, form$(), Fld()
     END IF
    END IF

    IF LEFT$(form$(5, 0), 6) = "Cash &" THEN
     Fld(6).Protected = False
     Fld(7).Protected = False
     IF PolledIt = 0 THEN
      Action = 1: PolledIt = 1: PayHow$ = LEFT$(form$(5, 0), 2)
 
     END IF
    END IF

    IF PolledIt = 1 AND LEFT$(form$(5, 0), 2) <> PayHow$ THEN
     PolledIt = 0: Action = 1
  '   CalcFields 0, 8, form$(), Fld()
    END IF

    IF Frm(1).FldNo > 9 AND Frm(1).PrevFld <= 9 AND VAL(form$(9, 0)) < 0 THEN
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
      LibFile2Scrn "UB.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
      PrintHelp help$
      GOSUB SetOperatorName
      Action = 1
      Done = False
    CASE F10Key
     RecAmtOwed# = CVD(MID$(form$(0, 0), Fld(4).Fields, 8))
     RecAmtRecd# = CVD(MID$(form$(0, 0), Fld(8).Fields, 8))
     MiscAmtRecd# = CVD(MID$(form$(0, 0), Fld(20).Fields, 8))
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
        GOSUB PostTransaction         ' Normal Posting
        GOSUB PostOldTransaction      ' Hohenwald Tenn Only
       Done = True
       ELSE
      help$ = "Misc Payment Entry"
      ShowCursor
      LibFile2Scrn "UB.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
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
  LOCATE 3, 55: COLOR 15: PRINT LEFT$(CMOperRec(1).operatorname, 19)
  LOCATE 4, 44: COLOR 11: PRINT "Post Date: "; : COLOR 15: PRINT PostDate$
  Operator = CMOperRec(1).OperatorNumber
  Operator$ = STR$(Operator)
  Operator$ = RIGHT$(Operator$, LEN(Operator$) - 1)
  CLOSE CMFile
  RETURN

PollMiscCodeEntry:
    IF Frm(1).FldNo < 10 THEN RETURN

    IF (Frm(1).FldNo = 11 AND Frm(1).PrevFld = 10) OR (Frm(1).FldNo = 13 AND Frm(1).PrevFld = 12) OR (Frm(1).FldNo = 15 AND Frm(1).PrevFld = 14) OR (Frm(1).FldNo = 17 AND Frm(1).PrevFld = 16) OR (Frm(1).FldNo = 19 AND Frm(1).PrevFld = 18) THEN

    MiscCodeValue$ = form$(Frm(1).PrevFld, 0)
    
    IF (form$(4, 0) = form$(20, 0)) AND (Frm(1).FldNo = 13 OR Frm(1).FldNo = 15 OR Frm(1).FldNo = 17 OR Frm(1).FldNo = 19) THEN Frm(1).FldNo = 21: Action = 1: RETURN

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
     form$(Frm(1).PrevFld, 0) = MiscCodeRec(1).MiscCode
     MiscRecNo((Frm(1).PrevFld) - 9) = RecNo
     Action = 1
     CLOSE MCFile
     END IF

     RETURN

StoreReceipt:
  RMFile = FREEFILE
  REDIM RMRec(1) AS RMReceiptRecType
  RMRecLen = LEN(RMRec(1))
  OPEN "CMRECPT.DAT" FOR RANDOM AS RMFile LEN = RMRecLen
    RMRec(1).RecName = form$(1, 0)
    RMRec(1).RecAddress = form$(2, 0)
    RMRec(1).RecDesc = form$(3, 0)
    RMRec(1).RecAmtOwed = CVD(MID$(form$(0, 0), Fld(4).Fields, 8))
    IF LEFT$(form$(5, 0), 6) = "Cash  " THEN
     RMRec(1).RecPayType = 1
    END IF
    IF LEFT$(form$(5, 0), 6) = "Check " THEN
     RMRec(1).RecPayType = 2
    END IF
    IF LEFT$(form$(5, 0), 6) = "Cash &" THEN
     RMRec(1).RecPayType = 3
    END IF
    RMRec(1).RecCashAmt = CVD(MID$(form$(0, 0), Fld(6).Fields, 8))
    RMRec(1).RecCheckAmt = CVD(MID$(form$(0, 0), Fld(7).Fields, 8))
    RMRec(1).RecChangeDue = CVD(MID$(form$(0, 0), Fld(9).Fields, 8))
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
  UBTRRec(1).TransType = 2
  UBTRRec(1).TransDesc = form$(1, 0)
  UBTRRec(1).TransAmt = CVD(MID$(form$(0, 0), Fld(4).Fields, 8))
  UBTRRec(1).RunBalance = 0
  FldFactor = 0
 FOR Cnt = 1 TO 5
  UBTRRec(1).RevAmt(Cnt) = CVD(MID$(form$(0, 0), Fld(Cnt + 10 + FldFactor).Fields, 8))
  FldFactor = FldFactor + 1
 NEXT Cnt
 FOR Cnt = 6 TO 10
  UBTRRec(1).RevAmt(Cnt) = 0
 NEXT Cnt

 FOR Cnt = 1 TO 10
  UBTRRec(1).CustRate(Cnt) = ""
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
UBTRRec(1).CashAmount = CVD(MID$(form$(0, 0), Fld(6).Fields, 8))
UBTRRec(1).CheckAmount = CVD(MID$(form$(0, 0), Fld(7).Fields, 8))
PUT Thandle, LOF(Thandle) \ UBTRRecLen + 1, UBTRRec(1)
CLOSE Thandle
RETURN

PostOldTransaction:
REM FOR HOHENWALD TENN.

   Thandle = FREEFILE
   OPEN "wbartran.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #Thandle LEN = 128
   FIELD Thandle, 8 AS tdate$, 8 AS tamt$, 2 AS ttype$, 15 AS tdesc$, 4 AS tmrec$, 4 AS tnexttr$, 4 AS tcuramt$
   FIELD Thandle, 45 AS tu1$, 4 AS wbrev1$, 4 AS wbrev2$, 4 AS wbrev3$, 4 AS wbrev4$, 4 AS wbrev5$, 4 AS wbrev6$, 4 AS wbrev7$, 4 AS wbrev8$
   FIELD Thandle, 77 AS tu2$, 1 AS testimate$, 8 AS tcurread$, 8 AS tpastread$
   FIELD Thandle, 94 AS TU3$, 8 AS tconsump$, 2 AS tbook$, 2 AS tcusttype$, 1 AS tcustloc$: REM tcustloc$ = 1 for cash and 2 for check in payment transaction
   FIELD Thandle, 107 AS TU4$, 1 AS tstatus$, 2 AS OPERNUMB$, 4 AS POSTTIME$, 1 AS tglpcode$, 13 AS wbblank$

  
  LSET tdate$ = LEFT$(PostDate$, 6) + RIGHT$(PostDate$, 2)
  LSET tamt$ = MKD$(CVD(MID$(form$(0, 0), Fld(6).Fields, 8)))
    IF LEFT$(form$(5, 0), 6) = "Cash  " THEN
     LSET ttype$ = MKI$(1)
    END IF
    IF LEFT$(form$(5, 0), 6) = "Check " THEN
     LSET ttype$ = MKI$(2)
    END IF
    IF LEFT$(form$(5, 0), 6) = "Cash &" THEN
     LSET ttype$ = MKI$(3)
    END IF

  
  LSET tdesc$ = form$(3, 0)
  LSET tmrec$ = MKS$(999999)
  LSET tnexttr$ = MKS$(0)
  LSET tcuramt$ = MKS$(CVD(MID$(form$(0, 0), Fld(4).Fields, 8)))
  LSET wbrev1$ = MKS$(MiscRecNo(1))
  LSET wbrev2$ = MKS$(MiscRecNo(3))
  LSET wbrev3$ = MKS$(MiscRecNo(5))
  LSET wbrev4$ = MKS$(MiscRecNo(7))
  LSET wbrev5$ = MKS$(MiscRecNo(9))
  MiscRev1# = CVD(MID$(form$(0, 0), Fld(11).Fields, 8))
  MiscRev1# = Round(MiscRev1#)
  LSET wbrev6$ = MKS$(MiscRev1#)
  IF MiscRecNo(3) > 0 THEN
   MiscRev2# = CVD(MID$(form$(0, 0), Fld(13).Fields, 8))
   MiscRev2# = Round(MiscRev2#)
  ELSE
   MiscRev2# = 0
  END IF
  LSET wbrev7$ = MKS$(MiscRev2#)
  IF MiscRecNo(5) > 0 THEN
   MiscRev3# = CVD(MID$(form$(0, 0), Fld(15).Fields, 8))
   MiscRev3# = Round(MiscRev3#)
   ELSE
   MiscRev3# = 0
  END IF
  LSET wbrev8$ = MKS$(MiscRev3#)
  LSET testimate$ = "N"
  IF MiscRecNo(7) > 0 THEN
  MiscRev4# = CVD(MID$(form$(0, 0), Fld(17).Fields, 8))
  MiscRev4# = Round(MiscRev4#)
  ELSE
  MiscRev4# = 0
  END IF
  LSET tcurread$ = MKD$(MiscRev4#)
  IF MiscRecNo(9) > 0 THEN
  MiscRev5# = CVD(MID$(form$(0, 0), Fld(19).Fields, 8))
  MiscRev5# = Round(MiscRev5#)
  ELSE
  MiscRev5# = 0
  END IF
  LSET tpastread$ = MKD$(MiscRev5#)
  LSET tconsump$ = MKD$(CVD(MID$(form$(0, 0), Fld(7).Fields, 8)))
  LSET tbook$ = MKI$(0)
  LSET tcusttype$ = MKI$(0)
    IF LEFT$(form$(5, 0), 6) = "Cash  " THEN
     LSET tcustloc$ = "1"
    END IF
    IF LEFT$(form$(5, 0), 6) = "Check " THEN
     LSET tcustloc$ = "2"
    END IF
    IF LEFT$(form$(5, 0), 6) = "Cash &" THEN
     LSET tcustloc$ = "3"
    END IF
  LSET tstatus$ = ""
  LSET OPERNUMB$ = Operator$
  LSET POSTTIME$ = LEFT$(TIME$, 2) + MID$(TIME$, 4, 2)
  LSET tglpcode$ = "N"
  LSET wbblank$ = form$(1, 0)
  PUT Thandle, LOF(Thandle) / 128 + 1
  CLOSE Thandle
  RETURN

END SUB

     SUB OldUtilPayEntry (OperRecNumber, PostDate$)
     DIM translink!(700), totm@(299), messline$(20), m$(20), Apply!(10)
     Scren$ = "4.00.00"
     SysDate$ = DATE$
  
10 CLS
    GOSUB 63500
    GOSUB OldSetOperatorName:
    
1000 REM
     FOR ll = 3 TO 23: LOCATE ll, 1: PRINT STRING$(79, 32): NEXT ll
     FOR JJ = 1 TO 8: Apply!(JJ) = 0: NEXT JJ
     pdesc$ = ""
     GOSUB openfile
     GET 14, 1
     max = CVI(REVTRACK$)
     CLOSE 14
1020 LOCATE 1, 30: COLOR 11: PRINT "Utility Payment Entry"; : LOCATE 3, 1
1025 DEPFLAG = 0
1030 COLOR 11: PRINT "Enter Acct # Please (<TAB>-Name Search : <ESC>-End) .. "; : fl = -5
     GOSUB 62000
     IF cf = 1 THEN EXIT SUB
     IF cf = 2 THEN GOSUB namesearch: IF GGO = 1 THEN GOTO 1080 ELSE CLOSE : GOTO 1000
     
1040 record = VAL(in$)
1080 CLOSE 1: GOSUB 60000
     IF record < 1 OR record > LOF(1) / 1024 THEN CLOSE : GOTO 1000
      GET 1, record: AcctRecord! = record
      GOSUB openfile: GET 14, 1
     IF VAL(LEFT$(meter$, 2)) = 0 THEN
      LOCATE 23, 1: COLOR 11: PRINT "Account Deleted : Press <ENTER> to Continue "; : fl = 1
      GOSUB 62000: CLOSE : GOTO 1000
     END IF
1090 LOCATE 3, 1: PRINT STRING$(79, 32)
     B = INSTR(meter$, "/"): IF B = 0 THEN m$ = meter$: GOTO 1095
     m$ = LEFT$(meter$, B - 1)
1095 COLOR 15: LOCATE 3, 1
      PRINT "  Account Number: "; USING "#####"; record;
     LSET IRev1$ = MKD$((INT((CVD(IRev1$) * 100) + .5) / 100))
     LSET IRev2$ = MKD$((INT((CVD(IRev2$) * 100) + .5) / 100))
     LSET IRev3$ = MKD$((INT((CVD(IRev3$) * 100) + .5) / 100))
     LSET IRev4$ = MKD$((INT((CVD(IRev4$) * 100) + .5) / 100))
     LSET IRev5$ = MKD$((INT((CVD(IRev5$) * 100) + .5) / 100))
     LSET IRev6$ = MKD$((INT((CVD(IRev6$) * 100) + .5) / 100))
     LSET irev7$ = MKD$((INT((CVD(irev7$) * 100) + .5) / 100))
     LSET IREV8$ = MKD$((INT((CVD(IREV8$) * 100) + .5) / 100))
     PUT 1, record
     LOCATE 3, 45
     PRINT "( CURRENT: "; USING "####,#.##"; CVD(CURBAL$); : PRINT ")"
     LOCATE 4, 1: PRINT "Account Location: "; m$;
     LOCATE 4, 45
     PRINT "(PAST DUE: "; USING "####,#.##"; CVD(pastdue$); : PRINT ")"
     LOCATE 5, 1
1100 IF CVS(messfirst$) > 0 THEN
      GOSUB reviewmessage
     END IF
     COLOR 11: PRINT "Customer Account of ..."; TAB(45); "Balance = "; USING "####.##"; CVD(CURBAL$) + CVD(pastdue$): owed# = CVD(CURBAL$) + CVD(pastdue$): owed# = INT((owed# * 100) + .5) / 100
1110 COLOR 14
     kk = INSTR(firstname$, "  "): IF kk = 0 THEN kk = 25
     IF LEFT$(firstname$, 1) = " " THEN
       nme$ = lname$
      ELSE
       nme$ = LEFT$(firstname$, kk) + lname$
     END IF
1120 IF messageflag = 1 THEN COLOR 27 ELSE COLOR 14
1135 PRINT TAB(5); LEFT$(nme$, 39);
     COLOR 14: PRINT TAB(45); "1-"; LEFT$(rev1name$, 15); USING "$####,#.##"; CVD(IRev1$)
1140 PRINT TAB(5); add1$; : IF max < 2 THEN PRINT  ELSE PRINT TAB(45); "2-"; LEFT$(rev2name$, 15); USING "$####,#.##"; CVD(IRev2$)
1150 PRINT TAB(5); add2$; : IF max < 3 THEN PRINT  ELSE PRINT TAB(45); "3-"; LEFT$(rev3name$, 15); USING "$####,#.##"; CVD(IRev3$)
1160 PRINT TAB(5); CITY$; " "; ST$; " "; ZP$; : IF max < 4 THEN PRINT  ELSE PRINT TAB(45); "4-"; LEFT$(rev4name$, 15); USING "$####,#.##"; CVD(IRev4$)
     IF max < 5 THEN 1165 ELSE PRINT TAB(45); "5-"; LEFT$(rev5name$, 15); USING "$####,#.##"; CVD(IRev5$)
     IF max < 6 THEN 1165 ELSE PRINT TAB(45); "6-"; LEFT$(rev6name$, 15); USING "$####,#.##"; CVD(IRev6$)
     IF max < 7 THEN 1165 ELSE PRINT TAB(45); "7-"; LEFT$(rev7name$, 15); USING "$####,#.##"; CVD(irev7$)
     IF max < 8 THEN 1165 ELSE PRINT TAB(45); "8-"; LEFT$(rev8name$, 15); USING "$####,#.##"; CVD(IREV8$)
1165 TOTREV# = CVD(IRev1$) + CVD(IRev2$) + CVD(IRev3$) + CVD(IRev4$) + CVD(IRev5$) + CVD(IRev6$) + CVD(irev7$) + CVD(IREV8$)
     TOTREV# = INT((TOTREV# * 100) + .5) / 100
1170 LOCATE 12, 1: COLOR 11: PRINT STRING$(44, "-")
1180 LOCATE 13, 1: COLOR 11: PRINT "Correct Acct (Y/N)(<F6>-Cash : <F7>-Check)";
     LOCATE 14, 1: COLOR 11: PRINT "(F4-Messages)(F5-History)(<TAB>-EDIT AMT):"; : fl = 1: GOSUB 62210:
     IF cf = 18 THEN
       LOCATE 19, 1
       PRINT "Enter New Cur Bal Amt .. "; : fl = -9: GOSUB 62000
       LSET CURBAL$ = MKD$(VAL(in$)): PUT 1, AcctRecord!
       LOCATE 19, 1: PRINT STRING$(79, 32)
       GOTO 1090
     END IF


     IF cf = 2 THEN GOSUB fixamts: GOTO 1090
     IF cf = 14 THEN
      GOSUB messagecenter
      Scren$ = "4.01.00": CLS : GOSUB 63500
      LOCATE 1, 34: COLOR 11: PRINT "Payment Entry"; : LOCATE 3, 1
      GOTO 1090
     END IF

     IF cf = 15 THEN
      GOSUB displayhistory
     Scren$ = "4.01.00": CLS : GOSUB 63500
     LOCATE 1, 34: COLOR 11: PRINT "Payment Entry"; : LOCATE 3, 1

      GOTO 1090
     END IF
     IF cf = 16 THEN amount@ = owed#: pcode = 1: GOTO 1285
     IF cf = 17 THEN
         amount@ = owed#: pcode = 2:
         LOCATE 15, 1: COLOR 11: PRINT "Enter Description (Ck #) ............ "; : fl = 15: GOSUB 62000: pdesc$ = in$
         GOTO 1285
     END IF
     IF NOT (in$ = "Y" OR in$ = "N") THEN 1180
1190 IF in$ = "Y" THEN 1200 ELSE CLOSE : GOTO 1000
1200 IF owed# = TOTREV# THEN 1205
      LOCATE 16, 1: PRINT "INDIVIDUAL REVENUES DO NOT EQUAL THE BALANCE"
      LOCATE 17, 1: COLOR 10
      PRINT "THE DIFFERENCE IS "; USING "$$#####,.##"; owed# - TOTREV#
      LOCATE 18, 1: COLOR 10
      PRINT "YOU SHOULD CORRECT THE REVENUES TO EQUAL THE BALANCE TO INSURE PROPER"
      PRINT "POSTING TO THE GENERAL LEDGER AND FOR A REVENUE REPORT THAT IS IN BALANCE"
1201 LOCATE 21, 1: COLOR 11
     PRINT "PRESS ANY KEY TO CONTINUE "; : fl = 1
     GOSUB 62000
     FOR JJ = 16 TO 23: LOCATE JJ, 1: PRINT STRING$(79, 32): NEXT JJ
     GOTO 1170
1205 LOCATE 13, 1: PRINT STRING$(44, 32)
     LOCATE 14, 1: PRINT STRING$(44, 32)
1210 IF CVS(waterdep$) > 0 THEN COLOR 10: LOCATE 23, 1: PRINT "Deposit on File in the Amount of :"; USING "###.##"; CVS(waterdep$);
1220 GOSUB 1900: LOCATE 13, 1: COLOR 11
     PRINT "Enter Payment Code (See Below 1-7) .. ";
     fl = -1: GOSUB 62000
     IF cf = 1 THEN CLOSE : GOTO 1000 ELSE pcode = VAL(in$)
     IF pcode < 1 OR pcode > 7 THEN 1220
1222 GOSUB 1950: COLOR 11: LOCATE 14, 55: PRINT dd$
1250 IF pcode = 3 THEN 1260
     LOCATE 14, 1: COLOR 11: PRINT "Enter Amount ... "; : PRINT "("; USING "$####.##"; owed#; : PRINT ") ......... "; : fl = -8: GOSUB 62000
     IF LEN(in$) = 0 THEN amount@ = owed#: LOCATE 14, 39: PRINT USING "$####.##"; amount@:  ELSE amount@ = VAL(in$): LOCATE 14, 39: PRINT USING "$####.##"; amount@
     GOTO 1260.0
1260 amount@ = CVS(waterdep$): amount@ = INT((amount@ * 100) + .5) / 100
1260.0 IF pcode = 4 OR pcode = 5 THEN
1260.1 LOCATE 15, 1: COLOR 11: PRINT "Enter Deposit Date (mmddyy) ......... "; : fl = 6: GOSUB 62000
       IF LEN(in$) <> 6 THEN GOTO 1260.1
       des$ = in$
       ELSE
1261 LOCATE 15, 1: COLOR 11: PRINT "Enter Description (Ck #) ............ "; : fl = 15: GOSUB 62000: pdesc$ = in$
     END IF
1265 LOCATE 16, 1: COLOR 11: PRINT "Enter Amount Tendered ............... "; : fl = -9: GOSUB 62000
     tendered@ = VAL(in$)
     LOCATE 16, 39:  PRINT USING "$####.##"; tendered@

     LOCATE 16, 55: COLOR 11
     PRINT "Change Due: "; : COLOR 15: PRINT USING "####.##"; tendered@ - amount@
1270 LOCATE 17, 1: COLOR 11: PRINT "Is the Above Correct (Y/N) .......... "; : fl = 1: GOSUB 62210: IF NOT (in$ = "Y" OR in$ = "N") THEN 1270 ELSE LOCATE 17, 39: PRINT in$
1280 IF in$ = "N" THEN
     FOR j = 13 TO 14: LOCATE j, 1: PRINT STRING$(44, 32): NEXT j
     FOR j = 14 TO 18: LOCATE j, 1: PRINT STRING$(79, 32): NEXT j
     GOTO 1200
     END IF
1285 IF amount@ = owed# THEN Apply!(1) = CVD(IRev1$): Apply!(2) = CVD(IRev2$): Apply!(3) = CVD(IRev3$): Apply!(4) = CVD(IRev4$): Apply!(5) = CVD(IRev5$): Apply!(6) = CVD(IRev6$): Apply!(7) = CVD(irev7$): Apply!(8) = CVD(IREV8$): GOTO 1295
     left@ = amount@
     IF pcode = 4 OR pcode = 5 THEN 1295
1286 FOR j = 12 TO 13: LOCATE j, 1: PRINT STRING$(44, 32): NEXT j
     FOR j = 14 TO 23: LOCATE j, 1: PRINT STRING$(79, 32): NEXT j
     GOSUB openfile
     GET 14, 1
     max = CVI(REVTRACK$)

     LOCATE 12, 1: COLOR 15: PRINT "Please Break Down the Payment"
     COLOR 10: LOCATE 13, 1: PRINT "Amount Paid by Customer: "; USING "$####.##"; amount@
     COLOR 11: LOCATE 14, 1: PRINT "1- Amount to Apply :"; LEFT$(rev1name$, 15)
      IF max < 2 THEN 1288 ELSE LOCATE 15, 1: PRINT "2- Amount to Apply :"; LEFT$(rev2name$, 15)
      IF max < 3 THEN 1288 ELSE LOCATE 16, 1: PRINT "3- Amount to Apply :"; LEFT$(rev3name$, 15)
      IF max < 4 THEN 1288 ELSE LOCATE 17, 1: PRINT "4- Amount to Apply :"; LEFT$(rev4name$, 15)
      IF max < 5 THEN 1288 ELSE LOCATE 18, 1: PRINT "5- Amount to Apply :"; LEFT$(rev5name$, 15)
      IF max < 6 THEN 1288 ELSE LOCATE 19, 1: PRINT "6- Amount to Apply :"; LEFT$(rev6name$, 15)
      IF max < 7 THEN 1288 ELSE LOCATE 20, 1: PRINT "7- Amount to Apply :"; LEFT$(rev7name$, 15)
      IF max < 8 THEN 1288 ELSE LOCATE 21, 1: PRINT "8- Amount to Apply :"; LEFT$(rev8name$, 15)

1288 GOSUB a1
     IF amount@ - left@ = 0 THEN LOCATE 23, 1: COLOR 13: PRINT "Amount has been fully distributed " ELSE LOCATE 23, 1: PRINT STRING$(65, 32)
     LOCATE 23, 1: COLOR 11: PRINT "Selection # (ESC-When Done) ... "; : fl = -1: GOSUB 62000
     IF cf = 1 THEN 1291
     LOCATE 16, 50: PRINT STRING$(28, 32)
     IF VAL(in$) < 1 OR VAL(in$) > 8 THEN 1288
     ON VAL(in$) GOSUB a2, a3, a4, a5, a6, a7, a8, a9: GOTO 1288
a1:
left@ = 0: FOR JJ = 1 TO 8
left@ = left@ + Apply!(JJ)
NEXT JJ: left@ = INT((left@ * 100) + .5) / 100
LOCATE 15, 50: COLOR 10: PRINT "Amount Left to Dist. "; USING "$####.##"; amount@ - left@: RETURN

a2: LOCATE 14, 36: fl = -9: GOSUB 62000: Apply!(1) = VAL(in$): LOCATE 14, 36: PRINT USING "$####.##"; VAL(in$): RETURN
a3: LOCATE 15, 36: fl = -9: GOSUB 62000: Apply!(2) = VAL(in$): LOCATE 15, 36: PRINT USING "$####.##"; VAL(in$): RETURN
a4: LOCATE 16, 36: fl = -9: GOSUB 62000: Apply!(3) = VAL(in$): LOCATE 16, 36: PRINT USING "$####.##"; VAL(in$): RETURN
a5: LOCATE 17, 36: fl = -9: GOSUB 62000: Apply!(4) = VAL(in$): LOCATE 17, 36: PRINT USING "$####.##"; VAL(in$): RETURN
a6: LOCATE 18, 36: fl = -9: GOSUB 62000: Apply!(5) = VAL(in$): LOCATE 18, 36: PRINT USING "$####.##"; VAL(in$): RETURN
a7: LOCATE 19, 36: fl = -9: GOSUB 62000: Apply!(6) = VAL(in$): LOCATE 19, 36: PRINT USING "$####.##"; VAL(in$): RETURN
a8: LOCATE 20, 36: fl = -9: GOSUB 62000: Apply!(7) = VAL(in$): LOCATE 20, 36: PRINT USING "$####.##"; VAL(in$): RETURN
a9: LOCATE 21, 36: fl = -9: GOSUB 62000: Apply!(8) = VAL(in$): LOCATE 21, 36: PRINT USING "$####.##"; VAL(in$): RETURN
1291
1292 IF left@ = amount@ THEN 1295 ELSE LOCATE 16, 50: PRINT CHR$(7); : COLOR 13: PRINT "Not Fully Distributed Yet!": GOTO 1288
1295 GOSUB printutilityreceipt
     d$ = LEFT$(PostDate$, 6) + RIGHT$(PostDate$, 2)
     pdate$ = d$: TM$ = LEFT$(TIME$, 2) + MID$(TIME$, 4, 2)
     postingdate$ = d$
     GOSUB 60000
4081 GET 1, AcctRecord!
      LOCK #1, AcctRecord!
       lockflag = 1

     IF pcode = 3 THEN LSET waterdep$ = MKS$(CVS(waterdep$) - amount@): PUT 1, AcctRecord!
     IF pcode = 4 OR pcode = 5 THEN
       LSET waterdep$ = MKS$(CVS(waterdep$) + amount@)
       LSET depdate$ = pdesc$
       PUT 1, AcctRecord!: GOTO 4100
     END IF

4090 GOSUB 4700
     newbalance@ = CVD(CURBAL$) + CVD(pastdue$)
     newbalance@ = INT((newbalance@ * 100) + .5) / 100
     IF newbalance@ = 0 AND acctstatus$ = "B" THEN LSET acctstatus$ = "I"
     PUT 1, AcctRecord!
4100 GOSUB 4600
4101 GOSUB 61000: REM post to transaction file
     LSET tdate$ = postingdate$
     LSET tdesc$ = pdesc$
     LSET tcuramt$ = MKS$(amount@)
     LSET POSTTIME$ = TM$
     LSET OPERNUMB$ = Operator$
     LSET tglpcode$ = "N"
     LSET tamt$ = MKD$(newbalance@)
     IF pcode = 1 OR pcode = 4 THEN LSET tcustloc$ = "1"
     IF pcode = 2 OR pcode = 5 THEN LSET tcustloc$ = "2"
4105 IF pcode < 3 THEN pcode = 3: GOTO 4110
4106 IF pcode = 3 THEN pcode = 4: GOTO 4110
     IF pcode = 4 OR pcode = 5 THEN pcode = 6: GOTO 4110
     IF pcode = 6 THEN pcode = 5
     IF pcode = 7 THEN pcode = 7
4110 LSET ttype$ = MKI$(pcode)
4120 LSET tmrec$ = MKS$(AcctRecord!)
     LSET tnexttr$ = MKS$(0)
     LSET wbrev1$ = MKS$(Apply!(1)): LSET wbrev2$ = MKS$(Apply!(2))
     LSET wbrev3$ = MKS$(Apply!(3)): LSET wbrev4$ = MKS$(Apply!(4))
     LSET wbrev5$ = MKS$(Apply!(5)): LSET wbrev6$ = MKS$(Apply!(6))
     LSET wbrev7$ = MKS$(Apply!(7)): LSET wbrev8$ = MKS$(Apply!(8))
4130 PUT 2, LOF(2) / 128 + 1: TRECORD! = LOF(2) / 128
     IF CVD(firsttr$) = 0 THEN
      LSET firsttr$ = MKD$(TRECORD!)
       LSET lasttr$ = MKD$(TRECORD!)
        PUT 1, AcctRecord!
         CLOSE 2
         GOTO 4200
      END IF
4150 prev! = CVD(lasttr$): LSET lasttr$ = MKD$(TRECORD!): PUT 1, AcctRecord!
     GET 2, prev!: LSET tnexttr$ = MKS$(TRECORD!): PUT 2, prev!: CLOSE 2
4200 IF lockflag = 1 THEN
      lockflag = 0: UNLOCK #1, AcctRecord!
     END IF
     CLOSE
     GOTO 10

4600 REM reduce individual amount
     GET 1, AcctRecord!
      LSET IRev1$ = MKD$(CVD(IRev1$) - Apply!(1))
      LSET IRev2$ = MKD$(CVD(IRev2$) - Apply!(2))
      LSET IRev3$ = MKD$(CVD(IRev3$) - Apply!(3))
      LSET IRev4$ = MKD$(CVD(IRev4$) - Apply!(4))
      LSET IRev5$ = MKD$(CVD(IRev5$) - Apply!(5))
      LSET IRev6$ = MKD$(CVD(IRev6$) - Apply!(6))
      LSET irev7$ = MKD$(CVD(irev7$) - Apply!(7))
      LSET IREV8$ = MKD$(CVD(IREV8$) - Apply!(8))
     PUT 1, AcctRecord!
     RETURN

4700 IF pcode = 6 THEN
      LSET CURBAL$ = MKD$(CVD(CURBAL$) - amount@)
      RETURN
     END IF
     a4# = amount@: E# = CVD(pastdue$): j# = E# - a4#: y = SGN(j#)
4710 IF y = -1 THEN LSET pastdue$ = MKD$(0): RB# = a4# - E#: GOTO 4740
4720 IF y = 0 THEN LSET pastdue$ = MKD$(0): RETURN
4730 IF y = 1 THEN LSET pastdue$ = MKD$(j#): RETURN
4740 E1# = CVD(CURBAL$): J1# = E1# - RB#: Y1 = SGN(J1#)
4750 IF Y1 = -1 THEN LSET CURBAL$ = MKD$(J1#): RETURN
4760 IF Y1 = 0 THEN LSET CURBAL$ = MKD$(0): RETURN
4770 IF Y1 = 1 THEN LSET CURBAL$ = MKD$(J1#): RETURN



1900 COLOR 11: LOCATE 19, 1: PRINT "Payment Codes Are As Follows: "
1910 COLOR 14: LOCATE 20, 1: PRINT "1= Reg Payment - CASH      2= Reg Payment - CHECK      3= Apply Deposit"
1920 LOCATE 21, 1: PRINT "4= Deposit Payment - CASH  5= Deposit Payment - CHECK  6= ADJUSTMENT"
     LOCATE 22, 1: PRINT "7= Direct Deposit"
1930 RETURN
1950 IF pcode = 1 THEN dd$ = "Reg Payment - CASH"
     IF pcode = 2 THEN dd$ = "Reg Payment - CHECK"
     IF pcode = 3 THEN dd$ = "Apply Deposit"
     IF pcode = 4 THEN dd$ = "Deposit Payment - CASH"
     IF pcode = 5 THEN dd$ = "Deposit Payment - CHECK"
     IF pcode = 6 THEN dd$ = "ADJUSTMENT/CREDIT"
     IF pcode = 7 THEN dd$ = "Direct/Draft"
1960 LOCATE 19, 1: PRINT STRING$(79, 32)
     PRINT STRING$(79, 32): PRINT STRING$(79, 32)
     PRINT STRING$(79, 32); : RETURN




fixamts:
GOSUB openfile
GET 14, 1
max = CVI(REVTRACK$)
fix1:
 LOCATE 15, 1: COLOR 11: PRINT "Enter Item Number to Edit (ESC-When Finished) .. "; : fl = -1: GOSUB 62000
 IF cf = 1 THEN LOCATE 15, 1: PRINT STRING$(79, 32): RETURN
 IF VAL(in$) < 1 OR VAL(in$) > max THEN GOTO fix1
 IF VAL(in$) = 1 THEN LOCATE 6, 63: fl = -10: GOSUB 62000: LSET IRev1$ = MKD$(VAL(in$)): PUT 1, AcctRecord!: GOTO fix1
 IF VAL(in$) = 2 THEN LOCATE 7, 63: fl = -10: GOSUB 62000: LSET IRev2$ = MKD$(VAL(in$)): PUT 1, AcctRecord!: GOTO fix1
 IF VAL(in$) = 3 THEN LOCATE 8, 63: fl = -10: GOSUB 62000: LSET IRev3$ = MKD$(VAL(in$)): PUT 1, AcctRecord!: GOTO fix1
 IF VAL(in$) = 4 THEN LOCATE 9, 63: fl = -10: GOSUB 62000: LSET IRev4$ = MKD$(VAL(in$)): PUT 1, AcctRecord!: GOTO fix1
 IF VAL(in$) = 5 THEN LOCATE 10, 63: fl = -10: GOSUB 62000: LSET IRev5$ = MKD$(VAL(in$)): PUT 1, AcctRecord!: GOTO fix1
 IF VAL(in$) = 6 THEN LOCATE 11, 63: fl = -10: GOSUB 62000: LSET IRev6$ = MKD$(VAL(in$)): PUT 1, AcctRecord!:  GOTO fix1
 IF VAL(in$) = 7 THEN LOCATE 12, 63: fl = -10: GOSUB 62000: LSET irev7$ = MKD$(VAL(in$)): PUT 1, AcctRecord!:  GOTO fix1
 IF VAL(in$) = 8 THEN LOCATE 13, 63: fl = -10: GOSUB 62000: LSET IREV8$ = MKD$(VAL(in$)): PUT 1, AcctRecord!:  GOTO fix1
60000 CLOSE 1: OPEN "wbcust.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #1 LEN = 1024
60010 FIELD 1, 25 AS lname$, 25 AS firstname$, 30 AS add1$, 30 AS add2$, 30 AS servadd$
      FIELD 1, 140 AS du1$, 20 AS CITY$, 2 AS ST$, 9 AS ZP$, 12 AS phone$
      FIELD 1, 183 AS du2$, 2 AS BOOK$, 12 AS meter$, 1 AS billcycle$
      FIELD 1, 198 AS DU3$, 8 AS firsttr$, 8 AS lasttr$, 2 AS CUSTTYPE$
      FIELD 1, 216 AS du4$, 8 AS CURBAL$, 8 AS pastdue$, 1 AS acctstatus$
      FIELD 1, 233 AS du5$, 12 AS ssn$, 12 AS drvlic$
      FIELD 1, 257 AS DU6$, 6 AS WMINSTALL$, 6 AS GMINSTALL$
      FIELD 1, 269 AS du7$, 4 AS wAVGUSE$, 2 AS wAVGCT$, 4 AS gavguse$, 2 AS gavgct$
      FIELD 1, 281 AS du8$, 8 AS wcurread$, 8 AS wpastread$, 6 AS wcurreaddate$
      FIELD 1, 303 AS DU9$, 6 AS wpastreaddate$, 10 AS wmeternum$, 2 AS wminimum$
      FIELD 1, 321 AS du10$, 1 AS wmeterest$, 1 AS gmeterest$
      FIELD 1, 323 AS du11$, 8 AS gcurread$, 8 AS gpastread$, 6 AS gcurreaddate$
      FIELD 1, 345 AS du12$, 6 AS gpastreaddate$, 10 AS gmeternum$, 2 AS gminimum$
      FIELD 1, 363 AS DU13$, 4 AS waterdep$, 4 AS gasdep$
      FIELD 1, 371 AS du14$, 8 AS IRev1$, 8 AS IRev2$, 8 AS IRev3$
      FIELD 1, 395 AS du15$, 8 AS IRev4$, 8 AS IRev5$, 8 AS IRev6$
      FIELD 1, 419 AS DU16$, 8 AS irev7$, 8 AS IREV8$
      FIELD 1, 435 AS du17$, 2 AS rate1$, 2 AS rate2$, 2 AS rate3$
      FIELD 1, 441 AS DU18$, 2 AS rate4$, 2 AS rate5$, 2 AS rate6$
      FIELD 1, 449 AS du19$, 2 AS rate7$, 2 AS rate8$
      FIELD 1, 453 AS DU20$, 1 AS workorder$, 4 AS workordernumber$
      FIELD 1, 458 AS du21$, 6 AS workorderbeg$, 6 AS workorderend$
      FIELD 1, 470 AS DU22$, 4 AS messfirst$, 4 AS messlast$
      FIELD 1, 478 AS du23$, 30 AS comment$, 4 AS seqnum$
      FIELD 1, 512 AS du24$, 4 AS firstwo$, 4 AS lastwo$
      FIELD 1, 520 AS du25$, 6 AS mtrtime$, 8 AS etime$, 3 AS prorate$
      FIELD 1, 537 AS du26$, 4 AS amtowed$, 4 AS amtpd$, 4 AS payment$, 2 AS rs$
      FIELD 1, 551 AS du27$, 4 AS meterseq$, 4 AS garbpickups$, 1 AS WTAXRATE$, 1 AS GTAXRATE$, 1 AS bankdraft$, 462 AS notused$
60030 RETURN

60100 CLOSE 6: OPEN "wbcust1.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #6 LEN = 256
60110 FIELD 6, 4 AS RAMT1$, 4 AS RAMT2$, 4 AS RAMT3$, 25 AS RDESC1$, 25 AS RDESC2$, 25 AS RDESC3$, 9 AS RACCT1$, 9 AS RACCT2$, 9 AS RACCT3$, 4 AS OAMT1$, 4 AS OAMT2$, 25 AS ODESC1$, 25 AS ODESC2$, 9 AS OACCT1$, 9 AS OACCT2$, 66 AS blank$
60120 RETURN
61000 CLOSE 2: OPEN "wbartran.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #2 LEN = 128
61010 FIELD 2, 8 AS tdate$, 8 AS tamt$, 2 AS ttype$, 15 AS tdesc$, 4 AS tmrec$, 4 AS tnexttr$, 4 AS tcuramt$
      FIELD 2, 45 AS tu1$, 4 AS wbrev1$, 4 AS wbrev2$, 4 AS wbrev3$, 4 AS wbrev4$, 4 AS wbrev5$, 4 AS wbrev6$, 4 AS wbrev7$, 4 AS wbrev8$
      FIELD 2, 77 AS tu2$, 1 AS testimate$, 8 AS tcurread$, 8 AS tpastread$
      FIELD 2, 94 AS TU3$, 8 AS tconsump$, 2 AS tbook$, 2 AS tcusttype$, 1 AS tcustloc$: REM tcustloc$ = 1 for cash and 2 for check in payment transaction
      FIELD 2, 107 AS TU4$, 1 AS tstatus$, 2 AS OPERNUMB$, 4 AS POSTTIME$, 1 AS tglpcode$, 13 AS wbblank$
61020 RETURN
61500 CLOSE 15: OPEN "wbmessag.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #15 LEN = 1600
61510 FOR xx = 0 TO 19: FIELD 15, (80 * xx) AS DUM$, 80 AS messline$(xx): NEXT xx: RETURN
      RETURN


62000 COLOR 14: cf = 0: in$ = "": RO = CSRLIN: PO = POS(X): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
62010 IF INKEY$ <> "" THEN 62010
62020 PRINT STRING$(ABS(fl), 95); : LOCATE RO, PO
62030 z$ = INKEY$: IF z$ = "" THEN 62030
62040 IF z$ <> CHR$(13) THEN 62060 ELSE PRINT STRING$(ABS(fl) - ZL, 32);
62050 GOSUB 62200: RETURN
62060 IF z$ = CHR$(27) OR z$ = CHR$(21) THEN cf = 1: RETURN
      IF z$ = CHR$(9) THEN cf = 2: RETURN
      IF LEN(z$) <> 2 THEN 62065
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 59 THEN cf = 11: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 60 THEN cf = 12: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 61 THEN cf = 13: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 62 THEN cf = 14: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 63 THEN cf = 15: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 64 THEN cf = 16: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 65 THEN cf = 17: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 66 THEN cf = 18: RETURN
62065 IF ASC(z$) = 0 THEN 62030
62070 IF LEN(z$) > 1 THEN cf = ASC(RIGHT$(z$, 1)) - 57: IF cf > 0 THEN RETURN ELSE cf = cf + 27: RETURN
62080 IF z$ = " " THEN 62130 ELSE IF z$ = CHR$(29) THEN PRINT STRING$(ZL, 29); : GOTO 62000
62090 IF z$ <> CHR$(8) THEN 62130 ELSE IF ZL = 0 THEN 62030 ELSE PRINT CHR$(29); : PRINT "_"; : PRINT CHR$(29); : IF fl > 0 THEN 62110 ELSE IF z$ = "," THEN 62120
62100 IF z$ = "." THEN ZD = 0: GOTO 62110 ELSE IF z$ = "+" OR z$ = "-" THEN ZS = 0
62110 in$ = LEFT$(in$, LEN(in$) - 1)
62120 ZL = ZL - 1: GOTO 62030
62130 IF ABS(fl) = ZL THEN 62030 ELSE IF fl > 0 THEN IF z$ >= " " AND z$ <= "z" THEN 62180
62140 IF z$ = "." AND INSTR(in$, ".") = 0 THEN GOTO 62180
62150 IF z$ = "," THEN PRINT ","; : ZL = ZL + 1: GOTO 62190
62160 IF (z$ = "-" OR z$ = "+") AND ZS = 0 AND ZL = 0 THEN ZS = 1: GOTO 62180
62170 IF z$ < "0" OR z$ > "9" THEN 62030
62180 PRINT z$; : in$ = in$ + z$: ZL = ZL + 1
62190 IF ABS(fl) = 1 THEN 62050 ELSE 62030
62200 LOCATE RO, PO: PRINT in$: RETURN
62210 GOSUB 62000: IF LEN(in$) THEN in$ = CHR$(ASC(in$) AND 223)
62220 RETURN
63000 cf = 0: in$ = "": COLOR 14: RO = CSRLIN: PO = POS(X): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
63010 IF INKEY$ <> "" THEN 63010
63020 PRINT STRING$(ABS(fl), 95); : LOCATE RO, PO
63030 z$ = INKEY$: IF z$ = "" THEN 63030
63040 IF z$ <> CHR$(13) THEN 63060 ELSE PRINT STRING$(ABS(fl) - ZL, 32);
63050 GOSUB 63200: RETURN
63060 IF z$ = CHR$(27) OR z$ = CHR$(21) THEN cf = 1: RETURN
      IF z$ = CHR$(9) THEN cf = 2: RETURN
63065 IF ASC(z$) = 0 THEN 63030
63070 IF LEN(z$) > 1 THEN cf = ASC(RIGHT$(z$, 1)) - 57: IF cf > 0 THEN RETURN ELSE cf = cf + 27: RETURN
63080 IF z$ = " " THEN 63130 ELSE IF z$ = CHR$(29) THEN PRINT STRING$(ZL, 29); : GOTO 63000
63090 IF z$ <> CHR$(8) THEN 63130 ELSE IF ZL = 0 THEN 63030 ELSE PRINT CHR$(29); : PRINT "_"; : PRINT CHR$(29); : IF fl > 0 THEN 63110 ELSE IF z$ = "," THEN 63120
63100 IF z$ = "." THEN ZD = 0: GOTO 63110 ELSE IF z$ = "+" OR z$ = "-" THEN ZS = 0
63110 in$ = LEFT$(in$, LEN(in$) - 1)
63120 ZL = ZL - 1: GOTO 63030
63130 IF ABS(fl) = ZL THEN 63030 ELSE IF fl > 0 THEN IF z$ >= " " AND z$ <= "z" THEN 63180
63140 IF z$ = "." AND ZD = 0 THEN ZD = 1: GOTO 63180
63150 IF z$ = "," THEN PRINT ","; : ZL = ZL + 1: GOTO 63190
63160 IF (z$ = "-" OR z$ = "+") AND ZS = 0 AND ZL = 0 THEN ZS = 1: GOTO 63180
63170 IF z$ < "0" OR z$ > "9" THEN 63030
63180 PRINT "*"; : in$ = in$ + z$: ZL = ZL + 1
63190 IF ABS(fl) = 1 THEN 63050 ELSE 63030
63200 PRINT "": RETURN

63210 GOSUB 63000: IF LEN(in$) THEN in$ = CHR$(ASC(in$) AND 223)
63220 RETURN

63500 LOCATE 1, 1: COLOR 10: PRINT "Sys. Date: "; SysDate$; TAB(70); Scren$: PRINT STRING$(79, "_"): RETURN

openfile:
CLOSE 14
OPEN "wbinfo.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #14 LEN = 512
FIELD 14, 30 AS dispname$, 20 AS defcity$, 2 AS defst$, 5 AS defzip$
FIELD 14, 57 AS du1$, 2 AS REVTRACK$
FIELD 14, 59 AS du2$, 22 AS rev1name$, 1 AS rev1code$
FIELD 14, 82 AS DU3$, 22 AS rev2name$, 1 AS rev2code$
FIELD 14, 105 AS du4$, 22 AS rev3name$, 1 AS rev3code$
FIELD 14, 128 AS du5$, 22 AS rev4name$, 1 AS rev4code$
FIELD 14, 151 AS DU6$, 22 AS rev5name$, 1 AS rev5code$
FIELD 14, 174 AS du7$, 22 AS rev6name$, 1 AS rev6code$
FIELD 14, 197 AS du8$, 22 AS rev7name$, 1 AS rev7code$
FIELD 14, 220 AS DU9$, 22 AS rev8name$, 1 AS rev8code$
FIELD 14, 243 AS du10$, 269 AS blank$
RETURN


displayhistory:
     dev$ = "scrn:": OPEN "o", 10, dev$: WIDTH #10, 255
     CLS : GOSUB 39000
31200 tr! = CVD(firsttr$): GOSUB 61000
     GOSUB translink
31300 IF counter = 0 THEN 32000
     firstone = counter
31350 FOR kkk = counter TO 1 STEP -1
31400 GET 2, translink!(kkk)
31450 REM
31460 IF lc >= 18 THEN
      firstone = kkk
      LOCATE 22, 1: COLOR 11
      PRINT "Press <F1>-Next Page  :  <F2>-Prev Page (ESC-Quit) .. "; : fl = 1: GOSUB 62000
      IF cf = 1 THEN CLOSE : RETURN
      IF cf = 11 THEN CLS : GOSUB 39000: GOTO 31500

      IF cf = 12 THEN
       counter = firstone + 24
       IF counter > maxcount THEN counter = maxcount
       CLS : GOSUB 39000: GOTO 31350
      END IF
      ELSE GOTO 31500
      END IF

31500 PRINT #10, tdate$; TAB(10); tdesc$; TAB(30);
31600 IF CVI(ttype$) = 1 THEN PRINT #10, "Charge"; : GOTO 31700
31620 IF CVI(ttype$) = 2 THEN PRINT #10, "Late Chrge"; : GOTO 31700
31640 IF CVI(ttype$) = 3 THEN PRINT #10, "Payment"; : GOTO 31700
31660 IF CVI(ttype$) = 4 THEN PRINT #10, "Pay by Dep"; : GOTO 31700
31670 IF CVI(ttype$) = 5 THEN PRINT #10, "Adjustment"; : GOTO 31700
     IF CVI(ttype$) = 6 THEN PRINT #10, "Dep. Pymt"; : GOTO 31700
     IF CVI(ttype$) = 7 THEN PRINT #10, "Dir. Dep."; : GOTO 31700
31680 PRINT #10, "";
31700 PRINT #10, TAB(45); USING "#####.##"; CVS(tcuramt$);
31710 PRINT #10, TAB(60); USING "$$#####,.##"; CVD(tamt$): lc = lc + 1
31720 NEXT kkk
32000 LOCATE 22, 1
      COLOR 11
      PRINT "That's All, Press any key to Continue "; : fl = 1: GOSUB 62000: CLOSE 10: RETURN


translink:
counter = 0
trlink1:
IF tr! = 0 THEN maxcount = counter: RETURN
 counter = counter + 1
 translink!(counter) = tr!: GET 2, tr!
 tr! = CVS(tnexttr$): GOTO trlink1


39000 PRINT #10, TAB(center); TOWN$
     PRINT #10, TAB(24); "Utility Customer Account History"
39100 PRINT #10, "Date: "; DATE$:
     PRINT #10, " Loc # "; m$; TAB(60); "Acct #"; record: PRINT #10, ""
39200 PRINT #10, "Account of ... "; : B = INSTR(firstname$, "  "): IF B = 0 THEN B = 20
39300 PRINT #10, LEFT$(firstname$, B); lname$; TAB(55); "Balance: "; USING "$####.##"; (CVD(CURBAL$) + CVD(pastdue$)): PRINT #10, ""
39400 PRINT #10, "  Date"; TAB(10); "Description"; TAB(30); "Type"; TAB(44); "  Amount"; TAB(63); "Balance": PRINT #10, STRING$(70, "="): lc = 6
     RETURN


namesearch:
LOCATE 3, 1:
COLOR 11: PRINT "Enter Last Name (Partial Accepted : ESC - End) .. "; : fl = 25
 GOSUB 62000: IF cf = 1 THEN GGO = 0: RETURN
 nn$ = in$
2 nn = LEN(nn$)
 CLOSE 4: OPEN "wblname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24: FIELD 4, 20 AS lastnam$, 4 AS mrec$
 CLOSE 1: GOSUB 60000
FOR ll = 1 TO LOF(4) / 24: GET 4, ll
IF LEFT$(lastnam$, nn) = nn$ THEN GOTO n1
n3:
NEXT ll
GGO = 0: RETURN
n1:
GET 1, CVS(mrec$)
LOCATE 6, 1: COLOR 11
kk = INSTR(firstname$, "  ")
IF kk = 0 THEN kk = 25
IF kk = 1 THEN nme$ = lname$ ELSE nme$ = LEFT$(firstname$, kk) + lname$
bb = INSTR(meter$, "/")
IF bb = 0 THEN bb = 9
m$ = LEFT$(meter$, bb - 1)
COLOR 10
PRINT "Full Account Name: "; nme$
PRINT "  Location Number: "; m$
PRINT "   Account Number: "; CVS(mrec$)
ACCTNUMBER$ = m$
PRINT
n2:
LOCATE 9, 1: COLOR 11
PRINT "Correct Account (Y/N/Q) .. "; : fl = 1: GOSUB 62210
IF NOT (in$ = "Y" OR in$ = "N" OR in$ = "Q") THEN GOTO n2
IF in$ = "Q" THEN GGO = 0: RETURN
IF in$ = "Y" THEN GGO = 1: record = CVS(mrec$): RETURN
IF in$ = "N" THEN GOTO n3
GOTO n2

printutilityreceipt:
LOCATE 23, 1: PRINT STRING$(79, 32)
LOCATE 23, 1: PRINT CHR$(7);
LOCATE 23, 1: COLOR 15
PRINT "Print RECEIPT (Y/N) .. "; : fl = 1
GOSUB 62210
IF NOT (in$ = "Y" OR in$ = "N") THEN GOTO printutilityreceipt
IF in$ = "N" THEN RETURN
LPRINT
LPRINT "CITY OF HOHENWALD"
LPRINT "UTILITIES RECEIPT"
LPRINT "    "; DATE$
LPRINT
LPRINT
LPRINT "Account#: "; AcctRecord!
LPRINT "Customer: "; LEFT$(nme$, 30)
LPRINT "Serv. At: "; LEFT$(servadd$, 30)
LPRINT
LPRINT
REM 14 LINES PRINTED TO HERE
IF pcode = 4 OR pcode = 5 THEN
 LPRINT "Amount Paid on"
 LPRINT "Deposit: "; USING "####,#.##"; amount@
 LPRINT
 LPRINT
ELSE
LPRINT "Amount Owed: "; USING "####,#.##"; owed#
IF pcode = 6 THEN
   LPRINT "Amount Adj.: "; USING "####,#.##"; amount@
  ELSE
   LPRINT "Amount Paid: "; USING "####,#.##"; amount@
END IF
LPRINT
  IF (owed# - amount@) >= 0 AND (owed# - amount@ < .01) THEN
    LPRINT "Account Paid in Full"
   ELSE
    LPRINT "Acct Balance: "; USING "####,#.##"; owed# - amount@
  END IF

END IF

LPRINT
LPRINT "LOC # "; LEFT$(meter$, 8)
LPRINT "Operator: "; filen$
LPRINT
LPRINT
LPRINT TAB(7); "T H A N K   Y O U !"
LPRINT
LPRINT
LPRINT
LPRINT
GOTO printutilityreceipt


messagecenter:

REM initialize values
 FOR ll = 1 TO 20: m$(ll) = "": NEXT ll: flag = 0
IF CVS(messfirst$) = 0 THEN
 nomessage = 1
 ELSE
 GOSUB 61500
 rec = CVS(messfirst$)
 GET 15, rec
 FOR ll = 0 TO 19: m$(ll + 1) = messline$(ll): NEXT ll
 flag = 1
END IF

CLS
COLOR 11
PRINT "* * * * * * * * * * * * *  M E S S A G E    C E N T E R    * * * * * * * * *"
FOR X = 1 TO 20
COLOR 11
PRINT USING "##"; X;
PRINT "-";
IF flag = 1 THEN COLOR 15: PRINT LEFT$(m$(X), 76) ELSE PRINT ""
COLOR 11
NEXT X
LOCATE 23, 1: COLOR 11
PRINT "Press <ENTER> to Continue "; : fl = 1: GOSUB 62000
CLOSE 15: RETURN

reviewmessage:
 GOSUB 61500
 rec = CVS(messfirst$)
 IF rec = 0 THEN RETURN
 GET 15, rec
 IF ASC(LEFT$(messline$(0), 1)) > 32 THEN messageflag = 1 ELSE messageflag = 0
 RETURN

OldSetOperatorName:
  Action = 1
  REDIM CMOperRec(1) AS CMOperRecType
  CMOperRecLen = LEN(CMOperRec(1))
  CMFile = FREEFILE
  OPEN "CMOPER.DAT" FOR RANDOM AS CMFile LEN = CMOperRecLen
  GET CMFile, OperRecNumber, CMOperRec(1)
  LOCATE 1, 55: COLOR 15: PRINT "Oper: "; LEFT$(CMOperRec(1).operatorname, 19)
  LOCATE 1, 1: COLOR 11: PRINT "Post Date: "; : COLOR 15: PRINT PostDate$
  Operator = CMOperRec(1).OperatorNumber
  Operator$ = STR$(Operator)
  Operator$ = RIGHT$(Operator$, LEN(Operator$) - 1)
  CLOSE CMFile
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
  Action = 1
  added = 0
  DisplayOperEntry
  LibName$ = "UB.QSL"
  ScrnName$ = "CMCHOSOP"
  help$ = "Cash Management Operator Entry"
  
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET form$(F, 0) = ""
  NEXT F
  
  DO
    
    EditForm form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F7KEY
      AddEditOperator
      added = 1: EXIT DO
      Done = False
    CASE F10Key
      OperNum = VAL(form$(1, 0))
      OperPassword$ = form$(2, 0)
      PostDate$ = form$(3, 0)
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
  LibName$ = "UB.QSL"
  ScrnName$ = "UBCMPOST"

  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)

  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), form$(), ErrCode

  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET form$(F, 0) = ""
  NEXT F

  LibFile2Scrn "UB.QSL", "UBCMPOST", MonoCode%, Attribute%, ErrorCode%
  PrintTitle User$
  PrintHelp help$

  form$(1, 0) = "Y"


  DO

    EditForm form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      IF form$(1, 0) = "Y" THEN
        PrintReceipt
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
  LibName$ = "UB.QSL"
  ScrnName$ = "PRNTJOUR"         'form to edit
  DevFld = 4    'Set to output device field number

  SHARED Choice$()

  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)

  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), form$(), ErrCode

  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET form$(F, 0) = ""
  NEXT F


  REDIM Choice$(0 TO 3, 0 TO 1)
  

  Choice$(0, 0) = "1"  'Choices for Output Order
  Choice$(1, 0) = "Entry Order"
  Choice$(2, 0) = "Name"

  Choice$(0, 1) = STR$(DevFld)  'Choices for Output Device
  Choice$(1, 1) = "Screen"
  Choice$(2, 1) = "Prn#1"
  Choice$(3, 1) = "Prn#2"

  LibFile2Scrn "UB.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  
  Action = 1

  DO

    EditForm form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode  'Check the user's key presses
    
    CASE F10Key 'SaveButton
    
      BegDate = Date2Num(form$(2, 0))
      BegDate$ = form$(2, 0)
      EndDate = Date2Num(form$(3, 0))
      EndDate$ = form$(3, 0)

      IF LEN(LTRIM$(RTRIM$(form$(DevFld, 0)))) = 0 THEN
        dev$ = "S"
        LPTNo = 1
      ELSE
        dev$ = LEFT$(form$(DevFld, 0), 1)
        LPTNo = VAL(RIGHT$(RTRIM$(form$(DevFld, 0)), 1))
      END IF

      IF BegDate > EndDate THEN
        Ok2Go = 0
        Frm(1).FldNo = 5
        Action = 1
      ELSE
        Ok2Go = 1
      END IF

    CASE ESC          'ESC Key
      EXIT SUB

    END SELECT

  LOOP UNTIL Ok2Go

  LOCATE , , 0



  
  IF LEFT$(form$(1, 0), 2) = "En" THEN
   SortOrder$ = "Entry Order"
   ELSE
   SortOrder$ = "Name Order"
  END IF

  ERASE form$, Fld, Choice$, Frm




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

    IF UBTRRec(1).CustAcctNo = 0 AND UBTRRec(1).OperatorNumber = OperRecNumb AND UBTRRec(1).TransType = 2 AND (TransDate >= BegDate AND TransDate <= EndDate) THEN
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
  
  LibName$ = "UB"
  ScrnName$ = "UBCMINPT"
  help$ = "Cash Management System V8.0"
  LOCATE 1, 1, 0
  
  
  ShowCursor
  LibFile2Scrn "UB.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  PrintHelp help$
  
  
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM form$(NumFlds, 2)
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
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET form$(F, 0) = ""
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
    EditForm form$(), Fld(), Frm(1), Cnf, Action
    
    source$ = form$(1, 0)
    IF ASC(LEFT$(source$, 1)) > 32 THEN EXIT SUB
    
    SELECT CASE Frm(1).KeyCode
    CASE F7KEY
      PrintJournal (OperRecNumber)
      Action = 1
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
  LibName$ = "UB"
  ScrnName$ = "UBCMUTIL"
  help$ = "Cash Management Operator Entry"
  LOCATE 1, 1, 0

  help$ = "Utility Payment Entry"
  ShowCursor
  LibFile2Scrn "UB.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  PrintHelp help$

  GOSUB UtilOperatorName


  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)

  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  'Rem Set Choice
  REDIM Choice$(3, 0)
  'Choice$(0, 0) = "5"
  'Choice$(1, 0) = "Cash"
  'Choice$(2, 0) = "Check"
  'Choice$(3, 0) = "Cash & Check"
   
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), form$(), ErrCode

  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET form$(F, 0) = ""
  NEXT F
 
   
  DO
    EditForm form$(), Fld(), Frm(1), Cnf, Action
  

    SELECT CASE Frm(1).KeyCode
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
  Action = 1
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
    RMRec(1).RecName = form$(1, 0)
    RMRec(1).RecAddress = form$(2, 0)
    RMRec(1).RecDesc = form$(3, 0)
    RMRec(1).RecAmtOwed = CVD(MID$(form$(0, 0), Fld(4).Fields, 8))
    IF LEFT$(form$(5, 0), 6) = "Cash  " THEN
     RMRec(1).RecPayType = 1
    END IF
    IF LEFT$(form$(5, 0), 6) = "Check " THEN
     RMRec(1).RecPayType = 2
    END IF
    IF LEFT$(form$(5, 0), 6) = "Cash &" THEN
     RMRec(1).RecPayType = 3
    END IF
    RMRec(1).RecCashAmt = CVD(MID$(form$(0, 0), Fld(6).Fields, 8))
    RMRec(1).RecCheckAmt = CVD(MID$(form$(0, 0), Fld(7).Fields, 8))
    RMRec(1).RecChangeDue = CVD(MID$(form$(0, 0), Fld(9).Fields, 8))
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
  UBTRRec(1).TransDesc = form$(1, 0)
  UBTRRec(1).TransAmt = CVD(MID$(form$(0, 0), Fld(4).Fields, 8))
  UBTRRec(1).RunBalance = 0
  FldFactor = 0
 FOR Cnt = 1 TO 5
  UBTRRec(1).RevAmt(Cnt) = CVD(MID$(form$(0, 0), Fld(Cnt + 10 + FldFactor).Fields, 8))
  FldFactor = FldFactor + 1
 NEXT Cnt
 FOR Cnt = 6 TO 10
  UBTRRec(1).RevAmt(Cnt) = 0
 NEXT Cnt

 FOR Cnt = 1 TO 10
  UBTRRec(1).CustRate(Cnt) = ""
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
UBTRRec(1).CashAmount = CVD(MID$(form$(0, 0), Fld(6).Fields, 8))
UBTRRec(1).CheckAmount = CVD(MID$(form$(0, 0), Fld(7).Fields, 8))
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
    LibName$ = "UB"
    ScrnName$ = "BADOPER"
    LibFile2Scrn "UB.QSL", "BADOPER", MonoCode%, Attribute%, ErrorCode%
    help$ = "Cash Management Operator Entry"
    
    ' Define Fields
    NumFlds = LibNumberOfFields(LibName$, ScrnName$)
    
    ' Define Quick Screen Form Editing Arrays
    REDIM Frm(1) AS FormInfo
    REDIM form$(NumFlds, 2)
    REDIM Fld(NumFlds) AS FieldInfo
    
    ' Get 1st & Last Fields
    StartEl = 0
    LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), form$(), ErrCode
    
    ' Clear Fields
    FOR F = 1 TO NumFlds
      LSET form$(F, 0) = ""
    NEXT F
    
    DO
      EditForm form$(), Fld(), Frm(1), Cnf, Action
      
      
      SELECT CASE Frm(1).KeyCode
        
      CASE ESC
        Done = True
      CASE ELSE
        Done = False
      END SELECT
    LOOP UNTIL Done
    
  END IF
  
END SUB

