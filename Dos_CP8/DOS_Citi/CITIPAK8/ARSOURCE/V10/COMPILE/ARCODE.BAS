DEFINT A-Z
DECLARE FUNCTION ubAcctLookUp% (Acct$)
DECLARE FUNCTION ubFmtAcct$ (AcctNum$, FundLen%, AcctLen%, DetLen%)
DECLARE FUNCTION ubStripAcct$ (AcctNum$)
DECLARE FUNCTION ubFindAcct% (AcctNum$)
DECLARE SUB ubGetAcctStruct (GLFundLen%, GLAcctLen%, GLDetLen%)
DECLARE SUB PrintCodes ()
DECLARE SUB addcode ()
DECLARE SUB OpenSetupFile (SetUpFileNum%)
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB printhelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB ShowCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB QPrint (X$, Colr%, page%)
DECLARE SUB QPrintRC (T$, r%, c%, clr%)
DECLARE SUB OpenAcctIdx (AcctIdxFileNum%, NumIdxRecs%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB HideCursor ()

DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION FindAcct% (AcctNum$)
DECLARE FUNCTION QPTrim$ (X$)


  '$INCLUDE: 'DefCnf.BI'
DECLARE SUB AcctPickList (Items() AS ANY, Choice, MaxLen%, BoxBot, Ky$, Action, Cnf AS Config)
DECLARE FUNCTION ASCII (Character$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValI% (Number$)
DECLARE FUNCTION QPValL& (Number$)
DECLARE FUNCTION QPStrL$ (Number&)

  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'AR.bi'                        'A/R FILE LAYOUTS
  '$INCLUDE: 'GL.bi'


  CONST False = 0, True = NOT False

TYPE Struct
 who AS STRING * 14
 RecNum AS INTEGER
END TYPE


DIM SHARED GLSETUP AS GLSetupRecType
DIM SHARED Acct AS GLAcctRecType
DIM SHARED AcctIdx AS GLAcctIndexType


 STACK 8000

   addcode

'   END

   RUN "armenu"

  SUB addcode
  SHARED MChoice$


MainBody:
  size = 400
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

  LibName$ = "AR"
  ScrnName$ = "ARCODE"

  LOCATE 1, 1, 0
  help$ = "Add AR Codes V8.0"
  ShowCursor

  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%
  'printhelp help$
  
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
        
  REDIM ARCatCodeRec(1) AS ARCatCodeRecType
  CatCodeRecLen = LEN(ARCatCodeRec(1))
  CatFile = FREEFILE
  OPEN "ARCODE.DAT" FOR RANDOM AS CatFile LEN = CatCodeRecLen
  NumOfCatRecs = LOF(CatFile) \ CatCodeRecLen

  IF NumOfCatRecs > 0 THEN
    GOSUB SortCatCodes
    REDIM MChoice$(1 TO NumOfCatRecs)
  FOR Cnt = 1 TO NumOfCatRecs
    GET CatFile, array(Cnt).RecNum, ARCatCodeRec(1)
    MChoice$(Cnt) = SPACE$(40)
    LSET MChoice$(Cnt) = ARCatCodeRec(1).CATCODE
    MID$(MChoice$(Cnt), 5) = ARCatCodeRec(1).CODEDESC
  NEXT Cnt

   MaxLen = 50     'Set menu width to zero
   BoxBot = 17    'limit the box length to go no lower than line 20
   Action = 0     '0 means stay in the menu until they select something
   Choice = 1     'Pre-load choice to highlight

   TText$ = SPACE$(MaxLen + 4)
   LSET TText$ = " Code   Description      ** Press <ESC> to Add Code"

   '--Center Menu within Screen
   Row = 8
   Col = 15

   DO

      '--Set upper left corner of menu, turn off the cursor
      help$ = "Add or Edit Catagories"

      LOCATE Row, Col, 0
      LibFile2Scrn "AR", "MENUBAK", MonoCode, -1, ErrorCode
            'printhelp help$
      ShowCursor
      QPrintRC TText$, Row - 1, Col, 112
      VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
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
     GET CatFile, RecNo, ARCatCodeRec(1)
     Form$(1, 0) = ARCatCodeRec(1).CATCODE
     Form$(2, 0) = ARCatCodeRec(1).CODEDESC
     Form$(3, 0) = ARCatCodeRec(1).REVGLNUM
     Form$(4, 0) = ARCatCodeRec(1).ARGLACCT
     Form$(5, 0) = ARCatCodeRec(1).CashAcct
  END IF

  ShowCursor
  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%
  'printhelp help$

 GOSUB LoadGLAcctInfo
 Action = 1

  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action

      IF Frm(1).FldNo <> Frm(1).PrevFld AND Fld(Frm(1).PrevFld).Protected = 0 THEN
       IF Frm(1).FldNo > 3 THEN
             TempAcct$ = ubStripAcct$(Form$(Frm(1).PrevFld, 0))
        IF LEN(TempAcct$) > 0 THEN
             TempAcct$ = ubFmtAcct$(TempAcct$, GLFundLen, GLAcctLen, GLDetLen)
             AcctOK = ubAcctLookUp(TempAcct$)
            IF AcctOK > 0 THEN
              LSET Form$(Frm(1).PrevFld, 0) = TempAcct$
              Action = 1
            ELSE

              Frm(1).FldNo = Frm(1).PrevFld
              LSET Form$(Frm(1).PrevFld, 0) = ""
              Action = 1
            END IF
         END IF
      END IF
      END IF



    SELECT CASE Frm(1).KeyCode


    CASE F7KEY
     CLOSE CatFile
     PrintCodes
     EXIT SUB

    CASE F10Key

      IF LEFT$(Form$(1, 0), 2) = "  " THEN EXIT SUB
      
      IF RecNo = 0 THEN
       GOSUB CheckCodes
       ELSE
       ValidCode = True
       NoProcess = False
      END IF

      IF ValidCode = False THEN
       LOCATE 14, 18: PRINT "DUPLICATE CODE ENTERED : PLEASE ENTER NEW CODE";
       PRINT CHR$(7);
       SLEEP 1
       Frm(1).FldNo = 1
       Action = 1
       LOCATE 14, 18: PRINT "                                               "
       Done = False
       NoProcess = True
      END IF



     IF NoProcess = False THEN
      ARCatCodeRec(1).CATCODE = Form$(1, 0)
      ARCatCodeRec(1).CODEDESC = Form$(2, 0)
      ARCatCodeRec(1).APPNUMB = 1
      ARCatCodeRec(1).BILLCODE = 1
      ARCatCodeRec(1).REVGLNUM = Form$(3, 0)
      ARCatCodeRec(1).ARGLACCT = Form$(4, 0)
      ARCatCodeRec(1).CashAcct = Form$(5, 0)
        CatFile = FREEFILE
        OPEN "ARCODE.DAT" FOR RANDOM AS CatFile LEN = CatCodeRecLen
      
    
         IF RecNo > 0 THEN
           PUT CatFile, RecNo, ARCatCodeRec(1)
         ELSE
           PUT CatFile, LOF(CatFile) \ CatCodeRecLen + 1, ARCatCodeRec(1)
         END IF
         CLOSE CatFile
         GOTO MainBody
       END IF
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT

  LOOP UNTIL Done

  EXIT SUB



CheckCodes:
 CatFile = FREEFILE
 OPEN "ARCODE.DAT" FOR RANDOM AS CatFile LEN = CatCodeRecLen

 FOR Cnt = 1 TO NumOfCatRecs
 GET CatFile, Cnt, ARCatCodeRec(1)
 IF Form$(1, 0) = ARCatCodeRec(1).CATCODE THEN ValidCode = False: RETURN
 NEXT Cnt
 ValidCode = True
 CLOSE
 RETURN

SortCatCodes:
 FOR Cnt = 1 TO NumOfCatRecs
    GET CatFile, Cnt, ARCatCodeRec(1)
    array(Cnt).who = ARCatCodeRec(1).CATCODE + STRING$(11, " ")
    array(Cnt).RecNum = Cnt
 NEXT Cnt
 SortT array(Start), NumOfCatRecs, Dir, SSize, MOff, MSize
RETURN

LoadGLAcctInfo:
 ubGetAcctStruct GLFundLen, GLAcctLen, GLDetLen
 RETURN



END SUB

SUB OpenCatCodeFile (NumOfCatRecs, CatFile)
  REDIM ARCatCodeRec(1) AS ARCatCodeRecType
  CatCodeRecLen = LEN(ARCatCodeRec(1))
  CatFile = FREEFILE
  OPEN "CatCode.DAT" FOR RANDOM AS CatFile LEN = CatCodeRecLen
  NumOfCatRecs = LOF(CatFile) \ CatCodeRecLen
  END SUB

SUB OpenSetupFile (SetUpFileNum)
   GLSetupRecLen = LEN(GLSETUP)
   SetUpFileNum = FREEFILE
   OPEN "GLSETUP.DAT" FOR RANDOM AS SetUpFileNum LEN = GLSetupRecLen
END SUB

SUB PrintCodes
  SHARED Choice$()


  ReportFile$ = "ARCODLST.PRN"  'Report File Name
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

  GOSUB SelectOutput: IF Canceled$ = "Y" THEN EXIT SUB

  REDIM ARCatCodeRec(1) AS ARCatCodeRecType     ' open transaction file
  CatCodeRecLen = LEN(ARCatCodeRec(1))
  TrHandle = FREEFILE
  OPEN "ARCODE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS TrHandle LEN = CatCodeRecLen
  TrNumRecs = LOF(TrHandle) \ CatCodeRecLen


  GOSUB GetReportInformation

  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle


  'PRINT #RptHandle, CHR$(27); CHR$(58); ' oki 320 12 cpi

  GOSUB PrintRptHeader

  FOR Cnt = 1 TO COUNT
   GET TrHandle, array(Cnt).RecNum, ARCatCodeRec(1)
   IF LEFT$(ARCatCodeRec(1).CATCODE, 1) <> " " THEN
   IF LineCnt >= MaxLines THEN
    PRINT #RptHandle, FF$
    GOSUB PrintRptHeader
   END IF
   PRINT #RptHandle, ARCatCodeRec(1).CATCODE; TAB(8); LEFT$(ARCatCodeRec(1).CODEDESC, 30);
     PRINT #RptHandle, TAB(40); QPTrim$(ARCatCodeRec(1).REVGLNUM);
     PRINT #RptHandle, TAB(55); QPTrim$(ARCatCodeRec(1).ARGLACCT);
     PRINT #RptHandle, TAB(70); QPTrim$(ARCatCodeRec(1).CashAcct)
    TotalCodes = TotalCodes + 1
    LineCnt = LineCnt + 1
   END IF
  NEXT Cnt
    GOSUB PrintRptEnding
    PRINT #RptHandle, CHR$(18); ' oki 320 10 cpi
    CLOSE                       'Close all open files now

  IF DevSpec$ = "S" THEN
       EntryPoint = 2
       ELSE
       EntryPoint = 5
  END IF

  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint

  KILL ReportFile$

  EXIT SUB


PrintRptHeader:
    page = page + 1
    PRINT #RptHandle, TAB(27); "A/R System : Catagory Code Listing "
    PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; page
    PRINT #RptHandle,
    PRINT #RptHandle, "Code "; TAB(8); "Description"; TAB(40); "Rev GL #"; TAB(55); "A/R GL #"; TAB(70); "Cash GL #"
    PRINT #RptHandle, STRING$(85, "=")
    LineCnt = 5
    RETURN

PrintRptEnding:
     PRINT #RptHandle, STRING$(80, "-")
     PRINT #RptHandle, "Number of Codes .. "; USING "####,#"; TotalCodes
     PRINT #RptHandle, FF$
     RETURN

GetReportInformation:
    
  FOR Cnt! = 1 TO TrNumRecs
      GET TrHandle, Cnt!, ARCatCodeRec(1)
      array(Cnt!).who = LEFT$(ARCatCodeRec(1).CATCODE, 3) + STRING$(11, " ")
      array(Cnt!).RecNum = Cnt!
  NEXT Cnt!

  COUNT = TrNumRecs

SortT array(Start), COUNT, Dir, SSize, MOff, MSize
RETURN

SelectOutput:
  LibName$ = "AR"
  ScrnName$ = "WHERPRNT"

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

  DIM Choice$(2, 0)

  Choice$(0, 0) = "1"
  Choice$(1, 0) = "SCREEN"
  Choice$(2, 0) = "PRINTER"


  Action = 1
  ShowCursor
  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%
  'printhelp help$
  Action = 1

  DO


    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode
     CASE F10Key
       DevSpec$ = LEFT$(Form$(1, 0), 1)
       RETURN
     CASE EscKey
      Canceled$ = "Y"
      RETURN
    END SELECT
 LOOP
  RETURN

END SUB

