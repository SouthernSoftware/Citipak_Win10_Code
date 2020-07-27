DEFINT A-Z
DECLARE SUB WaitForAction ()
DECLARE SUB DisplayDCScrn (ScrnName$)
DECLARE SUB ubGetAcctStruct (FundLen%, AcctLen%, DetLen%)
DECLARE SUB SaveScrn (ScrnArray())
DECLARE SUB RestScrn (ScrnArray())
DECLARE SUB ubOpenAcctFile (AcctFileNum%, NumAccts%)
DECLARE SUB ubOpenAcctIdx (AcctIdxFileNum%, NumIdxRecs%)
DECLARE SUB PrintCodes ()
DECLARE SUB addcode (EditFlag)
DECLARE SUB OpenSetupFile (SetUpFileNum%)
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB ShowCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB QPrint (x$, Colr%, page%)
DECLARE SUB QPrintRC (T$, r%, c%, clr%)
DECLARE SUB OpenAcctIdx (AcctIdxFileNum%, NumIdxRecs%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB HideCursor ()
  
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION ubGetAcctTitle$ (RecordNumber%)
DECLARE FUNCTION ubAcctLookUp% (Acct$)
DECLARE FUNCTION ubFindAcct% (AcctNum$)
DECLARE FUNCTION ubFmtAcct$ (AN$, FundLen%, AcctLen%, DetLen%)
DECLARE FUNCTION ubStripAcct$ (AcctNum$)
DECLARE FUNCTION QPTrim$ (x$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION FindAcct% (AcctNum$)
DECLARE FUNCTION QPTrim$ (x$)
  
  TYPE Struct
    who AS STRING * 14
    RecNum AS INTEGER
  END TYPE
  
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'DC.bi'                        'A/R FILE LAYOUTS
  '$INCLUDE: 'GL.bi'
  
  
  CONST False = 0, True = NOT False
  
  
  DIM SHARED GLSETUP AS GLSetupRecType
  DIM SHARED Acct AS GLAcctRecType
  DIM SHARED AcctIdx AS GLAcctIndexType
  
  STACK 8000

  REDIM MChoice$(1 TO 4)

  MChoice$(1) = "Add a New Decal Code"
  MChoice$(2) = "Edit an Existing Decal Code "
  MChoice$(3) = "Print Decal Code Listing"
  MChoice$(4) = "Exit to OS"

  MaxLen = 0    'Set menu width to zero
  BoxBot = 17   'limit the box length to go no lower than line 20
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight

  '--Find max menu width
  FOR Cnt = 1 TO UBOUND(MChoice$)
    TLen = LEN(MChoice$(Cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT

  '--Center Menu within Screen
  Row = ((25 - (UBOUND(MChoice$))) \ 2) - 1
  Col = ((80 - MaxLen) \ 2) - 1

  Help$ = "Vehicle Decals"

  DO

    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    LibFile2Scrn "DC.QSL", "MENUBAK", MonoCode, -1, ErrorCode

    TitleBox 3, Col, MaxLen + 3, "Decal Code Maintance", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

    ShowCursor

    VertMenu MChoice$(), Choice, MaxLen, BoxBot, ky$, Action, Cnf

    IF ky$ = CHR$(27) THEN EXIT DO              'choice = 0

    SELECT CASE Choice
    CASE 1
      addcode False
    CASE 2
      addcode True
    CASE 3
      PrintCodes
    CASE 4
      HideCursor
      CLS
      END
    END SELECT
  LOOP
'  IF Exist("CPMENU.EXE") THEN
'    RUN "cpmenu"
'  END IF

'  addcode
  RUN "dcmenu"

SUB addcode (EditFlag)

  SHARED MChoice$

  ubGetAcctStruct GLFundLen%, GLAcctLen%, GLDetLen%

MainBody:
  size = 500
  Start = 1     'start at array element 1
  Dir = 0       'sort direction - use anything else for descending
  SSize = 16    'total size of each TYPE element
  MOff = 0      'offset into the TYPE for the key element
  MSize = 5     'size of the key element - coded as follows:
  
  REDIM array(1 TO size)  AS Struct
  
  LibName$ = "DC"
  ScrnName$ = "DCCODE"
  
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
  FirstTime = True
  
  REDIM DCCatCodeRec(1) AS DCCatCodeRecType
  CatCodeRecLen = LEN(DCCatCodeRec(1))
  CatFile = FREEFILE
  OPEN "DCCODE.DAT" FOR RANDOM AS CatFile LEN = CatCodeRecLen
  NumOfCatRecs = LOF(CatFile) \ CatCodeRecLen

  IF EditFlag AND NumOfCatRecs > 0 THEN
    GOSUB SelectCatCode
  ELSEIF EditFlag AND NumOfCatRecs <= 0 THEN
    GOTO DoneInHere
  END IF

  IF EditFlag AND RecNo <= 0 THEN
    GOTO DoneInHere
  END IF

  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%
  
  Action = 1
  
  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      IF RecNo > 0 THEN
        GET CatFile, RecNo, DCCatCodeRec(1)
        LSET Form$(1, 0) = DCCatCodeRec(1).CATCODE
        LSET Form$(2, 0) = DCCatCodeRec(1).CODEDESC
        LSET Form$(3, 0) = DCCatCodeRec(1).REVGLNUM
        LSET Form$(4, 0) = DCCatCodeRec(1).CASHACCT
        LSET Form$(5, 0) = STR$(DCCatCodeRec(1).Fee)
        Action = 1
      END IF
    END IF

    IF Frm(1).FldNo <> Frm(1).PrevFld AND Fld(Frm(1).PrevFld).Protected = 0 THEN
      IF Frm(1).FldNo > 3 AND Frm(1).FldNo < 6 THEN
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
      
    CASE F10Key
      
      IF LEFT$(Form$(1, 0), 2) = "  " THEN EXIT SUB
      ' check for duplicate account number
      IF RecNo = 0 THEN
        GOSUB CheckCodes
      ELSE
        ValidCode = True
        NoProcess = False
        AcctRecNo = True
      END IF
      
      IF ValidCode = False THEN
        QPrintRC "DUPLICATE CODE ENTERED : PLEASE ENTER NEW CODE", 14, 18, 14
        SLEEP 1
        Frm(1).FldNo = 1
        Action = 1
        QPrintRC SPACE$(48), 14, 18, 7
        Done = False
        NoProcess = True
      END IF
      
      IF NoProcess = False THEN
        DCCatCodeRec(1).CATCODE = QPTrim$(Form$(1, 0))
        DCCatCodeRec(1).CODEDESC = QPTrim$(Form$(2, 0))
        DCCatCodeRec(1).REVGLNUM = QPTrim$(Form$(3, 0))
        DCCatCodeRec(1).CASHACCT = QPTrim$(Form$(4, 0))
        DCCatCodeRec(1).Fee = VAL(Form$(5, 0))
        
        IF RecNo > 0 THEN
          PUT CatFile, RecNo, DCCatCodeRec(1)
        ELSE
          PUT CatFile, LOF(CatFile) \ CatCodeRecLen + 1, DCCatCodeRec(1)
        END IF
        CLOSE CatFile
        Done = True
        DisplayDCScrn "UPDATEOK"
        WaitForAction
        'GOTO MainBody
      END IF
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done

DoneInHere:
  CLOSE
  EXIT SUB

CheckCodes:
  ValidCode = True 'assume this code is ok
  NewCode$ = QPTrim$(Form$(1, 0))
  FOR Cnt = 1 TO NumOfCatRecs
    GET CatFile, Cnt, DCCatCodeRec(1)
    ThisCode$ = QPTrim$(DCCatCodeRec(1).CATCODE)
    IF ThisCode$ = NewCode$ THEN
      ValidCode = False
      EXIT FOR
    END IF
  NEXT Cnt
RETURN
  
SortCatCodes:
  FOR Cnt = 1 TO NumOfCatRecs
    GET CatFile, Cnt, DCCatCodeRec(1)
    LSET array(Cnt).who = DCCatCodeRec(1).CATCODE
    array(Cnt).RecNum = Cnt
  NEXT Cnt
  SortT array(Start), NumOfCatRecs, Dir, SSize, MOff, MSize
RETURN

SelectCatCode:

  IF NumOfCatRecs > 0 THEN
    GOSUB SortCatCodes
    REDIM MChoice$(1 TO NumOfCatRecs)
    FOR Cnt = 1 TO NumOfCatRecs
      GET CatFile, array(Cnt).RecNum, DCCatCodeRec(1)
      MChoice$(Cnt) = SPACE$(50)
      LSET MChoice$(Cnt) = DCCatCodeRec(1).CATCODE
      MID$(MChoice$(Cnt), 5) = DCCatCodeRec(1).CODEDESC
    NEXT Cnt

    MaxLen = 50 'Set menu width to zero
    BoxBot = 17 'limit the box length to go no lower than line 20
    Action = 0  '0 means stay in the menu until they select something
    Choice = 1  'Pre-load choice to highlight

    TText$ = SPACE$(MaxLen + 4)
    LSET TText$ = " Code   Description                                "

    '--Center Menu within Screen
    Row = 8
    Col = 15


    DO

      '--Set upper left corner of menu, turn off the cursor
      LOCATE Row, Col, 0
      LibFile2Scrn "DC", "MENUBAK", MonoCode, -1, ErrorCode
      QPrintRC TText$, Row - 1, Col, 112
      VertMenu MChoice$(), Choice, MaxLen, BoxBot, ky$, Action, Cnf
      IF ky$ = CHR$(27) THEN
        RecNo = 0
        ExitFlag = True
      ELSE
        RecNo = array(Choice).RecNum
        ExitFlag = True
      END IF
    LOOP UNTIL ExitFlag
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
FUNCTION FindAcct% (AcctNum$)
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

SUB OpenAcctIdx (AcctIdxFileNum, NumIdxRecs)
  
  AcctIdxLen = LEN(AcctIdx)
  AcctIdxFileNum = FREEFILE
  OPEN "GLACCT.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS AcctIdxFileNum LEN = AcctIdxLen
  NumIdxRecs = LOF(AcctIdxFileNum) \ AcctIdxLen
  
END SUB

SUB OpenCatCodeFile (NumOfCatRecs, CatFile)
  REDIM DCCatCodeRec(1) AS DCCatCodeRecType
  CatCodeRecLen = LEN(DCCatCodeRec(1))
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
  
  ReportFile$ = "DCCODLST.PRN"  'Report File Name
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
  Start = 1     'start at array element 1
  Dir = 0       'sort direction - use anything else for descending
  SSize = 16    'total size of each TYPE element
  MOff = 0      'offset into the TYPE for the key element
  MSize = 16    'size of the key element - coded as follows:
  '   -1 = integer
  '   -2 = long integer
  '   -3 = single precision
  '   -4 = double precision
  '   +N = TYPE array/fixed-length string of length N
  
  REDIM array(1 TO size) AS Struct
  
  GOSUB SelectOutput: IF Canceled$ = "Y" THEN EXIT SUB
  
  REDIM DCCatCodeRec(1) AS DCCatCodeRecType     ' open transaction file
  CatCodeRecLen = LEN(DCCatCodeRec(1))
  TrHandle = FREEFILE
  OPEN "DCCODE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS TrHandle LEN = CatCodeRecLen
  TrNumRecs = LOF(TrHandle) \ CatCodeRecLen
  
  
  GOSUB GetReportInformation
  
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  
  'PRINT #RptHandle, CHR$(27); CHR$(58); ' oki 320 12 cpi
  
  GOSUB PrintRptHeader
  
  FOR Cnt = 1 TO COUNT
    GET TrHandle, array(Cnt).RecNum, DCCatCodeRec(1)
    IF LEFT$(DCCatCodeRec(1).CATCODE, 1) <> " " THEN
      IF LineCnt >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintRptHeader
      END IF
      PRINT #RptHandle, DCCatCodeRec(1).CATCODE; TAB(8); LEFT$(DCCatCodeRec(1).CODEDESC, 28);
      PRINT #RptHandle, TAB(40); USING "####.##"; DCCatCodeRec(1).Fee;
      PRINT #RptHandle, TAB(50); DCCatCodeRec(1).CASHACCT;
      PRINT #RptHandle, TAB(66); DCCatCodeRec(1).REVGLNUM
      TotalCodes = TotalCodes + 1
      LineCnt = LineCnt + 1
    END IF
  NEXT Cnt
  GOSUB PrintRptEnding
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi
  CLOSE         'Close all open files now
  
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
  PRINT #RptHandle, TAB(18); "Virginia Decal System Catagory Code Listing "
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; page
  PRINT #RptHandle, "Code "; TAB(8); "Description"; TAB(42); "Fee"; TAB(50); "Cash GL# (dr)"; TAB(66); "Rev GL# (cr)"
  PRINT #RptHandle, STRING$(80, "=")
  LineCnt = 5
  RETURN
  
PrintRptEnding:
  PRINT #RptHandle, STRING$(80, "-")
  PRINT #RptHandle, "Number of Codes .. "; USING "####,#"; TotalCodes
  PRINT #RptHandle, FF$
  RETURN
  
GetReportInformation:
  
  FOR Cnt! = 1 TO TrNumRecs
    GET TrHandle, Cnt!, DCCatCodeRec(1)
    array(Cnt!).who = LEFT$(DCCatCodeRec(1).CATCODE, 3) + STRING$(11, " ")
    array(Cnt!).RecNum = Cnt!
  NEXT Cnt!
  
  COUNT = TrNumRecs
  
  SortT array(Start), COUNT, Dir, SSize, MOff, MSize
  RETURN
  
SelectOutput:
  LibName$ = "DC"
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
  
  REDIM Choice$(2, 0)
  
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

