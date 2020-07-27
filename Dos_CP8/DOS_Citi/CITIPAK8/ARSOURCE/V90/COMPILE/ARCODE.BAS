DEFINT A-Z
DECLARE SUB ClearBack ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB AddEditCode (EdFlag%)
DECLARE SUB OpenCatCodeFile (NumOfCatRecs%, CatFile%)
DECLARE FUNCTION ubAcctLookUp% (Acct$)
DECLARE FUNCTION ubFmtAcct$ (AcctNum$, FundLen%, AcctLen%, DetLen%)
DECLARE FUNCTION ubStripAcct$ (AcctNum$)
DECLARE FUNCTION ubFindAcct% (AcctNum$)
DECLARE SUB ubGetAcctStruct (GLFundLen%, GLAcctLen%, GLDetLen%)
DECLARE SUB PrintCodes ()
DECLARE SUB addcode ()
DECLARE SUB OpenSetupFile (SetUpFileNum%)
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB PrintHelp (H$)
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
DECLARE SUB AcctPickList (Items() AS ANY, Choice, MaxLen%, BoxBot, ky$, Action, Cnf AS ANY)
DECLARE FUNCTION ASCII (Character$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValI% (Number$)
DECLARE FUNCTION QPValL& (Number$)
DECLARE FUNCTION QPStrL$ (Number&)

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'ARV95.bi'                        'A/R FILE LAYOUTS
  '$INCLUDE: 'GL.bi'
  
  CONST False = 0, True = NOT False
  
  TYPE Struct
    who AS STRING * 14
    RecNum AS INTEGER
  END TYPE
  
  DIM SHARED GLSETUP AS GLSetupRecType
  DIM SHARED Acct AS GLAcctRecType
  DIM SHARED AcctIdx AS GLAcctIndexType
  
  STACK 5000
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 8)

  MChoice$(1) = " Add a new Category Code"
  MChoice$(2) = " Edit an Existing Category "
  MChoice$(3) = " Category Listing Report"
  MChoice$(4) = " Exit to OS"

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
  Row = ((25 - (UBOUND(MChoice$))) \ 2) + 2
  Col = ((80 - MaxLen) \ 2) - 1
  Help$ = "Category Maintenance Menu"

  DO

    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    LibFile2Scrn "AR.QSL", "MENUBAK", MonoCode, -1, ErrorCode

    TitleBox 3, Col, MaxLen + 3, "Category Maintenance Menu ", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

    PrintTitle User$
    PrintHelp Help$

    ShowCursor

    VertMenu MChoice$(), Choice, MaxLen, BoxBot, ky$, Action, Cnf

    IF ky$ = CHR$(27) THEN EXIT DO              'choice = 0

    SELECT CASE Choice
    CASE 1
      AddEditCode False
    CASE 2
      AddEditCode True
    CASE 3
      PrintCodes
    CASE 4
      HideCursor
      CLS
      END
    END SELECT
  LOOP

  IF INSTR(COMMAND$, "TEST") = 0 THEN
    RUN "armenu"
  END IF

  HideCursor
  CLS
  END

SUB AddEditCode (EdFlag)

  SHARED MChoice$
  SHARED Choice$()
  
MainBody:
  size = 1000
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
  
  REDIM Array(1 TO size)  AS Struct
  
  LibName$ = "AR"
  ScrnName$ = "ARSCCODE"
  LOCATE , , 0
  'CursorOff

  Help$ = "Add AR Codes V8.5"
  ShowCursor
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Rem Set Choice
  REDIM Choice$(3, 0)
  Choice$(0, 0) = "2"
  Choice$(1, 0) = "Flat Rate"
  Choice$(2, 0) = "Multiplier"
  Choice$(3, 0) = "Step Rate"
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  Action = 1
  NoProcess = False
  GOSUB LoadGLAcctInfo
  
  REDIM ARCatCodeRec(1) AS ARNewCatCodeRecType
  CatCodeRecLen = LEN(ARCatCodeRec(1))

  IF EdFlag = False THEN
    GOTO JumpInHere
  END IF

  CatFile = FREEFILE
  OPEN "ARCODE.DAT" FOR RANDOM AS CatFile LEN = CatCodeRecLen
  NumOfCatRecs = LOF(CatFile) \ CatCodeRecLen
  
  IF NumOfCatRecs > 0 THEN
    GOSUB SortCatCodes
    REDIM MChoice$(1 TO NumOfCatRecs)
    FOR Cnt = 1 TO NumOfCatRecs
      GET CatFile, Array(Cnt).RecNum, ARCatCodeRec(1)
      MChoice$(Cnt) = SPACE$(40)
      LSET MChoice$(Cnt) = ARCatCodeRec(1).CATCODE
      MID$(MChoice$(Cnt), 7) = ARCatCodeRec(1).CODEDESC
    NEXT Cnt
    
    MaxLen = 50 'Set menu width to zero
    BoxBot = 17 'limit the box length to go no lower than line 20
    Action = 0  '0 means stay in the menu until they select something
    Choice = 1  'Pre-load choice to highlight
    
    TText$ = SPACE$(MaxLen + 4)
    LSET TText$ = " Code   Description"
    
    '--Center Menu within Screen
    Row = 8
    Col = 15
    
    DO
      
      '--Set upper left corner of menu, turn off the cursor
      Help$ = "Add or Edit Catagories"
      
      LOCATE Row, Col, 0
      LibFile2Scrn "AR", "MENUBAK", MonoCode, -1, ErrorCode
      ShowCursor
      QPrintRC TText$, Row - 1, Col, 112
      VertMenu MChoice$(), Choice, MaxLen, BoxBot, ky$, Action, Cnf
      IF ky$ = CHR$(27) THEN
        RecNo = 0
        ExitFlag = True
      ELSE
        RecNo = Array(Choice).RecNum
        ExitFlag = True
      END IF
      
    LOOP UNTIL ExitFlag
  
  END IF
  
  IF RecNo > 0 THEN
    GET CatFile, RecNo, ARCatCodeRec(1)
    Form$(1, 0) = ARCatCodeRec(1).CATCODE
    Form$(2, 0) = ARCatCodeRec(1).CodeType
    Form$(3, 0) = ARCatCodeRec(1).CODEDESC
    Form$(4, 0) = LTRIM$(STR$(ARCatCodeRec(1).Fee))
    
    IF ARCatCodeRec(1).CodeType = "M" THEN
      Form$(2, 0) = "Multiplier"
    END IF
    IF ARCatCodeRec(1).CodeType = "S" THEN
      Form$(2, 0) = "Step Rate"
    END IF
    Form$(5, 0) = STR$(ARCatCodeRec(1).BaseAmt1)
    Form$(6, 0) = STR$(ARCatCodeRec(1).Recpt1)
    Form$(7, 0) = STR$(ARCatCodeRec(1).Percent1)
    Form$(8, 0) = STR$(ARCatCodeRec(1).Maximum1)
    Form$(9, 0) = STR$(ARCatCodeRec(1).BaseAmt2)
    Form$(10, 0) = STR$(ARCatCodeRec(1).Recpt2)
    Form$(11, 0) = STR$(ARCatCodeRec(1).Percent2)
    Form$(12, 0) = STR$(ARCatCodeRec(1).Maximum2)
    Form$(13, 0) = STR$(ARCatCodeRec(1).BaseAmt3)
    Form$(14, 0) = STR$(ARCatCodeRec(1).Recpt3)
    Form$(15, 0) = STR$(ARCatCodeRec(1).Percent3)
    Form$(16, 0) = STR$(ARCatCodeRec(1).Maximum3)
    Form$(17, 0) = STR$(ARCatCodeRec(1).BaseAmt4)
    Form$(18, 0) = STR$(ARCatCodeRec(1).Recpt4)
    Form$(19, 0) = STR$(ARCatCodeRec(1).Percent4)
    Form$(20, 0) = STR$(ARCatCodeRec(1).Maximum4)
    Form$(21, 0) = STR$(ARCatCodeRec(1).BaseAmt5)
    Form$(22, 0) = STR$(ARCatCodeRec(1).Recpt5)
    Form$(23, 0) = STR$(ARCatCodeRec(1).Percent5)
    Form$(24, 0) = STR$(ARCatCodeRec(1).Maximum5)
    Form$(25, 0) = STR$(ARCatCodeRec(1).BaseAmt6)
    Form$(26, 0) = STR$(ARCatCodeRec(1).Recpt6)
    Form$(27, 0) = STR$(ARCatCodeRec(1).Percent6)
    Form$(28, 0) = STR$(ARCatCodeRec(1).Maximum6)
    Form$(29, 0) = STR$(ARCatCodeRec(1).RateStep)
    
    Form$(30, 0) = ARCatCodeRec(1).REVGLNUM
    Form$(31, 0) = ARCatCodeRec(1).ARGLACCT
    Form$(32, 0) = ARCatCodeRec(1).CASHACCT
    
    ARCatCodeRec(1).REVGLNUM = Form$(30, 0)
    ARCatCodeRec(1).ARGLACCT = Form$(31, 0)
    ARCatCodeRec(1).CASHACCT = Form$(32, 0)
  ELSE
    GOTO WereDoneInHere
  END IF
  CLOSE

JumpInHere:
  ShowCursor
  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%
  IF EdFlag THEN
    Help$ = "Editing an Existing Category"
  ELSE
    Help$ = "Adding a New Category"
  END IF
  PrintHelp Help$

  Action = 1
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF Frm(1).FldNo <> Frm(1).PrevFld AND Fld(Frm(1).PrevFld).Protected = 0 THEN
      IF Frm(1).FldNo > 30 THEN
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

      IF RecNo = 0 THEN
        GOSUB CheckCodes
      ELSE
        ValidCode = True
        NoProcess = False
      END IF
      
      IF ValidCode = False THEN
        LOCATE 11, 18
        PRINT "DUPLICATE CODE ENTERED : PLEASE ENTER NEW CODE";
        BEEP
        SLEEP 2
        Frm(1).FldNo = 1
        Action = 1
        LOCATE 11, 18
        PRINT SPACE$(48)
        Done = False
        NoProcess = True
      END IF
      
      IF NoProcess = False THEN
        ARCatCodeRec(1).CATCODE = Form$(1, 0)
        ARCatCodeRec(1).CODEDESC = Form$(3, 0)
        ARCatCodeRec(1).Fee = Value#(Form$(4, 0), ecode)
        ARCatCodeRec(1).CodeType = LEFT$(Form$(2, 0), 1)
        ARCatCodeRec(1).BaseAmt1 = Value#(Form$(5, 0), ecode)
        ARCatCodeRec(1).Recpt1 = Value#(Form$(6, 0), ecode)
        ARCatCodeRec(1).Percent1 = Value#(Form$(7, 0), ecode)
        ARCatCodeRec(1).Maximum1 = Value#(Form$(8, 0), ecode)
        ARCatCodeRec(1).BaseAmt2 = Value#(Form$(9, 0), ecode)
        ARCatCodeRec(1).Recpt2 = Value#(Form$(10, 0), ecode)
        ARCatCodeRec(1).Percent2 = Value#(Form$(11, 0), ecode)
        ARCatCodeRec(1).Maximum2 = Value#(Form$(12, 0), ecode)
        ARCatCodeRec(1).BaseAmt3 = Value#(Form$(13, 0), ecode)
        ARCatCodeRec(1).Recpt3 = Value#(Form$(14, 0), ecode)
        ARCatCodeRec(1).Percent3 = Value#(Form$(15, 0), ecode)
        ARCatCodeRec(1).Maximum3 = Value#(Form$(16, 0), ecode)
        ARCatCodeRec(1).BaseAmt4 = Value#(Form$(17, 0), ecode)
        ARCatCodeRec(1).Recpt4 = Value#(Form$(18, 0), ecode)
        ARCatCodeRec(1).Percent4 = Value#(Form$(19, 0), ecode)
        ARCatCodeRec(1).Maximum4 = Value#(Form$(20, 0), ecode)
        ARCatCodeRec(1).BaseAmt5 = Value#(Form$(21, 0), ecode)
        ARCatCodeRec(1).Recpt5 = Value#(Form$(22, 0), ecode)
        ARCatCodeRec(1).Percent5 = Value#(Form$(23, 0), ecode)
        ARCatCodeRec(1).Maximum5 = Value#(Form$(24, 0), ecode)
        ARCatCodeRec(1).BaseAmt6 = Value#(Form$(25, 0), ecode)
        ARCatCodeRec(1).Recpt6 = Value#(Form$(26, 0), ecode)
        ARCatCodeRec(1).Percent6 = Value#(Form$(27, 0), ecode)
        ARCatCodeRec(1).Maximum6 = Value#(Form$(28, 0), ecode)
        ARCatCodeRec(1).RateStep = Value#(Form$(29, 0), ecode)
        
        ARCatCodeRec(1).REVGLNUM = Form$(30, 0)
        ARCatCodeRec(1).ARGLACCT = Form$(31, 0)
        ARCatCodeRec(1).CASHACCT = Form$(32, 0)

        IF RecNo > 0 THEN
          CLOSE CatFile
          CatFile = FREEFILE
          OPEN "ARCODE.DAT" FOR RANDOM AS CatFile LEN = CatCodeRecLen
          PUT CatFile, RecNo, ARCatCodeRec(1)
        ELSE
          CatFile = FREEFILE
          OPEN "ARCODE.DAT" FOR RANDOM AS CatFile LEN = CatCodeRecLen
          PUT CatFile, (LOF(CatFile) \ CatCodeRecLen) + 1, ARCatCodeRec(1)
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

WereDoneInHere:
  CLOSE
  EXIT SUB
  
CheckCodes:
  ValidCode = True
  CatFile = FREEFILE
  OPEN "ARCODE.DAT" FOR RANDOM AS CatFile LEN = CatCodeRecLen
  NumOfCatRecs = LOF(CatFile) \ CatCodeRecLen
  
  FOR Cnt = 1 TO NumOfCatRecs
    GET CatFile, Cnt, ARCatCodeRec(1)
    IF Form$(1, 0) = ARCatCodeRec(1).CATCODE THEN
      ValidCode = False
      EXIT FOR
    END IF
  NEXT Cnt

  CLOSE CatFile

RETURN
  
SortCatCodes:
  FOR Cnt = 1 TO NumOfCatRecs
    GET CatFile, Cnt, ARCatCodeRec(1)
    Array(Cnt).who = ARCatCodeRec(1).CATCODE + STRING$(11, " ")
    Array(Cnt).RecNum = Cnt
  NEXT Cnt
  SortT Array(Start), NumOfCatRecs, Dir, SSize, MOff, MSize
RETURN
  
LoadGLAcctInfo:
  ubGetAcctStruct GLFundLen, GLAcctLen, GLDetLen
RETURN
  
END SUB

SUB OpenSetupFile (SetUpFileNum)
  GLSetupRecLen = LEN(GLSETUP)
  SetUpFileNum = FREEFILE
  OPEN "GLSETUP.DAT" FOR RANDOM AS SetUpFileNum LEN = GLSetupRecLen
END SUB

SUB PrintCodes

  SHARED Choice$()
  ReportFile$ = "ARCODLST.PRN"  'Report File Name
  FF$ = CHR$(12)
  MaxLines = 50
  LineCnt = 0
  SSize = 16    'total size of each TYPE element
  MOff = 0      'offset into the TYPE for the key element
  MSize = 16    'size of the key element - coded as follows:
  Header$ = "Category Codes Listing"

  ClearBack
  PrintTitle ""
  ShowProcessingScrn Header$
  
  REDIM ARCatCodeRec(1) AS ARNewCatCodeRecType  ' open transaction file
  CatCodeRecLen = LEN(ARCatCodeRec(1))
  TrHandle = FREEFILE
  OPEN "ARCODE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS TrHandle LEN = CatCodeRecLen
  TRNumRecs = LOF(TrHandle) \ CatCodeRecLen

  REDIM Array(1 TO TRNumRecs) AS Struct
  FOR Cnt = 1 TO TRNumRecs
    GET TrHandle, Cnt, ARCatCodeRec(1)
    Array(Cnt).who = LEFT$(ARCatCodeRec(1).CATCODE, 3) + STRING$(11, " ")
    Array(Cnt).RecNum = Cnt
  NEXT
  SortT Array(1), TRNumRecs, 0, SSize, MOff, MSize
  
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  'PRINT #RptHandle, CHR$(27); CHR$(58); ' oki 320 12 cpi
  GOSUB PrintRptHeader
  
  FOR Cnt = 1 TO TRNumRecs
    GET TrHandle, Array(Cnt).RecNum, ARCatCodeRec(1)
    IF LEFT$(ARCatCodeRec(1).CATCODE, 1) <> " " THEN
      IF LEN(QPTrim$(ARCatCodeRec(1).CodeType)) = 0 THEN
        ARCatCodeRec(1).CodeType = "F"
        PUT TrHandle, Array(Cnt).RecNum, ARCatCodeRec(1)
      END IF
      IF LineCnt >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintRptHeader
      END IF
      PRINT #RptHandle, ARCatCodeRec(1).CATCODE;
      PRINT #RptHandle, TAB(8); LEFT$(ARCatCodeRec(1).CODEDESC, 30);
      PRINT #RptHandle, TAB(40); QPTrim$(ARCatCodeRec(1).REVGLNUM);
      PRINT #RptHandle, TAB(55); QPTrim$(ARCatCodeRec(1).ARGLACCT);
      PRINT #RptHandle, TAB(70); QPTrim$(ARCatCodeRec(1).CASHACCT);
      PRINT #RptHandle, TAB(85); USING "$$#####,#.##"; ARCatCodeRec(1).Fee;
      PRINT #RptHandle, TAB(99); ARCatCodeRec(1).CodeType
      IF ARCatCodeRec(1).CodeType = "S" THEN
        PRINT #RptHandle, "Base Amt"; TAB(15); "Receipts Up To"; TAB(40); "  Plus %"; TAB(55); "On Amount Over"
        PRINT #RptHandle, USING "#####.##"; ARCatCodeRec(1).BaseAmt1;
        PRINT #RptHandle, TAB(15); USING "###########,#"; ARCatCodeRec(1).Recpt1;
        PRINT #RptHandle, TAB(40); USING "##.###%"; ARCatCodeRec(1).Percent1;
        PRINT #RptHandle, TAB(55); USING "#########,#"; ARCatCodeRec(1).Maximum1
        PRINT #RptHandle, USING "#####.##"; ARCatCodeRec(1).BaseAmt2;
        PRINT #RptHandle, TAB(15); USING "###########,#"; ARCatCodeRec(1).Recpt2;
        PRINT #RptHandle, TAB(40); USING "##.###%"; ARCatCodeRec(1).Percent2;
        PRINT #RptHandle, TAB(55); USING "#########,#"; ARCatCodeRec(1).Maximum2
        PRINT #RptHandle, USING "#####.##"; ARCatCodeRec(1).BaseAmt3;
        PRINT #RptHandle, TAB(15); USING "###########,#"; ARCatCodeRec(1).Recpt3;
        PRINT #RptHandle, TAB(40); USING "##.###%"; ARCatCodeRec(1).Percent3;
        PRINT #RptHandle, TAB(55); USING "#########,#"; ARCatCodeRec(1).Maximum3
        PRINT #RptHandle, USING "#####.##"; ARCatCodeRec(1).BaseAmt4;
        PRINT #RptHandle, TAB(15); USING "###########,#"; ARCatCodeRec(1).Recpt4;
        PRINT #RptHandle, TAB(40); USING "##.###%"; ARCatCodeRec(1).Percent4;
        PRINT #RptHandle, TAB(55); USING "#########,#"; ARCatCodeRec(1).Maximum4
        PRINT #RptHandle, USING "#####.##"; ARCatCodeRec(1).BaseAmt5;
        PRINT #RptHandle, TAB(15); USING "###########,#"; ARCatCodeRec(1).Recpt5;
        PRINT #RptHandle, TAB(40); USING "##.###%"; ARCatCodeRec(1).Percent5;
        PRINT #RptHandle, TAB(55); USING "#########,#"; ARCatCodeRec(1).Maximum5
        PRINT #RptHandle, USING "#####.##"; ARCatCodeRec(1).BaseAmt6;
        PRINT #RptHandle, TAB(15); USING "###########,#"; ARCatCodeRec(1).Recpt6;
        PRINT #RptHandle, TAB(40); USING "##.###%"; ARCatCodeRec(1).Percent6;
        PRINT #RptHandle, TAB(55); USING "#########,#"; ARCatCodeRec(1).Maximum6
        PRINT #RptHandle, STRING$(79, "-")
        LineCnt = LineCnt + 7
      END IF
      TotalCodes = TotalCodes + 1
      LineCnt = LineCnt + 1
    END IF
    ShowPctComp Cnt, TRNumRecs
  NEXT
  GOSUB PrintRptEnding
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi
  CLOSE         'Close all open files now
  
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  KILL ReportFile$
  
  EXIT SUB

PrintRptHeader:
  page = page + 1
  PRINT #RptHandle, TAB(27); "A/R System : Category Code Listing "
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; page
  PRINT #RptHandle,
  PRINT #RptHandle, "Code "; TAB(8); "Description"; TAB(40); "Rev GL #"; TAB(55); "A/R GL #"; TAB(70); "Cash GL #"; TAB(87); "FEE AMOUNT  TYP"
  PRINT #RptHandle, STRING$(100, "=")
  LineCnt = 5
RETURN
  
PrintRptEnding:
  PRINT #RptHandle, STRING$(100, "-")
  PRINT #RptHandle, "Number of Codes .. "; USING "####,#"; TotalCodes
  PRINT #RptHandle, FF$
RETURN
  
END SUB

