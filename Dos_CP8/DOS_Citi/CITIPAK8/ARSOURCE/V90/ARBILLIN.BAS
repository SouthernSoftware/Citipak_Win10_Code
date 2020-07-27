DEFINT A-Z
DECLARE SUB AppListing ()
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ARPenaltyProc ()
DECLARE SUB CreatePenalty ()
DECLARE SUB PrintPenRpt ()
DECLARE SUB PostPenalty ()
DECLARE FUNCTION MakeDay$ (DayNum%)
DECLARE FUNCTION MakeLongDate$ (PDate$)
DECLARE FUNCTION MakeMonth$ (TDate$)
DECLARE FUNCTION Num2Day% (DayNum%)
DECLARE SUB PrintLateNotice ()
DECLARE SUB SETFEE ()
DECLARE SUB PrintLicenses ()
DECLARE SUB LicenseRegister ()
DECLARE SUB OpenARCustFile (NumOfArRecs%, ARFile%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB QPrint (X$, Colr%, Page%)
DECLARE SUB QPrintRC (T$, R%, C%, clr%)
DECLARE SUB ARAPPLICATION ()
DECLARE SUB ShowNoCode ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB HideCursor ()
DECLARE SUB ShowCursor ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB WaitForAction ()
DECLARE SUB SaveScrn (TempScrn())
DECLARE SUB RestScrn (TempScrn())
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE FUNCTION Round# (DlbNum#)
DECLARE FUNCTION QPStrI$ (A%)
  
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'ARV95.bi'                        'A/R FILE LAYOUTS
  '$INCLUDE: 'GL.bi'
  
  CONST False = 0, True = NOT False
  
  DIM SHARED ARCustRec(1) AS ARCustRecType
  DIM SHARED ARCustIdxRec(1) AS ARCustIDXRecType

  REDIM SHARED PctC(1) AS STRING * 3

  TYPE Struct
    who AS STRING * 14
    RecNum AS INTEGER
  END TYPE
  
  STACK 8000
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 7)
  'Mchoice$(1) = "Set Fee's on Accounts"
  MChoice$(1) = " Penalty Processing"
  MChoice$(2) = " Application Listing"
  MChoice$(3) = " Print Applications/Renewals"
  MChoice$(4) = " Licenses to Print Register"
  MChoice$(5) = " Print License Forms"
  MChoice$(6) = " Print Delinquent Notices"
  MChoice$(7) = " Exit to OS"
  
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
  Row = ((25 - (UBOUND(MChoice$))) \ 2)
  Col = ((80 - MaxLen) \ 2) - 1
  Help$ = "Billing Menu"

'  SETFEE

  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    'DisplayARScrn "MENUBAK"
    LibFile2Scrn "AR.QSL", "MENUBAK", MonoCode, -1, ErrorCode
    
    TitleBox 3, Col, MaxLen + 3, "Issue License ", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    PrintTitle User$
    PrintHelp Help$
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
'      SETFEE
      ARPenaltyProc
    CASE 2
      AppListing
    CASE 3
      ARAPPLICATION
    CASE 4
      LicenseRegister
    CASE 5
      PrintLicenses
    CASE 6
      PrintLateNotice
    CASE 7
      HideCursor
      CLS
      END
    END SELECT
  LOOP
  
  RUN "armenu"

SUB ARAPPLICATION
  PrintAtmore$ = "Y"
  'PrintDumfries$ = "Y"
  'PrintClintWood$ = "Y"
  'PrintStephens$ = "Y"
  'PrintExmore$ = "Y"
  
  SHARED Choice$()
  
  LibName$ = "AR"
  ScrnName$ = "ARAPPPRN"
  Help$ = "Print Applications!!!"
  LOCATE 1, 1, 0
  
  REDIM ARCatCodeRec(1) AS ARNewCatCodeRecType
  ARCatCodeRecLen = LEN(ARCatCodeRec(1))
  ARCatFile = FREEFILE
  OPEN "ARCODE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARCatFile LEN = ARCatCodeRecLen
  NumOFARCatRecs = LOF(ARCatFile) \ ARCatCodeRecLen
  
  IF NumOFARCatRecs = 0 THEN
    CLOSE ARCatFile
    ShowNoCode
    EXIT SUB
  END IF
  
  NumOFARCatRecs = NumOFARCatRecs + 1
  
  
  REDIM Choice$(0 TO NumOFARCatRecs, 0 TO 1)
  
  Choice$(0, 0) = "1"
  FOR Cnt! = 1 TO (NumOFARCatRecs - 1)
    GET ARCatFile, Cnt!, ARCatCodeRec(1)
    Choice$(Cnt!, 0) = ARCatCodeRec(1).CATCODE
  NEXT Cnt!
  Choice$(NumOFARCatRecs, 0) = "ALL"
  
  CLOSE ARCatFile
  
  Choice$(0, 1) = "3"
  Choice$(1, 1) = "Standard"
  Choice$(2, 1) = "Custom"
  
  
  ShowCursor
  LibFile2Scrn "AR.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
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
  
  REM check for code file
  
  DO
    
MainBody:
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key

      IF ASC(LEFT$(Form$(1, 0), 1)) = 32 THEN
        Done = False
        COLOR 12
        LOCATE 9, 34
        PRINT "Catagory Not Set"
        PRINT CHR$(7);
        COLOR 15
        SLEEP 3
        LOCATE 9, 34
        PRINT STRING$(23, 32)
        Action = 1
        Frm(1).FldNo = 1
        GOTO MainBody
      END IF
      IF ASC(LEFT$(Form$(2, 0), 1)) = 32 THEN
        Done = False
        COLOR 12
        LOCATE 10, 34
        PRINT "Year Not Set"
        PRINT CHR$(7);
        SLEEP 3
        LOCATE 10, 34
        COLOR 15
        PRINT STRING$(23, 32)
        Action = 1
        Frm(1).FldNo = 2
        GOTO MainBody
      END IF
      IF ASC(LEFT$(Form$(3, 0), 1)) = 32 THEN
        Done = False
        COLOR 12
        LOCATE 11, 34
        PRINT "Form Type Not Set"
        PRINT CHR$(7);
        SLEEP 3
        LOCATE 11, 34
        COLOR 15
        PRINT STRING$(23, 32)
        Action = 1
        Frm(1).FldNo = 3
        GOTO MainBody
      END IF
      Code$ = QPTrim$(Form$(1, 0))
      TownName$ = UCASE$(QPTrim$(Form$(4, 0)))
      IF INSTR(TownName$, "CLINTWOOD") THEN PrintClintWood$ = "Y"               'Hard Coded for ClintWood's App's
      IF INSTR(TownName$, "ATMORE") THEN PrintAtmore$ = "Y"
      IF INSTR(TownName$, "EXMORE") THEN PrintExmore$ = "Y"
      IF LEFT$(Form$(3, 0), 6) = "Custom" THEN GOSUB PrintCustom
      IF LEFT$(Form$(3, 0), 8) = "Standard" THEN GOSUB PrintStandard
      
    CASE EscKey
      EXIT SUB
    END SELECT
    
  LOOP
  
  
PrintCustom:
  
  Year$ = Form$(2, 0)
  
  ReportFile$ = "CUSTAPPS.RPT"
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  ARCatFile = FREEFILE
  OPEN "ARCODE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARCatFile LEN = ARCatCodeRecLen
  NumOFARCatRecs = LOF(ARCatFile) \ ARCatCodeRecLen
  
  CustRecLen = LEN(ARCustRec(1))
  TrHandle = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS TrHandle LEN = CustRecLen
  TrNumRecs = LOF(TrHandle) \ CustRecLen
  
  IdxCustRecLen = LEN(ARCustIdxRec(1))
  IdxTrHandle = FREEFILE
  OPEN "ARCUST.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS IdxTrHandle LEN = IdxCustRecLen
  IdxTrNumRecs = LOF(IdxTrHandle) \ IdxCustRecLen
  
  FOR Cnt = 1 TO IdxTrNumRecs
    GET IdxTrHandle, Cnt, ARCustIdxRec(1)
    GET TrHandle, ARCustIdxRec(1).IDXRECORD, ARCustRec(1)
    IF (ARCustRec(1).Deleted <> "Y") AND (ARCustRec(1).Inactive <> "Y") THEN
      IF Code$ = "ALL" THEN
        GOSUB PrintCustomForm
      ELSE
        ThisCode$ = QPTrim$(ARCustRec(1).BILLCAT1)
        IF ThisCode$ = Code$ THEN
          GOSUB PrintCustomForm
          GOTO DoneEm
        END IF
        ThisCode$ = QPTrim$(ARCustRec(1).BILLCAT2)
        IF ThisCode$ = Code$ THEN
          GOSUB PrintCustomForm
          GOTO DoneEm
        END IF
        ThisCode$ = QPTrim$(ARCustRec(1).BILLCAT3)
        IF ThisCode$ = Code$ THEN
          GOSUB PrintCustomForm
          GOTO DoneEm
        END IF
        ThisCode$ = QPTrim$(ARCustRec(1).BILLCAT4)
        IF ThisCode$ = Code$ THEN
          GOSUB PrintCustomForm
          GOTO DoneEm
        END IF
        ThisCode$ = QPTrim$(ARCustRec(1).BILLCAT5)
        IF ThisCode$ = Code$ THEN
          GOSUB PrintCustomForm
          GOTO DoneEm
        END IF
      END IF

    END IF
'    IF TotalCust >= 1 THEN
'      EXIT FOR
'    END IF
DoneEm:
  NEXT
  
  CLOSE         'Close all open files now
  EntryPoint = 1
  LPTPort% = 1
  Header$ = "Applications"
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  'KILL ReportFile$
  EXIT SUB
  
PrintCustomForm:


''$INCLUDE: 'ARRIVRSD.BI'
''$INCLUDE: 'ARELLORE.BI'

''$INCLUDE: 'ARSUNBCH.BI'               'Sunset Beach


' '$INCLUDE: 'ARLOUISA.BI'               'Louisa
'  IF PrintStephens$ = "Y" THEN
'    '$INCLUDE: 'ARSTEPHN.BI'               'Stephens City
'  END IF
'  IF PrintDumfries$ = "Y" THEN
'    '$INCLUDE: 'ARDUMFRI.BI2'               'Dumfires
'  END IF
'  IF PrintClintWood$ = "Y" THEN
  '$INCLUDE: 'ARCLNTWD.BI'               'Clintwood
'  END IF
'  IF PrintAtmore$ = "Y" THEN
'    '$INCLUDE: 'ARATMORE.BI'               'Atmore
'  END IF
'  IF PrintExmore$ = "Y" THEN
'    '$INCLUDE: 'AREXMORE.BI'               'Exmore
'  END IF

RETURN

PrintStandard:
  ReportFile$ = "CUSTAPPS.RPT"
  RptHandle = FREEFILE
  
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  CustRecLen = LEN(ARCustRec(1))
  TrHandle = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS TrHandle LEN = CustRecLen
  TrNumRecs = LOF(TrHandle) \ CustRecLen
  
  'REDIM ARCustIdxRec(1) AS ARCustIdxRecType     ' open customer file
  
  IdxCustRecLen = LEN(ARCustIdxRec(1))
  IdxTrHandle = FREEFILE
  OPEN "ARCUST.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS IdxTrHandle LEN = IdxCustRecLen
  IdxTrNumRecs = LOF(IdxTrHandle) \ IdxCustRecLen
  
  FOR Cnt = 1 TO IdxTrNumRecs
    GET IdxTrHandle, Cnt, ARCustIdxRec(1)
    GET TrHandle, ARCustIdxRec(1).IDXRECORD, ARCustRec(1)
    IF ARCustRec(1).Deleted <> "Y" THEN
      IF Code$ = "ALL" THEN
        GOSUB PrintSTDForm
      ELSE
'dale

        IF QPTrim$(ARCustRec(1).BILLCAT1) = Code$ OR QPTrim$(ARCustRec(1).BILLCAT2) = Code$ OR QPTrim$(ARCustRec(1).BILLCAT3) = Code$ OR QPTrim$(ARCustRec(1).BILLCAT4) = Code$ OR QPTrim$(ARCustRec(1).BILLCAT5) = Code$ THEN
          GOSUB PrintSTDForm
        END IF
      END IF
      
    END IF
  NEXT Cnt
  
  CLOSE         'Close all open files now
  EntryPoint = 1
  LPTPort% = 2
  Header$ = "Applications"
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  KILL ReportFile$
  EXIT SUB
  
PrintSTDForm:
  FF$ = CHR$(12)
  MaxLines = 53
  LineCnt = 0
  FOR ll = 1 TO 5
    PRINT #RptHandle, ""
  NEXT
  PRINT #RptHandle, TAB(37 - tab1); Heading1$
  PRINT #RptHandle, TAB(37 - tab2); Heading2$
  PRINT #RptHandle, TAB(37 - Tab3); Heading3$
  PRINT #RptHandle, TAB(37 - Tab4); Heading4$
  PRINT #RptHandle, TAB(66); Form$(2, 0)
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(11); ARCustRec(1).BILLNAME
  PRINT #RptHandle, TAB(11); ARCustRec(1).Address1
  PRINT #RptHandle, TAB(11); ARCustRec(1).Address2
  PRINT #RptHandle, TAB(11); RTRIM$(ARCustRec(1).City); "  "; ARCustRec(1).State; " "; ARCustRec(1).ZipCode
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(11); ARCustRec(1).CUSTNAME
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(5); ARCustRec(1).BILLCAT1;
  PRINT #RptHandle, TAB(15); ARCustRec(1).Desc1;
  PRINT #RptHandle, TAB(62); USING "#####.##"; ARCustRec(1).Fee1
  SCnt = 24
  IF VAL(ARCustRec(1).BILLCAT2) = 0 THEN GOTO ExitFormPrint
  PRINT #RptHandle, TAB(5); ARCustRec(1).BILLCAT2;
  PRINT #RptHandle, TAB(15); ARCustRec(1).Desc2;
  PRINT #RptHandle, TAB(62); USING "#####.##"; ARCustRec(1).Fee2
  SCnt = 25
  IF VAL(ARCustRec(1).BILLCAT3) = 0 THEN GOTO ExitFormPrint
  PRINT #RptHandle, TAB(5); ARCustRec(1).BILLCAT3;
  PRINT #RptHandle, TAB(15); ARCustRec(1).Desc3;
  PRINT #RptHandle, TAB(62); USING "#####.##"; ARCustRec(1).Fee3
  SCnt = 26
  IF VAL(ARCustRec(1).BILLCAT4) = 0 THEN GOTO ExitFormPrint
  PRINT #RptHandle, TAB(5); ARCustRec(1).BILLCAT4;
  PRINT #RptHandle, TAB(15); ARCustRec(1).Desc4;
  PRINT #RptHandle, TAB(62); USING "#####.##"; ARCustRec(1).Fee4
  SCnt = 27
  IF VAL(ARCustRec(1).BILLCAT5) = 0 THEN GOTO ExitFormPrint
  PRINT #RptHandle, TAB(5); ARCustRec(1).BILLCAT5;
  PRINT #RptHandle, TAB(15); ARCustRec(1).Desc5;
  PRINT #RptHandle, TAB(62); USING "#####.##"; ARCustRec(1).Fee5
  SCnt = 28
  
  
  
ExitFormPrint:
  FOR lcnt = SCnt TO 35
    PRINT #RptHandle, ""
  NEXT
  PRINT #RptHandle, TAB(62); USING "#####.##"; ARCustRec(1).FeeAmt
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(62); USING "#####.##"; ARCustRec(1).FeeAmt
  PRINT #RptHandle,
  PRINT #RptHandle,
  TotalCust = TotalCust + 1
RETURN
  
  
GetCode:
  FOR Snt& = 1 TO NumOFARCatRecs
    GET ARCatFile, Snt&, ARCatCodeRec(1)
    IF QPTrim$(ARCatCodeRec(1).CATCODE) = CATCODE$ THEN
      CODEDESC$ = QPTrim$(ARCatCodeRec(1).CODEDESC)
      SELECT CASE ARCatCodeRec(1).CodeType
      CASE "F"
        Amt# = ARCatCodeRec(1).Fee
        CodeType$ = ARCatCodeRec(1).CodeType
        
      CASE "M"
        Desc1$ = "Per Each"
        Amt# = ARCatCodeRec(1).Fee
        CodeType$ = ARCatCodeRec(1).CodeType
      CASE IS = "S"
        BaseAmt1# = ARCatCodeRec(1).BaseAmt1
        Revenue1# = ARCatCodeRec(1).Recpt1
        Percent1# = ARCatCodeRec(1).Percent1
        Maximum1# = ARCatCodeRec(1).Maximum1
        BaseAmt2# = ARCatCodeRec(1).BaseAmt2
        Revenue2# = ARCatCodeRec(1).Recpt2
        Percent2# = ARCatCodeRec(1).Percent2
        Maximum2# = ARCatCodeRec(1).Maximum2
        BaseAmt3# = ARCatCodeRec(1).BaseAmt3
        Revenue3# = ARCatCodeRec(1).Recpt3
        Percent3# = ARCatCodeRec(1).Percent3
        Maximum3# = ARCatCodeRec(1).Maximum3
        BaseAmt4# = ARCatCodeRec(1).BaseAmt4
        Revenue4# = ARCatCodeRec(1).Recpt4
        Percent4# = ARCatCodeRec(1).Percent4
        Maximum4# = ARCatCodeRec(1).Maximum4
        BaseAmt5# = ARCatCodeRec(1).BaseAmt5
        Revenue5# = ARCatCodeRec(1).Recpt5
        Percent5# = ARCatCodeRec(1).Percent5
        Maximum5# = ARCatCodeRec(1).Maximum5
        BaseAmt6# = ARCatCodeRec(1).BaseAmt6
        Revenue6# = ARCatCodeRec(1).Recpt6
        Percent6# = ARCatCodeRec(1).Percent6
        Maximum6# = ARCatCodeRec(1).Maximum6
        CodeType$ = ARCatCodeRec(1).CodeType
      CASE ELSE
        CodeType$ = "N"
      END SELECT
      EXIT FOR
    END IF
  NEXT Snt&
  

GotCode:
  RETURN
  
END SUB

SUB ARPenaltyProc

  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 4)
  MChoice$(1) = "Create Penalty File"
  MChoice$(2) = "Print Penalty Report"
  MChoice$(3) = "Post Penalty transactions"
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
  Row = ((25 - (UBOUND(MChoice$))) \ 2)
  Col = ((80 - MaxLen) \ 2) - 1
  Help$ = "Billing Menu"

  DO

    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    'DisplayARScrn "MENUBAK"
    LibFile2Scrn "AR.QSL", "MENUBAK", MonoCode, -1, ErrorCode

    TitleBox 3, Col, MaxLen + 3, "Penalty Processing", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

    PrintTitle User$
    PrintHelp Help$

    ShowCursor

    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf

    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0

    SELECT CASE Choice
    CASE 1
      CreatePenalty
    CASE 2
      PrintPenRpt
    CASE 3
      PostPenalty
    CASE 4
      HideCursor
      CLS
      END
    END SELECT
  LOOP

END SUB

SUB CreatePenalty
  SHARED Choice$()

  LibName$ = "AR"
  ScrnName$ = "ARPENINF"

  REDIM Choice$(0 TO 2, 0 TO 2)

  Choice$(0, 1) = "1"
  Choice$(1, 1) = "Total Balance"
  Choice$(2, 1) = "License Only"

  Choice$(0, 2) = "2"
  Choice$(1, 2) = "Fixed Amount"
  Choice$(2, 2) = "Percentage"

  ShowCursor
  LibFile2Scrn "AR.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
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

MainCreate:
  Action = 1

  REM check for code file

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode

    CASE F10Key
      SELECT CASE LEFT$(Form$(1, 0), 1)
      CASE "T"
        TBalFlag = True
      CASE "L"
        TBalFlag = False
      CASE ELSE
        NGFlag = True
        Frm(1).FldNo = 1
        EXIT DO
      END SELECT


      SELECT CASE LEFT$(Form$(2, 0), 1)
      CASE "F"
        PctFlag = False
      CASE "P"
        PctFlag = True
      CASE ELSE
        NGFlag = True
        Frm(1).FldNo = 2
        EXIT DO
      END SELECT
      
      Amt# = VAL(Form$(3, 0))
      IF Amt# <= 0 THEN
        NGFlag = True
        Frm(1).FldNo = 3
        EXIT DO
      END IF
      EXIT DO
    CASE EscKey
      EXIT SUB
    END SELECT

  LOOP

  IF NGFlag THEN
    NGFlag = False
    GOTO MainCreate
  END IF
  IF PctFlag THEN
    Amt# = Amt# / 100
  END IF

  ShowProcessingScrn "Creating Penalty Transactions"

  REDIM ARTransRec(1) AS ARTransRecType
  ARTransRecLen = LEN(ARTransRec(1))

  CustRecLen = LEN(ARCustRec(1))

  CALL KillFile("ARTPENTR.DAT")
  ARTransFile = FREEFILE
  OPEN "ARTPENTR.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARTransFile LEN = ARTransRecLen

  CUHandle = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS CUHandle LEN = CustRecLen
  TrNumRecs = LOF(CUHandle) \ CustRecLen
  FOR Cnt = 1 TO TrNumRecs
    REDIM ARTransRec(1) AS ARTransRecType
    GET #CUHandle, Cnt, ARCustRec(1)
    IF ARCustRec(1).Deleted <> "Y" AND ARCustRec(1).Inactive <> "Y" THEN
      IF TBalFlag THEN
        CustBal# = ARCustRec(1).AcctBal
      ELSE
        CustBal# = ARCustRec(1).LicBal
      END IF
      IF CustBal# > 0 THEN
        IF PctFlag THEN
          CustPen# = Round#(CustBal# * Amt#)
        ELSE
          CustPen# = Amt#
        END IF
        ARTransRec(1).CustomerNumber = STR$(Cnt)
        ARTransRec(1).TransDate = Date2Num(DATE$)
        ARTransRec(1).TransType = 6
        ARTransRec(1).TransAmount = CustPen#
        ARTransRec(1).TransDesc = "PENALTY"
        ARTransRec(1).CashAmount = 0
        ARTransRec(1).ChkAmount = 0
        ARTransRec(1).PenAMT = CustPen#
        PUT #ARTransFile, Cnt, ARTransRec(1)
      END IF
    ELSE
      PUT #ARTransFile, Cnt, ARTransRec(1)
    END IF

    ShowPctComp Cnt, TrNumRecs
  NEXT

  'NumOfTransRecs = LOF(ARTransFile) \ ARTransRecLen
  'NextTransRec = NumOfTransRecs + 1

  CLOSE

  LibFile2Scrn "AR.QSL", "MENUBAK", MonoCode, -1, ErrorCode
  LibFile2Scrn "AR.QSL", "UPDATEOK", MonoCode, -1, ErrorCode

  WaitForAction
  EXIT SUB


'      CustomerNumber AS STRING * 10
'      TransDate AS INTEGER
'      TransAmount AS DOUBLE precision
'      TransType AS INTEGER
'      TransDesc AS STRING * 35
'      CashAmount AS DOUBLE precision
'      ChkAmount AS DOUBLE precision
'      BalanceAfterTrans AS DOUBLE precision
'      Posted2GL AS STRING * 1
'      CatCodeRec1 AS INTEGER
'      CatCodeRec2 AS INTEGER
'      CatCodeRec3 AS INTEGER
'      CatCodeRec4 AS INTEGER
'      CatCodeRec5 AS INTEGER
'      LicAmt AS DOUBLE precision
'      FeeAmt AS DOUBLE precision
'      PenAmt AS DOUBLE precision
'      ExtraRoom AS STRING * 8
'      NextTrans AS LONG integer



END SUB

SUB LicenseRegister
  
  SHARED Choice$()

  SETFEE

  ReportFile$ = "ARLICREG.PRN"  'Report File Name
  FF$ = CHR$(12)
  MaxLines = 53
  LineCnt = 0
  Start = 1     'start at array element 1
  '   -1 = integer
  '   -2 = long integer
  '   -3 = single precision
  '   -4 = double precision
  '   +N = TYPE array/fixed-length string of length N
  
  
  GOSUB SelectLicRegOutput
  IF Canceled$ = "Y" THEN EXIT SUB
  
  'GOSUB GetReportLicRegInformation
  
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  
  'PRINT #RptHandle, CHR$(27); CHR$(58); ' oki 320 12 cpi
  
  GOSUB PrintLicRegRptHeader
  
  ' Print Main Body
  OpenARCustFile NumOfArRecs, ARFile
  REDIM ARIdxRec(1) AS ARCustIDXRecType
  IdxFile = FREEFILE
  OPEN "ARCUST.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS #IdxFile LEN = LEN(ARIdxRec(1))
  
  IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN NumOfArRecs = LOF(IdxFile) \ LEN(ARIdxRec(1))
  
  FOR Cnt& = 1 TO NumOfArRecs
    'IF Cnt& = 18 THEN STOP
    IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN
      GET IdxFile, Cnt&, ARIdxRec(1)
      GET ARFile, ARIdxRec(1).IDXRECORD, ARCustRec(1)
      CustomerNumber = ARIdxRec(1).IDXRECORD
    ELSE
      GET ARFile, Cnt&, ARCustRec(1)
      CustomerNumber = Cnt&
    END IF
    
    '082399 Added deleted check
    DelFlag$ = UCASE$(ARCustRec(1).Deleted)
    
    IF (DelFlag$ = "Y") OR (ARCustRec(1).Inactive = "Y") THEN
      GOTO DelSkip
    END IF
    
    IF ARCustRec(1).IssueLicense = "Y" THEN
      IF LineCnt >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintLicRegRptHeader
      END IF
      CustFee# = Round#(ARCustRec(1).Fee1 + ARCustRec(1).Fee2 + ARCustRec(1).Fee3 + ARCustRec(1).Fee4 + ARCustRec(1).Fee5)
      IF CustFee# > 0 THEN
        CCode$ = " " + QPTrim$(ARCustRec(1).BILLCAT1)
        TC$ = QPTrim$(ARCustRec(1).BILLCAT2)
        IF LEN(TC$) > 0 THEN
          CCode$ = CCode$ + "/" + TC$
        END IF
        TC$ = QPTrim$(ARCustRec(1).BILLCAT3)
        IF LEN(TC$) > 0 THEN
          CCode$ = CCode$ + "/" + TC$
        END IF
        TC$ = QPTrim$(ARCustRec(1).BILLCAT4)
        IF LEN(TC$) > 0 THEN
          CCode$ = CCode$ + "/" + TC$
        END IF
        TC$ = QPTrim$(ARCustRec(1).BILLCAT5)
        IF LEN(TC$) > 0 THEN
          CCode$ = CCode$ + "/" + TC$
        END IF

        PRINT #RptHandle, CustomerNumber; TAB(10); LEFT$(ARCustRec(1).BILLNAME, 30); CCode$; TAB(70); USING "$$####,#.##"; CustFee#
        TotalFee# = Round#(TotalFee# + CustFee#)
        LineCnt = LineCnt + 1
      END IF
    END IF
    
DelSkip:
  NEXT
  
  GOSUB PrintLicRegRptEnding
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
  
  
PrintLicRegRptHeader:
  Page = Page + 1
  PRINT #RptHandle, TAB(18); "Business License System : License's to Print Register"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
  PRINT #RptHandle, ""
  PRINT #RptHandle, "Cust#"; TAB(10); "Customer Name"; TAB(70); "Fee Amount"
  PRINT #RptHandle, STRING$(80, "=")
  LineCnt = 5
  RETURN
  
PrintLicRegRptEnding:
  PRINT #RptHandle, STRING$(80, "-")
  PRINT #RptHandle, "Total Fees to Bill  .. "; USING "$######,#.##"; TotalFee#
  PRINT #RptHandle, FF$
  RETURN
  
  'GetReportLicRegInformation:
  '      RETURN
  
SelectLicRegOutput:
  LibName$ = "AR"
  ScrnName$ = "ARLICREG"
  
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
  
  REDIM Choice$(100, 2)         ' assume maximum of 99 codes
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Name"
  Choice$(2, 0) = "Number"
  Choice$(0, 2) = "2"
  Choice$(1, 2) = "SCREEN"
  Choice$(2, 2) = "PRINTER"
  
  Action = 1
  ShowCursor
  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%
  'printhelp help$
  Action = 1
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      Order$ = Form$(1, 0)
      DevSpec$ = LEFT$(Form$(2, 0), 1)
      RETURN
    CASE EscKey
      Canceled$ = "Y"
      RETURN
    END SELECT
  LOOP
  RETURN
  
END SUB

FUNCTION MakeDay$ (DayNum)
  SELECT CASE DayNum
  CASE 1
    MakeDay$ = "Sunday"
  CASE 2
    MakeDay$ = "Monday"
  CASE 3
    MakeDay$ = "Tuesday"
  CASE 4
    MakeDay$ = "Wednesday"
  CASE 5
    MakeDay$ = "Thursday"
  CASE 6
    MakeDay$ = "Friday"
  CASE 7
    MakeDay$ = "Saturday"
  END SELECT
END FUNCTION

FUNCTION MakeLongDate$ (PDate$)
  DayNum = Num2Day(Date2Num(PDate$))
  DayName$ = MakeDay$(DayNum)
  MonthName$ = MakeMonth$(PDate$)
  MakeLongDate$ = DayName$ + ", " + MonthName$ + " " + MID$(PDate$, 4, 2) + ", " + RIGHT$(PDate$, 4)
END FUNCTION

FUNCTION MakeMonth$ (TDate$)
  Month = VAL(LEFT$(TDate$, 2))
  SELECT CASE Month
  CASE 1
    MakeMonth$ = "January"
  CASE 2
    MakeMonth$ = "February"
  CASE 3
    MakeMonth$ = "March"
  CASE 4
    MakeMonth$ = "April"
  CASE 5
    MakeMonth$ = "May"
  CASE 6
    MakeMonth$ = "June"
  CASE 7
    MakeMonth$ = "July"
  CASE 8
    MakeMonth$ = "August"
  CASE 9
    MakeMonth$ = "September"
  CASE 10
    MakeMonth$ = "October"
  CASE 11
    MakeMonth$ = "November"
  CASE 12
    MakeMonth$ = "December"
  END SELECT
END FUNCTION

SUB OpenARCustFile (NumOfArRecs, ARFile)
  CLOSE ARFile
  
  ARCustRecLen = LEN(ARCustRec(1))
  ARFile = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARFile LEN = ARCustRecLen
  NumOfArRecs = LOF(ARFile) \ ARCustRecLen
  
END SUB

SUB PostPenalty

  LibFile2Scrn "AR.QSL", "MENUBAK", MonoCode, -1, ErrorCode
  Ok = MsgBox("AR", "RDY2POST")

  IF Ok = 1 THEN
    EXIT SUB
  END IF

  ShowProcessingScrn "Posting Penalty Transactions."

  ARCustRecLen = LEN(ARCustRec(1))
  ARFile = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM SHARED AS ARFile LEN = ARCustRecLen
  NumOfArRecs = LOF(ARFile) \ ARCustRecLen

  REDIM ARTransRec(1) AS ARTransRecType
  ARTransRecLen = LEN(ARTransRec(1))
  ARTransFile = FREEFILE
  OPEN "ARTRANS.DAT" FOR RANDOM SHARED AS ARTransFile LEN = ARTransRecLen
  NumOfTransRecs = LOF(ARTransFile) \ ARTransRecLen
  NextTransRec = NumOfTransRecs + 1

  ARPenFile = FREEFILE
  OPEN "ARTPENTR.DAT" FOR RANDOM SHARED AS ARPenFile LEN = ARTransRecLen
  NumOfPen = LOF(ARPenFile) \ ARTransRecLen
  FOR PCnt = 1 TO NumOfPen
    GET #ARPenFile, PCnt, ARTransRec(1)
'    IF PCnt = 18 THEN STOP
    CustRecNum$ = QPTrim$(ARTransRec(1).CustomerNumber)
    IF LEN(CustRecNum$) > 0 THEN
      CustomerNumber = VAL(CustRecNum$)
      GET #ARFile, CustomerNumber, ARCustRec(1)
      ARCustRec(1).PenBal = Round#(ARCustRec(1).PenBal + ARTransRec(1).TransAmount)
      ARCustRec(1).AcctBal = Round#(ARCustRec(1).AcctBal + ARTransRec(1).TransAmount)
      ARTransRec(1).BalanceAfterTrans = ARCustRec(1).AcctBal

      PUT ARTransFile, NextTransRec, ARTransRec(1)

      'unrem
      IF ARCustRec(1).FirstTrans = 0 THEN
        ARCustRec(1).FirstTrans = NextTransRec
        ARCustRec(1).LastTrans = NextTransRec
        PUT ARFile, CustomerNumber, ARCustRec(1)
        'unrem
      ELSE
        Prev! = ARCustRec(1).LastTrans
        ARCustRec(1).LastTrans = NextTransRec
        PUT ARFile, CustomerNumber, ARCustRec(1)
        'unrem
        GET ARTransFile, Prev!, ARTransRec(1)
        ARTransRec(1).NextTrans = NextTransRec
        PUT ARTransFile, Prev!, ARTransRec(1)
        'unrem
      END IF
      NextTransRec = NextTransRec + 1
    END IF
    ShowPctComp PCnt, NumOfPen
  NEXT
  CLOSE
  CALL KillFile("ARTPENTR.DAT")
  LibFile2Scrn "AR.QSL", "MENUBAK", MonoCode, -1, ErrorCode
  LibFile2Scrn "AR.QSL", "UPDATEOK", MonoCode, -1, ErrorCode

  WaitForAction

END SUB

SUB PrintLateNotice
  
  SHARED Choice$()
  
  ReportFile$ = "AREXPLIC.PRN"  'Report File Name
  FF$ = CHR$(12)
  DevSpec$ = "S"
  MaxLines = 53
  LineCnt = 0
  
  LibName$ = "AR"
  ScrnName$ = "ARLATNOT"
  
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
  
  REDIM Choice$(3, 0)           ' assume maximum of 99 codes
  
  REDIM ARCatCodeRec(1) AS ARNewCatCodeRecType
  ARCatCodeRecLen = LEN(ARCatCodeRec(1))
  ARCatFile = FREEFILE
  OPEN "ARCODE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARCatFile LEN = ARCatCodeRecLen
  NumOFARCatRecs = LOF(ARCatFile) \ ARCatCodeRecLen
  CLOSE ARCatFile
  
  IF NumOFARCatRecs = 0 THEN
    ShowNoCode
    EXIT SUB
  END IF
  
  'NumOFARCatRecs = NumOFARCatRecs + 1
  
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Name"
  Choice$(2, 0) = "Number"
  
  'Choice$(0, 1) = "2"
  'FOR Cnt! = 1 TO (NumOFARCatRecs - 1):
  '  GET ARCatFile, Cnt!, ARCatCodeRec(1)
  '  Choice$(Cnt!, 1) = ARCatCodeRec(1).CATCODE
  'NEXT Cnt!
  'Choice$(NumOFARCatRecs, 1) = "ALL"
  
  Today$ = DATE$
  Year = VAL(RIGHT$(Today$, 4))
  Year = Year - 1
  ExpDate$ = "12-31-" + QPTrim$(STR$(Year))
  PayDate$ = "03-31-" + RIGHT$(Today$, 4)
  FirstTime = True
  
  Action = 1
  ShowCursor
  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%
  Action = 1
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      LSET Form$(3, 0) = Today$
      LSET Form$(4, 0) = PayDate$
      LSET Form$(5, 0) = ExpDate$
      Action = 1
    END IF
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      Order$ = Form$(1, 0)
      Catagory$ = QPTrim$(Form$(2, 0))
      LNotDate$ = MakeLongDate$(Form$(3, 0))
      LPayDate$ = MakeLongDate$(Form$(4, 0))
      LNDLen = LEN(LNotDate$)
      LNDTab = (40 - (LNDLen / 2))
      LNDTab = LNDTab - 1
      ExpDate = Date2Num%(Form$(5, 0))
      PenPct$ = QPTrim$(Form$(6, 0)) + "%"
      IF ExpDate = -32768 THEN
        PRINT CHR$(7);
        LOCATE 17, 37
        COLOR 12
        PRINT "NO DATE ENTERED"
        SLEEP 2
        LOCATE 17, 37
        PRINT STRING$(15, 32);
        Frm(1).FldNo = 5
        Action = 1
      ELSE
        EXIT DO
      END IF
    CASE EscKey
      Canceled$ = "Y"
      EXIT DO
    END SELECT
  LOOP
  
  IF Canceled$ = "Y" THEN EXIT SUB
  
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  'GOSUB PrintExpiredLicenseListRptHeader
  
  ' Print Main Body
  OpenARCustFile NumOfArRecs, ARFile
  REDIM ARIdxRec(1) AS ARCustIDXRecType
  IdxFile = FREEFILE
  OPEN "ARCUST.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS #IdxFile LEN = LEN(ARIdxRec(1))
  
  IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN NumOfArRecs = LOF(IdxFile) \ LEN(ARIdxRec(1))
  
  FOR Cnt! = 1 TO NumOfArRecs
    IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN
      GET IdxFile, Cnt!, ARIdxRec(1)
      GET ARFile, ARIdxRec(1).IDXRECORD, ARCustRec(1)
      CustomerNumber = ARIdxRec(1).IDXRECORD
    ELSE
      GET ARFile, Cnt!, ARCustRec(1)
      CustomerNumber = Cnt!
    END IF
    
    IF UCASE$(ARCustRec(1).Deleted) <> "Y" THEN
      IF (ARCustRec(1).VALID > -32768) AND (ARCustRec(1).VALID <= ExpDate) THEN
        IF ARCustRec(1).AcctBal > 0 THEN
          IF UCASE$(ARCustRec(1).Inactive) <> "Y" THEN
            PrintHelp ARCustRec(1).CUSTNAME
            ''$INCLUDE: 'ARATMOLN.BI'
            ''$INCLUDE: 'ARcantln.BI'
            '$INCLUDE: 'ARSTDLN.BI'
            ''$INCLUDE: 'ARPENGLN.BI'
            ''$INCLUDE: 'ARHAMLN.BI'
            ''$INCLUDE: 'ARMAXTLN.BI'
            ''$INCLUDE: 'ARCBEACH.BI'
            ''$INCLUDE: 'ARXMORLN.BI'

  'PRINT #RptHandle, TAB(31); "TOWN OF LELAND"
  'PRINT #RptHandle, TAB(31); "102 TOWN HALL DRIVE"
  'PRINT #RptHandle, TAB(31); "LELAND, NC 28451"
  'PRINT #RptHandle, TAB(31); "TEL 910-371-0148"

          'ELSE
          '  ccc = ccc + 1
          '  'STOP
          END IF
        END IF
      END IF
    END IF
  NEXT Cnt!
  
  CLOSE         'Close all open files now
  
  Header$ = "Delinquent Notice Printing"
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, 0
  
  'KILL ReportFile$
  
  EXIT SUB
  
  
MakeCatLine:
  CatLine$ = ""
  C$ = QPTrim$(ARCustRec(1).BILLCAT1)
  IF LEN(C$) > 0 THEN
    CatLine$ = CatLine$ + C$
  END IF
  
  C$ = QPTrim$(ARCustRec(1).BILLCAT2)
  IF LEN(C$) > 0 THEN
    IF LEN(CatLine$) > 0 THEN
      CatLine$ = CatLine$ + ", " + C$
    ELSE
      CatLine$ = CatLine$ + C$
    END IF
  END IF
  C$ = QPTrim$(ARCustRec(1).BILLCAT3)
  IF LEN(C$) > 0 THEN
    IF LEN(CatLine$) > 0 THEN
      CatLine$ = CatLine$ + ", " + C$
    ELSE
      CatLine$ = CatLine$ + C$
    END IF
  END IF
  C$ = QPTrim$(ARCustRec(1).BILLCAT4)
  IF LEN(C$) > 0 THEN
    IF LEN(CatLine$) > 0 THEN
      CatLine$ = CatLine$ + ", " + C$
    ELSE
      CatLine$ = CatLine$ + C$
    END IF
  END IF
  C$ = QPTrim$(ARCustRec(1).BILLCAT5)
  IF LEN(C$) > 0 THEN
    IF LEN(CatLine$) > 0 THEN
      CatLine$ = CatLine$ + ", " + C$
    ELSE
      CatLine$ = CatLine$ + C$
    END IF
  END IF
  
RETURN
  
END SUB

SUB PrintLicenses
  SHARED Choice$()

  ReportFile$ = "ARLICENS.PRN"  'Report File Name
  MaxLines = 53
  LineCnt = 0
  
  GOSUB SelectLicenseOutput
  IF Canceled$ = "Y" THEN EXIT SUB
  Help$ = "Now Printing the License's"
  PrintHelp Help$
  
  ReportFile$ = "ARLIC.PRN"
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  'PRINT #RptHandle, CHR$(27); CHR$(58); ' oki 320 12 cpi
  ' Print Main Body
  OpenARCustFile NumOfArRecs, ARFile
  REDIM ARIdxRec(1) AS ARCustIDXRecType
  IdxFile = FREEFILE
  OPEN "ARCUST.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS #IdxFile LEN = LEN(ARIdxRec(1))
  
  IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN
    NumOfArRecs = LOF(IdxFile) \ LEN(ARIdxRec(1))
  END IF
  
  FOR Cnt& = 1 TO NumOfArRecs
    IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN
      GET IdxFile, Cnt&, ARIdxRec(1)
      GET ARFile, ARIdxRec(1).IDXRECORD, ARCustRec(1)
      CustomerNumber = ARIdxRec(1).IDXRECORD
    ELSE
      GET ARFile, Cnt&, ARCustRec(1)
      CustomerNumber = Cnt&
    END IF
    
    'HERE DALE
    IF (ARCustRec(1).IssueLicense = "Y") AND (ARCustRec(1).Inactive <> "Y") THEN
      '(ARCustRec(1).Deleted <> "Y") AND (ARCustRec(1).Inactive <> "Y")
      'unrem
      'IF ARCustRec(1).IssueLicense <> "Y" THEN
      'Get Catagory Desc First
      CatagoryDesc$ = ""
      REDIM ARCatCodeRec(1) AS ARNewCatCodeRecType
      ARCatCodeRecLen = LEN(ARCatCodeRec(1))
      ARCatFile = FREEFILE
      OPEN "ARCODE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARCatFile LEN = ARCatCodeRecLen
      NumOFARCatRecs = LOF(ARCatFile) \ ARCatCodeRecLen
      CatagoryRecord1 = 0
      FOR Cnt = 1 TO NumOFARCatRecs
        GET ARCatFile, Cnt, ARCatCodeRec(1)
        IF RTRIM$(ARCustRec(1).BILLCAT1) = RTRIM$(ARCatCodeRec(1).CATCODE) THEN
          CatagoryDesc1$ = ARCatCodeRec(1).CODEDESC
          CatagoryRecord1 = Cnt
          EXIT FOR
        END IF
      NEXT Cnt
      CatagoryRecord2 = 0
      FOR Cnt = 1 TO NumOFARCatRecs
        GET ARCatFile, Cnt, ARCatCodeRec(1)
        IF RTRIM$(ARCustRec(1).BILLCAT2) = RTRIM$(ARCatCodeRec(1).CATCODE) THEN
          CatagoryDesc2$ = ARCatCodeRec(1).CODEDESC
          CatagoryRecord2 = Cnt
          EXIT FOR
        END IF
      NEXT Cnt
      CatagoryRecord3 = 0
      FOR Cnt = 1 TO NumOFARCatRecs
        GET ARCatFile, Cnt, ARCatCodeRec(1)
        IF RTRIM$(ARCustRec(1).BILLCAT3) = RTRIM$(ARCatCodeRec(1).CATCODE) THEN
          CatagoryDesc3$ = ARCatCodeRec(1).CODEDESC
          CatagoryRecord3 = Cnt
          EXIT FOR
        END IF
      NEXT Cnt
      CatagoryRecord4 = 0
      FOR Cnt = 1 TO NumOFARCatRecs
        GET ARCatFile, Cnt, ARCatCodeRec(1)
        IF RTRIM$(ARCustRec(1).BILLCAT4) = RTRIM$(ARCatCodeRec(1).CATCODE) THEN
          CatagoryDesc4$ = ARCatCodeRec(1).CODEDESC
          CatagoryRecord4 = Cnt
          EXIT FOR
        END IF
      NEXT Cnt
      CatagoryRecord5 = 0
      FOR Cnt = 1 TO NumOFARCatRecs
        GET ARCatFile, Cnt, ARCatCodeRec(1)
        IF RTRIM$(ARCustRec(1).BILLCAT5) = RTRIM$(ARCatCodeRec(1).CATCODE) THEN
          CatagoryDesc5$ = ARCatCodeRec(1).CODEDESC
          CatagoryRecord5 = Cnt
          EXIT FOR
        END IF
      NEXT Cnt
      
      CLOSE ARCatFile
      
      FOR ll = 1 TO 5
        PRINT #RptHandle,
      NEXT ll
      DidCnt = DidCnt + 1
      PRINT #RptHandle, TAB(37 - tab1); Heading1$
      PRINT #RptHandle, TAB(37 - tab2); Heading2$
      PRINT #RptHandle, TAB(37 - Tab3); Heading3$
      PRINT #RptHandle, TAB(37 - Tab4); Heading4$
      PRINT #RptHandle, TAB(66); Year$
      'IF CustomerNumber = 83 THEN STOP
      PRINT #RptHandle, TAB(11); USING "#####"; CustomerNumber
      PRINT #RptHandle, TAB(11); ARCustRec(1).BILLNAME
      PRINT #RptHandle, TAB(11); ARCustRec(1).Address1; TAB(58); USING "########"; LICENSE#
      PRINT #RptHandle, TAB(11); ARCustRec(1).Address2
      PRINT #RptHandle, TAB(11); RTRIM$(ARCustRec(1).City); "  "; ARCustRec(1).State; " "; ARCustRec(1).ZipCode
      PRINT #RptHandle, TAB(55); IssueDate$;
      PRINT #RptHandle, TAB(64); ExpireDate$
      PRINT #RptHandle,
      PRINT #RptHandle,
      PRINT #RptHandle,
      PRINT #RptHandle, TAB(11); ARCustRec(1).CUSTNAME
      PRINT #RptHandle,
      PRINT #RptHandle,
      PRINT #RptHandle, TAB(5); ARCustRec(1).BILLCAT1;
      PRINT #RptHandle, TAB(15); ARCustRec(1).Desc1;
      PRINT #RptHandle, TAB(62); USING "#####.##"; ARCustRec(1).Fee1
      SCnt = 24
      IF CatagoryRecord2 = 0 THEN GOTO ExitFormPrint1
      PRINT #RptHandle, TAB(5); ARCustRec(1).BILLCAT2;
      PRINT #RptHandle, TAB(15); ARCustRec(1).Desc2;
      PRINT #RptHandle, TAB(62); USING "#####.##"; ARCustRec(1).Fee2
      SCnt = 25
      IF CatagoryRecord3 = 0 THEN GOTO ExitFormPrint1
      PRINT #RptHandle, TAB(5); ARCustRec(1).BILLCAT3;
      PRINT #RptHandle, TAB(15); ARCustRec(1).Desc3;
      PRINT #RptHandle, TAB(62); USING "#####.##"; ARCustRec(1).Fee3
      SCnt = 26
      IF CatagoryRecord4 = 0 THEN GOTO ExitFormPrint1
      PRINT #RptHandle, TAB(5); ARCustRec(1).BILLCAT4;
      PRINT #RptHandle, TAB(15); ARCustRec(1).Desc4;
      PRINT #RptHandle, TAB(62); USING "#####.##"; ARCustRec(1).Fee4
      SCnt = 27
      IF CatagoryRecord5 = 0 THEN GOTO ExitFormPrint1
      PRINT #RptHandle, TAB(5); ARCustRec(1).BILLCAT5;
      PRINT #RptHandle, TAB(15); ARCustRec(1).Desc5;
      PRINT #RptHandle, TAB(62); USING "#####.##"; ARCustRec(1).Fee5
      SCnt = 28
      
ExitFormPrint1:
      IF ARCustRec(1).IssuanceFee > .01 THEN
        PRINT #RptHandle, TAB(15); "Issue Fee"; TAB(62); USING "#####.##"; ARCustRec(1).IssuanceFee
        SCnt = SCnt + 1
      END IF
      
      FOR lcnt = SCnt TO 31
        PRINT #RptHandle,
      NEXT
      PRINT #RptHandle, ""
      
      FOR lcnt = 33 TO 35
        PRINT #RptHandle, ""
      NEXT lcnt
      
      'Calc Total License Amount Here
      TotalBillAmt# = Round#(ARCustRec(1).Fee1 + ARCustRec(1).Fee2 + ARCustRec(1).Fee3 + ARCustRec(1).Fee4 + ARCustRec(1).Fee5)

      'TotalBillAmt# = Round#(TotalBillAmt# + ARCustRec(1).AcctBal)      'add past due here

      TotalBillAmt# = Round#(TotalBillAmt# + ARCustRec(1).IssuanceFee)

'011404
      PRINT #RptHandle, TAB(62); USING "#####.##"; TotalBillAmt#

'PRINT #RptHandle, TAB(62); USING "#####.##"; TotalBillAmt# - ARCustRec(1).AcctBal

      PRINT #RptHandle,
      PRINT #RptHandle,
      PRINT #RptHandle,
      PRINT #RptHandle, TAB(62); USING "#####.##"; TotalBillAmt#
      PRINT #RptHandle,
      PRINT #RptHandle, "~"

'dale
      GOSUB Post2Account

      LICENSE# = LICENSE# + 1
    END IF
    'here dale

    'IF DidCnt > 2 THEN EXIT FOR
  NEXT Cnt&
  PRINT #RptHandle, CHR$(12);
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi
  CLOSE         'Close all open files now
  
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  'KILL ReportFile$
  
  EXIT SUB
  
  
Post2Account:
  REDIM ARTransRec(1) AS ARTransRecType
  ARTransRecLen = LEN(ARTransRec(1))
  ARTransFile = FREEFILE
  OPEN "ARTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARTransFile LEN = ARTransRecLen
  NumOfTransRecs = LOF(ARTransFile) \ ARTransRecLen
  NextTransRec = NumOfTransRecs + 1
  Help$ = "Posting: " + LEFT$(ARCustRec(1).BILLNAME, 30)
  PrintHelp Help$
  
  ' Update Customer Information First
  ARCustRec(1).LICENSE = LTRIM$(STR$(LICENSE#))
  ARCustRec(1).VALID = Date2Num%(StoreExpireDate$)
  ARCustRec(1).IssueLicense$ = "N"
  
  PUT ARFile, CustomerNumber, ARCustRec(1)
  'unrem
  IF ChargeAcct$ = "Y" THEN
    
    ' Calc New Running Balance
    'CurBal# = ARCustRec(1).AcctBal

    'AcctBal# = TotalBillAmt#
    ARCustRec(1).LicBal = Round#(ARCustRec(1).LicBal + TotalBillAmt#)
    ARCustRec(1).AcctBal = Round#(ARCustRec(1).AcctBal + TotalBillAmt#)
    
    ' Post Transaction Record First
    ARTransRec(1).CustomerNumber = LTRIM$(STR$(CustomerNumber))
    ARTransRec(1).TransDate = Date2Num%(PostDate$)
    ARTransRec(1).TransAmount = TotalBillAmt#
    'ARTransRec(1).TransAmount = Round#(TotalBillAmt# - CurBal#) ???
    ARTransRec(1).TransType = 1 ' Type 1 = CHARGE
    ARTransRec(1).TransDesc = "LICENSE"
    ARTransRec(1).CashAmount = 0
    ARTransRec(1).ChkAmount = 0
    ARTransRec(1).BalanceAfterTrans = ARCustRec(1).AcctBal
    ARTransRec(1).ExtraRoom = ""
    ARTransRec(1).Posted2GL = "N"
    ARTransRec(1).NextTrans = 0
    ARTransRec(1).CatCodeRec1 = CatagoryRecord1
    
    PUT ARTransFile, NextTransRec, ARTransRec(1)
    'unrem
    IF ARCustRec(1).FirstTrans = 0 THEN
      ARCustRec(1).FirstTrans = NextTransRec
      ARCustRec(1).LastTrans = NextTransRec
      PUT ARFile, CustomerNumber, ARCustRec(1)
      'unrem
    ELSE
      Prev! = ARCustRec(1).LastTrans
      ARCustRec(1).LastTrans = NextTransRec
      PUT ARFile, CustomerNumber, ARCustRec(1)
      'unrem
      GET ARTransFile, Prev!, ARTransRec(1)
      ARTransRec(1).NextTrans = NextTransRec
      PUT ARTransFile, Prev!, ARTransRec(1)
      'unrem
    END IF
  END IF
  CLOSE ARTransFile
  RETURN
  
  
SelectLicenseOutput:
  LibName$ = "AR"
  ScrnName$ = "ARPRTLIC"
  
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
  
  REDIM Choice$(2, 2)           ' assume maximum of 99 codes
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Name"
  Choice$(2, 0) = "Number"
  Form$(1, 0) = "Name"
  Form$(5, 0) = "Y"
  Action = 1
  ShowCursor
  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%
  'printhelp help$
  Action = 1
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F3Key
      Help$ = " ** NOW PRINTING YOUR ALIGNMENT **"
      PrintHelp Help$
      Heading1$ = RTRIM$(Form$(6, 0))
      Heading2$ = RTRIM$(Form$(7, 0))
      Heading3$ = RTRIM$(Form$(8, 0))
      Heading4$ = RTRIM$(Form$(9, 0))
      
      Heading1 = LEN(Heading1$)
      Heading2 = LEN(Heading2$)
      Heading3 = LEN(Heading3$)
      Heading4 = LEN(Heading4$)
      
      IF LEN(Heading1) > 0 THEN tab1 = Heading1 / 2 ELSE tab1 = 0
      IF LEN(Heading2) > 0 THEN tab2 = Heading2 / 2 ELSE tab2 = 0
      IF LEN(Heading3) > 0 THEN Tab3 = Heading3 / 2 ELSE Tab3 = 0
      IF LEN(Heading4) > 0 THEN Tab4 = Heading4 / 2 ELSE Tab4 = 0
      
      ' Print Form Test
      LPRINT "TOP"
      FOR lcnt = 1 TO 4
        LPRINT ""
      NEXT lcnt
      LPRINT TAB(37 - tab1); Heading1$
      LPRINT TAB(37 - tab2); Heading2$
      LPRINT TAB(37 - Tab3); Heading3$
      LPRINT TAB(37 - Tab4); Heading4$
      LPRINT TAB(66); RIGHT$(Form$(2, 0), 4)
      LPRINT
      LPRINT TAB(11); "Name of Some Business"
      LPRINT TAB(11); "Address Line 1"; TAB(58); "########"
      LPRINT TAB(11); "Address Line 2"
      LPRINT TAB(11); "Address Line 3"
      LPRINT TAB(55); LEFT$(Form$(4, 0), 6) + RIGHT$(Form$(4, 0), 2);
      LPRINT TAB(64); LEFT$(Form$(2, 0), 6) + RIGHT$(Form$(2, 0), 2)
      LPRINT
      LPRINT
      LPRINT
      LPRINT TAB(11); STRING$(35, "X")
      LPRINT
      LPRINT
      LPRINT TAB(5); "XXXXXXXX"; TAB(15); STRING$(30, "X"); TAB(62); "XXXXX.XX"
      FOR lcnt = 24 TO 35
        LPRINT ""
      NEXT lcnt
      LPRINT TAB(62); "XXXXX.XX"
      LPRINT
      LPRINT
      LPRINT
      LPRINT TAB(62); "XXXXX.XX"
      LPRINT
      LPRINT "BOTTOM"
      Help$ = " Did It Line Up ??"
      PrintHelp Help$
      
    CASE F10Key
      
      Heading1$ = RTRIM$(Form$(6, 0))
      Heading2$ = RTRIM$(Form$(7, 0))
      Heading3$ = RTRIM$(Form$(8, 0))
      Heading4$ = RTRIM$(Form$(9, 0))
      
      Heading1 = LEN(Heading1$)
      Heading2 = LEN(Heading2$)
      Heading3 = LEN(Heading3$)
      Heading4 = LEN(Heading4$)
      
      IF LEN(Heading1) > 0 THEN tab1 = Heading1 / 2 ELSE tab1 = 0
      IF LEN(Heading2) > 0 THEN tab2 = Heading2 / 2 ELSE tab2 = 0
      IF LEN(Heading3) > 0 THEN Tab3 = Heading3 / 2 ELSE Tab3 = 0
      IF LEN(Heading4) > 0 THEN Tab4 = Heading4 / 2 ELSE Tab4 = 0
      Order$ = Form$(1, 0)
      LICENSE# = VAL(Form$(3, 0))
      IssueDate$ = LEFT$(Form$(4, 0), 6) + RIGHT$(Form$(4, 0), 2)
'corrects a bug in post date
      PostDate$ = Form$(4, 0)
'***************************
      StoreExpireDate$ = Form$(2, 0)
      ExpireDate$ = LEFT$(Form$(2, 0), 6) + RIGHT$(Form$(2, 0), 2)

      Year$ = RIGHT$(Form$(2, 0), 4)
      Year$ = QPTrim$(STR$(VAL(Year$) - 1))

      ChargeAcct$ = Form$(5, 0)
      DevSpec$ = "P"
      
      ' test here for items that need to be done
      IF LICENSE# <> 0 AND LEN(RTRIM$(ExpireDate$)) > 0 AND LEN(RTRIM$(IssueDate$)) > 0 THEN
        RETURN
      ELSE
        Action = 1
      END IF
      
    CASE EscKey
      Canceled$ = "Y"
      RETURN
    END SELECT
  LOOP
  RETURN
  
END SUB

SUB PrintPenRpt

  SHARED Choice$()

  ReportFile$ = "ARPENLST.PRN"  'Report File Name
  LPTPort% = 1
  FF$ = CHR$(12)
  MaxLines = 58
  LineCnt = 0
  CustCnt = 0

  ShowProcessingScrn "Creating Report"

  REDIM ARTransRec(1) AS ARTransRecType
  ARTransRecLen = LEN(ARTransRec(1))

  CustRecLen = LEN(ARCustRec(1))

  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle

  GOSUB PrintCustBalRptHeader

  ARTransFile = FREEFILE
  OPEN "ARTPENTR.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARTransFile LEN = ARTransRecLen
  TrNumRecs = LOF(ARTransFile) \ ARTransRecLen

  CUHandle = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS CUHandle LEN = CustRecLen

  FOR Cnt = 1 TO TrNumRecs
    GET #ARTransFile, Cnt, ARTransRec(1)
    TempNum$ = QPTrim$(ARTransRec(1).CustomerNumber)
    IF LEN(TempNum$) = 0 THEN
      GOTO NoPenalty
    END IF
    CustRec = VAL(TempNum$)
    GET #CUHandle, CustRec, ARCustRec(1)

    PRINT #RptHandle, USING "#####"; CustRec;
    PRINT #RptHandle, TAB(10); ARCustRec(1).BILLNAME;
    PRINT #RptHandle, TAB(50); USING "$$####,#.##"; ARTransRec(1).TransAmount
    CustCnt = CustCnt + 1
    TotalBal# = Round#(TotalBal# + ARTransRec(1).TransAmount)
    LineCnt = LineCnt + 1
    IF LineCnt > MaxLines THEN
      LineCnt = 1
      PRINT #RptHandle, CHR$(12)
      GOSUB PrintCustBalRptHeader
    END IF

NoPenalty:
    ShowPctComp Cnt, TrNumRecs
  NEXT
  GOSUB PrintCustBalRptEnding
  CLOSE         'Close all open files now

  EntryPoint = 1

  PrintRptFile "Penalty Report", ReportFile$, LPTPort%, RetCode%, EntryPoint

  KILL ReportFile$

  EXIT SUB

PrintCustBalRptHeader:
  Page = Page + 1
  PRINT #RptHandle, TAB(20); "Business License's : Penalty Transaction Listing"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
  PRINT #RptHandle, "Cust #"; TAB(10); "Customer Name"; TAB(50); "Penalty Amount"
  PRINT #RptHandle, STRING$(80, "=")
  LineCnt = 5
RETURN

PrintCustBalRptEnding:
  PRINT #RptHandle, STRING$(79, "-")
  PRINT #RptHandle, "Total Customers Printed: "; USING "#####"; CustCnt
  PRINT #RptHandle, "          Total Penalty: "; USING "$$#####,#.##"; TotalBal#
  PRINT #RptHandle, FF$
RETURN

END SUB

  '****************************************************************************
  'Rounds a double precision value to nearest hundreth
  '****************************************************************************
FUNCTION Round# (N#)
  Round# = INT(N# * 100 + .5000001) / 100
END FUNCTION

SUB SETFEE

  LOCATE , , 0
  ShowProcessingScrn ("Processing Fees")

  REDIM ARCatCodeRec(1) AS ARNewCatCodeRecType
  ARCatCodeRecLen = LEN(ARCatCodeRec(1))
  ARCatFile = FREEFILE
  OPEN "ARCODE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARCatFile LEN = ARCatCodeRecLen
  NumOFARCatRecs = LOF(ARCatFile) \ ARCatCodeRecLen
  
  OpenARCustFile NumOfArRecs, ARFile
  REDIM ARIdxRec(1) AS ARCustIDXRecType
  IdxFile = FREEFILE
  OPEN "ARCUST.IDX" FOR RANDOM SHARED AS #IdxFile LEN = LEN(ARIdxRec(1))
  
  IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN NumOfArRecs = LOF(IdxFile) \ LEN(ARIdxRec(1))
  
  FOR Cnt! = 1 TO NumOfArRecs
    IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN
      GET IdxFile, Cnt!, ARIdxRec(1)
      CustomerNumber = ARIdxRec(1).IDXRECORD
    ELSE
      CustomerNumber = Cnt!
    END IF

    GET ARFile, CustomerNumber, ARCustRec(1)
    ProRateFlag = False

    ARCustRec(1).Fee1 = 0
    ARCustRec(1).Fee2 = 0
    ARCustRec(1).Fee3 = 0
    ARCustRec(1).Fee4 = 0
    ARCustRec(1).Fee5 = 0
    PUT ARFile, CustomerNumber, ARCustRec(1)
    
    'Help$ = ARCustRec(1).BILLNAME
    'PrintHelp Help$
    'Calc Fee Amount Here

    IF ARCustRec(1).Deleted = "Y" OR ARCustRec(1).Inactive = "Y" THEN
      GOTO SkipEm
    END IF

    ProRate# = ARCustRec(1).ProRate
    IF ProRate# > 100 OR ProRate# < 0 THEN
      ProRate# = 100
    END IF
    ProAT# = ProRate#
    IF ProRate# < 100 THEN
      ProRateFlag = True
      ProRate# = Round#(ProRate# * .01)
    END IF


    CATCODE$ = QPTrim$(ARCustRec(1).BILLCAT1)
    IF LEN(CATCODE$) > 0 THEN
      FOR Snt& = 1 TO NumOFARCatRecs
        GET ARCatFile, Snt&, ARCatCodeRec(1)
        IF QPTrim$(ARCatCodeRec(1).CATCODE) = CATCODE$ THEN
          ARCustRec(1).Desc1 = ARCatCodeRec(1).CODEDESC         'Reset Code Descriptions
          IF ARCatCodeRec(1).CodeType = "F" THEN
            ARCustRec(1).Fee1 = ARCatCodeRec(1).Fee
            FeeAmt# = ARCatCodeRec(1).Fee
            GOSUB PrintLine
            GOTO C2
          END IF
          IF ARCatCodeRec(1).CodeType = "M" THEN
            Mult = ARCustRec(1).REV1
            ARCustRec(1).Fee1 = Mult * ARCatCodeRec(1).Fee
            FeeAmt# = Mult * ARCatCodeRec(1).Fee
            GOSUB PrintLine
            GOTO C2
          END IF
          IF ARCatCodeRec(1).CodeType = "S" THEN
            Revenue# = ARCustRec(1).REV1
            IF Revenue# <= ARCatCodeRec(1).Recpt1 THEN
              ARCustRec(1).Fee1 = ARCatCodeRec(1).BaseAmt1 + (ARCatCodeRec(1).Percent1 / 100) * (Revenue# - ARCatCodeRec(1).Maximum1)
              FeeAmt# = ARCatCodeRec(1).BaseAmt1 + (ARCatCodeRec(1).Percent1 / 100) * (Revenue# - ARCatCodeRec(1).Maximum1)
              GOSUB PrintLine
              GOTO C2
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt2 THEN
              ARCustRec(1).Fee1 = ARCatCodeRec(1).BaseAmt2 + (ARCatCodeRec(1).Percent2 / 100) * (Revenue# - ARCatCodeRec(1).Maximum2)
              FeeAmt# = ARCatCodeRec(1).BaseAmt2 + (ARCatCodeRec(1).Percent2 / 100) * (Revenue# - ARCatCodeRec(1).Maximum2)
              GOSUB PrintLine
              GOTO C2
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt3 THEN
              ARCustRec(1).Fee1 = ARCatCodeRec(1).BaseAmt3 + (ARCatCodeRec(1).Percent3 / 100) * (Revenue# - ARCatCodeRec(1).Maximum3)
              FeeAmt# = ARCatCodeRec(1).BaseAmt3 + (ARCatCodeRec(1).Percent3 / 100) * (Revenue# - ARCatCodeRec(1).Maximum3)
              GOSUB PrintLine
              GOTO C2
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt4 THEN
              ARCustRec(1).Fee1 = ARCatCodeRec(1).BaseAmt4 + (ARCatCodeRec(1).Percent4 / 100) * (Revenue# - ARCatCodeRec(1).Maximum4)
              FeeAmt# = ARCatCodeRec(1).BaseAmt4 + (ARCatCodeRec(1).Percent4 / 100) * (Revenue# - ARCatCodeRec(1).Maximum4)
              GOSUB PrintLine
              GOTO C2
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt5 THEN
              ARCustRec(1).Fee1 = ARCatCodeRec(1).BaseAmt5 + (ARCatCodeRec(1).Percent5 / 100) * (Revenue# - ARCatCodeRec(1).Maximum5)
              FeeAmt# = ARCatCodeRec(1).BaseAmt5 + (ARCatCodeRec(1).Percent5 / 100) * (Revenue# - ARCatCodeRec(1).Maximum5)
              GOSUB PrintLine
              GOTO C2
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt6 THEN
              ARCustRec(1).Fee1 = ARCatCodeRec(1).BaseAmt6 + (ARCatCodeRec(1).Percent6 / 100) * (Revenue# - ARCatCodeRec(1).Maximum6)
              FeeAmt# = ARCatCodeRec(1).BaseAmt6 + (ARCatCodeRec(1).Percent6 / 100) * (Revenue# - ARCatCodeRec(1).Maximum6)
              GOSUB PrintLine
              GOTO C2
            END IF
            
          END IF
        END IF  'End Test for Code
      NEXT Snt&
    ELSE
      ARCustRec(1).Fee1 = 0
    END IF      'End Test for Cat 1
    
    
C2:             'Catagory #2
    
    CATCODE$ = QPTrim$(ARCustRec(1).BILLCAT2)
    IF LEN(CATCODE$) > 0 THEN
      FOR Snt& = 1 TO NumOFARCatRecs
        GET ARCatFile, Snt&, ARCatCodeRec(1)
        IF QPTrim$(ARCatCodeRec(1).CATCODE) = CATCODE$ THEN
          IF ARCatCodeRec(1).CodeType = "F" THEN
            ARCustRec(1).Fee2 = ARCatCodeRec(1).Fee
            FeeAmt# = ARCatCodeRec(1).Fee
            GOSUB PrintLine
            GOTO C3
          END IF
          IF ARCatCodeRec(1).CodeType = "M" THEN
            Mult = ARCustRec(1).REV2
            ARCustRec(1).Fee2 = Mult * ARCatCodeRec(1).Fee
            FeeAmt# = Mult * ARCatCodeRec(1).Fee
            GOSUB PrintLine
            GOTO C3
          END IF
          IF ARCatCodeRec(1).CodeType = "S" THEN
            Revenue# = ARCustRec(1).REV2
            IF Revenue# <= ARCatCodeRec(1).Recpt1 THEN
              ARCustRec(1).Fee2 = ARCatCodeRec(1).BaseAmt1 + (ARCatCodeRec(1).Percent1 / 100) * (Revenue# - ARCatCodeRec(1).Maximum1)
              FeeAmt# = ARCatCodeRec(1).BaseAmt1 + (ARCatCodeRec(1).Percent1 / 100) * (Revenue# - ARCatCodeRec(1).Maximum1)
              GOSUB PrintLine
              GOTO C3
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt2 THEN
              ARCustRec(1).Fee2 = ARCatCodeRec(1).BaseAmt2 + (ARCatCodeRec(1).Percent2 / 100) * (Revenue# - ARCatCodeRec(1).Maximum2)
              FeeAmt# = ARCatCodeRec(1).BaseAmt2 + (ARCatCodeRec(1).Percent2 / 100) * (Revenue# - ARCatCodeRec(1).Maximum2)
              GOSUB PrintLine
              GOTO C3
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt3 THEN
              ARCustRec(1).Fee2 = ARCatCodeRec(1).BaseAmt3 + (ARCatCodeRec(1).Percent3 / 100) * (Revenue# - ARCatCodeRec(1).Maximum3)
              FeeAmt# = ARCatCodeRec(1).BaseAmt3 + (ARCatCodeRec(1).Percent3 / 100) * (Revenue# - ARCatCodeRec(1).Maximum3)
              GOSUB PrintLine
              GOTO C3
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt4 THEN
              ARCustRec(1).Fee2 = ARCatCodeRec(1).BaseAmt4 + (ARCatCodeRec(1).Percent4 / 100) * (Revenue# - ARCatCodeRec(1).Maximum4)
              FeeAmt# = ARCatCodeRec(1).BaseAmt4 + (ARCatCodeRec(1).Percent4 / 100) * (Revenue# - ARCatCodeRec(1).Maximum4)
              GOSUB PrintLine
              GOTO C3
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt5 THEN
              ARCustRec(1).Fee2 = ARCatCodeRec(1).BaseAmt5 + (ARCatCodeRec(1).Percent5 / 100) * (Revenue# - ARCatCodeRec(1).Maximum5)
              FeeAmt# = ARCatCodeRec(1).BaseAmt5 + (ARCatCodeRec(1).Percent5 / 100) * (Revenue# - ARCatCodeRec(1).Maximum5)
              GOSUB PrintLine
              GOTO C3
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt6 THEN
              ARCustRec(1).Fee2 = ARCatCodeRec(1).BaseAmt6 + (ARCatCodeRec(1).Percent6 / 100) * (Revenue# - ARCatCodeRec(1).Maximum6)
              FeeAmt# = ARCatCodeRec(1).BaseAmt6 + (ARCatCodeRec(1).Percent6 / 100) * (Revenue# - ARCatCodeRec(1).Maximum6)
              GOSUB PrintLine
              GOTO C3
            END IF
          END IF
        END IF  'End Test for Code
      NEXT Snt&
    ELSE
      ARCustRec(1).Fee2 = 0
    END IF      'End Test for Cat 1
    
    
C3:
    CATCODE$ = QPTrim$(ARCustRec(1).BILLCAT3)
    IF LEN(CATCODE$) > 0 THEN
      FOR Snt& = 1 TO NumOFARCatRecs
        GET ARCatFile, Snt&, ARCatCodeRec(1)
        IF QPTrim$(ARCatCodeRec(1).CATCODE) = CATCODE$ THEN
          IF ARCatCodeRec(1).CodeType = "F" THEN
            ARCustRec(1).Fee3 = ARCatCodeRec(1).Fee
            FeeAmt# = ARCatCodeRec(1).Fee
            GOSUB PrintLine
            GOTO c4
          END IF
          IF ARCatCodeRec(1).CodeType = "M" THEN
            Mult = ARCustRec(1).REV3
            ARCustRec(1).Fee3 = Mult * ARCatCodeRec(1).Fee
            FeeAmt# = Mult * ARCatCodeRec(1).Fee
            GOSUB PrintLine
            GOTO c4
          END IF
          IF ARCatCodeRec(1).CodeType = "S" THEN
            Revenue# = ARCustRec(1).REV3
            IF Revenue# <= ARCatCodeRec(1).Recpt1 THEN
              ARCustRec(1).Fee3 = ARCatCodeRec(1).BaseAmt1 + (ARCatCodeRec(1).Percent1 / 100) * (Revenue# - ARCatCodeRec(1).Maximum1)
              FeeAmt# = ARCatCodeRec(1).BaseAmt1 + (ARCatCodeRec(1).Percent1 / 100) * (Revenue# - ARCatCodeRec(1).Maximum1)
              GOSUB PrintLine
              GOTO c4
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt2 THEN
              ARCustRec(1).Fee3 = ARCatCodeRec(1).BaseAmt2 + (ARCatCodeRec(1).Percent2 / 100) * (Revenue# - ARCatCodeRec(1).Maximum2)
              FeeAmt# = ARCatCodeRec(1).BaseAmt2 + (ARCatCodeRec(1).Percent2 / 100) * (Revenue# - ARCatCodeRec(1).Maximum2)
              GOSUB PrintLine
              GOTO c4
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt3 THEN
              ARCustRec(1).Fee3 = ARCatCodeRec(1).BaseAmt3 + (ARCatCodeRec(1).Percent3 / 100) * (Revenue# - ARCatCodeRec(1).Maximum3)
              FeeAmt# = ARCatCodeRec(1).BaseAmt3 + (ARCatCodeRec(1).Percent3 / 100) * (Revenue# - ARCatCodeRec(1).Maximum3)
              GOSUB PrintLine
              GOTO c4
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt4 THEN
              ARCustRec(1).Fee3 = ARCatCodeRec(1).BaseAmt4 + (ARCatCodeRec(1).Percent4 / 100) * (Revenue# - ARCatCodeRec(1).Maximum4)
              FeeAmt# = ARCatCodeRec(1).BaseAmt4 + (ARCatCodeRec(1).Percent4 / 100) * (Revenue# - ARCatCodeRec(1).Maximum4)
              GOSUB PrintLine
              GOTO c4
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt5 THEN
              ARCustRec(1).Fee3 = ARCatCodeRec(1).BaseAmt5 + (ARCatCodeRec(1).Percent5 / 100) * (Revenue# - ARCatCodeRec(1).Maximum5)
              FeeAmt# = ARCatCodeRec(1).BaseAmt5 + (ARCatCodeRec(1).Percent5 / 100) * (Revenue# - ARCatCodeRec(1).Maximum5)
              GOSUB PrintLine
              GOTO c4
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt6 THEN
              ARCustRec(1).Fee3 = ARCatCodeRec(1).BaseAmt6 + (ARCatCodeRec(1).Percent6 / 100) * (Revenue# - ARCatCodeRec(1).Maximum6)
              FeeAmt# = ARCatCodeRec(1).BaseAmt6 + (ARCatCodeRec(1).Percent6 / 100) * (Revenue# - ARCatCodeRec(1).Maximum6)
              GOSUB PrintLine
              GOTO c4
            END IF
          END IF
        END IF  'End Test for Code
      NEXT Snt&
    ELSE
      ARCustRec(1).Fee3 = 0
    END IF      'End Test for Cat 3
    
c4:
    CATCODE$ = QPTrim$(ARCustRec(1).BILLCAT4)
    IF LEN(CATCODE$) > 0 THEN
      FOR Snt& = 1 TO NumOFARCatRecs
        GET ARCatFile, Snt&, ARCatCodeRec(1)
        IF QPTrim$(ARCatCodeRec(1).CATCODE) = CATCODE$ THEN
          IF ARCatCodeRec(1).CodeType = "F" THEN
            ARCustRec(1).Fee4 = ARCatCodeRec(1).Fee
            FeeAmt# = ARCatCodeRec(1).Fee
            GOSUB PrintLine
            GOTO c5
          END IF
          IF ARCatCodeRec(1).CodeType = "M" THEN
            Mult = ARCustRec(1).REV4
            ARCustRec(1).Fee4 = Mult * ARCatCodeRec(1).Fee
            FeeAmt# = Mult * ARCatCodeRec(1).Fee
            GOSUB PrintLine
            GOTO c5
          END IF
          IF ARCatCodeRec(1).CodeType = "S" THEN
            Revenue# = ARCustRec(1).REV4
            IF Revenue# <= ARCatCodeRec(1).Recpt1 THEN
              ARCustRec(1).Fee4 = ARCatCodeRec(1).BaseAmt1 + (ARCatCodeRec(1).Percent1 / 100) * (Revenue# - ARCatCodeRec(1).Maximum1)
              FeeAmt# = ARCatCodeRec(1).BaseAmt1 + (ARCatCodeRec(1).Percent1 / 100) * (Revenue# - ARCatCodeRec(1).Maximum1)
              GOSUB PrintLine
              GOTO c5
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt2 THEN
              ARCustRec(1).Fee4 = ARCatCodeRec(1).BaseAmt2 + (ARCatCodeRec(1).Percent2 / 100) * (Revenue# - ARCatCodeRec(1).Maximum2)
              FeeAmt# = ARCatCodeRec(1).BaseAmt2 + (ARCatCodeRec(1).Percent2 / 100) * (Revenue# - ARCatCodeRec(1).Maximum2)
              GOSUB PrintLine
              GOTO c5
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt3 THEN
              ARCustRec(1).Fee4 = ARCatCodeRec(1).BaseAmt3 + (ARCatCodeRec(1).Percent3 / 100) * (Revenue# - ARCatCodeRec(1).Maximum3)
              FeeAmt# = ARCatCodeRec(1).BaseAmt3 + (ARCatCodeRec(1).Percent3 / 100) * (Revenue# - ARCatCodeRec(1).Maximum3)
              GOSUB PrintLine
              GOTO c5
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt4 THEN
              ARCustRec(1).Fee4 = ARCatCodeRec(1).BaseAmt4 + (ARCatCodeRec(1).Percent4 / 100) * (Revenue# - ARCatCodeRec(1).Maximum4)
              FeeAmt# = ARCatCodeRec(1).BaseAmt4 + (ARCatCodeRec(1).Percent4 / 100) * (Revenue# - ARCatCodeRec(1).Maximum4)
              GOSUB PrintLine
              GOTO c5
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt5 THEN
              ARCustRec(1).Fee4 = ARCatCodeRec(1).BaseAmt5 + (ARCatCodeRec(1).Percent5 / 100) * (Revenue# - ARCatCodeRec(1).Maximum5)
              FeeAmt# = ARCatCodeRec(1).BaseAmt5 + (ARCatCodeRec(1).Percent5 / 100) * (Revenue# - ARCatCodeRec(1).Maximum5)
              GOSUB PrintLine
              GOTO c5
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt6 THEN
              ARCustRec(1).Fee4 = ARCatCodeRec(1).BaseAmt6 + (ARCatCodeRec(1).Percent6 / 100) * (Revenue# - ARCatCodeRec(1).Maximum6)
              FeeAmt# = ARCatCodeRec(1).BaseAmt6 + (ARCatCodeRec(1).Percent6 / 100) * (Revenue# - ARCatCodeRec(1).Maximum6)
              GOSUB PrintLine
              GOTO c5
            END IF
            
          END IF
        END IF  'End Test for Code
      NEXT Snt&
    ELSE
      ARCustRec(1).Fee4 = 0
    END IF      'End Test for Cat 1
    
c5:
    CATCODE$ = QPTrim$(ARCustRec(1).BILLCAT5)
    IF LEN(CATCODE$) > 0 THEN
      FOR Snt& = 1 TO NumOFARCatRecs
        GET ARCatFile, Snt&, ARCatCodeRec(1)
        IF QPTrim$(ARCatCodeRec(1).CATCODE) = CATCODE$ THEN
          IF ARCatCodeRec(1).CodeType = "F" THEN
            ARCustRec(1).Fee5 = ARCatCodeRec(1).Fee
            FeeAmt# = ARCatCodeRec(1).Fee
            GOSUB PrintLine
            GOTO SkipEm
          END IF
          IF ARCatCodeRec(1).CodeType = "M" THEN
            Mult = ARCustRec(1).REV5
            ARCustRec(1).Fee5 = Mult * ARCatCodeRec(1).Fee
            FeeAmt# = Mult * ARCatCodeRec(1).Fee
            GOSUB PrintLine
            GOTO SkipEm
          END IF
          IF ARCatCodeRec(1).CodeType = "S" THEN
            Revenue# = ARCustRec(1).REV5
            IF Revenue# <= ARCatCodeRec(1).Recpt1 THEN
              ARCustRec(1).Fee5 = ARCatCodeRec(1).BaseAmt1 + (ARCatCodeRec(1).Percent1 / 100) * (Revenue# - ARCatCodeRec(1).Maximum1)
              FeeAmt# = ARCatCodeRec(1).BaseAmt1 + (ARCatCodeRec(1).Percent1 / 100) * (Revenue# - ARCatCodeRec(1).Maximum1)
              GOSUB PrintLine
              GOTO SkipEm
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt2 THEN
              ARCustRec(1).Fee5 = ARCatCodeRec(1).BaseAmt2 + (ARCatCodeRec(1).Percent2 / 100) * (Revenue# - ARCatCodeRec(1).Maximum2)
              FeeAmt# = ARCatCodeRec(1).BaseAmt2 + (ARCatCodeRec(1).Percent2 / 100) * (Revenue# - ARCatCodeRec(1).Maximum2)
              GOSUB PrintLine
              GOTO SkipEm
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt3 THEN
              ARCustRec(1).Fee5 = ARCatCodeRec(1).BaseAmt3 + (ARCatCodeRec(1).Percent3 / 100) * (Revenue# - ARCatCodeRec(1).Maximum3)
              FeeAmt# = ARCatCodeRec(1).BaseAmt3 + (ARCatCodeRec(1).Percent3 / 100) * (Revenue# - ARCatCodeRec(1).Maximum3)
              GOSUB PrintLine
              GOTO SkipEm
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt4 THEN
              ARCustRec(1).Fee5 = ARCatCodeRec(1).BaseAmt4 + (ARCatCodeRec(1).Percent4 / 100) * (Revenue# - ARCatCodeRec(1).Maximum4)
              FeeAmt# = ARCatCodeRec(1).BaseAmt4 + (ARCatCodeRec(1).Percent4 / 100) * (Revenue# - ARCatCodeRec(1).Maximum4)
              GOSUB PrintLine
              GOTO SkipEm
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt5 THEN
              ARCustRec(1).Fee5 = ARCatCodeRec(1).BaseAmt5 + (ARCatCodeRec(1).Percent5 / 100) * (Revenue# - ARCatCodeRec(1).Maximum5)
              FeeAmt# = ARCatCodeRec(1).BaseAmt5 + (ARCatCodeRec(1).Percent5 / 100) * (Revenue# - ARCatCodeRec(1).Maximum5)
              GOSUB PrintLine
              GOTO SkipEm
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt6 THEN
              ARCustRec(1).Fee5 = ARCatCodeRec(1).BaseAmt6 + (ARCatCodeRec(1).Percent6 / 100) * (Revenue# - ARCatCodeRec(1).Maximum6)
              FeeAmt# = ARCatCodeRec(1).BaseAmt6 + (ARCatCodeRec(1).Percent6 / 100) * (Revenue# - ARCatCodeRec(1).Maximum6)
              GOSUB PrintLine
              GOTO SkipEm
            END IF
            
          END IF
        END IF  'End Test for Code
      NEXT Snt&
    ELSE
      ARCustRec(1).Fee5 = 0
    END IF      'End Test for Cat 1
    
SkipEm:

    'ARCustRec(1).LicBal = Round#(ARCustRec(1).LicBal + FeeAmt#)
    'ARCustRec(1).AcctBal = Round#(ARCustRec(1).AcctBal + FeeAmt#)
    PUT ARFile, CustomerNumber, ARCustRec(1)
    ShowPctComp Cnt!, NumOfArRecs
  NEXT Cnt!
  
  GOSUB PrintSetFeeRptEnding
  CLOSE         'Close all open files now
  
  EXIT SUB
  
  
PrintLine:
RETURN
  
PrintSetFeeRptHeader:
RETURN
  
PrintSetFeeRptEnding:
RETURN

  
  
END SUB

SUB ShowNoCode
  Ok = MsgBox("AR", "ARNOCODE")
END SUB

SUB ShowPctComp (BYVAL RecNo, BYVAL NumOfRecs) STATIC
  RSET PctC(1) = QPStrI$(INT((RecNo / NumOfRecs) * 100))
  QPrintRC PctC(1), 13, 40, Cnf.HiLite
  '  QPrintRC STR$(FRE("")), 25, 1, Cnf.HiLite
END SUB

SUB ShowProcessingScrn (RptTitle$)

  TitleRow = 9
  TitleCol = 40 - (LEN(RptTitle$) \ 2) + 1

  'CursorOff
  'BlockClear
  LibFile2Scrn "AR.QSL", "MENUBAK", MonoCode, -1, ErrorCode
  LibFile2Scrn "AR.QSL", "PRORPT", MonoCode, -1, ErrorCode
  QPrintRC RptTitle$, TitleRow, TitleCol, 126
  QPrintRC "Processing:    % Completed.", 13, 28, Cnf.HiLite

END SUB

