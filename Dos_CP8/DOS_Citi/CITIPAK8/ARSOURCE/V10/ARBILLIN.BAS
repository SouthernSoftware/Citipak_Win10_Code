DEFINT A-Z
DECLARE SUB PrintLicenses ()
DECLARE SUB LicenseRegister ()
DECLARE SUB OpenARCustFile (NumOfArRecs%, ARFile%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB QPrint (X$, Colr%, page%)
DECLARE SUB QPrintRC (T$, r%, c%, clr%)
DECLARE SUB ARAPPLICATION ()
DECLARE SUB ShowNoCode ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB printhelp (H$)
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
  
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'AR.bi'                        'A/R FILE LAYOUTS
  '$INCLUDE: 'GL.bi'
  
  CONST False = 0, True = NOT False
  DIM SHARED ARCustRec(1) AS ARCustRecType
  DIM SHARED ARCustIdxRec(1) AS ARCustIDXRecType
  
  TYPE Struct
    who AS STRING * 14
    RecNum AS INTEGER
  END TYPE
  
  STACK 8000
  
  '--Dim the choice array to the number of menu items
  REDIM Mchoice$(1 TO 5)
  
  Mchoice$(1) = "Print Applications/Renewals"
  Mchoice$(2) = "Licenses to Print Register"
  Mchoice$(3) = "Print License Forms"
  Mchoice$(4) = "Exit to OS"
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 17   'limit the box length to go no lower than line 20
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  '--Find max menu width
  FOR Cnt = 1 TO UBOUND(Mchoice$)
    TLen = LEN(Mchoice$(Cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT
  
  '--Center Menu within Screen
  Row = ((25 - (UBOUND(Mchoice$))) \ 2)
  Col = ((80 - MaxLen) \ 2)
  help$ = "Billing Menu"
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    LibFile2Scrn "AR.QSL", "MENUBAK", MonoCode, -1, ErrorCode
    
    TitleBox 3, Col, MaxLen + 3, " Issue License ", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    'PrintTitle user$
    'printhelp help$
    
    ShowCursor
    
    VertMenu Mchoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      ARAPPLICATION
    CASE 2
      LicenseRegister
    CASE 3
      PrintLicenses
    CASE 4
      END
    END SELECT
  LOOP
  RUN "armenu"

SUB ARAPPLICATION
  
  SHARED Choice$()
  
  LibName$ = "AR"
  ScrnName$ = "ARAPPPRN"
  help$ = "Print Applications!!!"
  LOCATE 1, 1, 0
  
  REDIM ARCatCodeRec(1) AS ARCatCodeRecType
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
  
  FirstTime = True
  REM check for code file
  
  DO
    
MainBody:
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      LSET Form$(1, 0) = "ALL"
      LSET Form$(2, 0) = RIGHT$(DATE$, 4)
      LSET Form$(3, 0) = "Custom"
      Action = 2
    END IF

    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      IF ASC(LEFT$(Form$(1, 0), 1)) = 32 THEN
        Done = False
        COLOR 12
        LOCATE 9, 34
        PRINT "Catagory Not Set"
        PRINT CHR$(7);
        COLOR 15
        SLEEP 3: LOCATE 9, 34
        PRINT STRING$(23, 32)
        Action = 1
        Frm(1).FldNo = 1
        GOTO MainBody
      END IF
      IF ASC(LEFT$(Form$(2, 0), 1)) = 32 THEN
        Done = False
        COLOR 12
        LOCATE 10, 34: PRINT "Year Not Set": PRINT CHR$(7);
        SLEEP 3: LOCATE 10, 34
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
        PRINT "Form Type Not Set": PRINT CHR$(7);
        SLEEP 3: LOCATE 11, 34
        COLOR 15
        PRINT STRING$(23, 32)
        Action = 1
        Frm(1).FldNo = 3
        GOTO MainBody
      END IF
      Code$ = LEFT$(Form$(1, 0), 3)
      IF LEFT$(Form$(3, 0), 6) = "Custom" THEN GOSUB PrintCustom
      IF LEFT$(Form$(3, 0), 8) = "Standard" THEN GOSUB PrintStandard
      
    CASE EscKey
      EXIT SUB
    END SELECT
    
  LOOP
  
  
PrintCustom:
  
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
        GOSUB PrintCustomForm
'Testing
'        IF PrnCnt& > 0 THEN
'          EXIT FOR
'        END IF
      ELSE
        IF LEFT$(ARCustRec(1).BILLCAT, 3) = Code$ THEN
          GOSUB PrintCustomForm
        END IF
      END IF
      
    END IF
  NEXT Cnt
  
  CLOSE         'Close all open files now
  EntryPoint = 1
  Header$ = "Applications"
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  KILL ReportFile$
  EXIT SUB
  
PrintCustomForm:
  ''$INCLUDE: 'ARHEMIWY.BI'
  '$INCLUDE: 'ARDUMFRI.BI'
  ''$INCLUDE: 'ARRIVRSD.BI'
  ''$INCLUDE: 'ARJACKSN.BI'
  TotalCust = TotalCust + 1
'  PrnCnt& = PrnCnt& + 1

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
        IF LEFT$(ARCustRec(1).BILLCAT, 3) = Code$ THEN
          GOSUB PrintSTDForm
        END IF
      END IF
      
    END IF
  NEXT Cnt
  
  CLOSE         'Close all open files now
  EntryPoint = 1
  Header$ = "Applications"
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  KILL ReportFile$
  EXIT SUB
  
PrintSTDForm:
  Heading1$ = RTRIM$(Form$(4, 0))
  Heading2$ = RTRIM$(Form$(5, 0))
  Heading3$ = RTRIM$(Form$(6, 0))
  
  Heading1 = LEN(Heading1$)
  Heading2 = LEN(Heading2$)
  Heading3 = LEN(Heading3$)
  
  IF LEN(Heading1) > 0 THEN tab1 = Heading1 / 2 ELSE tab1 = 0
  IF LEN(Heading2) > 0 THEN tab2 = Heading2 / 2 ELSE tab2 = 0
  IF LEN(Heading3) > 0 THEN Tab3 = Heading3 / 2 ELSE Tab3 = 0
  IF LEN(Heading4) > 0 THEN Tab4 = Heading4 / 2 ELSE Tab4 = 0
  
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
  FOR LL = 1 TO 5
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
  PRINT #RptHandle, TAB(11); ARCustRec(1).ADDRESS2
  PRINT #RptHandle, TAB(11); RTRIM$(ARCustRec(1).CITY); "  "; ARCustRec(1).STATE; " "; ARCustRec(1).ZIPCODE
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(11); ARCustRec(1).CustName
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(5); ARCustRec(1).BILLCAT;
  PRINT #RptHandle, TAB(15); CatagoryDesc$;
  PRINT #RptHandle, TAB(62); USING "#####.##"; ARCustRec(1).FeeAmt
  FOR LCnt = 24 TO 35
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
  
  
  
END SUB

SUB LicenseRegister
  
  SHARED Choice$()
  
  
  ReportFile$ = "ARLICREG.PRN"  'Report File Name
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
  
  GOSUB SelectLicRegOutput: IF Canceled$ = "Y" THEN EXIT SUB
  GOSUB GetReportLicRegInformation
  
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
  
  FOR Cnt! = 1 TO NumOfArRecs
    IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN
      GET IdxFile, Cnt!, ARIdxRec(1)
      GET ARFile, ARIdxRec(1).IDXRECORD, ARCustRec(1)
      CustomerNumber = ARIdxRec(1).IDXRECORD
    ELSE
      GET ARFile, Cnt!, ARCustRec(1)
      CustomerNumber = Cnt!
    END IF
    IF ARCustRec(1).IssueLicense = "Y" THEN
      
      IF LineCnt >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintLicRegRptHeader
      END IF
      PRINT #RptHandle, CustomerNumber; TAB(10); ARCustRec(1).BILLNAME; TAB(70); USING "$$####,#.##"; ARCustRec(1).FeeAmt
      
      TotalFee# = TotalFee# + ARCustRec(1).FeeAmt
      LineCnt = LineCnt + 1
    END IF
    
  NEXT Cnt!
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
  page = page + 1
  PRINT #RptHandle, TAB(18); "Business License System : License's to Print Register"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; page
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
  
GetReportLicRegInformation:
  
  
  RETURN
  
  
  
  
  
  
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

SUB OpenARCustFile (NumOfArRecs, ARFile)
  CLOSE ARFile
  
  ARCustRecLen = LEN(ARCustRec(1))
  ARFile = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARFile LEN = ARCustRecLen
  NumOfArRecs = LOF(ARFile) \ ARCustRecLen
  
END SUB

SUB PrintLicenses

  SHARED Choice$()

  ReportFile$ = "ARLICENS.PRN"  'Report File Name
  CommaFmt$ = "########,.##"    'format takes 13 chars
  TotalFmt$ = "#########,.##"   'format takes 14 chars
  SumLine$ = STRING$(13, "-")   'column summary line
  DivLine$ = STRING$(77, "-")   'dashed line
  DivLine2$ = STRING$(77, "=")  'Double Line
  FF$ = CHR$(12)
  LPTPort% = 1
  MaxLines = 53
  LineCnt = 0
  TotDr# = 0
  TotCr# = 0
  'size = 2500
  Start = 1     'start at array element 1
  'Dir = 0       'sort direction - use anything else for descending
  'SSize = 16    'total size of each TYPE element
  MOff = 0      'offset into the TYPE for the key element
  MSize = 16    'size of the key element - coded as follows:
  '   -1 = integer
  '   -2 = long integer
  '   -3 = single precision
  '   -4 = double precision
  '   +N = TYPE array/fixed-length string of length N
  
  'REDIM array(1 TO size) AS Struct
  
  GOSUB SelectLicenseOutput: IF Canceled$ = "Y" THEN EXIT SUB
  help$ = "Now Printing the License's"
  'printhelp help$
  
  ReportFile$ = "ARLIC.PRN"
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  'PRINT #RptHandle, CHR$(27); CHR$(58); ' oki 320 12 cpi
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
    IF ARCustRec(1).IssueLicense = "Y" THEN
      
      'Get Catagory Desc First
      CatagoryDesc$ = ""
      REDIM ARCatCodeRec(1) AS ARCatCodeRecType
      ARCatCodeRecLen = LEN(ARCatCodeRec(1))
      ARCatFile = FREEFILE
      OPEN "ARCODE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARCatFile LEN = ARCatCodeRecLen
      NumOFARCatRecs = LOF(ARCatFile) \ ARCatCodeRecLen
      CatagoryRecord = 0
      FOR Cnt = 1 TO NumOFARCatRecs
        GET ARCatFile, Cnt, ARCatCodeRec(1)
        IF RTRIM$(ARCustRec(1).BILLCAT) = RTRIM$(ARCatCodeRec(1).CATCODE) THEN
          CatagoryDesc$ = ARCatCodeRec(1).CODEDESC
          CatagoryRecord = Cnt
          EXIT FOR
        END IF
      NEXT Cnt
      CLOSE ARCatFile
      
      
      FOR LL = 1 TO 5
        PRINT #RptHandle, ""
      NEXT
        PRINT #RptHandle, TAB(37 - tab1); Heading1$
        PRINT #RptHandle, TAB(37 - tab2); Heading2$
        PRINT #RptHandle, TAB(37 - Tab3); Heading3$
        PRINT #RptHandle, TAB(37 - Tab4); Heading4$
        PRINT #RptHandle, TAB(66); Year$
        PRINT #RptHandle,
        PRINT #RptHandle, TAB(11); ARCustRec(1).BILLNAME
        PRINT #RptHandle, TAB(11); ARCustRec(1).Address1; TAB(58); USING "########"; LICENSE#
        PRINT #RptHandle, TAB(11); ARCustRec(1).ADDRESS2
        PRINT #RptHandle, TAB(11); RTRIM$(ARCustRec(1).CITY); "  "; ARCustRec(1).STATE; " "; ARCustRec(1).ZIPCODE
        PRINT #RptHandle, TAB(55); LEFT$(IssueDate$, 6); RIGHT$(IssueDate$, 2);
        PRINT #RptHandle, TAB(64); LEFT$(ExpireDate$, 6); RIGHT$(ExpireDate$, 2)
        PRINT #RptHandle,
        PRINT #RptHandle,
        PRINT #RptHandle,
        PRINT #RptHandle, TAB(11); ARCustRec(1).CustName
        PRINT #RptHandle,
        PRINT #RptHandle,
        PRINT #RptHandle, TAB(5); ARCustRec(1).BILLCAT;
        PRINT #RptHandle, TAB(15); CatagoryDesc$;
        PRINT #RptHandle, TAB(62); USING "#####.##"; ARCustRec(1).FeeAmt
        FOR LCnt = 24 TO 31: PRINT #RptHandle, "": NEXT LCnt
          IF ARCustRec(1).AcctBal >= .01 THEN
            PRINT #RptHandle, ""
            'PRINT #RptHandle, TAB(15); "Previous Balance Owed";
            'PRINT #RptHandle, TAB(62); USING "#####.##"; ARCustRec(1).AcctBal
          ELSE
            PRINT #RptHandle, ""
          END IF
          FOR LCnt = 33 TO 35: PRINT #RptHandle, "": NEXT LCnt
            'IF ARCustRec(1).AcctBal >= .01 THEN
            '  PRINT #RptHandle, TAB(62); USING "#####.##"; ARCustRec(1).FeeAmt + ARCustRec(1).AcctBal
            'ELSE
              PRINT #RptHandle, TAB(62); USING "#####.##"; ARCustRec(1).FeeAmt
            'END IF
            PRINT #RptHandle,
            PRINT #RptHandle,
            PRINT #RptHandle,
            'IF ARCustRec(1).AcctBal >= .01 THEN
              'PRINT #RptHandle, TAB(62); USING "#####.##"; ARCustRec(1).FeeAmt + ARCustRec(1).AcctBal
            'ELSE
              PRINT #RptHandle, TAB(62); USING "#####.##"; ARCustRec(1).FeeAmt
            'END IF
            PRINT #RptHandle,
            PRINT #RptHandle,

            GOSUB Post2Account

            LICENSE# = LICENSE# + 1
          END IF
          
        NEXT Cnt!
        PRINT #RptHandle, CHR$(12);
        PRINT #RptHandle, CHR$(18);             ' oki 320 10 cpi
        CLOSE   'Close all open files now
        
        IF DevSpec$ = "S" THEN
          EntryPoint = 2
        ELSE
          EntryPoint = 5
        END IF
        
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
        help$ = "Posting: " + LEFT$(ARCustRec(1).BILLNAME, 30)
        'printhelp help$
        
        ' Update Customer Information First
        ARCustRec(1).LICENSE = LTRIM$(STR$(LICENSE#))
        ARCustRec(1).VALID = ExpDateNum
        ARCustRec(1).IssueLicense$ = "N"
        
        PUT ARFile, CustomerNumber, ARCustRec(1)
        
        IF ChargeAcct$ = "Y" THEN
          ' Calc New Running Balance
          CurBal# = ARCustRec(1).AcctBal
          AcctBal# = CurBal# + ARCustRec(1).FeeAmt
          ARCustRec(1).AcctBal = AcctBal#
          ' Post Transaction Record First
          ARTransRec(1).CustomerNumber = LTRIM$(STR$(CustomerNumber))
          ARTransRec(1).TransDate = Date2Num%(IssueDate$)
          ARTransRec(1).TransAmount = ARCustRec(1).FeeAmt
          ARTransRec(1).TransType = 1           ' Type 1 = CHARGE
          ARTransRec(1).TransDesc = "LICENSE"
          ARTransRec(1).CashAmount = 0
          ARTransRec(1).ChkAmount = 0
          ARTransRec(1).BalanceAfterTrans = AcctBal#
          ARTransRec(1).ExtraRoom = ""
          'ARTransRec(1).Posted2GL = "N"
          ARTransRec(1).NextTrans = 0
          ' ARTransRec(1).CatCodeRec = CatagoryRecord
          
          PUT ARTransFile, NextTransRec, ARTransRec(1)
          
          IF ARCustRec(1).FirstTrans = 0 THEN
            ARCustRec(1).FirstTrans = NextTransRec
            ARCustRec(1).LastTrans = NextTransRec
            PUT ARFile, CustomerNumber, ARCustRec(1)
          ELSE
            Prev! = ARCustRec(1).LastTrans
            ARCustRec(1).LastTrans = NextTransRec
            PUT ARFile, CustomerNumber, ARCustRec(1)
            GET ARTransFile, Prev!, ARTransRec(1)
            ARTransRec(1).NextTrans = NextTransRec
            PUT ARTransFile, Prev!, ARTransRec(1)
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
        
        REDIM Choice$(2, 2)     ' assume maximum of 99 codes
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
            help$ = " ** NOW PRINTING YOUR ALIGNMENT **"
            'printhelp help$
            Heading1$ = RTRIM$(Form$(6, 0))
            Heading2$ = RTRIM$(Form$(7, 0))
            Heading3$ = RTRIM$(Form$(8, 0))
            Heading4$ = RTRIM$(Form$(9, 0))
            
            Heading1 = LEN(Heading1$): Heading2 = LEN(Heading2$)
            Heading3 = LEN(Heading3$): Heading4 = LEN(Heading4$)
            
            IF LEN(Heading1) > 0 THEN tab1 = Heading1 / 2 ELSE tab1 = 0
            IF LEN(Heading2) > 0 THEN tab2 = Heading2 / 2 ELSE tab2 = 0
            IF LEN(Heading3) > 0 THEN Tab3 = Heading3 / 2 ELSE Tab3 = 0
            IF LEN(Heading4) > 0 THEN Tab4 = Heading4 / 2 ELSE Tab4 = 0
            
            ' Print Form Test
            FOR LCnt = 1 TO 5
              LPRINT ""
            NEXT LCnt
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
            FOR LCnt = 24 TO 35
              LPRINT ""
            NEXT LCnt
            LPRINT TAB(62); "XXXXX.XX"
            LPRINT
            LPRINT
            LPRINT
            LPRINT TAB(62); "XXXXX.XX"
            LPRINT
            LPRINT
            help$ = " Did It Line Up ??"
            'printhelp help$
            
          CASE F10Key
            
            Heading1$ = RTRIM$(Form$(6, 0))
            Heading2$ = RTRIM$(Form$(7, 0))
            Heading3$ = RTRIM$(Form$(8, 0))
            Heading4$ = RTRIM$(Form$(9, 0))
            
            Heading1 = LEN(Heading1$): Heading2 = LEN(Heading2$)
            Heading3 = LEN(Heading3$): Heading4 = LEN(Heading4$)
            
            IF LEN(Heading1) > 0 THEN tab1 = Heading1 / 2 ELSE tab1 = 0
            IF LEN(Heading2) > 0 THEN tab2 = Heading2 / 2 ELSE tab2 = 0
            IF LEN(Heading3) > 0 THEN Tab3 = Heading3 / 2 ELSE Tab3 = 0
            IF LEN(Heading4) > 0 THEN Tab4 = Heading4 / 2 ELSE Tab4 = 0
            Order$ = Form$(1, 0)
            LICENSE# = VAL(Form$(3, 0))

            IssueDate$ = Form$(4, 0)

            ExpireDate$ = Form$(2, 0)
            ExpDateNum = Date2Num(ExpireDate$)
            Year$ = RIGHT$(Form$(2, 0), 4)

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

SUB ShowNoCode
        LibName$ = "AR"
        ScrnName$ = "ARNOCODE"
        help$ = "NEW A/R Customer Entry"
        LOCATE 1, 1, 0
        
        
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
        
        
        PRINT CHR$(7);
        ShowCursor
        LibFile2Scrn "AR.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
        Action = 1
        DO
          EditForm Form$(), Fld(), Frm(1), Cnf, Action
          
          SELECT CASE Frm(1).KeyCode
          CASE EscKey
            EXIT SUB
          END SELECT
          
        LOOP
        
END SUB

