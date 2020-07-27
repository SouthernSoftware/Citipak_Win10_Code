DEFINT A-Z
DECLARE SUB ClearBack ()
DECLARE SUB CustomerFeesListing ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB SmallPause ()
DECLARE SUB LookUp (RecNo&, Text$, ChkBalFlag%, CLSFlag%, CNumFlag%)
DECLARE SUB QuickCustomer ()
DECLARE SUB AppListing ()
DECLARE SUB CustomerBalance ()
DECLARE SUB ExpiredLicenseListing ()
DECLARE SUB LicenseListing ()
DECLARE SUB OpenARCustIdxFile (NumOfARIdxRecs%, ARIdxFile%)
DECLARE SUB ARFixMess ()
DECLARE SUB ShowNoCode ()
DECLARE SUB OpenARCustFile (NumOfArRecs, ARFile)
DECLARE SUB TransactionJournal ()
DECLARE SUB CustomerListing ()
DECLARE SUB CatListing ()
DECLARE SUB CustomerInquiry (AccountRec&)
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB ShowCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB HideCursor ()
DECLARE SUB QPrint (x$, Colr%, Page%)
DECLARE SUB QPrintRC (t$, r%, c%, clr%)
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE FUNCTION Round# (N#)
DECLARE FUNCTION PDQTimer& ()
  
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ARV95.bi'                        'A/R FILE LAYOUTS
  ''$INCLUDE: 'ARV85.bi'                        'A/R FILE LAYOUTS
  '$INCLUDE: 'GL.bi'
  
  TYPE Struct
    who AS STRING * 14
    RecNum AS INTEGER
  END TYPE
  
  DIM SHARED ARCustRec(1) AS ARCustRecType
  DIM SHARED ARCustIdxRec(1) AS ARCustIDXRecType
  
  CONST False = 0, True = NOT False
  
  STACK 5000
  
  '  ARFixMess
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 11)
  
  MChoice$(1) = " Customer Inquiry"
  MChoice$(2) = " Customer Balance Listing "
  MChoice$(3) = " Transaction Journal"
  MChoice$(4) = " Detailed Category Listing "
  MChoice$(5) = " Detailed Customer Listing "
  MChoice$(6) = " License Listing"
  MChoice$(7) = " Expired License Listing"
  MChoice$(8) = " Application Listing"
  MChoice$(9) = " Quick Customer Listing"
  MChoice$(10) = " Exit to OS"
  
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
  Row = ((25 - (UBOUND(MChoice$))) \ 2) + 1
  Col = ((80 - MaxLen) \ 2) - 1
  Help$ = "Main Reports Menu"
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    LibFile2Scrn "AR.QSL", "MENUBAK", MonoCode, -1, ErrorCode
    
    TitleBox 3, Col, MaxLen + 3, "Reports Menu ", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    PrintTitle User$
    PrintHelp Help$
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      LookUp AccountRec&, "Customer", False, True, False
      IF AccountRec& > 0 THEN
        CustomerInquiry AccountRec&
      END IF
    CASE 2
      CustomerBalance
    CASE 3
      TransactionJournal
    CASE 4
      CatListing
    CASE 5
      CustomerListing
    CASE 6
      LicenseListing
    CASE 7
      ExpiredLicenseListing
    CASE 8
      AppListing
    CASE 9
      QuickCustomer
    CASE 10
      END
    END SELECT
  LOOP
  RUN "armenu"

SUB ARFixMess
  
  CLS
  
  OpenARCustFile NumOfArRecs, ARFile
  NumOfCust = LOF(ARFile) \ LEN(ARCustRec(1))
  FOR Cnt = 1 TO NumOfCust
    GET ARFile, Cnt, ARCustRec(1)
    ARCustRec(1).ProRate = 100
    '    ARCustRec(1).AcctBal = 0
    ARCustRec(1).LicBal = ARCustRec(1).AcctBal
    ARCustRec(1).FeeBal = 0
    ARCustRec(1).PenBal = 0
    'ARCustRec(1).FirstTrans = 0
    'ARCustRec(1).LastTrans = 0
    '    ARCustRec(1).IssueLicense = "Y"
    PUT ARFile, Cnt, ARCustRec(1)
  NEXT
  CLOSE
  'CALL KillFile("artrans.dat")
  
  '  REDIM ARTransRec(1) AS ARTransRecType
  '
  '  ARTRFile = FREEFILE
  '  OPEN "ARTrans.Dat" FOR RANDOM ACCESS READ WRITE SHARED AS #ARTRFile LEN = LEN(ARTransRec(1))
  '  TrNumRecs = LOF(ARTRFile) / LEN(ARTransRec(1))
  '
  '  FOR Cnt = 1 TO 104
  '    GET ARTRFile, Cnt, ARTransRec(1)
  '
  '    'Get Customer
  '    OpenARCustFile NumOfArRecs, ARFile
  '    GET ARFile, VAL(ARTransRec(1).CustomerNumber), ARCustRec(1)
  '    ARCustRec(1).IssueLicense = "Y"
  '    ARCustRec(1).FeeAmt = ARTransRec(1).TransAmount
  '    PUT ARFile, VAL(ARTransRec(1).CustomerNumber), ARCustRec(1)
  '
  '
  '    IF ARTransRec(1).TransType = 1 THEN PRINT "Charge"; Cnt: STOP
  '    IF ARTransRec(1).TransType = 2 THEN PRINT "Payment"; Cnt
  '
  '  NEXT Cnt
  '  CLOSE         'Close all open files now
  
  
END SUB

SUB CatListing
  SHARED Choice$()
  
  
  ReportFile$ = "ARCODLST.PRN"  'Report File Name
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
  
  GOSUB SelectOutput
  IF Canceled$ = "Y" THEN EXIT SUB
  
  
  REDIM ARCatCodeRec(1) AS ARNewCatCodeRecType  ' open transaction file
  CatCodeRecLen = LEN(ARCatCodeRec(1))
  TrHandle = FREEFILE
  OPEN "ARCODE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS TrHandle LEN = CatCodeRecLen
  TRNumRecs = LOF(TrHandle) \ CatCodeRecLen
  
  
  GOSUB GetReportInformation
  
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  
  PRINT #RptHandle, CHR$(27); CHR$(58);         ' oki 320 12 cpi
  
  GOSUB PrintRptHeader
  
  FOR Cnt = 1 TO Count
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
  Page = Page + 1
  PRINT #RptHandle, TAB(20); "Business License's  : Category Code Listing "
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
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
  
  FOR Cnt! = 1 TO TRNumRecs
    GET TrHandle, Cnt!, ARCatCodeRec(1)
    array(Cnt!).who = LEFT$(ARCatCodeRec(1).CATCODE, 3) + STRING$(11, " ")
    array(Cnt!).RecNum = Cnt!
  NEXT Cnt!
  
  Count = TRNumRecs
  
  SortT array(Start), Count, Dir, SSize, MOff, MSize
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
  
  REDIM Choice$(2, 0)
  
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "SCREEN"
  Choice$(2, 0) = "PRINTER"
  
  
  Action = 1
  ShowCursor
  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%
  'printhelp Help$
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

SUB CustomerBalance
  SHARED Choice$()
  
  ReportFile$ = "ARCusBal.PRN"  'Report File Name
  LPTPort% = 1
  FF$ = CHR$(12)
  MaxLines = 58
  LineCnt = 0
  CustCnt = 0
  
  GOSUB SelectCustBalOutput
  IF Canceled$ = "Y" THEN
    EXIT SUB
  END IF
  ShowProcessingScrn "Customer Balance Listing"
  
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  'PRINT #RptHandle, CHR$(27); CHR$(58); ' oki 320 12 cpi
  
  GOSUB PrintCustBalRptHeader
  
  ' Print Main Body
  OpenARCustFile NumOfArRecs, ARFile
  REDIM ARIdxRec(1) AS ARCustIDXRecType
  IdxFile = FREEFILE
  OPEN "ARCUST.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS #IdxFile LEN = LEN(ARIdxRec(1))
  
  IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN
    NumOfArRecs = LOF(IdxFile) \ LEN(ARIdxRec(1))
  END IF
  
  FOR Cnt! = 1 TO NumOfArRecs
    IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN
      GET IdxFile, Cnt!, ARIdxRec(1)
      GET ARFile, ARIdxRec(1).IDXRECORD, ARCustRec(1)
      CustomerNumber = ARIdxRec(1).IDXRECORD
    ELSE
      GET ARFile, Cnt!, ARCustRec(1)
      CustomerNumber = Cnt!
    END IF
    IF ARCustRec(1).Deleted <> "Y" AND ARCustRec(1).Inactive <> "Y" THEN
      TCat$ = RTRIM$(Catagory$)
      IF TCat$ = RTRIM$(ARCustRec(1).BILLCAT1) OR TCat$ = RTRIM$(ARCustRec(1).BILLCAT2) OR TCat$ = RTRIM$(ARCustRec(1).BILLCAT3) OR TCat$ = RTRIM$(ARCustRec(1).BILLCAT4) OR TCat$ = RTRIM$(ARCustRec(1).BILLCAT5) OR TCat$ = "ALL" THEN
        'IF RTRIM$(Catagory$) = RTRIM$(ARCustRec(1).BillCat1) OR RTRIM$(Catagory$) = RTRIM$(ARCustRec(1).BillCat2)
        'OR RTRIM$(Catagory$) = RTRIM$(ARCustRec(1).BillCat3) OR RTRIM$(Catagory$) = RTRIM$(ARCustRec(1).BillCat4) OR RTRIM$(Catagory$) = RTRIM$(ARCustRec(1).BillCat5) OR RTRIM$(Catagory$) = "ALL" THEN
        IF ARCustRec(1).AcctBal > 0 THEN
          IF LineCnt >= MaxLines THEN
            PRINT #RptHandle, FF$
            GOSUB PrintCustBalRptHeader
          END IF
          PRINT #RptHandle, USING "#####"; CustomerNumber;
          PRINT #RptHandle, TAB(10); ARCustRec(1).BILLNAME;
          PRINT #RptHandle, TAB(50); USING "$$####,#.##"; ARCustRec(1).AcctBal
          CustCnt = CustCnt + 1
          TotalBal# = Round#(TotalBal# + ARCustRec(1).AcctBal)
          LineCnt = LineCnt + 1
        END IF
      END IF
    END IF
    ShowPctComp Cnt!, NumOfArRecs
  NEXT Cnt!
  GOSUB PrintCustBalRptEnding
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
  
PrintCustBalRptHeader:
  Page = Page + 1
  PRINT #RptHandle, TAB(20); "Business License's : Customer Balance Listing"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
  PRINT #RptHandle, "Category: "; Catagory$
  PRINT #RptHandle, "Cust #"; TAB(10); "Customer Name"; TAB(50); "Account Balance"
  PRINT #RptHandle, STRING$(80, "=")
  LineCnt = 5
  RETURN
  
PrintCustBalRptEnding:
  PRINT #RptHandle, STRING$(79, "-")
  PRINT #RptHandle, "Total Customers Printed: "; USING "#####"; CustCnt
  PRINT #RptHandle, "  Total Balance on File: "; USING "$$#####,#.##"; TotalBal#
  PRINT #RptHandle, FF$
  RETURN
  
SelectCustBalOutput:
  LibName$ = "AR"
  ScrnName$ = "ARCUSBAL"
  
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
  
  'REDIM Choice$(650, 2)         ' assume maximum of 99 codes
  
  REDIM ARCatCodeRec(1) AS ARNewCatCodeRecType
  ARCatCodeRecLen = LEN(ARCatCodeRec(1))
  ARCatFile = FREEFILE
  OPEN "ARCODE.DAT" FOR RANDOM SHARED AS ARCatFile LEN = ARCatCodeRecLen
  NumOFARCatRecs = LOF(ARCatFile) \ ARCatCodeRecLen
  
  IF NumOFARCatRecs = 0 THEN
    CLOSE ARCatFile
    ShowNoCode
    EXIT SUB
  END IF
  
  NumOFARCatRecs = NumOFARCatRecs + 1
  REDIM Choice$(NumOFARCatRecs, 2)          ' assume maximum of 99 codes
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Name"
  Choice$(2, 0) = "Number"
  Choice$(0, 1) = "2"
  
  FOR Cnt! = 1 TO (NumOFARCatRecs - 1):
    GET ARCatFile, Cnt!, ARCatCodeRec(1)
    Choice$(Cnt!, 1) = ARCatCodeRec(1).CATCODE
  NEXT Cnt!
  Choice$(NumOFARCatRecs, 1) = "ALL"
  CLOSE ARCatFile
  
  'SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
  'SortT2 Choice$(0, 1), NumOFARCatRecs, 1, 3, 0, 3
  'Choice$(NumOFARCatRecs, 1)
  
  Choice$(0, 2) = "3"
  Choice$(1, 2) = "SCREEN"
  Choice$(2, 2) = "PRINTER"
  
  Action = 1
  ShowCursor
  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%
  'printhelp Help$
  Action = 1
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      Order$ = Form$(1, 0)
      Catagory$ = Form$(2, 0)
      DevSpec$ = LEFT$(Form$(3, 0), 1)
      RETURN
    CASE EscKey
      Canceled$ = "Y"
      RETURN
    END SELECT
  LOOP
  
END SUB

SUB CustomerInquiry (AccountRec&)
  
  
CIMainBody:
  SHARED MChoice$
  LibName$ = "AR"
  ScrnName$ = "ARCUSTIQ"
  'Help$ = "Customer Inquiry"
  'LOCATE 1, 1, 0
  LPTPort% = 1
  ShowCursor
  LibFile2Scrn "AR.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  'printhelp Help$
  
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
  NEXT
  
  Action = 1
  
  CLOSE ARFile
  OpenARCustFile NumOfArRecs, ARFile
  OpenARCustIdxFile NumOfARIdxRecs, ARIdxFile
  
  
  FirstTime = True
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      FirstTime = False
      CustomerGrabed = 0
      GOSUB LoadCustInfo
      Action = 1
    END IF
    
    SELECT CASE Frm(1).KeyCode
    CASE F4KEY
      GOSUB PrintHistory
      LibFile2Scrn "AR.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
      OpenARCustFile NumOfArRecs, ARFile
      OpenARCustIdxFile NumOfARIdxRecs, ARIdxFile
      Frm(1).FldNo = 26
      Action = 1
    CASE F10Key
      EXIT DO
    CASE EscKey
      EXIT DO
    END SELECT
  LOOP
  
  CLOSE ARFile
  EXIT SUB
  
LoadCustInfo:
  IF AccountRec& > 0 AND AccountRec& <= NumOfArRecs THEN
    GET ARFile, AccountRec&, ARCustRec(1)
    IF ARCustRec(1).Deleted = "Y" THEN GOTO CustomerDeleted
    Form$(1, 0) = ARCustRec(1).CUSTNUMB
    Form$(2, 0) = ARCustRec(1).SORTNAME
    Form$(3, 0) = ARCustRec(1).BILLNAME
    Form$(4, 0) = ARCustRec(1).ADDRESS1
    Form$(5, 0) = ARCustRec(1).ADDRESS2
    Form$(6, 0) = ARCustRec(1).CITY
    Form$(7, 0) = ARCustRec(1).STATE
    Form$(8, 0) = ARCustRec(1).ZIPCODE
    Form$(9, 0) = ARCustRec(1).CustName
    Form$(10, 0) = ARCustRec(1).BILLCAT1
    Form$(20, 0) = ARCustRec(1).WPHONE
    Form$(21, 0) = STR$(ARCustRec(1).Fee1 + ARCustRec(1).Fee2 + ARCustRec(1).Fee3 + ARCustRec(1).Fee4 + ARCustRec(1).Fee5)
    Form$(22, 0) = ARCustRec(1).LICENSE
    Form$(23, 0) = Num2Date$(ARCustRec(1).VALID)
    Form$(24, 0) = ARCustRec(1).IssueLicense
    Form$(25, 0) = STR$(ARCustRec(1).AcctBal)
    Form$(26, 0) = ARCustRec(1).Inactive
    FOR Flds = 1 TO 26
      Fld(Flds).Protected = True
    NEXT Flds
    CustomerGrabed = 1
    Action = 1
    
    QPrintRC STRING$(32, 32), 4, 40, 15
    RETURN
    
  ELSE
    
    LibName$ = "AR"
    ScrnName$ = "ARBADCUS"
    Help$ = "Edit A/R Customer Entry"
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
    'printhelp Help$
    
    Done = False
    Action = 1
    
    
    DO
      
      EditForm Form$(), Fld(), Frm(1), Cnf, Action
      
      SELECT CASE Frm(1).KeyCode
      CASE EscKey
        Done = True
      END SELECT
      IF Done = True THEN
        EXIT SUB
        'GOTO CIMainBody
      END IF
    LOOP
    
  END IF
  
CustomerDeleted:
  LibName$ = "AR"
  ScrnName$ = "ARDELCUS"
  Help$ = "Edit Customer"
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
  
  Done = False
  Action = 1
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      GOTO CIMainBody
    END SELECT
  LOOP
  
  
PrintHistory:
  RunBalance# = 0
  Page = 0
  REDIM ARTransRec(1) AS ARTransRecType
  ARTransRecLen = LEN(ARTransRec(1))
  ARTransFile = FREEFILE
  OPEN "ARTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARTransFile LEN = ARTransRecLen
  
  MaxLines = 58
  ReportFile$ = "ARHIST.PRN"
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  GOSUB PrintCustInqRptHeader
  
  TransRec& = ARCustRec(1).FirstTrans
  
  WHILE TransRec& > 0
    GET ARTransFile, TransRec&, ARTransRec(1)
    
    IF LineCnt >= MaxLines THEN
      PRINT #RptHandle, FF$
      GOSUB PrintCustInqRptHeader
    END IF
    PRINT #RptHandle, Num2Date$(ARTransRec(1).TransDate);
    PRINT #RptHandle, TAB(12); RTRIM$(ARTransRec(1).TransDesc);
    PRINT #RptHandle, TAB(50); "";
    SELECT CASE ARTransRec(1).TransType
    CASE 1      'lic
      PRINT #RptHandle, "Charge";
      RunBalance# = RunBalance# + ARTransRec(1).TransAmount
      PRINT #RptHandle, TAB(60); USING "######.##"; ARTransRec(1).TransAmount;
    CASE 2      'payment
      PRINT #RptHandle, "Payment";
      RunBalance# = RunBalance# - ARTransRec(1).TransAmount
      PRINT #RptHandle, TAB(59); USING "#######.##"; -ARTransRec(1).TransAmount;
    CASE 9
      PRINT #RptHandle, "Beg Bal";
      RunBalance# = RunBalance# + ARTransRec(1).TransAmount
      PRINT #RptHandle, TAB(60); USING "######.##"; ARTransRec(1).TransAmount;
    CASE 100    'dn adj
      PRINT #RptHandle, "DN Adj.";
      RunBalance# = RunBalance# - ARTransRec(1).TransAmount
      PRINT #RptHandle, TAB(59); USING "#######.##"; -ARTransRec(1).TransAmount;
      
    CASE 101    'up adj
      PRINT #RptHandle, "UP Adj.";
      RunBalance# = RunBalance# + ARTransRec(1).TransAmount
      PRINT #RptHandle, TAB(59); USING "#######.##"; ARTransRec(1).TransAmount;
      
    END SELECT
    
    PRINT #RptHandle, TAB(71); USING "######.##"; ARTransRec(1).BalanceAfterTrans               'RunBalance#
    LineCnt = LineCnt + 1
    TransRec& = ARTransRec(1).NextTrans
    
  WEND
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi
  PRINT #RptHandle, CHR$(12);
  CLOSE         'Close all open files now
  
  PrintRptFile "Customer Account History", ReportFile$, LPTPort%, RetCode%, EntryPoint
  KILL ReportFile$
  RETURN
  
PrintCustInqRptHeader:
  Page = Page + 1
  PRINT #RptHandle, TAB(18); "Business License : Detailed Customer History"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
  PRINT #RptHandle, "Customer Name: "; ARCustRec(1).BILLNAME
  PRINT #RptHandle, "  Date"; TAB(10); "Description "; TAB(50); "Type"; TAB(60); "  Amount"; TAB(73); "Balance"
  PRINT #RptHandle, STRING$(80, "=")
  LineCnt = 3
  
  RETURN
  
  
END SUB

SUB CustomerListing
  
  SHARED Choice$()
  
  ReportFile$ = "ARDetCus.PRN"  'Report File Name
  FF$ = CHR$(12)
  MaxLines = 58
  LineCnt = 0
  CustCnt = 0
  
  LibName$ = "AR"
  ScrnName$ = "ARCUSREP"
  
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
  REDIM Choice$(NumOFARCatRecs + 3, 2)    ' assume maximum of 99 codes
  
  NumOFARCatRecs = NumOFARCatRecs + 1
  
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Name"
  Choice$(2, 0) = "Number"
  
  Choice$(0, 1) = "2"
  FOR Cnt! = 1 TO (NumOFARCatRecs - 1):
    GET ARCatFile, Cnt!, ARCatCodeRec(1)
    Choice$(Cnt!, 1) = ARCatCodeRec(1).CATCODE
  NEXT Cnt!
  Choice$(NumOFARCatRecs, 1) = "ALL"
  CLOSE ARCatFile
  
  Choice$(0, 2) = "4"
  Choice$(1, 2) = "SCREEN"
  Choice$(2, 2) = "PRINTER"
  
  FirstTime = True
  
  Action = 1
  ShowCursor
  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%
  'printhelp Help$
  Action = 1
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      LSET Form$(3, 0) = "N"
      Action = 1
    END IF
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      Order$ = Form$(1, 0)
      Catagory$ = Form$(2, 0)
      ShowFee = Form$(3, 0) = "Y"
      DevSpec$ = LEFT$(Form$(4, 0), 1)
      EXIT DO
    CASE EscKey
      Canceled$ = "Y"
      EXIT DO
    END SELECT
  LOOP
  
  IF Canceled$ = "Y" THEN EXIT SUB
  
  ClearBack
  ShowProcessingScrn "Detailed Customer Listing"
  
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  ARCatFile = FREEFILE
  OPEN "ARCODE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARCatFile LEN = ARCatCodeRecLen
  NumOFARCatRecs = LOF(ARCatFile) \ ARCatCodeRecLen
  
  'PRINT #RptHandle, CHR$(27); CHR$(58); ' oki 320 12 cpi
  
  GOSUB PrintDetailCustomerRptHeader
  
  ' Print Main Body
  OpenARCustFile NumOfArRecs, ARFile
  REDIM ARIdxRec(1) AS ARCustIDXRecType
  IdxFile = FREEFILE
  OPEN "ARCUST.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS #IdxFile LEN = LEN(ARIdxRec(1))
  
  IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN
    NumOfArRecs = LOF(IdxFile) \ LEN(ARIdxRec(1))
  END IF
  
  FOR Cnt! = 1 TO NumOfArRecs
    IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN
      GET IdxFile, Cnt!, ARIdxRec(1)
      GET ARFile, ARIdxRec(1).IDXRECORD, ARCustRec(1)
      CustomerNumber = ARIdxRec(1).IDXRECORD
    ELSE
      GET ARFile, Cnt!, ARCustRec(1)
      CustomerNumber = Cnt!
    END IF
    
    IF ARCustRec(1).Deleted <> "Y" THEN
      TCat$ = QPTrim$(Catagory$)
      IF TCat$ = QPTrim$(ARCustRec(1).BILLCAT1) OR TCat$ = QPTrim$(ARCustRec(1).BILLCAT2) OR TCat$ = QPTrim$(ARCustRec(1).BILLCAT3) OR TCat$ = QPTrim$(ARCustRec(1).BILLCAT4) OR TCat$ = QPTrim$(ARCustRec(1).BILLCAT5) OR TCat$ = "ALL" THEN
        IF LineCnt >= MaxLines THEN
          PRINT #RptHandle, FF$
          GOSUB PrintDetailCustomerRptHeader
        END IF
        IF ShowFee THEN
          GOSUB GetCustFee
        END IF
        CustNum$ = SPACE$(6)
        RSET CustNum$ = QPTrim$(ARCustRec(1).CUSTNUMB)
        PRINT #RptHandle, CustNum$;
        PRINT #RptHandle, TAB(13); ARCustRec(1).Inactive;
        PRINT #RptHandle, TAB(25); USING "%###"; ARCustRec(1).ProRate;
        PRINT #RptHandle, TAB(39); "Lic # "; ARCustRec(1).LICENSE; TAB(60); "Valid to: "; Num2Date$(ARCustRec(1).VALID)
        PRINT #RptHandle, ARCustRec(1).BILLNAME; TAB(40); "Category: "; ARCustRec(1).BILLCAT1; " / "; ARCustRec(1).BILLCAT2; " / "; ARCustRec(1).BILLCAT3; " / "; ARCustRec(1).BILLCAT4; " / "; ARCustRec(1).BILLCAT5
        PRINT #RptHandle, RTRIM$(ARCustRec(1).ADDRESS1); TAB(50); " W Phone: "; ARCustRec(1).WPHONE
        PRINT #RptHandle, RTRIM$(ARCustRec(1).ADDRESS2);
        PRINT #RptHandle, RTRIM$(ARCustRec(1).CITY); ", "; ARCustRec(1).STATE; " "; ARCustRec(1).ZIPCODE;
        IF ShowFee THEN
          PRINT #RptHandle, TAB(50); "Lic Fees: $"; USING "#####,#.##"; CustFee#
        ELSE
          PRINT #RptHandle,
        END IF
        PRINT #RptHandle, STRING$(80, "-")
        CustCnt = CustCnt + 1
        LineCnt = LineCnt + 6
      END IF
    END IF
    ShowPctComp Cnt!, NumOfArRecs
  NEXT Cnt!
  GOSUB PrintDetailCustomerRptEnding
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
  
  
PrintDetailCustomerRptHeader:
  Page = Page + 1
  PRINT #RptHandle, TAB(20); "Business License's  : Detailed Customer Listing"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
  PRINT #RptHandle, "Acct No   Inactive   Prorate             Lic No"
  PRINT #RptHandle, STRING$(80, "=")
  LineCnt = 3
  RETURN
  
PrintDetailCustomerRptEnding:
  PRINT #RptHandle, "Total Customers Printed: "; USING "#####"; CustCnt
  PRINT #RptHandle,
  PRINT #RptHandle, FF$
  RETURN
  
GetCustFee:
  SmallPause
  CustFee# = 0
  FeeAmt# = 0
  ProRateFlag = False
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
        ARCustRec(1).Desc1 = ARCatCodeRec(1).CODEDESC           'Reset Code Descriptions
        IF ARCatCodeRec(1).CodeType = "F" THEN
          FeeAmt# = ARCatCodeRec(1).Fee
          GOTO C2
        END IF
        IF ARCatCodeRec(1).CodeType = "M" THEN
          Mult = ARCustRec(1).REV1
          FeeAmt# = Round#(Mult * ARCatCodeRec(1).Fee)
          GOTO C2
        END IF
        IF ARCatCodeRec(1).CodeType = "S" THEN
          Revenue# = ARCustRec(1).REV1
          IF Revenue# <= ARCatCodeRec(1).Recpt1 THEN
            FeeAmt# = ARCatCodeRec(1).BaseAmt1 + (ARCatCodeRec(1).Percent1 / 100) * (Revenue# - ARCatCodeRec(1).Maximum1)
            GOTO C2
          END IF
          IF Revenue# <= ARCatCodeRec(1).Recpt2 THEN
            FeeAmt# = ARCatCodeRec(1).BaseAmt2 + (ARCatCodeRec(1).Percent2 / 100) * (Revenue# - ARCatCodeRec(1).Maximum2)
            GOTO C2
          END IF
          IF Revenue# <= ARCatCodeRec(1).Recpt3 THEN
            FeeAmt# = ARCatCodeRec(1).BaseAmt3 + (ARCatCodeRec(1).Percent3 / 100) * (Revenue# - ARCatCodeRec(1).Maximum3)
            GOTO C2
          END IF
          IF Revenue# <= ARCatCodeRec(1).Recpt4 THEN
            FeeAmt# = ARCatCodeRec(1).BaseAmt4 + (ARCatCodeRec(1).Percent4 / 100) * (Revenue# - ARCatCodeRec(1).Maximum4)
            GOTO C2
          END IF
          IF Revenue# <= ARCatCodeRec(1).Recpt5 THEN
            FeeAmt# = ARCatCodeRec(1).BaseAmt5 + (ARCatCodeRec(1).Percent5 / 100) * (Revenue# - ARCatCodeRec(1).Maximum5)
            GOTO C2
          END IF
          IF Revenue# <= ARCatCodeRec(1).Recpt6 THEN
            FeeAmt# = ARCatCodeRec(1).BaseAmt6 + (ARCatCodeRec(1).Percent6 / 100) * (Revenue# - ARCatCodeRec(1).Maximum6)
            GOTO C2
          END IF
          
        END IF
      END IF    'End Test for Code
    NEXT Snt&
  END IF        'End Test for Cat 1
  
  
C2:             'Catagory #2
  
  CustFee# = Round#(CustFee# + FeeAmt#)
  FeeAmt# = 0
  CATCODE$ = QPTrim$(ARCustRec(1).BILLCAT2)
  IF LEN(CATCODE$) > 0 THEN
    FOR Snt& = 1 TO NumOFARCatRecs
      GET ARCatFile, Snt&, ARCatCodeRec(1)
      IF QPTrim$(ARCatCodeRec(1).CATCODE) = CATCODE$ THEN
        IF ARCatCodeRec(1).CodeType = "F" THEN
          FeeAmt# = ARCatCodeRec(1).Fee
          GOTO C3
        END IF
        IF ARCatCodeRec(1).CodeType = "M" THEN
          Mult = ARCustRec(1).REV2
          FeeAmt# = Round#(Mult * ARCatCodeRec(1).Fee)
          GOTO C3
        END IF
        IF ARCatCodeRec(1).CodeType = "S" THEN
          Revenue# = ARCustRec(1).REV2
          IF Revenue# <= ARCatCodeRec(1).Recpt1 THEN
            FeeAmt# = ARCatCodeRec(1).BaseAmt1 + (ARCatCodeRec(1).Percent1 / 100) * (Revenue# - ARCatCodeRec(1).Maximum1)
            GOTO C3
          END IF
          IF Revenue# <= ARCatCodeRec(1).Recpt2 THEN
            FeeAmt# = ARCatCodeRec(1).BaseAmt2 + (ARCatCodeRec(1).Percent2 / 100) * (Revenue# - ARCatCodeRec(1).Maximum2)
            GOTO C3
          END IF
          IF Revenue# <= ARCatCodeRec(1).Recpt3 THEN
            FeeAmt# = ARCatCodeRec(1).BaseAmt3 + (ARCatCodeRec(1).Percent3 / 100) * (Revenue# - ARCatCodeRec(1).Maximum3)
            GOTO C3
          END IF
          IF Revenue# <= ARCatCodeRec(1).Recpt4 THEN
            FeeAmt# = ARCatCodeRec(1).BaseAmt4 + (ARCatCodeRec(1).Percent4 / 100) * (Revenue# - ARCatCodeRec(1).Maximum4)
            GOTO C3
          END IF
          IF Revenue# <= ARCatCodeRec(1).Recpt5 THEN
            FeeAmt# = ARCatCodeRec(1).BaseAmt5 + (ARCatCodeRec(1).Percent5 / 100) * (Revenue# - ARCatCodeRec(1).Maximum5)
            GOTO C3
          END IF
          IF Revenue# <= ARCatCodeRec(1).Recpt6 THEN
            FeeAmt# = ARCatCodeRec(1).BaseAmt6 + (ARCatCodeRec(1).Percent6 / 100) * (Revenue# - ARCatCodeRec(1).Maximum6)
            GOTO C3
          END IF
        END IF
      END IF    'End Test for Code
    NEXT Snt&
  END IF        'End Test for Cat 1
  
  
C3:
  CustFee# = Round#(CustFee# + FeeAmt#)
  FeeAmt# = 0
  CATCODE$ = QPTrim$(ARCustRec(1).BILLCAT3)
  IF LEN(CATCODE$) > 0 THEN
    FOR Snt& = 1 TO NumOFARCatRecs
      GET ARCatFile, Snt&, ARCatCodeRec(1)
      IF QPTrim$(ARCatCodeRec(1).CATCODE) = CATCODE$ THEN
        IF ARCatCodeRec(1).CodeType = "F" THEN
          FeeAmt# = ARCatCodeRec(1).Fee
          GOTO c4
        END IF
        IF ARCatCodeRec(1).CodeType = "M" THEN
          Mult = ARCustRec(1).REV3
          FeeAmt# = Round#(Mult * ARCatCodeRec(1).Fee)
          GOTO c4
        END IF
        IF ARCatCodeRec(1).CodeType = "S" THEN
          Revenue# = ARCustRec(1).REV3
          IF Revenue# <= ARCatCodeRec(1).Recpt1 THEN
            FeeAmt# = ARCatCodeRec(1).BaseAmt1 + (ARCatCodeRec(1).Percent1 / 100) * (Revenue# - ARCatCodeRec(1).Maximum1)
            GOTO c4
          END IF
          IF Revenue# <= ARCatCodeRec(1).Recpt2 THEN
            FeeAmt# = ARCatCodeRec(1).BaseAmt2 + (ARCatCodeRec(1).Percent2 / 100) * (Revenue# - ARCatCodeRec(1).Maximum2)
            GOTO c4
          END IF
          IF Revenue# <= ARCatCodeRec(1).Recpt3 THEN
            FeeAmt# = ARCatCodeRec(1).BaseAmt3 + (ARCatCodeRec(1).Percent3 / 100) * (Revenue# - ARCatCodeRec(1).Maximum3)
            GOTO c4
          END IF
          IF Revenue# <= ARCatCodeRec(1).Recpt4 THEN
            FeeAmt# = ARCatCodeRec(1).BaseAmt4 + (ARCatCodeRec(1).Percent4 / 100) * (Revenue# - ARCatCodeRec(1).Maximum4)
            GOTO c4
          END IF
          IF Revenue# <= ARCatCodeRec(1).Recpt5 THEN
            FeeAmt# = ARCatCodeRec(1).BaseAmt5 + (ARCatCodeRec(1).Percent5 / 100) * (Revenue# - ARCatCodeRec(1).Maximum5)
            GOTO c4
          END IF
          IF Revenue# <= ARCatCodeRec(1).Recpt6 THEN
            FeeAmt# = ARCatCodeRec(1).BaseAmt6 + (ARCatCodeRec(1).Percent6 / 100) * (Revenue# - ARCatCodeRec(1).Maximum6)
            GOTO c4
          END IF
        END IF
      END IF    'End Test for Code
    NEXT Snt&
  END IF        'End Test for Cat 3
  
c4:
  CustFee# = Round#(CustFee# + FeeAmt#)
  FeeAmt# = 0
  CATCODE$ = QPTrim$(ARCustRec(1).BILLCAT4)
  IF LEN(CATCODE$) > 0 THEN
    FOR Snt& = 1 TO NumOFARCatRecs
      GET ARCatFile, Snt&, ARCatCodeRec(1)
      IF QPTrim$(ARCatCodeRec(1).CATCODE) = CATCODE$ THEN
        IF ARCatCodeRec(1).CodeType = "F" THEN
          FeeAmt# = ARCatCodeRec(1).Fee
          GOTO c5
        END IF
        IF ARCatCodeRec(1).CodeType = "M" THEN
          Mult = ARCustRec(1).REV4
          FeeAmt# = Round#(Mult * ARCatCodeRec(1).Fee)
          GOTO c5
        END IF
        IF ARCatCodeRec(1).CodeType = "S" THEN
          Revenue# = ARCustRec(1).REV4
          IF Revenue# <= ARCatCodeRec(1).Recpt1 THEN
            FeeAmt# = ARCatCodeRec(1).BaseAmt1 + (ARCatCodeRec(1).Percent1 / 100) * (Revenue# - ARCatCodeRec(1).Maximum1)
            GOTO c5
          END IF
          IF Revenue# <= ARCatCodeRec(1).Recpt2 THEN
            FeeAmt# = ARCatCodeRec(1).BaseAmt2 + (ARCatCodeRec(1).Percent2 / 100) * (Revenue# - ARCatCodeRec(1).Maximum2)
            GOTO c5
          END IF
          IF Revenue# <= ARCatCodeRec(1).Recpt3 THEN
            FeeAmt# = ARCatCodeRec(1).BaseAmt3 + (ARCatCodeRec(1).Percent3 / 100) * (Revenue# - ARCatCodeRec(1).Maximum3)
            GOTO c5
          END IF
          IF Revenue# <= ARCatCodeRec(1).Recpt4 THEN
            FeeAmt# = ARCatCodeRec(1).BaseAmt4 + (ARCatCodeRec(1).Percent4 / 100) * (Revenue# - ARCatCodeRec(1).Maximum4)
            GOTO c5
          END IF
          IF Revenue# <= ARCatCodeRec(1).Recpt5 THEN
            FeeAmt# = ARCatCodeRec(1).BaseAmt5 + (ARCatCodeRec(1).Percent5 / 100) * (Revenue# - ARCatCodeRec(1).Maximum5)
            GOTO c5
          END IF
          IF Revenue# <= ARCatCodeRec(1).Recpt6 THEN
            FeeAmt# = ARCatCodeRec(1).BaseAmt6 + (ARCatCodeRec(1).Percent6 / 100) * (Revenue# - ARCatCodeRec(1).Maximum6)
            GOTO c5
          END IF
          
        END IF
      END IF    'End Test for Code
    NEXT Snt&
  END IF        'End Test for Cat 1
  
c5:
  CustFee# = Round#(CustFee# + FeeAmt#)
  FeeAmt# = 0
  CATCODE$ = QPTrim$(ARCustRec(1).BILLCAT5)
  IF LEN(CATCODE$) > 0 THEN
    FOR Snt& = 1 TO NumOFARCatRecs
      GET ARCatFile, Snt&, ARCatCodeRec(1)
      IF QPTrim$(ARCatCodeRec(1).CATCODE) = CATCODE$ THEN
        IF ARCatCodeRec(1).CodeType = "F" THEN
          FeeAmt# = ARCatCodeRec(1).Fee
          GOTO SkipEm
        END IF
        IF ARCatCodeRec(1).CodeType = "M" THEN
          Mult = ARCustRec(1).REV5
          FeeAmt# = Round#(Mult * ARCatCodeRec(1).Fee)
          GOTO SkipEm
        END IF
        IF ARCatCodeRec(1).CodeType = "S" THEN
          Revenue# = ARCustRec(1).REV5
          IF Revenue# <= ARCatCodeRec(1).Recpt1 THEN
            FeeAmt# = ARCatCodeRec(1).BaseAmt1 + (ARCatCodeRec(1).Percent1 / 100) * (Revenue# - ARCatCodeRec(1).Maximum1)
            GOTO SkipEm
          END IF
          IF Revenue# <= ARCatCodeRec(1).Recpt2 THEN
            FeeAmt# = ARCatCodeRec(1).BaseAmt2 + (ARCatCodeRec(1).Percent2 / 100) * (Revenue# - ARCatCodeRec(1).Maximum2)
            GOTO SkipEm
          END IF
          IF Revenue# <= ARCatCodeRec(1).Recpt3 THEN
            FeeAmt# = ARCatCodeRec(1).BaseAmt3 + (ARCatCodeRec(1).Percent3 / 100) * (Revenue# - ARCatCodeRec(1).Maximum3)
            GOTO SkipEm
          END IF
          IF Revenue# <= ARCatCodeRec(1).Recpt4 THEN
            FeeAmt# = ARCatCodeRec(1).BaseAmt4 + (ARCatCodeRec(1).Percent4 / 100) * (Revenue# - ARCatCodeRec(1).Maximum4)
            GOTO SkipEm
          END IF
          IF Revenue# <= ARCatCodeRec(1).Recpt5 THEN
            FeeAmt# = ARCatCodeRec(1).BaseAmt5 + (ARCatCodeRec(1).Percent5 / 100) * (Revenue# - ARCatCodeRec(1).Maximum5)
            GOTO SkipEm
          END IF
          IF Revenue# <= ARCatCodeRec(1).Recpt6 THEN
            FeeAmt# = ARCatCodeRec(1).BaseAmt6 + (ARCatCodeRec(1).Percent6 / 100) * (Revenue# - ARCatCodeRec(1).Maximum6)
            GOTO SkipEm
          END IF
          
        END IF
      END IF    'End Test for Code
    NEXT Snt&
  END IF        'End Test for Cat 1
  
SkipEm:
  CustFee# = Round#(CustFee# + FeeAmt#)
  FeeAmt# = 0
  
  RETURN
END SUB

SUB ExpiredLicenseListing
  
  SHARED Choice$()
  
  ReportFile$ = "AREXPLIC.PRN"  'Report File Name
  FF$ = CHR$(12)
  DevSpec$ = "S"
  MaxLines = 53
  LineCnt = 0
  
  GOSUB SelectExpiredLicenseListOutput
  IF Canceled$ = "Y" THEN
    EXIT SUB
  END IF
  
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  'PRINT #RptHandle, CHR$(27); CHR$(58); ' oki 320 12 cpi
  
  GOSUB PrintExpiredLicenseListRptHeader
  
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
    IF ARCustRec(1).Deleted <> "Y" THEN
      TCat$ = RTRIM$(Catagory$)
      IF TCat$ = RTRIM$(ARCustRec(1).BILLCAT1) OR TCat$ = RTRIM$(ARCustRec(1).BILLCAT2) OR TCat$ = RTRIM$(ARCustRec(1).BILLCAT3) OR TCat$ = RTRIM$(ARCustRec(1).BILLCAT4) OR TCat$ = RTRIM$(ARCustRec(1).BILLCAT5) OR TCat$ = "ALL" THEN
        'IF RTRIM$(Catagory$) = RTRIM$(ARCustRec(1).BillCat1) OR RTRIM$(Catagory$) = RTRIM$(ARCustRec(1).BillCat2)
        'OR RTRIM$(Catagory$) = RTRIM$(ARCustRec(1).BillCat3) OR RTRIM$(Catagory$) = RTRIM$(ARCustRec(1).BillCat4) OR RTRIM$(Catagory$) = RTRIM$(ARCustRec(1).BillCat5) OR RTRIM$(Catagory$) = "ALL" THEN
        IF ARCustRec(1).VALID < ExpDate THEN
          IF LineCnt >= MaxLines THEN
            PRINT #RptHandle, FF$
            GOSUB PrintExpiredLicenseListRptHeader
          END IF
          PRINT #RptHandle, CustomerNumber; TAB(10); ARCustRec(1).BILLNAME; TAB(50); ARCustRec(1).LICENSE; TAB(65); Num2Date$(ARCustRec(1).VALID)
          CustCnt = CustCnt + 1
          LineCnt = LineCnt + 1
        END IF
      END IF
    END IF
  NEXT Cnt!
  GOSUB PrintExpiredLicenseListRptEnding
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi
  CLOSE         'Close all open files now
  
  IF DevSpec$ = "S" THEN
    EntryPoint = 2
  ELSE
    EntryPoint = 5
  END IF
  Header$ = "Expired License Listing"
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  KILL ReportFile$
  
  EXIT SUB
  
  
PrintExpiredLicenseListRptHeader:
  Page = Page + 1
  PRINT #RptHandle, TAB(28); "Expired License Listing"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
  PRINT #RptHandle, "Category "; Catagory$
  PRINT #RptHandle, "Cust #"; TAB(10); "Customer Name"; TAB(50); "License #"; TAB(65); "Valid Thru"
  PRINT #RptHandle, STRING$(80, "=")
  LineCnt = 3
  RETURN
  
PrintExpiredLicenseListRptEnding:
  PRINT #RptHandle, STRING$(80, "-")
  PRINT #RptHandle, "Customers Counted: "; USING "####,#"; CustCnt
  PRINT #RptHandle, FF$
  RETURN
  
  
SelectExpiredLicenseListOutput:
  LibName$ = "AR"
  ScrnName$ = "AREXPLIC"
  
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
  
  'REDIM Choice$(350, 3)         ' assume maximum of 99 codes
  
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
  REDIM Choice$(NumOFARCatRecs, 2)         ' assume maximum of 99 codes

  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Name"
  Choice$(2, 0) = "Number"
  
  Choice$(0, 1) = "2"
  FOR Cnt! = 1 TO (NumOFARCatRecs - 1):
    GET ARCatFile, Cnt!, ARCatCodeRec(1)
    Choice$(Cnt!, 1) = ARCatCodeRec(1).CATCODE
  NEXT Cnt!
  Choice$(NumOFARCatRecs, 1) = "ALL"
  CLOSE ARCatFile
  
  
  Choice$(0, 2) = "3"
  Choice$(1, 2) = "SCREEN"
  Choice$(2, 2) = "PRINTER"
  
  Action = 1
  ShowCursor
  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%
  'printhelp Help$
  Action = 1
  
  DO
    
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      Order$ = Form$(1, 0)
      Catagory$ = Form$(2, 0)
      DevSpec$ = LEFT$(Form$(3, 0), 1)
      ExpDate = Date2Num%(Form$(4, 0))
      IF ExpDate = -32768 THEN
        PRINT CHR$(7);
        LOCATE 16, 37: COLOR 12: PRINT "NO DATE ENTERED"
        SLEEP 2
        LOCATE 16, 37: PRINT STRING$(15, 32);
        Frm(1).FldNo = 4
        Action = 1
      ELSE
        RETURN
      END IF
    CASE EscKey
      Canceled$ = "Y"
      RETURN
    END SELECT
  LOOP
  
END SUB

SUB LicenseListing
  
  SHARED Choice$()
  
  ReportFile$ = "ARLICLST.PRN"  'Report File Name
  FF$ = CHR$(12)
  MaxLines = 58
  LineCnt = 0
  CustCnt = 0
  
  GOSUB SelectLicenseListOutput
  IF Canceled$ = "Y" THEN
    GOTO AbortExit1
  END IF
  
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  
  'PRINT #RptHandle, CHR$(27); CHR$(58); ' oki 320 12 cpi
  
  GOSUB PrintLicenseListRptHeader
  
  ' Print Main Body
  OpenARCustFile NumOfArRecs, ARFile
  REDIM ARIdxRec(1) AS ARCustIDXRecType
  IdxFile = FREEFILE
  OPEN "ARCUST.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS #IdxFile LEN = LEN(ARIdxRec(1))
  
  IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN NumOfArRecs = LOF(IdxFile) \ LEN(ARIdxRec(1))
  
  FOR ZCnt& = 1 TO NumOfArRecs
    IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN
      GET IdxFile, ZCnt&, ARIdxRec(1)
      GET ARFile, ARIdxRec(1).IDXRECORD, ARCustRec(1)
      CustomerNumber = ARIdxRec(1).IDXRECORD
    ELSE
      GET ARFile, ZCnt&, ARCustRec(1)
      CustomerNumber = ZCnt&
    END IF
    IF (ARCustRec(1).Deleted <> "Y") AND (ARCustRec(1).Inactive <> "Y") THEN
      TCat$ = RTRIM$(Catagory$)
      IF TCat$ = RTRIM$(ARCustRec(1).BILLCAT1) OR TCat$ = RTRIM$(ARCustRec(1).BILLCAT2) OR TCat$ = RTRIM$(ARCustRec(1).BILLCAT3) OR TCat$ = RTRIM$(ARCustRec(1).BILLCAT4) OR TCat$ = RTRIM$(ARCustRec(1).BILLCAT5) OR TCat$ = "ALL" THEN
        'IF RTRIM$(Catagory$) = RTRIM$(ARCustRec(1).BillCat1) OR RTRIM$(Catagory$) = RTRIM$(ARCustRec(1).BillCat2)
        'OR RTRIM$(Catagory$) = RTRIM$(ARCustRec(1).BillCat3) OR RTRIM$(Catagory$) = RTRIM$(ARCustRec(1).BillCat4) OR RTRIM$(Catagory$) = RTRIM$(ARCustRec(1).BillCat5) OR RTRIM$(Catagory$) = "ALL" THEN
        
        IF VAL(ARCustRec(1).LICENSE) > 0 THEN
          
          IF LineCnt >= MaxLines THEN
            PRINT #RptHandle, FF$
            GOSUB PrintLicenseListRptHeader
          END IF
          PRINT #RptHandle, CustomerNumber; TAB(10); ARCustRec(1).BILLNAME; TAB(50); ARCustRec(1).LICENSE; TAB(65); Num2Date$(ARCustRec(1).VALID)
          CustCnt = CustCnt + 1
          LineCnt = LineCnt + 1
        END IF
      END IF
    END IF
    
  NEXT ZCnt&
  
  GOSUB PrintLicenseListRptEnding
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi
  CLOSE         'Close all open files now
  
  IF DevSpec$ = "S" THEN
    EntryPoint = 2
  ELSE
    EntryPoint = 5
  END IF
  Header$ = "Current License Listing"
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  KILL ReportFile$
  
AbortExit1:
  EXIT SUB
  
  
PrintLicenseListRptHeader:
  Page = Page + 1
  PRINT #RptHandle, TAB(20); "Business License Current License Listing"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
  PRINT #RptHandle, "Cust #"; TAB(10); "Customer Name"; TAB(50); "License #"; TAB(65); "Valid Thru"
  PRINT #RptHandle, STRING$(80, "=")
  LineCnt = 3
  RETURN
  
PrintLicenseListRptEnding:
  PRINT #RptHandle, STRING$(80, "-")
  PRINT #RptHandle, "Customers Counted: "; USING "#####,#"; CustCnt
  PRINT #RptHandle, FF$
  RETURN
  
SelectLicenseListOutput:
  LibName$ = "AR"
  ScrnName$ = "ARLICREP"
  
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
  
  'REDIM Choice$(350, 2)         ' assume maximum of 99 codes
  
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
  REDIM Choice$(NumOFARCatRecs, 2)         ' assume maximum of 99 codes
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Name"
  Choice$(2, 0) = "Number"
  
  Choice$(0, 1) = "2"
  FOR ZCnt& = 1 TO (NumOFARCatRecs - 1):
    GET ARCatFile, ZCnt&, ARCatCodeRec(1)
    Choice$(ZCnt&, 1) = ARCatCodeRec(1).CATCODE
  NEXT ZCnt&
  Choice$(NumOFARCatRecs, 1) = "ALL"
  CLOSE ARCatFile
  
  Choice$(0, 2) = "3"
  Choice$(1, 2) = "SCREEN"
  Choice$(2, 2) = "PRINTER"
  
  Action = 1
  ShowCursor
  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%
  'printhelp Help$
  Action = 1
  
  DO
    
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      Order$ = Form$(1, 0)
      Catagory$ = Form$(2, 0)
      DevSpec$ = LEFT$(Form$(3, 0), 1)
      RETURN
    CASE EscKey
      Canceled$ = "Y"
      RETURN
    END SELECT
  LOOP
  
END SUB

SUB OpenARCustFile (NumOfArRecs, ARFile)
  CLOSE ARFile
  
  ARCustRecLen = LEN(ARCustRec(1))
  ARFile = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARFile LEN = ARCustRecLen
  NumOfArRecs = LOF(ARFile) \ ARCustRecLen
  'FOR x = 1 TO NumOfArRecs
  'GET ARFile, x, ARCust(1)
  'PRINT ARCust(1).Custnumb; TAB(15); ARCust(1).FirstTrans
  'SLEEP 1
  'NEXT x
  'STOP
END SUB

SUB OpenARCustIdxFile (NumOfARIdxRecs, ARIdxFile)
  CLOSE ARIdxFile
  ARCustIdxRecLen = LEN(ARCustIdxRec(1))
  ARIdxFile = FREEFILE
  OPEN "ARCUST.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS ARIdxFile LEN = ARCustIdxRecLen
  NumOfARIdxRecs = LOF(ARIdxFile) \ ARCustIdxRecLen
END SUB

SUB QuickCustomer
  
  SHARED Choice$()
  
  ReportFile$ = "ARDetCus.PRN"  'Report File Name
  FF$ = CHR$(12)
  MaxLines = 58
  LineCnt = 0
  CustCnt = 0
  
  GOSUB SelectQCustomerOutput
  IF Canceled$ = "Y" THEN
    GOTO QuickCustExit
  END IF
  
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  'PRINT #RptHandle, CHR$(27); CHR$(58); ' oki 320 12 cpi
  
  GOSUB PrintQCustomerRptHeader
  
  ' Print Main Body
  OpenARCustFile NumOfArRecs, ARFile
  REDIM ARIdxRec(1) AS ARCustIDXRecType
  IdxFile = FREEFILE
  OPEN "ARCUST.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS #IdxFile LEN = LEN(ARIdxRec(1))
  'here
  IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN
    NumOfArRecs = LOF(IdxFile) \ LEN(ARIdxRec(1))
  END IF
  
  FOR Cnt! = 1 TO NumOfArRecs
    IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN
      GET IdxFile, Cnt!, ARIdxRec(1)
      GET ARFile, ARIdxRec(1).IDXRECORD, ARCustRec(1)
      CustomerNumber = ARIdxRec(1).IDXRECORD
    ELSE
      GET ARFile, Cnt!, ARCustRec(1)
      CustomerNumber = Cnt!
    END IF
    IF ARCustRec(1).Deleted <> "Y" THEN
      TCat$ = RTRIM$(Catagory$)
      IF TCat$ = RTRIM$(ARCustRec(1).BILLCAT1) OR TCat$ = RTRIM$(ARCustRec(1).BILLCAT2) OR TCat$ = RTRIM$(ARCustRec(1).BILLCAT3) OR TCat$ = RTRIM$(ARCustRec(1).BILLCAT4) OR TCat$ = RTRIM$(ARCustRec(1).BILLCAT5) OR TCat$ = "ALL" THEN
        IF LineCnt >= MaxLines THEN
          PRINT #RptHandle, FF$
          GOSUB PrintQCustomerRptHeader
        END IF
        PRINT #RptHandle, "Cust #"; CustomerNumber; TAB(25); "License # "; ARCustRec(1).LICENSE; TAB(50); "Valid to: "; Num2Date$(ARCustRec(1).VALID)
        PRINT #RptHandle, ARCustRec(1).BILLNAME; TAB(40); "Category: "; ARCustRec(1).BILLCAT1; " / "; ARCustRec(1).BILLCAT2; " / "; ARCustRec(1).BILLCAT3; " / "; ARCustRec(1).BILLCAT4; " / "; ARCustRec(1).BILLCAT5
        PRINT #RptHandle, RTRIM$(ARCustRec(1).ADDRESS1)
        PRINT #RptHandle, RTRIM$(ARCustRec(1).ADDRESS2); TAB(50); " W Phone: "; ARCustRec(1).WPHONE
        PRINT #RptHandle, RTRIM$(ARCustRec(1).CITY); ", "; ARCustRec(1).STATE; " "; ARCustRec(1).ZIPCODE
        PRINT #RptHandle, STRING$(80, "-")
        CustCnt = CustCnt + 1
        
        LineCnt = LineCnt + 6
      END IF
    END IF
  NEXT Cnt!
  
  GOSUB PrintQCustomerRptEnding
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi
  CLOSE         'Close all open files now
  
  IF DevSpec$ = "S" THEN
    EntryPoint = 2
  ELSE
    EntryPoint = 5
  END IF
  
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  KILL ReportFile$
  
QuickCustExit:
  EXIT SUB
  
  
PrintQCustomerRptHeader:
  Page = Page + 1
  PRINT #RptHandle, TAB(20); "Business License's  : Quick Customer Listing"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
  PRINT #RptHandle,
  PRINT #RptHandle, STRING$(80, "=")
  LineCnt = 3
  RETURN
  
PrintQCustomerRptEnding:
  PRINT #RptHandle, "Total Customers Printed: "; USING "#####"; CustCnt
  PRINT #RptHandle,
  PRINT #RptHandle, FF$
  RETURN
  
SelectQCustomerOutput:
  LibName$ = "AR"
  ScrnName$ = "ARQCUSRP"
  
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
  
  
  'REDIM Choice$(350, 3)         ' assume maximum of 99 codes
  
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
  REDIM Choice$(NumOFARCatRecs, 2)         ' assume maximum of 99 codes
  
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Name"
  Choice$(2, 0) = "Number"
  
  
  Choice$(0, 1) = "2"
  FOR Cnt! = 1 TO (NumOFARCatRecs - 1):
    GET ARCatFile, Cnt!, ARCatCodeRec(1)
    Choice$(Cnt!, 1) = ARCatCodeRec(1).CATCODE
  NEXT Cnt!
  Choice$(NumOFARCatRecs, 1) = "ALL"
  CLOSE ARCatFile
  
  
  Choice$(0, 2) = "3"
  Choice$(1, 2) = "SCREEN"
  Choice$(2, 2) = "PRINTER"
  
  Action = 1
  ShowCursor
  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%
  
  'printhelp Help$
  Action = 1
  
  DO
    
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      Order$ = Form$(1, 0)
      Catagory$ = Form$(2, 0)
      DevSpec$ = LEFT$(Form$(3, 0), 1)
      RETURN
    CASE EscKey
      Canceled$ = "Y"
      RETURN
    END SELECT
  LOOP
  
END SUB

SUB ShowNoCode
  LibName$ = "AR"
  ScrnName$ = "ARNOCODE"
  Help$ = "NEW A/R Customer Entry"
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

SUB TransactionJournal
  
  SHARED Choice$()
  ReportFile$ = "ARTRANS.PRN"   'Report File Name
  LPTPort% = 1
  FF$ = CHR$(12)
  MaxLines = 53
  LineCnt = 0
  size = 4000
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
  REDIM ARTransRec(1) AS ARTransRecType
  REDIM Cat$(300), CatAmt#(300), TotalAmt#(101)
  
  LibName$ = "AR"
  ScrnName$ = "ARTRREP"
  
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
  
  REDIM Choice$(4, 1)
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "All"
  Choice$(2, 0) = "Charges"
  Choice$(3, 0) = "Payments"
  Choice$(4, 0) = "UP/DN Adj."
  Choice$(0, 1) = "4"
  Choice$(1, 1) = "SCREEN"
  Choice$(2, 1) = "PRINTER"
  
  Action = 1
  ShowCursor
  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%
  
  FirstTime = True
  
  Action = 1
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      FirstTime = False
      LSET Form$(3, 0) = DATE$
      Action = 1
    END IF
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      DevSpec$ = LEFT$(Form$(4, 0), 1)
      Type$ = Form$(1, 0)
      BegDate = Date2Num%(Form$(2, 0))
      BegDate$ = Form$(2, 0)
      EndDate = Date2Num%(Form$(3, 0))
      EndDate$ = Form$(3, 0)
      IF EndDate < BegDate THEN CLOSE : EXIT SUB
      EXIT DO
    CASE EscKey
      Canceled$ = "Y"
      EXIT DO
    END SELECT
  LOOP
  
  'if they have canceled get out
  IF Canceled$ = "Y" THEN
    EXIT SUB
  END IF
  
  GOSUB GetReportInformation2
  
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  'PRINT #RptHandle, CHR$(27); CHR$(58); ' oki 320 12 cpi
  
  GOSUB PrintRptHeader2
  
  FOR Cnt = 1 TO Count
    GET ARTRFile, array(Cnt).RecNum, ARTransRec(1)
    IF VAL(ARTransRec(1).CustomerNumber) = 0 THEN
      GOTO BadCustSkip
    END IF
    
    IF LineCnt >= MaxLines THEN
      PRINT #RptHandle, FF$
      GOSUB PrintRptHeader2
    END IF
    
    'Get Customer
    OpenARCustFile NumOfArRecs, ARFile
    
    GET ARFile, VAL(ARTransRec(1).CustomerNumber), ARCustRec(1)
    Help$ = ARCustRec(1).CustName
    
    PRINT #RptHandle, Num2Date$(ARTransRec(1).TransDate);
    PRINT #RptHandle, TAB(13); LEFT$(ARCustRec(1).CustName, 25);
    PRINT #RptHandle, TAB(40); "";
    
    SELECT CASE ARTransRec(1).TransType
    CASE 1
      PRINT #RptHandle, "Charge";
    CASE 2
      PRINT #RptHandle, "Payment";
    CASE 6
      PRINT #RptHandle, "Penalty";
    'CASE 9
    ' PRINT #RptHandle, "Beg Bal";
    CASE 101
      PRINT #RptHandle, "UP Adj.";
    CASE 100
      PRINT #RptHandle, "DN Adj.";
    END SELECT
    
    'print
    
    PRINT #RptHandle, TAB(50); LEFT$(ARTransRec(1).TransDesc, 18);
    PRINT #RptHandle, TAB(69); USING "$$#####,#.##"; ARTransRec(1).TransAmount
    
    CLOSE ARFile
    TotalTrans = TotalTrans + 1
    LineCnt = LineCnt + 1
    TotalAmt# = TotalAmt# + ARTransRec(1).TransAmount
    REM total by catagory
    
TotalUp:
    TotalAmt#(ARTransRec(1).TransType) = TotalAmt#(ARTransRec(1).TransType) + ARTransRec(1).TransAmount
    
    IF ARTransRec(1).TransType = 2 THEN
      BILLCAT1$ = QPTrim$(ARCustRec(1).BILLCAT1)
      Fee1# = ARCustRec(1).Fee1
      BILLCAT2$ = QPTrim$(ARCustRec(1).BILLCAT2)
      Fee2# = ARCustRec(1).Fee2
      BILLCAT3$ = QPTrim$(ARCustRec(1).BILLCAT3)
      Fee3# = ARCustRec(1).Fee3
      BILLCAT4$ = QPTrim$(ARCustRec(1).BILLCAT4)
      Fee4# = ARCustRec(1).Fee4
      BILLCAT5$ = QPTrim$(ARCustRec(1).BILLCAT5)
      Fee5# = ARCustRec(1).Fee5
      
      TotalPaid# = ARTransRec(1).TransAmount
      
      IF TotalPaid# > Fee1# THEN
        FeePd# = Fee1#
        Catagory$ = BILLCAT1$
        GOSUB UpdateSummary
        LeftOver# = TotalPaid# - Fee1#
        LeftOver# = INT((LeftOver# * 100) + .5) / 100
      ELSE
        FeePd# = TotalPaid#
        Catagory$ = BILLCAT1$
        GOSUB UpdateSummary
        GOTO FinishSummary
      END IF
      
      IF LeftOver# > Fee2# THEN
        FeePd# = Fee2#
        Catagory$ = BILLCAT2$
        GOSUB UpdateSummary
        LeftOver# = LeftOver# - Fee2#
        LeftOver# = INT((LeftOver# * 100) + .5) / 100
      ELSE
        FeePd# = LeftOver#
        Catagory$ = BILLCAT2$
        GOSUB UpdateSummary
        GOTO FinishSummary
      END IF
      IF LeftOver# > Fee3# THEN
        FeePd# = Fee3#
        Catagory$ = BILLCAT3$
        GOSUB UpdateSummary
        LeftOver# = LeftOver# - Fee3#
        LeftOver# = INT((LeftOver# * 100) + .5) / 100
      ELSE
        FeePd# = LeftOver#
        Catagory$ = BILLCAT3$
        GOSUB UpdateSummary
        GOTO FinishSummary
      END IF
      IF LeftOver# > Fee4# THEN
        FeePd# = Fee4#
        Catagory$ = BILLCAT4$
        GOSUB UpdateSummary
        LeftOver# = LeftOver# - Fee4#
        LeftOver# = INT((LeftOver# * 100) + .5) / 100
      ELSE
        FeePd# = LeftOver#
        Catagory$ = BILLCAT4$
        GOSUB UpdateSummary
        GOTO FinishSummary
      END IF
      LeftOver# = INT((LeftOver# * 100) + .5) / 100
      FeePd# = LeftOver#
      Catagory$ = BILLCAT5$
      GOSUB UpdateSummary
      
      
FinishSummary:
      
    END IF
    
BadCustSkip:
    
  NEXT Cnt
  
  GOSUB PrintRptEnding2
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi
  CLOSE         'Close all open files now
  
  IF DevSpec$ = "P" THEN
    EntryPoint = 5
  ELSE
    EntryPoint = 2
  END IF
  Header$ = "Transaction Journal"
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  'KILL ReportFile$
  
  EXIT SUB
  
UpdateSummary:
  IF FeePd# <= 0 THEN RETURN
  IF CatCnt! = 0 THEN
    CatCnt! = 1
    Cat$(1) = Catagory$
    CatAmt#(1) = FeePd#
    RETURN
  END IF
  FOR CatFnd! = 1 TO CatCnt!
    IF Cat$(CatFnd!) = Catagory$ THEN
      CatAmt#(CatFnd!) = CatAmt#(CatFnd!) + FeePd#
      RETURN
    END IF
  NEXT CatFnd!
  CatCnt! = CatCnt! + 1
  Cat$(CatCnt!) = Catagory$
  CatAmt#(CatCnt!) = FeePd#
  RETURN
  
  
PrintRptHeader2:
  Page = Page + 1
  PRINT #RptHandle, TAB(22); "Business License : Transactions Journal"
  PRINT #RptHandle, ""
  PRINT #RptHandle, "               Report Date: "; DATE$; TAB(65); "Page #"; Page
  PRINT #RptHandle, "Beginning Transaction Date: "; BegDate$
  PRINT #RptHandle, "   Ending Transaction Date: "; EndDate$
  PRINT #RptHandle, ""
  PRINT #RptHandle, "  Date"; TAB(13); "Customer Name"; TAB(40); "Type"; TAB(50); "Description"; TAB(70); "Amount"
  PRINT #RptHandle, STRING$(80, "=")
  LineCnt = 5
  RETURN
  
PrintRptEnding2:
  PRINT #RptHandle, STRING$(80, "-")
  PRINT #RptHandle, TAB(69); USING "$$#####,#.##"; TotalAmt#
  FOR Cnt = 1 TO 101
    IF TotalAmt#(Cnt) <> 0 THEN
      PRINT #RptHandle, "Trans Type : ";
      SELECT CASE Cnt
      CASE 1
        PRINT #RptHandle, "Charges  ";
      CASE 2
        PRINT #RptHandle, "Payments ";
      CASE 6
        PRINT #RptHandle, "Penalty  ";
      CASE 9
        PRINT #RptHandle, "Beg Bal  ";
      CASE 101
        PRINT #RptHandle, "UP Adj.  ";
      CASE 100
        PRINT #RptHandle, "DN Adj.  ";
      END SELECT
      PRINT #RptHandle, "     Total Amount: "; USING "$$######,#.##"; TotalAmt#(Cnt)
    END IF
  NEXT Cnt
  PRINT #RptHandle, FF$
  
  IF CatCnt! > 0 THEN
    Page = Page + 1
    PRINT #RptHandle, TAB(20); "Business License : Transactions Journal "
    PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
    PRINT #RptHandle, "Total Payments by Category"
    LCnt = 1
    
    FOR ll = 1 TO CatCnt!
      IF LCnt > 55 THEN
        Page = Page + 1
        PRINT #RptHandle, TAB(20); "Business License : Transactions Journal "
        PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
        PRINT #RptHandle, "Total Payments by Category"
        LCnt = 3
      END IF
      
      'Get Catagory Desc First
      CatagoryDesc$ = ""
      REDIM ARCatCodeRec(1) AS ARNewCatCodeRecType
      ARCatCodeRecLen = LEN(ARCatCodeRec(1))
      ARCatFile = FREEFILE
      OPEN "ARCODE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARCatFile LEN = ARCatCodeRecLen
      NumOFARCatRecs = LOF(ARCatFile) \ ARCatCodeRecLen
      FOR Cnt = 1 TO NumOFARCatRecs
        GET ARCatFile, Cnt, ARCatCodeRec(1)
        IF Cat$(ll) = RTRIM$(ARCatCodeRec(1).CATCODE) THEN
          CatagoryDesc$ = ARCatCodeRec(1).CODEDESC
          EXIT FOR
        END IF
      NEXT Cnt
      
      CLOSE ARCatFile
      
      PRINT #RptHandle, Cat$(ll); TAB(10); CatagoryDesc$; TAB(50); USING "$$#######,#.##"; CatAmt#(ll)
      LCnt = LCnt + 1
    NEXT ll
    PRINT #RptHandle, FF$
  END IF
  RETURN
  
GetReportInformation2:
  ARTRFile = FREEFILE
  OPEN "ARTrans.Dat" FOR RANDOM ACCESS READ WRITE SHARED AS #ARTRFile LEN = LEN(ARTransRec(1))
  TRNumRecs& = LOF(ARTRFile) / LEN(ARTransRec(1))
  
  FOR Cnt! = 1 TO TRNumRecs&
    GET ARTRFile, Cnt!, ARTransRec(1)
    IF LEFT$(Type$, 1) = "A" THEN
      IF ARTransRec(1).TransDate >= BegDate AND ARTransRec(1).TransDate <= EndDate THEN
        Count = Count + 1
        array(Count).who = LTRIM$(STR$(ARTransRec(1).TransDate)) + ARTransRec(1).CustomerNumber + STRING$(4, " ")
        array(Count).RecNum = Cnt!
      ELSE
        'STOP
      END IF
    END IF
    IF LEFT$(Type$, 1) = "U" THEN
      IF ARTransRec(1).TransType = 100 OR ARTransRec(1).TransType = 101 THEN
        IF ARTransRec(1).TransDate >= BegDate AND ARTransRec(1).TransDate <= EndDate THEN
          Count = Count + 1
          array(Count).who = LTRIM$(STR$(ARTransRec(1).TransDate)) + ARTransRec(1).CustomerNumber + STRING$(4, " ")
          array(Count).RecNum = Cnt!
        END IF
      END IF
    END IF
    IF LEFT$(Type$, 1) = "C" THEN
      IF ARTransRec(1).TransType = 1 THEN
        IF ARTransRec(1).TransDate >= BegDate AND ARTransRec(1).TransDate <= EndDate THEN
          Count = Count + 1
          array(Count).who = LTRIM$(STR$(ARTransRec(1).TransDate)) + ARTransRec(1).CustomerNumber + STRING$(4, " ")
          array(Count).RecNum = Cnt!
        END IF
      END IF
    END IF
    IF LEFT$(Type$, 1) = "P" THEN
      IF ARTransRec(1).TransType = 2 THEN
        IF ARTransRec(1).TransDate >= BegDate AND ARTransRec(1).TransDate <= EndDate THEN
          Count = Count + 1
          array(Count).who = LTRIM$(STR$(ARTransRec(1).TransDate)) + ARTransRec(1).CustomerNumber + STRING$(4, " ")
          array(Count).RecNum = Cnt!
        END IF
      END IF
    END IF
  NEXT Cnt!
  
  SortT array(Start), Count, Dir, SSize, MOff, MSize
  RETURN
  
  
  
END SUB

