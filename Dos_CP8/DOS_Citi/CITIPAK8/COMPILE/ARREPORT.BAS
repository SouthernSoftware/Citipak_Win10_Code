DEFINT A-Z
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
DECLARE SUB CustomerInquiry ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB printhelp (H$)
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

  '$INCLUDE: 'DefCnf.BI'

  DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)
  
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'AR.bi'                        'A/R FILE LAYOUTS
  '$INCLUDE: 'GL.bi'

TYPE Struct
 who AS STRING * 14
 RecNum AS INTEGER
END TYPE



   DIM SHARED ARCustRec(1) AS ARCustRecType
   DIM SHARED ARCustIdxRec(1) AS ARCustIdxRecType
   CONST False = 0, True = NOT False

 STACK 8000

   '--Dim the choice array to the number of menu items
   REDIM MChoice$(1 TO 10)

   MChoice$(1) = "Customer Inquiry"
   MChoice$(2) = "Customer Balance Listing"
   MChoice$(3) = "Transaction Journal"
   MChoice$(4) = "Detailed Catagory Listing"
   MChoice$(5) = "Detailed Customer Listing"
   MChoice$(6) = "License Listing"
   MChoice$(7) = "Expired License Listing"
   MChoice$(8) = "Application Listing"
   MChoice$(9) = "Quick Customer Listing"
   MChoice$(10) = "Exit to OS"

   MaxLen = 0     'Set menu width to zero
   BoxBot = 17    'limit the box length to go no lower than line 20
   Action = 0     '0 means stay in the menu until they select something
   Choice = 1     'Pre-load choice to highlight

   '--Find max menu width
   FOR Cnt = 1 TO UBOUND(MChoice$)
     TLen = LEN(MChoice$(Cnt))
     IF TLen > MaxLen THEN
       MaxLen = TLen
     END IF
   NEXT

   '--Center Menu within Screen
   Row = ((25 - (UBOUND(MChoice$))) \ 2)
   Col = ((80 - MaxLen) \ 2) - 2
   help$ = "Main Reports Menu"
   
   DO

      '--Set upper left corner of menu, turn off the cursor
      LOCATE Row, Col, 0
      LibFile2Scrn "AR.QSL", "MENUBAK", MonoCode, -1, ErrorCode

      TitleBox 3, Col, MaxLen + 3, " Reports Menu ", Cnf
      TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

      'PrintTitle user$
     ' PrintHelp Help$

      ShowCursor

      VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf


      IF Ky$ = CHR$(27) THEN EXIT DO 'choice = 0

      SELECT CASE Choice
          CASE 1
          CustomerInquiry
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
          CASE IS = 10
          END
      END SELECT
   LOOP
   RUN "armenu"

SUB AppListing


  SHARED Choice$()


  ReportFile$ = "ARAPPLST.PRN"  'Report File Name
  CommaFmt$ = "########,.##"    'format takes 13 chars
  TotalFmt$ = "#########,.##"   'format takes 14 chars
  SumLine$ = STRING$(13, "-")   'column summary line
  DivLine$ = STRING$(77, "-")   'dashed line
  DivLine2$ = STRING$(77, "=")  'Double Line
  FF$ = CHR$(12)
  MaxLines = 58
  LineCnt = 0
  CustCnt = 0

  GOSUB SelectAppListOutput: IF Canceled$ = "Y" THEN EXIT SUB



  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle


 'PRINT #RptHandle, CHR$(27); CHR$(58); ' oki 320 12 cpi

  GOSUB PrintAppListHeader

 ' Print Main Body
  OpenARCustFile NumOfArRecs, ARFile
  REDIM ARIdxRec(1) AS ARCustIdxRecType
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
   IF RTRIM$(Catagory$) = RTRIM$(ARCustRec(1).BillCat) OR RTRIM$(Catagory$) = "ALL" THEN

   IF LineCnt >= MaxLines THEN
    PRINT #RptHandle, FF$
    GOSUB PrintAppListHeader
   END IF
   PRINT #RptHandle, "Cust #"; CustomerNumber; TAB(15); "Bus. Name: "; ARCustRec(1).CustName
   PRINT #RptHandle, "         Applicant Name: "; ARCustRec(1).BillName
   PRINT #RptHandle, "                Address: ";
   PRINT #RptHandle, TAB(26); RTRIM$(ARCustRec(1).Address1)
   PRINT #RptHandle, TAB(26); RTRIM$(ARCustRec(1).Address2)
   PRINT #RptHandle, TAB(26); RTRIM$(ARCustRec(1).City); ", "; ARCustRec(1).State; " "; ARCustRec(1).ZipCode
   PRINT #RptHandle, STRING$(80, "-")
    CustCnt = CustCnt + 1

    LineCnt = LineCnt + 6
   END IF
  END IF
    NEXT Cnt!
    GOSUB PrintAppListEnding
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


PrintAppListHeader:
    Page = Page + 1
    PRINT #RptHandle, TAB(20); "Business License Application Customer Listing"
    PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
    PRINT #RptHandle,
    PRINT #RptHandle, STRING$(80, "=")
    LineCnt = 3
    RETURN

PrintAppListEnding:
     PRINT #RptHandle, "Total Customers Printed: "; USING "#####"; CustCnt
     PRINT #RptHandle,
     PRINT #RptHandle, FF$
     RETURN







SelectAppListOutput:
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


  REDIM Choice$(300, 3)      ' assume maximum of 99 codes

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




  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Name"
  Choice$(2, 0) = "Number"


   Choice$(0, 1) = "2"
  FOR Cnt! = 1 TO (NumOFARCatRecs - 1):
  GET ARCatFile, Cnt!, ARCatCodeRec(1)

   Choice$(Cnt!, 1) = ARCatCodeRec(1).CatCode
  NEXT Cnt!
   Choice$(NumOFARCatRecs, 1) = "ALL"
  CLOSE ARCatFile


  Choice$(0, 2) = "3"
  Choice$(1, 2) = "SCREEN"
  Choice$(2, 2) = "PRINTER"

  Action = 1
  ShowCursor
  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%
  'printhelp help$
  Action = 1
  COLOR 14
  LOCATE 8, 24: PRINT "Application Listing "; : COLOR 10: PRINT "]"
  DO


    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode
     CASE F10Key
       Order$ = Form$(1, 0)
       Catagory$ = Form$(2, 0)
       DevSpec$ = LEFT$(Form$(3, 0), 1)
       RETURN
     CASE ESCKey
      Canceled$ = "Y"
      RETURN
    END SELECT
 LOOP

END SUB

SUB ARFixMess

 CLS

  REDIM ARTransRec(1) AS ARTransRecType

  ARTRFile = FREEFILE
  OPEN "ARTrans.Dat" FOR RANDOM ACCESS READ WRITE SHARED AS #ARTRFile LEN = LEN(ARTransRec(1))
  TrNumRecs = LOF(ARTRFile) / LEN(ARTransRec(1))
  


  FOR Cnt = 1 TO 104
   GET ARTRFile, Cnt, ARTransRec(1)
   
   'Get Customer
   OpenARCustFile NumOfArRecs, ARFile
   GET ARFile, VAL(ARTransRec(1).CustomerNumber), ARCustRec(1)
   ARCustRec(1).IssueLicense = "Y"
   ARCustRec(1).FeeAmt = ARTransRec(1).TransAmount
   PUT ARFile, VAL(ARTransRec(1).CustomerNumber), ARCustRec(1)


      IF ARTransRec(1).TransType = 1 THEN PRINT "Charge"; Cnt: STOP
      IF ARTransRec(1).TransType = 2 THEN PRINT "Payment"; Cnt
  
  NEXT Cnt
    CLOSE                       'Close all open files now


END SUB

SUB CatListing
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

  GOSUB SelectOutput
  IF Canceled$ = "Y" THEN EXIT SUB


  REDIM ARCatCodeRec(1) AS ARCatCodeRecType     ' open transaction file
  CatCodeRecLen = LEN(ARCatCodeRec(1))
  TrHandle = FREEFILE
  OPEN "ARCODE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS TrHandle LEN = CatCodeRecLen
  TrNumRecs = LOF(TrHandle) \ CatCodeRecLen


  GOSUB GetReportInformation

  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle


  PRINT #RptHandle, CHR$(27); CHR$(58); ' oki 320 12 cpi

  GOSUB PrintRptHeader

  FOR Cnt = 1 TO Count
   GET TrHandle, array(Cnt).RecNum, ARCatCodeRec(1)
   IF LEFT$(ARCatCodeRec(1).CatCode, 1) <> " " THEN
   IF LineCnt >= MaxLines THEN
    PRINT #RptHandle, FF$
    GOSUB PrintRptHeader
   END IF
   PRINT #RptHandle, ARCatCodeRec(1).CatCode; TAB(8); LEFT$(ARCatCodeRec(1).CodeDesc, 30);
     PRINT #RptHandle, TAB(40); QPTrim$(ARCatCodeRec(1).RevGLNum);
     PRINT #RptHandle, TAB(55); QPTrim$(ARCatCodeRec(1).ARGLAcct);
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
    Page = Page + 1
    PRINT #RptHandle, TAB(20); "Business License's  : Catagory Code Listing "
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

  FOR Cnt! = 1 TO TrNumRecs
      GET TrHandle, Cnt!, ARCatCodeRec(1)
      array(Cnt!).who = LEFT$(ARCatCodeRec(1).CatCode, 3) + STRING$(11, " ")
      array(Cnt!).RecNum = Cnt!
  NEXT Cnt!

  Count = TrNumRecs

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
  'printhelp help$
  Action = 1

  DO


    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode
     CASE F10Key
       DevSpec$ = LEFT$(Form$(1, 0), 1)
       RETURN
     CASE ESCKey
      Canceled$ = "Y"
      RETURN
    END SELECT
 LOOP
  RETURN

END SUB

SUB CustomerBalance


  SHARED Choice$()


  ReportFile$ = "ARCusBal.PRN"  'Report File Name
  CommaFmt$ = "########,.##"    'format takes 13 chars
  TotalFmt$ = "#########,.##"   'format takes 14 chars
  SumLine$ = STRING$(13, "-")   'column summary line
  DivLine$ = STRING$(77, "-")   'dashed line
  DivLine2$ = STRING$(77, "=")  'Double Line
  LPTPort% = 1
  FF$ = CHR$(12)
  MaxLines = 58
  LineCnt = 0
  CustCnt = 0

  GOSUB SelectCustBalOutput: IF Canceled$ = "Y" THEN EXIT SUB



  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle


 'PRINT #RptHandle, CHR$(27); CHR$(58); ' oki 320 12 cpi

  GOSUB PrintCustBalRptHeader

 ' Print Main Body
  OpenARCustFile NumOfArRecs, ARFile
  REDIM ARIdxRec(1) AS ARCustIdxRecType
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
   IF RTRIM$(Catagory$) = RTRIM$(ARCustRec(1).BillCat) OR RTRIM$(Catagory$) = "ALL" THEN
   IF ARCustRec(1).AcctBal <> 0 THEN
   IF LineCnt >= MaxLines THEN
    PRINT #RptHandle, FF$
    GOSUB PrintCustBalRptHeader
   END IF
   PRINT #RptHandle, USING "#####"; CustomerNumber;
   PRINT #RptHandle, TAB(10); ARCustRec(1).BillName;
   PRINT #RptHandle, TAB(50); USING "$$####,#.##"; ARCustRec(1).AcctBal
   CustCnt = CustCnt + 1
    TotalBal# = TotalBal# + ARCustRec(1).AcctBal
    LineCnt = LineCnt + 1
    END IF
   END IF
  END IF
    NEXT Cnt!
    GOSUB PrintCustBalRptEnding
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


PrintCustBalRptHeader:
    Page = Page + 1
    PRINT #RptHandle, TAB(20); "Business License's : Customer Balance Listing"
    PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
    PRINT #RptHandle, "Catagory: "; Catagory$
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


  REDIM Choice$(200, 3)      ' assume maximum of 99 codes

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
  
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Name"
  Choice$(2, 0) = "Number"
  Choice$(0, 1) = "2"

  FOR Cnt! = 1 TO (NumOFARCatRecs - 1):
    GET ARCatFile, Cnt!, ARCatCodeRec(1)
    Choice$(Cnt!, 1) = ARCatCodeRec(1).CatCode
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
  'printhelp help$
  Action = 1

  DO


    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode
     CASE F10Key
       Order$ = Form$(1, 0)
       Catagory$ = Form$(2, 0)
       DevSpec$ = LEFT$(Form$(3, 0), 1)
       RETURN
     CASE ESCKey
      Canceled$ = "Y"
      RETURN
    END SELECT
 LOOP

END SUB

SUB CustomerInquiry

CIMainBody:
  SHARED MChoice$
  LibName$ = "AR"
  ScrnName$ = "ARCUSTIQ"
  help$ = "Customer Inquiry"
  LOCATE 1, 1, 0
  LPTPort% = 1
  ShowCursor
  LibFile2Scrn "AR.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  printhelp help$



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

  CLOSE ARFile
  OpenARCustFile NumOfArRecs, ARFile

  OpenARCustIdxFile NumOfARIdxRecs, ARIdxFile

  IF NumOfArRecs = 0 THEN
   LibName$ = "AR"
   ScrnName$ = "ARNOCUST"
   help$ = "Edit A/R Customer Entry"
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
  printhelp help$

  DO
  Done = False
  EditForm Form$(), Fld(), Frm(1), Cnf, Action

  SELECT CASE Frm(1).KeyCode
   CASE ESCKey
    Done = True
  END SELECT
    IF Done = True THEN EXIT SUB
  LOOP

  END IF
  COLOR 15
  LOCATE 5, 40: PRINT "PRESS <enter> FOR CUSTOMER LIST"




  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF Frm(1).PrevFld = 2 AND CustomerGrabed = 0 THEN
     GOSUB GetCustomer
     LOCK #ARFile, AccountRecord
     Frm(1).FldNo = 26
    END IF


    SELECT CASE Frm(1).KeyCode
    CASE F4Key
     GOSUB PrintHistory
     LibFile2Scrn "AR.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
     OpenARCustFile NumOfArRecs, ARFile
     OpenARCustIdxFile NumOfARIdxRecs, ARIdxFile
     GOSUB GetCustomer
     Frm(1).FldNo = 26
    CASE F10Key
      CLOSE ARFile
      EXIT SUB
    CASE ESCKey
      CLOSE ARFile
      EXIT SUB
    END SELECT

  LOOP




FindCatagory:

  REDIM ARCatCodeRec(1) AS ARCatCodeRecType
  ARCatCodeRecLen = LEN(ARCatCodeRec(1))
  ARCatFile = FREEFILE
  OPEN "ARCODE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARCatFile LEN = ARCatCodeRecLen
  NumOFARCatRecs = LOF(ARCatFile) \ ARCatCodeRecLen

  REDIM MChoice$(1 TO NumOFARCatRecs)
  FOR Cnt = 1 TO NumOFARCatRecs
    GET ARCatFile, Cnt, ARCatCodeRec(1)
    MChoice$(Cnt) = SPACE$(50)
    LSET MChoice$(Cnt) = ARCatCodeRec(1).CatCode
    MID$(MChoice$(Cnt), 5) = ARCatCodeRec(1).CodeDesc
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

  '--Set upper left corner of menu, turn off the cursor
   LOCATE Row, Col, 0
   QPrintRC TText$, Row - 1, Col, 112
   VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
   GET ARCatFile, Choice, ARCatCodeRec(1)

   Form$(10, 0) = ARCatCodeRec(1).CatCode
   Form$(17, 0) = STR$(ARCatCodeRec(1).APPNUMB)
   Form$(18, 0) = STR$(ARCatCodeRec(1).BILLCODE)

   Fld(17).Protected = True
   Fld(18).Protected = True
   Frm(1).FldNo = 11
RETURN





GetCustomer:

CustomerGrabed = 0
AccountRecord = VAL(Form$(1, 0))

REM **************************************************************************

IF AccountRecord = 0 THEN

   MaxLen = 50     'Set menu width to zero
   BoxBot = 17    'limit the box length to go no lower than line 20
   Action = 0     '0 means stay in the menu until they select something
   Choice = 1     'Pre-load choice to highlight

   TText$ = SPACE$(MaxLen + 4)
   LSET TText$ = " Cust #    Customer Sort Name"

   '--Center Menu within Screen
   Row = 8
   Col = 15

   REDIM MChoice$(1 TO NumOfARIdxRecs)

  ChoiceCounter = 0
  FOR Cnt = 1 TO NumOfARIdxRecs
    GET ARIdxFile, Cnt, ARCustIdxRec(1)
    IF LEFT$(ARCustIdxRec(1).IdxName, 7) <> "DELETED" THEN
     ChoiceCounter = ChoiceCounter + 1
     MChoice$(ChoiceCounter) = SPACE$(50)
     LSET MChoice$(ChoiceCounter) = STR$(ARCustIdxRec(1).IDXRECORD)
     MID$(MChoice$(ChoiceCounter), 10) = ARCustIdxRec(1).IdxName
    END IF
  NEXT Cnt

   DO

      '--Set upper left corner of menu, turn off the cursor
      LOCATE Row, Col, 0
      LibFile2Scrn "AR.QSL", "MENUBAK", MonoCode, -1, ErrorCode
      ShowCursor
      QPrintRC TText$, Row - 1, Col, 112
      VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        AccountRecord = 0
        ExitFlag = True
      ELSE
        AccountRecord = VAL(LEFT$(MChoice$(Choice), 8))
        ExitFlag = True
      END IF

   LOOP UNTIL ExitFlag

  LibName$ = "AR"
  ScrnName$ = "ARCUSTIQ"
  help$ = "CUSTOMER INQUIRY"
  LOCATE 1, 1, 0

  ShowCursor
  LibFile2Scrn "AR.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%

   END IF

REM ************************************************************************
IF AccountRecord > 0 AND AccountRecord <= NumOfArRecs THEN
    GET ARFile, AccountRecord, ARCustRec(1)
    IF ARCustRec(1).Deleted = "Y" THEN GOTO CustomerDeleted
    Form$(1, 0) = ARCustRec(1).Custnumb
    Form$(2, 0) = ARCustRec(1).SortName
    Form$(3, 0) = ARCustRec(1).BillName
    Form$(4, 0) = ARCustRec(1).Address1
    Form$(5, 0) = ARCustRec(1).Address2
    Form$(6, 0) = ARCustRec(1).City
    Form$(7, 0) = ARCustRec(1).State
    Form$(8, 0) = ARCustRec(1).ZipCode
    Form$(9, 0) = ARCustRec(1).CustName
    Form$(10, 0) = ARCustRec(1).BillCat
    Form$(11, 0) = ARCustRec(1).SOSEC
    Form$(12, 0) = ARCustRec(1).DRVLIC
    Form$(13, 0) = Num2Date(ARCustRec(1).DATEOPED)
    Form$(14, 0) = ARCustRec(1).BILLCMT
    Form$(15, 0) = ARCustRec(1).PAYCMT
    Form$(16, 0) = ARCustRec(1).CASHONLY
    Form$(17, 0) = STR$(ARCustRec(1).APPNUMB)
    Form$(18, 0) = STR$(ARCustRec(1).BILLFORM)
    Form$(19, 0) = ARCustRec(1).HPhone
    Form$(20, 0) = ARCustRec(1).WPhone
    Form$(21, 0) = STR$(ARCustRec(1).FeeAmt)
    Form$(22, 0) = ARCustRec(1).License
    Form$(23, 0) = Num2Date$(ARCustRec(1).Valid)
    Form$(24, 0) = ARCustRec(1).IssueLicense
    Form$(25, 0) = STR$(ARCustRec(1).AcctBal)
    FOR Flds = 1 TO 25
     Fld(Flds).Protected = True
    NEXT Flds
    CustomerGrabed = 1
    Action = 1
    COLOR 15
    LOCATE 5, 40: PRINT STRING$(32, 32)
    RETURN

   ELSE

    LibName$ = "AR"
    ScrnName$ = "ARBADCUS"
    help$ = "Edit A/R Customer Entry"
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
    printhelp help$

    Done = False
    Action = 1


  DO

   EditForm Form$(), Fld(), Frm(1), Cnf, Action

   SELECT CASE Frm(1).KeyCode
    CASE ESCKey
     Done = True
     END SELECT
     IF Done = True THEN GOTO CIMainBody
  LOOP

 END IF

CustomerDeleted:
    LibName$ = "AR"
    ScrnName$ = "ARDELCUS"
    help$ = "Edit Customer"
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
    printhelp help$

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
   IF ARTransRec(1).TransType >= 1 AND ARTransRec(1).TransType <= 9 THEN
   PRINT #RptHandle, Num2Date$(ARTransRec(1).TransDate);
    PRINT #RptHandle, TAB(12); RTRIM$(ARTransRec(1).TransDesc);
     PRINT #RptHandle, TAB(50); "";
      IF ARTransRec(1).TransType = 1 THEN
       PRINT #RptHandle, "Charge";
       RunBalance# = RunBalance# + ARTransRec(1).TransAmount
      END IF
      IF ARTransRec(1).TransType = 2 THEN
       PRINT #RptHandle, "Payment";
       RunBalance# = RunBalance# - ARTransRec(1).TransAmount
      END IF
      IF ARTransRec(1).TransType = 9 THEN
       PRINT #RptHandle, "Beg Bal";
       RunBalance# = RunBalance# + ARTransRec(1).TransAmount
      END IF
      IF ARTransRec(1).TransType = 2 THEN
      PRINT #RptHandle, TAB(60); USING "####,#.##"; ARTransRec(1).TransAmount * -1;
      ELSE
      PRINT #RptHandle, TAB(60); USING "####,#.##"; ARTransRec(1).TransAmount;
      END IF
     PRINT #RptHandle, TAB(70); USING "####,#.##"; RunBalance#
    LineCnt = LineCnt + 1
   END IF
    TransRec& = ARTransRec(1).NextTrans

   WEND
   PRINT #RptHandle, CHR$(18); ' oki 320 10 cpi
   PRINT #RptHandle, CHR$(12);
   CLOSE                 'Close all open files now
                 
   PrintRptFile "Customer Account History", ReportFile$, LPTPort%, RetCode%, EntryPoint
   KILL ReportFile$
   RETURN

PrintCustInqRptHeader:
    Page = Page + 1
    PRINT #RptHandle, TAB(18); "Business License : Detailed Customer History"
    PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
    PRINT #RptHandle, "Customer Name: "; ARCustRec(1).BillName
    PRINT #RptHandle, "  Date"; TAB(10); "Description "; TAB(50); "Type"; TAB(60); "  Amount"; TAB(73); "Balance"
    PRINT #RptHandle, STRING$(80, "=")
    LineCnt = 3
    RETURN


RETURN


END SUB

SUB CustomerListing


  SHARED Choice$()


  ReportFile$ = "ARDetCus.PRN"  'Report File Name
  CommaFmt$ = "########,.##"    'format takes 13 chars
  TotalFmt$ = "#########,.##"   'format takes 14 chars
  SumLine$ = STRING$(13, "-")   'column summary line
  DivLine$ = STRING$(77, "-")   'dashed line
  DivLine2$ = STRING$(77, "=")  'Double Line
  FF$ = CHR$(12)
  MaxLines = 58
  LineCnt = 0
  CustCnt = 0

  GOSUB SelectDetailCustomerOutput: IF Canceled$ = "Y" THEN EXIT SUB

  

  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle


 'PRINT #RptHandle, CHR$(27); CHR$(58); ' oki 320 12 cpi

  GOSUB PrintDetailCustomerRptHeader

 ' Print Main Body
  OpenARCustFile NumOfArRecs, ARFile
  REDIM ARIdxRec(1) AS ARCustIdxRecType
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
   IF RTRIM$(Catagory$) = RTRIM$(ARCustRec(1).BillCat) OR RTRIM$(Catagory$) = "ALL" THEN

   IF LineCnt >= MaxLines THEN
    PRINT #RptHandle, FF$
    GOSUB PrintDetailCustomerRptHeader
   END IF
   PRINT #RptHandle, "Cust #"; CustomerNumber; TAB(25); "License # "; ARCustRec(1).License; TAB(50); "Valid to: "; Num2Date$(ARCustRec(1).Valid)
   PRINT #RptHandle, ARCustRec(1).BillName; TAB(50); "Catagory: "; ARCustRec(1).BillCat
   PRINT #RptHandle, RTRIM$(ARCustRec(1).Address1); TAB(50); " H Phone: "; ARCustRec(1).HPhone
   PRINT #RptHandle, RTRIM$(ARCustRec(1).Address2); TAB(50); " W Phone: "; ARCustRec(1).WPhone
   PRINT #RptHandle, RTRIM$(ARCustRec(1).City); ", "; ARCustRec(1).State; " "; ARCustRec(1).ZipCode; TAB(50); " Fee Amt: $"; USING "#####,#.##"; ARCustRec(1).FeeAmt
   PRINT #RptHandle, STRING$(80, "-")
    CustCnt = CustCnt + 1

    LineCnt = LineCnt + 6
   END IF
  END IF
    NEXT Cnt!
    GOSUB PrintDetailCustomerRptEnding
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


PrintDetailCustomerRptHeader:
    Page = Page + 1
    PRINT #RptHandle, TAB(20); "Business License's  : Detailed Customer Listing"
    PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
    PRINT #RptHandle,
    PRINT #RptHandle, STRING$(80, "=")
    LineCnt = 3
    RETURN

PrintDetailCustomerRptEnding:
     PRINT #RptHandle, "Total Customers Printed: "; USING "#####"; CustCnt
     PRINT #RptHandle,
     PRINT #RptHandle, FF$
     RETURN







SelectDetailCustomerOutput:
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

  
  REDIM Choice$(200, 3)      ' assume maximum of 99 codes

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
  

  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Name"
  Choice$(2, 0) = "Number"


   Choice$(0, 1) = "2"
  FOR Cnt! = 1 TO (NumOFARCatRecs - 1):
  GET ARCatFile, Cnt!, ARCatCodeRec(1)

   Choice$(Cnt!, 1) = ARCatCodeRec(1).CatCode
  NEXT Cnt!
   Choice$(NumOFARCatRecs, 1) = "ALL"
  CLOSE ARCatFile

  
  Choice$(0, 2) = "3"
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
       Catagory$ = Form$(2, 0)
       DevSpec$ = LEFT$(Form$(3, 0), 1)
       RETURN
     CASE ESCKey
      Canceled$ = "Y"
      RETURN
    END SELECT
 LOOP

END SUB

SUB ExpiredLicenseListing
  SHARED Choice$()


  ReportFile$ = "AREXPLIC.PRN"  'Report File Name
  CommaFmt$ = "########,.##"    'format takes 13 chars
  TotalFmt$ = "#########,.##"   'format takes 14 chars
  SumLine$ = STRING$(13, "-")   'column summary line
  DivLine$ = STRING$(77, "-")   'dashed line
  DivLine2$ = STRING$(77, "=")  'Double Line
  FF$ = CHR$(12)
  DevSpec$ = "S"
  MaxLines = 53
  LineCnt = 0

  GOSUB SelectExpiredLicenseListOutput: IF Canceled$ = "Y" THEN EXIT SUB



  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle


 'PRINT #RptHandle, CHR$(27); CHR$(58); ' oki 320 12 cpi

  GOSUB PrintExpiredLicenseListRptHeader

 ' Print Main Body
  OpenARCustFile NumOfArRecs, ARFile
  REDIM ARIdxRec(1) AS ARCustIdxRecType
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
  IF RTRIM$(Catagory$) = RTRIM$(ARCustRec(1).BillCat) OR RTRIM$(Catagory$) = "ALL" THEN

   IF ARCustRec(1).Valid < ExpDate THEN


   IF LineCnt >= MaxLines THEN
    PRINT #RptHandle, FF$
    GOSUB PrintExpiredLicenseListRptHeader
   END IF
   PRINT #RptHandle, CustomerNumber; TAB(10); ARCustRec(1).BillName; TAB(50); ARCustRec(1).License; TAB(65); Num2Date$(ARCustRec(1).Valid)
   CustCnt = CustCnt + 1
   LineCnt = LineCnt + 1
  END IF
 END IF
 END IF
    NEXT Cnt!
    GOSUB PrintExpiredLicenseListRptEnding
    PRINT #RptHandle, CHR$(18); ' oki 320 10 cpi
    CLOSE                       'Close all open files now

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
    PRINT #RptHandle, "Catagory "; Catagory$
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


  REDIM Choice$(300, 3)      ' assume maximum of 99 codes

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




  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Name"
  Choice$(2, 0) = "Number"


   Choice$(0, 1) = "2"
  FOR Cnt! = 1 TO (NumOFARCatRecs - 1):
  GET ARCatFile, Cnt!, ARCatCodeRec(1)

   Choice$(Cnt!, 1) = ARCatCodeRec(1).CatCode
  NEXT Cnt!
   Choice$(NumOFARCatRecs, 1) = "ALL"
  CLOSE ARCatFile


  Choice$(0, 2) = "3"
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
     CASE ESCKey
      Canceled$ = "Y"
      RETURN
    END SELECT
 LOOP

END SUB

SUB LicenseListing

  SHARED Choice$()


  ReportFile$ = "ARLICLST.PRN"  'Report File Name
  CommaFmt$ = "########,.##"    'format takes 13 chars
  TotalFmt$ = "#########,.##"   'format takes 14 chars
  SumLine$ = STRING$(13, "-")   'column summary line
  DivLine$ = STRING$(77, "-")   'dashed line
  DivLine2$ = STRING$(77, "=")  'Double Line
  FF$ = CHR$(12)
  MaxLines = 58
  LineCnt = 0
  CustCnt = 0

  GOSUB SelectLicenseListOutput: IF Canceled$ = "Y" THEN EXIT SUB



  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle


 'PRINT #RptHandle, CHR$(27); CHR$(58); ' oki 320 12 cpi

  GOSUB PrintLicenseListRptHeader

 ' Print Main Body
  OpenARCustFile NumOfArRecs, ARFile
  REDIM ARIdxRec(1) AS ARCustIdxRecType
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
  IF RTRIM$(Catagory$) = RTRIM$(ARCustRec(1).BillCat) OR RTRIM$(Catagory$) = "ALL" THEN

   IF VAL(ARCustRec(1).License) > 0 THEN

   IF LineCnt >= MaxLines THEN
    PRINT #RptHandle, FF$
    GOSUB PrintLicenseListRptHeader
   END IF
   PRINT #RptHandle, CustomerNumber; TAB(10); ARCustRec(1).BillName; TAB(50); ARCustRec(1).License; TAB(65); Num2Date$(ARCustRec(1).Valid)
   CustCnt = CustCnt + 1
   LineCnt = LineCnt + 1
   END IF
  END IF
  END IF

    NEXT Cnt!
    GOSUB PrintLicenseListRptEnding
    PRINT #RptHandle, CHR$(18); ' oki 320 10 cpi
    CLOSE                       'Close all open files now

  IF DevSpec$ = "S" THEN
       EntryPoint = 2
       ELSE
       EntryPoint = 5
  END IF
  Header$ = "Current License Listing"
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint

  KILL ReportFile$

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


  REDIM Choice$(300, 3)      ' assume maximum of 99 codes

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




  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Name"
  Choice$(2, 0) = "Number"


   Choice$(0, 1) = "2"
  FOR Cnt! = 1 TO (NumOFARCatRecs - 1):
   GET ARCatFile, Cnt!, ARCatCodeRec(1)
   Choice$(Cnt!, 1) = ARCatCodeRec(1).CatCode
  NEXT Cnt!
   Choice$(NumOFARCatRecs, 1) = "ALL"
  CLOSE ARCatFile


  Choice$(0, 2) = "3"
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
       Catagory$ = Form$(2, 0)
       DevSpec$ = LEFT$(Form$(3, 0), 1)
       RETURN
     CASE ESCKey
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
  CommaFmt$ = "########,.##"    'format takes 13 chars
  TotalFmt$ = "#########,.##"   'format takes 14 chars
  SumLine$ = STRING$(13, "-")   'column summary line
  DivLine$ = STRING$(77, "-")   'dashed line
  DivLine2$ = STRING$(77, "=")  'Double Line
  FF$ = CHR$(12)
  MaxLines = 58
  LineCnt = 0
  CustCnt = 0

  GOSUB SelectQCustomerOutput: IF Canceled$ = "Y" THEN EXIT SUB



  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle


 'PRINT #RptHandle, CHR$(27); CHR$(58); ' oki 320 12 cpi

  GOSUB PrintQCustomerRptHeader

 ' Print Main Body
  OpenARCustFile NumOfArRecs, ARFile
  REDIM ARIdxRec(1) AS ARCustIdxRecType
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
   IF RTRIM$(Catagory$) = RTRIM$(ARCustRec(1).BillCat) OR RTRIM$(Catagory$) = "ALL" THEN

   IF LineCnt >= MaxLines THEN
    PRINT #RptHandle, FF$
    GOSUB PrintQCustomerRptHeader
   END IF
   PRINT #RptHandle, "Cust #"; CustomerNumber; TAB(25); "License # "; ARCustRec(1).License; TAB(50); "Valid to: "; Num2Date$(ARCustRec(1).Valid)
   PRINT #RptHandle, ARCustRec(1).BillName; TAB(50); "Catagory: "; ARCustRec(1).BillCat
   PRINT #RptHandle, RTRIM$(ARCustRec(1).Address1); TAB(50); " H Phone: "; ARCustRec(1).HPhone
   PRINT #RptHandle, RTRIM$(ARCustRec(1).Address2); TAB(50); " W Phone: "; ARCustRec(1).WPhone
   PRINT #RptHandle, RTRIM$(ARCustRec(1).City); ", "; ARCustRec(1).State; " "; ARCustRec(1).ZipCode
   PRINT #RptHandle, STRING$(80, "-")
    CustCnt = CustCnt + 1

    LineCnt = LineCnt + 6
   END IF
  END IF
    NEXT Cnt!
    GOSUB PrintQCustomerRptEnding
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


  REDIM Choice$(200, 3)      ' assume maximum of 99 codes

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


  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Name"
  Choice$(2, 0) = "Number"


   Choice$(0, 1) = "2"
  FOR Cnt! = 1 TO (NumOFARCatRecs - 1):
  GET ARCatFile, Cnt!, ARCatCodeRec(1)

   Choice$(Cnt!, 1) = ARCatCodeRec(1).CatCode
  NEXT Cnt!
   Choice$(NumOFARCatRecs, 1) = "ALL"
  CLOSE ARCatFile


  Choice$(0, 2) = "3"
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
       Catagory$ = Form$(2, 0)
       DevSpec$ = LEFT$(Form$(3, 0), 1)
       RETURN
     CASE ESCKey
      Canceled$ = "Y"
      RETURN
    END SELECT
 LOOP

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
    CASE ESCKey
     EXIT SUB
    END SELECT

  LOOP
END SUB

SUB TransactionJournal

  SHARED Choice$()
  ReportFile$ = "ARTRANS.PRN"  'Report File Name
  CommaFmt$ = "########,.##"    'format takes 13 chars
  TotalFmt$ = "#########,.##"   'format takes 14 chars
  SumLine$ = STRING$(13, "-")   'column summary line
  DivLine$ = STRING$(77, "-")   'dashed line
  DivLine2$ = STRING$(77, "=")  'Double Line
  LPTPort% = 1
  FF$ = CHR$(12)
  MaxLines = 53
  LineCnt = 0
  TotDr# = 0
  TotCr# = 0
  size = 4000
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
  REDIM ARTransRec(1) AS ARTransRecType
  REDIM TotalAmt#(99), ChrgeCat$(199), ChrgeAmt#(199), PayCat$(199), PayAmt#(199)


  GOSUB SelectOutput2
  IF Canceled$ = "Y" THEN EXIT SUB

  GOSUB GetReportInformation2

  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle


  'PRINT #RptHandle, CHR$(27); CHR$(58); ' oki 320 12 cpi

  GOSUB PrintRptHeader2

  FOR Cnt = 1 TO Count
   GET ARTRFile, array(Cnt).RecNum, ARTransRec(1)
   
   IF LineCnt >= MaxLines THEN
    PRINT #RptHandle, FF$
    GOSUB PrintRptHeader2
   END IF


   'Get Customer
   OpenARCustFile NumOfArRecs, ARFile
   GET ARFile, VAL(ARTransRec(1).CustomerNumber), ARCustRec(1)
   help$ = ARCustRec(1).CustName
   printhelp help$
   BillCat$ = QPTrim$(ARCustRec(1).BillCat)

   PRINT #RptHandle, Num2Date$(ARTransRec(1).TransDate);
    PRINT #RptHandle, TAB(13); LEFT$(ARCustRec(1).CustName, 25);
     PRINT #RptHandle, TAB(40); "";
      IF ARTransRec(1).TransType = 1 THEN PRINT #RptHandle, "Charge";
      IF ARTransRec(1).TransType = 2 THEN PRINT #RptHandle, "Payment";
      IF ARTransRec(1).TransType = 9 THEN PRINT #RptHandle, "Beg Bal";
       PRINT #RptHandle, TAB(50); LEFT$(ARTransRec(1).TransDesc, 18);
       PRINT #RptHandle, TAB(69); USING "$$#####,#.##"; ARTransRec(1).TransAmount

    CLOSE ARFile
    TotalTrans = TotalTrans + 1
    LineCnt = LineCnt + 1
    TotalAmt# = TotalAmt# + ARTransRec(1).TransAmount
   REM total by catagory
   IF ARTransRec(1).TransType = 1 THEN
    IF ChrgeCat = 0 THEN
       ChrgeCat = 1: ChrgeCat$(1) = BillCat$: ChrgeAmt#(1) = ARTransRec(1).TransAmount
       GOTO TotalUp
    END IF
    FOR Snt! = 1 TO ChrgeCat
     IF BillCat$ = ChrgeCat$(Snt!) THEN
      ChrgeAmt#(Snt!) = ChrgeAmt#(Snt!) + ARTransRec(1).TransAmount
      GOTO TotalUp
     END IF
    NEXT Snt!
    ChrgeCat = ChrgeCat + 1
    ChrgeCat$(ChrgeCat) = BillCat$: ChrgeAmt#(ChrgeCat) = ARTransRec(1).TransAmount
    GOTO TotalUp
   END IF
   IF ARTransRec(1).TransType = 2 THEN
    IF PayCat = 0 THEN
       PayCat = 1: PayCat$(1) = BillCat$: PayAmt#(1) = ARTransRec(1).TransAmount
       GOTO TotalUp
    END IF
    FOR Snt! = 1 TO PayCat
     IF BillCat$ = PayCat$(Snt!) THEN
      PayAmt#(Snt!) = PayAmt#(Snt!) + ARTransRec(1).TransAmount
      GOTO TotalUp
     END IF
    NEXT Snt!
    PayCat = PayCat + 1
    PayCat$(PayCat) = BillCat$: PayAmt#(PayCat) = ARTransRec(1).TransAmount
    GOTO TotalUp

   END IF





TotalUp:
    TotalAmt#(ARTransRec(1).TransType) = TotalAmt#(ARTransRec(1).TransType) + ARTransRec(1).TransAmount
   
  NEXT Cnt

    GOSUB PrintRptEnding2
    PRINT #RptHandle, CHR$(18); ' oki 320 10 cpi
    CLOSE                       'Close all open files now

  IF DevSpec$ = "S" THEN
       EntryPoint = 2
       ELSE
       EntryPoint = 5
  END IF
  Header$ = "Transaction Journal"
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint

  KILL ReportFile$

  EXIT SUB


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
     FOR Cnt! = 1 TO 99
     IF TotalAmt#(Cnt!) <> 0 THEN
      PRINT #RptHandle, "Trans Type : ";
      IF Cnt! = 1 THEN PRINT #RptHandle, "Charges  ";
      IF Cnt! = 2 THEN PRINT #RptHandle, "Payments ";
      IF Cnt! = 9 THEN PRINT #RptHandle, "Beg Bal  ";
      PRINT #RptHandle, "     Total Amount: "; USING "$$######,#.##"; TotalAmt#(Cnt!)
     END IF
     NEXT Cnt!
     PRINT #RptHandle, FF$
     IF ChrgeCat > 0 THEN
      Page = Page + 1
      PRINT #RptHandle, TAB(20); "Business License : Transactions Journal "
      PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
      PRINT #RptHandle, "Total Charges by Catagory"
      LCnt = 1
  FOR ll = 1 TO ChrgeCat

    IF LCnt > 55 THEN
     Page = Page + 1
     PRINT #RptHandle, TAB(20); "Business License : Transactions Journal "
     PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
     PRINT #RptHandle, "Total Charges by Catagory"
     LCnt = 3
    END IF



   'Get Catagory Desc First
     CatagoryDesc$ = ""
     REDIM ARCatCodeRec(1) AS ARCatCodeRecType
     ARCatCodeRecLen = LEN(ARCatCodeRec(1))
     ARCatFile = FREEFILE
     OPEN "ARCODE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARCatFile LEN = ARCatCodeRecLen
     NumOFARCatRecs = LOF(ARCatFile) \ ARCatCodeRecLen
     FOR Cnt = 1 TO NumOFARCatRecs
       GET ARCatFile, Cnt, ARCatCodeRec(1)
        IF ChrgeCat$(ll) = RTRIM$(ARCatCodeRec(1).CatCode) THEN
         CatagoryDesc$ = ARCatCodeRec(1).CodeDesc
         EXIT FOR
        END IF
     NEXT Cnt
     CLOSE ARCatFile

    PRINT #RptHandle, ChrgeCat$(ll); TAB(10); CatagoryDesc$; TAB(50); USING "$$#######,#.##"; ChrgeAmt#(ll)
    LCnt = LCnt + 1
     NEXT ll
     PRINT #RptHandle, FF$
     END IF
     IF PayCat > 0 THEN
      Page = Page + 1
      PRINT #RptHandle, TAB(20); "Business License : Transactions Journal "
      PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
      PRINT #RptHandle, "Total Payments by Catagory"
      LCnt = 1
  FOR ll = 1 TO PayCat

    IF LCnt > 55 THEN
     Page = Page + 1
     PRINT #RptHandle, TAB(20); "Business License : Transactions Journal "
     PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
     PRINT #RptHandle, "Total Payments by Catagory"
     LCnt = 3
    END IF



   'Get Catagory Desc First
     CatagoryDesc$ = ""
     REDIM ARCatCodeRec(1) AS ARCatCodeRecType
     ARCatCodeRecLen = LEN(ARCatCodeRec(1))
     ARCatFile = FREEFILE
     OPEN "ARCODE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARCatFile LEN = ARCatCodeRecLen
     NumOFARCatRecs = LOF(ARCatFile) \ ARCatCodeRecLen
     FOR Cnt = 1 TO NumOFARCatRecs
       GET ARCatFile, Cnt, ARCatCodeRec(1)
        IF PayCat$(ll) = RTRIM$(ARCatCodeRec(1).CatCode) THEN
         CatagoryDesc$ = ARCatCodeRec(1).CodeDesc
         EXIT FOR
        END IF
     NEXT Cnt
     CLOSE ARCatFile

    PRINT #RptHandle, PayCat$(ll); TAB(10); CatagoryDesc$; TAB(50); USING "$$#######,#.##"; PayAmt#(ll)
    LCnt = LCnt + 1
     NEXT ll
     PRINT #RptHandle, FF$
     END IF

     RETURN

GetReportInformation2:
  ARTRFile = FREEFILE
  OPEN "ARTrans.Dat" FOR RANDOM ACCESS READ WRITE SHARED AS #ARTRFile LEN = LEN(ARTransRec(1))
  TrNumRecs = LOF(ARTRFile) / LEN(ARTransRec(1))

  FOR Cnt! = 1 TO TrNumRecs
      GET ARTRFile, Cnt!, ARTransRec(1)
      IF LEFT$(Type$, 1) = "A" THEN
      IF ARTransRec(1).TransDate >= BegDate AND ARTransRec(1).TransDate <= EndDate THEN
       Count = Count + 1
       array(Count).who = LTRIM$(STR$(ARTransRec(1).TransDate)) + ARTransRec(1).CustomerNumber + STRING$(4, " ")
       array(Count).RecNum = Cnt!
      END IF
      END IF
      IF LEFT$(Type$, 1) = "B" THEN
      IF ARTransRec(1).TransType = 9 THEN
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






SelectOutput2:
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
  Choice$(4, 0) = "Beg Bal"
  Choice$(0, 1) = "4"
  Choice$(1, 1) = "SCREEN"
  Choice$(2, 1) = "PRINTER"




  Action = 1
  ShowCursor
  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%
  'printhelp help$
  Form$(2, 0) = DATE$
  Form$(3, 0) = DATE$
  Action = 1

  DO


    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode
     CASE F10Key
       DevSpec$ = LEFT$(Form$(4, 0), 1)
       Type$ = Form$(1, 0)
       BegDate = Date2Num%(Form$(2, 0))
       BegDate$ = Form$(2, 0)
       EndDate = Date2Num%(Form$(3, 0))
       EndDate$ = Form$(3, 0)
       IF EndDate < BegDate THEN CLOSE : EXIT SUB
       RETURN
     CASE ESCKey
      Canceled$ = "Y"
      RETURN
    END SELECT
 LOOP
  RETURN


END SUB

