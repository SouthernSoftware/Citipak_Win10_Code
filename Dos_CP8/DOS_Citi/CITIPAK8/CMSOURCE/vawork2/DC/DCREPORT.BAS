DEFINT A-Z
DECLARE SUB ResidentReport ()
DECLARE SUB OwnerReport ()
DECLARE SUB NonOwnerReport ()
DECLARE SUB CustomerBalance ()
DECLARE SUB ExpiredLicenseListing ()
DECLARE SUB LicenseListing ()
DECLARE SUB OpenDCCustIdxFile (NumOfDCIdxRecs%, DCIdxFile%)
DECLARE SUB ARFixMess ()
DECLARE SUB ShowNoCode ()
DECLARE SUB OpenDCCustFile (NumOfDcRecs, DCFile)
DECLARE SUB TransactionJournal ()
DECLARE SUB CustomerListing ()
DECLARE SUB CatListing ()
DECLARE SUB CustomerInquiry ()
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

  '$INCLUDE: 'DefCnf.BI'

  DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)
  
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'DC.bi'                        'A/R FILE LAYOUTS
  '$INCLUDE: 'GL.bi'

TYPE Struct
 who AS STRING * 14
 RecNum AS INTEGER
END TYPE



   DIM SHARED DCCustRec(1) AS DCCustRecType
   DIM SHARED DCCustIdxRec(1) AS DCCustIdxRecType
   CONST False = 0, True = NOT False

 STACK 8000

   '--Dim the choice array to the number of menu items
   REDIM MChoice$(1 TO 11)

   MChoice$(1) = "Customer Inquiry"
   MChoice$(2) = "Customer Balance Listing"
   MChoice$(3) = "Transaction Journal"
   MChoice$(4) = "Detailed Catagory Listing"
   MChoice$(5) = "Detailed Customer Listing"
   MChoice$(6) = "Decal Listing"
   MChoice$(7) = "Expired Decal Listing"
   MChoice$(8) = "Resident Report"
   MChoice$(9) = "Owner Report"
   MChoice$(10) = "Non-Owner Report"
   MChoice$(11) = "Exit to OS"

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
      LibFile2Scrn "DC.QSL", "MENUBAK", MonoCode, -1, ErrorCode

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
          CASE IS = 8
          ResidentReport
          CASE IS = 9
          OwnerReport
          CASE IS = 10
          NonOwnerReport
          CASE IS = 11
          END
      END SELECT
   LOOP
   RUN "dcmenu"

SUB CatListing
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


  REDIM DCCatCodeRec(1) AS DCCatCodeRecType     ' open transaction file
  CatCodeRecLen = LEN(DCCatCodeRec(1))
  TrHandle = FREEFILE
  OPEN "DCCODE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS TrHandle LEN = CatCodeRecLen
  TrNumRecs = LOF(TrHandle) \ CatCodeRecLen


  GOSUB GetReportInformation

  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle


  PRINT #RptHandle, CHR$(27); CHR$(58); ' oki 320 12 cpi

  GOSUB PrintRptHeader

  FOR Cnt = 1 TO Count
   GET TrHandle, array(Cnt).RecNum, DCCatCodeRec(1)
   IF LEFT$(DCCatCodeRec(1).CATCODE, 1) <> " " THEN
   IF LineCnt >= MaxLines THEN
    PRINT #RptHandle, FF$
    GOSUB PrintRptHeader
   END IF
   PRINT #RptHandle, DCCatCodeRec(1).CATCODE; TAB(8); LEFT$(DCCatCodeRec(1).CODEDESC, 28);
    PRINT #RptHandle, TAB(42); USING "####.##"; DCCatCodeRec(1).Fee;
    PRINT #RptHandle, TAB(50); DCCatCodeRec(1).CASHACCT; TAB(66); DCCatCodeRec(1).REVGLNUM
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
    PRINT #RptHandle, TAB(20); "Va. Decal System : Decal Catagory Code Listing "
    PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
    PRINT #RptHandle,
    PRINT #RptHandle, "Code "; TAB(8); "Description"; TAB(41); "Std Fee"; TAB(50); "Cash GL #"; TAB(66); "Revenue GL #"
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

  Count = TrNumRecs

SortT array(Start), Count, Dir, SSize, MOff, MSize
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
     CASE ESCKey
      Canceled$ = "Y"
      RETURN
    END SELECT
 LOOP
  RETURN

END SUB

SUB CustomerBalance


  SHARED Choice$()


  ReportFile$ = "DCCusBal.PRN"  'Report File Name
  CommaFmt$ = "########,.##"    'format takes 13 chars
  TotalFmt$ = "#########,.##"   'format takes 14 chars
  SumLine$ = STRING$(13, "-")   'column summary line
  DivLine$ = STRING$(77, "-")   'dashed line
  DivLine2$ = STRING$(77, "=")  'Double Line
  FF$ = CHR$(12)
  MaxLines = 58
  LineCnt = 0
  CustCnt = 0
  LPTPort% = 1

  GOSUB SelectCustBalOutput: IF Canceled$ = "Y" THEN EXIT SUB



  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle


 'PRINT #RptHandle, CHR$(27); CHR$(58); ' oki 320 12 cpi

  GOSUB PrintCustBalRptHeader

 ' Print Main Body
  OpenDCCustFile NumOfDcRecs, DCFile
  REDIM DCIdxRec(1) AS DCCustIdxRecType
  IdxFile = FREEFILE
  OPEN "DCCust.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS #IdxFile LEN = LEN(DCIdxRec(1))

  IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN NumOfDcRecs = LOF(IdxFile) \ LEN(DCIdxRec(1))

  FOR Cnt! = 1 TO NumOfDcRecs
  IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN
   GET IdxFile, Cnt!, DCIdxRec(1)
   GET DCFile, DCIdxRec(1).IDXRECORD, DCCustRec(1)
   CustomerNumber = DCIdxRec(1).IDXRECORD
   ELSE
   GET DCFile, Cnt!, DCCustRec(1)
   CustomerNumber = Cnt!
  END IF
  IF DCCustRec(1).Deleted <> "Y" THEN
   IF DCCustRec(1).AcctBal <> 0 THEN
   IF LineCnt >= MaxLines THEN
    PRINT #RptHandle, FF$
    GOSUB PrintCustBalRptHeader
   END IF
   PRINT #RptHandle, USING "#####"; CustomerNumber;
   PRINT #RptHandle, TAB(10); DCCustRec(1).BillName;
   PRINT #RptHandle, TAB(50); USING "$$####,#.##"; DCCustRec(1).AcctBal
   CustCnt = CustCnt + 1
    TotalBal# = TotalBal# + DCCustRec(1).AcctBal
    LineCnt = LineCnt + 1
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
    PRINT #RptHandle, TAB(21); "Va. Decal System : Detailed Customer Listing"
    PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
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
  LibName$ = "DC"
  ScrnName$ = "DCCUSBAL"

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


  REDIM Choice$(100, 3)      ' assume maximum of 99 codes

  REDIM DCCatCodeRec(1) AS DCCatCodeRecType
  DCCatCodeRecLen = LEN(DCCatCodeRec(1))
  DCCatFile = FREEFILE
  OPEN "DCCODE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCCatFile LEN = DCCatCodeRecLen
  NumOFDCCatRecs = LOF(DCCatFile) \ DCCatCodeRecLen

  IF NumOFDCCatRecs = 0 THEN
   CLOSE DCCatFile
   ShowNoCode
   EXIT SUB
  END IF

  NumOFDCCatRecs = NumOFDCCatRecs + 1




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
     CASE ESCKey
      Canceled$ = "Y"
      RETURN
    END SELECT
 LOOP

END SUB

SUB CustomerInquiry
 LPTPort% = 1
CIMainBody:
  SHARED MChoice$
  LibName$ = "DC"
  ScrnName$ = "DCCUSTIQ"
  help$ = "Customer Inquiry"
  LOCATE 1, 1, 0

  ShowCursor
  LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  PrintHelp help$



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

  CLOSE DCFile
  OpenDCCustFile NumOfDcRecs, DCFile

  OpenDCCustIdxFile NumOfDCIdxRecs, DCIdxFile

  IF NumOfDcRecs = 0 THEN
   LibName$ = "DC"
   ScrnName$ = "DCNOCUST"
   help$ = "Customer Inquiry"
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
  LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  PrintHelp help$

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
     LOCK #DCFile, AccountRecord
     Frm(1).FldNo = 23
    END IF


    SELECT CASE Frm(1).KeyCode
    CASE F4Key
     GOSUB PrintHistory

     LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
     OpenDCCustFile NumOfDcRecs, DCFile
     OpenDCCustIdxFile NumOfDCIdxRecs, DCIdxFile
     GOSUB GetCustomer
     Action = 1
     Frm(1).FldNo = 23
    CASE F10Key
      CLOSE DCFile
      EXIT SUB
    CASE ESCKey
      CLOSE DCFile
      EXIT SUB
    END SELECT

  LOOP




FindCatagory:

  REDIM DCCatCodeRec(1) AS DCCatCodeRecType
  DCCatCodeRecLen = LEN(DCCatCodeRec(1))
  DCCatFile = FREEFILE
  OPEN "DCCODE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCCatFile LEN = DCCatCodeRecLen
  NumOFDCCatRecs = LOF(DCCatFile) \ DCCatCodeRecLen

  REDIM MChoice$(1 TO NumOFDCCatRecs)
  FOR Cnt = 1 TO NumOFDCCatRecs
    GET DCCatFile, Cnt, DCCatCodeRec(1)
    MChoice$(Cnt) = SPACE$(50)
    LSET MChoice$(Cnt) = DCCatCodeRec(1).CATCODE
    MID$(MChoice$(Cnt), 5) = DCCatCodeRec(1).CODEDESC
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
   GET DCCatFile, Choice, DCCatCodeRec(1)

   Form$(11, 0) = DCCatCodeRec(1).CATCODE
   
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

   REDIM MChoice$(1 TO NumOfDCIdxRecs)

  ChoiceCounter = 0
  FOR Cnt = 1 TO NumOfDCIdxRecs
    GET DCIdxFile, Cnt, DCCustIdxRec(1)
    IF LEFT$(DCCustIdxRec(1).IDXNAME, 7) <> "DELETED" THEN
     ChoiceCounter = ChoiceCounter + 1
     MChoice$(ChoiceCounter) = SPACE$(50)
     LSET MChoice$(ChoiceCounter) = STR$(DCCustIdxRec(1).IDXRECORD)
     MID$(MChoice$(ChoiceCounter), 10) = DCCustIdxRec(1).IDXNAME
    END IF
  NEXT Cnt

   DO

      '--Set upper left corner of menu, turn off the cursor
      LOCATE Row, Col, 0
      LibFile2Scrn "DC.QSL", "MENUBAK", MonoCode, -1, ErrorCode
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

  LibName$ = "DC"
  ScrnName$ = "DCCUSTIQ"
  help$ = "CUSTOMER INQUIRY"
  LOCATE 1, 1, 0

  ShowCursor
  LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%

   END IF

REM ************************************************************************
IF AccountRecord > 0 AND AccountRecord <= NumOfDcRecs THEN
    GET DCFile, AccountRecord, DCCustRec(1)
    IF DCCustRec(1).Deleted = "Y" THEN GOTO CustomerDeleted
    Form$(1, 0) = DCCustRec(1).CUSTNUMB
    Form$(2, 0) = DCCustRec(1).SORTNAME
    Form$(3, 0) = DCCustRec(1).BillName
    Form$(4, 0) = DCCustRec(1).ADDRESS1
    Form$(5, 0) = DCCustRec(1).ADDRESS2
    Form$(6, 0) = DCCustRec(1).CITY
    Form$(7, 0) = DCCustRec(1).STATE
    Form$(8, 0) = DCCustRec(1).ZIPCODE
    Form$(9, 0) = STR$(DCCustRec(1).AcctBal)
    Form$(10, 0) = DCCustRec(1).SOSEC
    Form$(12, 0) = DCCustRec(1).DRVLIC
    Form$(14, 0) = Num2Date(DCCustRec(1).DATEOPED)
    Form$(16, 0) = DCCustRec(1).CashOnly
    Form$(11, 0) = DCCustRec(1).Resident
    Form$(13, 0) = DCCustRec(1).Owner
    Form$(15, 0) = DCCustRec(1).HPHONE
    Form$(17, 0) = DCCustRec(1).WPHONE
    CustomerGrabed = 1
    Action = 1
    COLOR 15
    LOCATE 5, 40: PRINT STRING$(32, 32)
    RETURN

   ELSE

    LibName$ = "DC"
    ScrnName$ = "DCBADCUS"
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
    LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
    PrintHelp help$

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
    LibName$ = "DC"
    ScrnName$ = "DCDELCUS"
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
    LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
    PrintHelp help$

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
 REDIM DCTransRec(1) AS DCTransRecType
 DCTransRecLen = LEN(DCTransRec(1))
 DCTransFile = FREEFILE
 OPEN "DCTrans.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCTransFile LEN = DCTransRecLen
 
  MaxLines = 58
  ReportFile$ = "DCHIST.PRN"
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle

  GOSUB PrintCustInqRptHeader

  TransRec& = DCCustRec(1).FirstTrans

   WHILE TransRec& > 0
   GET DCTransFile, TransRec&, DCTransRec(1)

   IF LineCnt >= MaxLines THEN
    PRINT #RptHandle, FF$
    GOSUB PrintCustInqRptHeader
   END IF
   IF DCTransRec(1).TransType >= 1 AND DCTransRec(1).TransType <= 9 THEN
   PRINT #RptHandle, Num2Date$(DCTransRec(1).TransDate);
    PRINT #RptHandle, TAB(12); RTRIM$(DCTransRec(1).TransDesc);
     PRINT #RptHandle, TAB(50); "";
      IF DCTransRec(1).TransType = 1 THEN
       PRINT #RptHandle, "Charge";
       RunBalance# = RunBalance# + DCTransRec(1).TransAmount
      END IF
      IF DCTransRec(1).TransType = 2 THEN
       PRINT #RptHandle, "Payment";
       RunBalance# = RunBalance# - DCTransRec(1).TransAmount
      END IF
      IF DCTransRec(1).TransType = 2 THEN
      PRINT #RptHandle, TAB(60); USING "####,#.##"; DCTransRec(1).TransAmount * -1;
      ELSE
      PRINT #RptHandle, TAB(60); USING "####,#.##"; DCTransRec(1).TransAmount;
      END IF
     PRINT #RptHandle, TAB(70); USING "####,#.##"; RunBalance#
     IF DCTransRec(1).TransType = 1 THEN
      PRINT #RptHandle, "Veh Desc: "; RTRIM$(DCTransRec(1).MakeModel);
      PRINT #RptHandle, "  Decal # "; RTRIM$(DCTransRec(1).Sticker);
      PRINT #RptHandle, "  Exp Date: "; Num2Date$(DCTransRec(1).ExpireDate)
      LineCnt = LineCnt + 1
     END IF

    PRINT #RptHandle, STRING$(79, "-")
    LineCnt = LineCnt + 2
   END IF
    TransRec& = DCTransRec(1).NextTrans

   WEND
   PRINT #RptHandle, CHR$(18); ' oki 320 10 cpi
   PRINT #RptHandle, CHR$(12);
   CLOSE                 'Close all open files now
                 
   PrintRptFile "Customer Account History", ReportFile$, LPTPort%, RetCode%, EntryPoint
   KILL ReportFile$
   RETURN

PrintCustInqRptHeader:
    Page = Page + 1
    PRINT #RptHandle, TAB(18); "Va. Vehicle Decals : Detailed Customer History"
    PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
    PRINT #RptHandle, "Customer Name: "; DCCustRec(1).BillName
    PRINT #RptHandle, "  Date"; TAB(10); "Description "; TAB(50); "Type"; TAB(60); "  Amount"; TAB(73); "Balance"
    PRINT #RptHandle, STRING$(80, "=")
    LineCnt = 3
    RETURN


RETURN


END SUB

SUB CustomerListing


  SHARED Choice$()


  ReportFile$ = "DCDetCus.PRN"  'Report File Name
  CommaFmt$ = "########,.##"    'format takes 13 chars
  TotalFmt$ = "#########,.##"   'format takes 14 chars
  SumLine$ = STRING$(13, "-")   'column summary line
  DivLine$ = STRING$(77, "-")   'dashed line
  DivLine2$ = STRING$(77, "=")  'Double Line
  FF$ = CHR$(12)
  MaxLines = 58
  LineCnt = 0
  CustCnt = 0
  LPTPort% = 1

  GOSUB SelectDetailCustomerOutput: IF Canceled$ = "Y" THEN EXIT SUB


  

  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle


 'PRINT #RptHandle, CHR$(27); CHR$(58); ' oki 320 12 cpi

  GOSUB PrintDetailCustomerRptHeader

 ' Print Main Body
  OpenDCCustFile NumOfDcRecs, DCFile
  REDIM DCIdxRec(1) AS DCCustIdxRecType
  IdxFile = FREEFILE
  OPEN "DCCust.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS #IdxFile LEN = LEN(DCIdxRec(1))

  'Open Vehicle File
  REDIM DCVRec(1) AS DCVehType
  DCVehRecLen = LEN(DCVRec(1))
  DCVFile = FREEFILE
  OPEN "DCVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCVFile LEN = DCVehRecLen
  NumOfVrecs = LOF(DCVFile) \ DCVehRecLen



  IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN NumOfDcRecs = LOF(IdxFile) \ LEN(DCIdxRec(1))

  FOR Cnt! = 1 TO NumOfDcRecs
  IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN
   GET IdxFile, Cnt!, DCIdxRec(1)
   GET DCFile, DCIdxRec(1).IDXRECORD, DCCustRec(1)
   CustomerNumber = DCIdxRec(1).IDXRECORD
   ELSE
   GET DCFile, Cnt!, DCCustRec(1)
   CustomerNumber = Cnt!
  END IF
  IF DCCustRec(1).Deleted <> "Y" THEN
  IF LineCnt >= MaxLines THEN
    PRINT #RptHandle, FF$
    GOSUB PrintDetailCustomerRptHeader
   END IF

   IF RTRIM$(Catagory$) <> "ALL" THEN
     VehRecord! = DCCustRec(1).FirstCar
     WHILE VehRecord! <> 0
     GET DCVFile, VehRecord!, DCVRec(1)
     IF RTRIM$(DCVRec(1).DecalCat) = RTRIM$(Catagory$) THEN GOTO ProcessThisOne
     VehRecord! = DCVRec(1).NextRec
     WEND
     GOTO SkipRecord
   END IF

ProcessThisOne:
   PRINT #RptHandle, "Cust #"; CustomerNumber
   PRINT #RptHandle, DCCustRec(1).BillName
   PRINT #RptHandle, RTRIM$(DCCustRec(1).ADDRESS1); TAB(50); " H Phone: "; DCCustRec(1).HPHONE
   PRINT #RptHandle, RTRIM$(DCCustRec(1).ADDRESS2); TAB(50); " W Phone: "; DCCustRec(1).WPHONE
   PRINT #RptHandle, RTRIM$(DCCustRec(1).CITY); ", "; DCCustRec(1).STATE; " "; DCCustRec(1).ZIPCODE
   PRINT #RptHandle, STRING$(80, "-")
    CustCnt = CustCnt + 1

    LineCnt = LineCnt + 6
   END IF

SkipRecord:
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
    PRINT #RptHandle, TAB(22); "Va. Decal System : Detailed Customer Listing"
    PRINT #RptHandle, "Catagory: "; Catagory$
    PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
    PRINT #RptHandle, STRING$(80, "=")
    LineCnt = 4
    RETURN

PrintDetailCustomerRptEnding:
     PRINT #RptHandle, "Total Customers Printed: "; USING "#####"; CustCnt
     PRINT #RptHandle,
     PRINT #RptHandle, FF$
     RETURN







SelectDetailCustomerOutput:
  LibName$ = "DC"
  ScrnName$ = "DCCUSREP"

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

  
  REDIM Choice$(100, 3)      ' assume maximum of 99 codes

  REDIM DCCatCodeRec(1) AS DCCatCodeRecType
  DCCatCodeRecLen = LEN(DCCatCodeRec(1))
  DCCatFile = FREEFILE
  OPEN "DCCODE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCCatFile LEN = DCCatCodeRecLen
  NumOFDCCatRecs = LOF(DCCatFile) \ DCCatCodeRecLen

  IF NumOFDCCatRecs = 0 THEN
   CLOSE DCCatFile
   ShowNoCode
   EXIT SUB
  END IF

  NumOFDCCatRecs = NumOFDCCatRecs + 1


   

  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Name"
  Choice$(2, 0) = "Number"


   Choice$(0, 1) = "2"
  FOR Cnt! = 1 TO (NumOFDCCatRecs - 1):
  GET DCCatFile, Cnt!, DCCatCodeRec(1)

   Choice$(Cnt!, 1) = DCCatCodeRec(1).CATCODE
  NEXT Cnt!
   Choice$(NumOFDCCatRecs, 1) = "ALL"
  CLOSE DCCatFile

  
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


  ReportFile$ = "DCEXPLIC.PRN"  'Report File Name
  CommaFmt$ = "########,.##"    'format takes 13 chars
  TotalFmt$ = "#########,.##"   'format takes 14 chars
  SumLine$ = STRING$(13, "-")   'column summary line
  DivLine$ = STRING$(77, "-")   'dashed line
  DivLine2$ = STRING$(77, "=")  'Double Line
  FF$ = CHR$(12)
  DevSpec$ = "S"
  MaxLines = 53
  LineCnt = 0
  LPTPort% = 1

  GOSUB SelectExpiredLicenseListOutput: IF Canceled$ = "Y" THEN EXIT SUB



  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle

  'Open Vehicle File
  REDIM DCVRec(1) AS DCVehType
  DCVehRecLen = LEN(DCVRec(1))
  DCVFile = FREEFILE
  OPEN "DCVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCVFile LEN = DCVehRecLen
  NumOfVrecs = LOF(DCVFile) \ DCVehRecLen


 'PRINT #RptHandle, CHR$(27); CHR$(58); ' oki 320 12 cpi

  GOSUB PrintExpiredLicenseListRptHeader

 ' Print Main Body
  OpenDCCustFile NumOfDcRecs, DCFile
  REDIM DCIdxRec(1) AS DCCustIdxRecType
  IdxFile = FREEFILE
  OPEN "DCCust.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS #IdxFile LEN = LEN(DCIdxRec(1))

  IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN NumOfDcRecs = LOF(IdxFile) \ LEN(DCIdxRec(1))

  FOR Cnt! = 1 TO NumOfDcRecs
  IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN
   GET IdxFile, Cnt!, DCIdxRec(1)
   GET DCFile, DCIdxRec(1).IDXRECORD, DCCustRec(1)
   CustomerNumber = DCIdxRec(1).IDXRECORD
   ELSE
   GET DCFile, Cnt!, DCCustRec(1)
   CustomerNumber = Cnt!
  END IF

   IF LineCnt >= MaxLines THEN
    PRINT #RptHandle, FF$
    GOSUB PrintExpiredLicenseListRptHeader
   END IF

  IF DCCustRec(1).Deleted <> "Y" THEN
   
     VehRecord! = DCCustRec(1).FirstCar
     WHILE VehRecord! <> 0
     GET DCVFile, VehRecord!, DCVRec(1)
     IF RTRIM$(DCVRec(1).DecalCat) = RTRIM$(Catagory$) OR RTRIM$(Catagory$) = "ALL" THEN
      IF DCVRec(1).ExpireDate < ExpDate THEN
        PRINT #RptHandle, CustomerNumber; TAB(10); DCCustRec(1).BillName;
        PRINT #RptHandle, TAB(50); LTRIM$(DCVRec(1).Sticker);
        PRINT #RptHandle, TAB(70); Num2Date$(DCVRec(1).ExpireDate)
        LineCnt = LineCnt + 1
        CustCnt = CustCnt + 1
      END IF
     END IF
     VehRecord! = DCVRec(1).NextRec
     WEND


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
   Header$ = "Expired Decal Listing"
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint

  KILL ReportFile$

  EXIT SUB


PrintExpiredLicenseListRptHeader:
    Page = Page + 1
    PRINT #RptHandle, TAB(29); "Expired Decal Listing"
    PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
    PRINT #RptHandle, "Catagory "; Catagory$; TAB(50); "Expiration Date: "; Num2Date$(ExpDate)
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
  LibName$ = "DC"
  ScrnName$ = "DCEXPLIC"

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


  REDIM Choice$(100, 3)      ' assume maximum of 99 codes

  REDIM DCCatCodeRec(1) AS DCCatCodeRecType
  DCCatCodeRecLen = LEN(DCCatCodeRec(1))
  DCCatFile = FREEFILE
  OPEN "DCCODE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCCatFile LEN = DCCatCodeRecLen
  NumOFDCCatRecs = LOF(DCCatFile) \ DCCatCodeRecLen

  IF NumOFDCCatRecs = 0 THEN
   CLOSE DCCatFile
   ShowNoCode
   EXIT SUB
  END IF

  NumOFDCCatRecs = NumOFDCCatRecs + 1




  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Name"
  Choice$(2, 0) = "Number"


   Choice$(0, 1) = "2"
  FOR Cnt! = 1 TO (NumOFDCCatRecs - 1):
  GET DCCatFile, Cnt!, DCCatCodeRec(1)

   Choice$(Cnt!, 1) = DCCatCodeRec(1).CATCODE
  NEXT Cnt!
   Choice$(NumOFDCCatRecs, 1) = "ALL"
  CLOSE DCCatFile


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
        LOCATE 15, 47: COLOR 12: PRINT "NO DATE ENTERED"
        SLEEP 2
        LOCATE 15, 47: PRINT STRING$(15, 32);
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


  ReportFile$ = "DCLICLST.PRN"  'Report File Name
  CommaFmt$ = "########,.##"    'format takes 13 chars
  TotalFmt$ = "#########,.##"   'format takes 14 chars
  SumLine$ = STRING$(13, "-")   'column summary line
  DivLine$ = STRING$(77, "-")   'dashed line
  DivLine2$ = STRING$(77, "=")  'Double Line
  FF$ = CHR$(12)
  MaxLines = 58
  LineCnt = 0
  CustCnt = 0
  LPTPort% = 1

  GOSUB SelectLicenseListOutput: IF Canceled$ = "Y" THEN EXIT SUB

  'Open Vehicle File
  REDIM DCVRec(1) AS DCVehType
  DCVehRecLen = LEN(DCVRec(1))
  DCVFile = FREEFILE
  OPEN "DCVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCVFile LEN = DCVehRecLen
  NumOfVrecs = LOF(DCVFile) \ DCVehRecLen


  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle


 'PRINT #RptHandle, CHR$(27); CHR$(58); ' oki 320 12 cpi

  GOSUB PrintLicenseListRptHeader

 ' Print Main Body
  OpenDCCustFile NumOfDcRecs, DCFile
  REDIM DCIdxRec(1) AS DCCustIdxRecType
  IdxFile = FREEFILE
  OPEN "DCCust.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS #IdxFile LEN = LEN(DCIdxRec(1))

  IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN NumOfDcRecs = LOF(IdxFile) \ LEN(DCIdxRec(1))

  FOR Cnt! = 1 TO NumOfDcRecs
  IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN
   GET IdxFile, Cnt!, DCIdxRec(1)
   GET DCFile, DCIdxRec(1).IDXRECORD, DCCustRec(1)
   CustomerNumber = DCIdxRec(1).IDXRECORD
   ELSE
   GET DCFile, Cnt!, DCCustRec(1)
   CustomerNumber = Cnt!
  END IF

   IF LineCnt >= MaxLines THEN
    PRINT #RptHandle, FF$
    GOSUB PrintLicenseListRptHeader
   END IF

  IF DCCustRec(1).Deleted <> "Y" THEN
  
  
   
     VehRecord! = DCCustRec(1).FirstCar
     WHILE VehRecord! <> 0
     GET DCVFile, VehRecord!, DCVRec(1)
     IF RTRIM$(DCVRec(1).DecalCat) = RTRIM$(Catagory$) OR RTRIM$(Catagory$) = "ALL" THEN
      IF DCVRec(1).Active <> "N" THEN
      PRINT #RptHandle, CustomerNumber; TAB(10); DCCustRec(1).BillName;
      PRINT #RptHandle, TAB(50); LTRIM$(DCVRec(1).Sticker);
      PRINT #RptHandle, TAB(70); Num2Date$(DCVRec(1).ExpireDate)
       LineCnt = LineCnt + 1
       CustCnt = CustCnt + 1
      END IF
     END IF
     VehRecord! = DCVRec(1).NextRec
     WEND
     
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
    PRINT #RptHandle, TAB(20); "Va. Vehicle Decal Current Sticker Listing"
    PRINT #RptHandle, "Catagory: "; Catagory$
    PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
    PRINT #RptHandle, "Cust #"; TAB(10); "Customer Name"; TAB(50); "Sticker #"; TAB(65); "Valid Thru"
    PRINT #RptHandle, STRING$(80, "=")
    LineCnt = 3
    RETURN

PrintLicenseListRptEnding:
     PRINT #RptHandle, STRING$(80, "=")
     PRINT #RptHandle, "Stickers Counted: "; USING "#####,#"; CustCnt
     PRINT #RptHandle, FF$
     RETURN







SelectLicenseListOutput:
  LibName$ = "DC"
  ScrnName$ = "DCLICREP"

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


  REDIM Choice$(100, 3)      ' assume maximum of 99 codes

  REDIM DCCatCodeRec(1) AS DCCatCodeRecType
  DCCatCodeRecLen = LEN(DCCatCodeRec(1))
  DCCatFile = FREEFILE
  OPEN "DCCODE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCCatFile LEN = DCCatCodeRecLen
  NumOFDCCatRecs = LOF(DCCatFile) \ DCCatCodeRecLen

  IF NumOFDCCatRecs = 0 THEN
   CLOSE DCCatFile
   ShowNoCode
   EXIT SUB
  END IF

  NumOFDCCatRecs = NumOFDCCatRecs + 1




  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Name"
  Choice$(2, 0) = "Number"


   Choice$(0, 1) = "2"
  FOR Cnt! = 1 TO (NumOFDCCatRecs - 1):
  GET DCCatFile, Cnt!, DCCatCodeRec(1)

   Choice$(Cnt!, 1) = DCCatCodeRec(1).CATCODE
  NEXT Cnt!
   Choice$(NumOFDCCatRecs, 1) = "ALL"
  CLOSE DCCatFile


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

SUB NonOwnerReport
END SUB

SUB OpenDCCustFile (NumOfDcRecs, DCFile)
  CLOSE DCFile

  DCCustRecLen = LEN(DCCustRec(1))
  DCFile = FREEFILE
  OPEN "DCCust.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCFile LEN = DCCustRecLen
  NumOfDcRecs = LOF(DCFile) \ DCCustRecLen
  'FOR x = 1 TO NumOfDcRecs
  'GET DCFile, x, DCCust(1)
  'PRINT DCCust(1).Custnumb; TAB(15); DCCust(1).FirstTrans
  'SLEEP 1
  'NEXT x
  'STOP
   END SUB

SUB OpenDCCustIdxFile (NumOfDCIdxRecs, DCIdxFile)
  CLOSE DCIdxFile
  DCCustIdxRecLen = LEN(DCCustIdxRec(1))
  DCIdxFile = FREEFILE
  OPEN "DCCust.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS DCIdxFile LEN = DCCustIdxRecLen
  NumOfDCIdxRecs = LOF(DCIdxFile) \ DCCustIdxRecLen
END SUB

SUB OwnerReport
END SUB

SUB ResidentReport
END SUB

SUB ShowNoCode
  LibName$ = "DC"
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
  LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
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
  ReportFile$ = "DCTrans.PRN"  'Report File Name
  CommaFmt$ = "########,.##"    'format takes 13 chars
  TotalFmt$ = "#########,.##"   'format takes 14 chars
  SumLine$ = STRING$(13, "-")   'column summary line
  DivLine$ = STRING$(77, "-")   'dashed line
  DivLine2$ = STRING$(77, "=")  'Double Line
  FF$ = CHR$(12)
  MaxLines = 53
  LPTPort% = 1
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
  REDIM DCTransRec(1) AS DCTransRecType
  REDIM TotalAmt#(99)

  GOSUB SelectOutput2
  IF Canceled$ = "Y" THEN EXIT SUB

  GOSUB GetReportInformation2

  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle


  'PRINT #RptHandle, CHR$(27); CHR$(58); ' oki 320 12 cpi

  GOSUB PrintRptHeader2

  FOR Cnt = 1 TO Count
   GET DCTRFile, array(Cnt).RecNum, DCTransRec(1)
   
   IF LineCnt >= MaxLines THEN
    PRINT #RptHandle, FF$
    GOSUB PrintRptHeader2
   END IF


   'Get Customer
   OpenDCCustFile NumOfDcRecs, DCFile
   GET DCFile, VAL(DCTransRec(1).CustomerNumber), DCCustRec(1)
       
   PRINT #RptHandle, Num2Date$(DCTransRec(1).TransDate);
    PRINT #RptHandle, TAB(13); LEFT$(DCCustRec(1).BillName, 25);
     PRINT #RptHandle, TAB(40); "";
      IF DCTransRec(1).TransType = 1 THEN PRINT #RptHandle, "Charge";
      IF DCTransRec(1).TransType = 2 THEN PRINT #RptHandle, "Payment";
       PRINT #RptHandle, TAB(50); LEFT$(DCTransRec(1).TransDesc, 18);
       PRINT #RptHandle, TAB(69); USING "$$#####,#.##"; DCTransRec(1).TransAmount

    CLOSE DCFile
    TotalTrans = TotalTrans + 1
    LineCnt = LineCnt + 1
    TotalAmt# = TotalAmt# + DCTransRec(1).TransAmount
    
    TotalAmt#(DCTransRec(1).TransType) = TotalAmt#(DCTransRec(1).TransType) + DCTransRec(1).TransAmount
   
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
    PRINT #RptHandle, TAB(20); "Va. Vehicle Decals : Transactions Journal "
    PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
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
      PRINT #RptHandle, "     Total Amount: "; USING "$$######,#.##"; TotalAmt#(Cnt!)
     END IF
     NEXT Cnt!



     PRINT #RptHandle, FF$
     RETURN

GetReportInformation2:
  
  DCTRFile = FREEFILE
  OPEN "DCTrans.Dat" FOR RANDOM ACCESS READ WRITE SHARED AS #DCTRFile LEN = LEN(DCTransRec(1))
  TrNumRecs = LOF(DCTRFile) / LEN(DCTransRec(1))

  FOR Cnt! = 1 TO TrNumRecs
      GET DCTRFile, Cnt!, DCTransRec(1)
      IF LEFT$(Type$, 1) = "A" THEN
      Count = Count + 1
      array(Count).who = LTRIM$(STR$(DCTransRec(1).TransDate)) + DCTransRec(1).CustomerNumber + STRING$(4, " ")
      array(Count).RecNum = Cnt!
      END IF
      IF LEFT$(Type$, 1) = "B" THEN
      IF DCTransRec(1).TransType = 9 THEN
      Count = Count + 1
      array(Count).who = LTRIM$(STR$(DCTransRec(1).TransDate)) + DCTransRec(1).CustomerNumber + STRING$(4, " ")
      array(Count).RecNum = Cnt!
      END IF
      END IF
      IF LEFT$(Type$, 1) = "C" THEN
      IF DCTransRec(1).TransType = 1 THEN
      Count = Count + 1
      array(Count).who = LTRIM$(STR$(DCTransRec(1).TransDate)) + DCTransRec(1).CustomerNumber + STRING$(4, " ")
      array(Count).RecNum = Cnt!
      END IF
      END IF
      IF LEFT$(Type$, 1) = "P" THEN
      IF DCTransRec(1).TransType = 2 THEN
      Count = Count + 1
      array(Count).who = LTRIM$(STR$(DCTransRec(1).TransDate)) + DCTransRec(1).CustomerNumber + STRING$(4, " ")
      array(Count).RecNum = Cnt!
      END IF
      END IF
NEXT Cnt!

  

SortT array(Start), Count, Dir, SSize, MOff, MSize
RETURN






SelectOutput2:
  LibName$ = "DC"
  ScrnName$ = "DCTRREP"

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

  REDIM Choice$(3, 1)
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "All"
  Choice$(2, 0) = "Charges"
  Choice$(3, 0) = "Payments"
  Choice$(0, 1) = "2"
  Choice$(1, 1) = "SCREEN"
  Choice$(2, 1) = "PRINTER"




  Action = 1
  ShowCursor
  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%
  'printhelp help$
  Action = 1

  DO


    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode
     CASE F10Key
       DevSpec$ = LEFT$(Form$(2, 0), 1)
       Type$ = Form$(1, 0)
       BegDate = Date2Num%(Form$(3, 0))
       EndDate = Date2Num%(Form$(4, 0))
       RETURN
     CASE ESCKey
      Canceled$ = "Y"
      RETURN
    END SELECT
 LOOP
  RETURN


END SUB

