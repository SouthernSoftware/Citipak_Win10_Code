DEFINT A-Z
DECLARE SUB LookUp (RecNo&, Text$, ChkBalFlag%, CLSFlag%, SSNFlag%)
DECLARE SUB DisplayDCScrn (ScrnName$)
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
DECLARE SUB CustomerInquiry (RecNo&)
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB ShowCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB HideCursor ()
DECLARE SUB QPrint (x$, Colr%, page%)
DECLARE SUB QPrintRC (t$, r%, c%, clr%)
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE SUB ClearBack ()
DECLARE FUNCTION GotDCCusts ()
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE FUNCTION Round# (N#)

  '$INCLUDE: 'DefCnf.BI'
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
  REDIM Mchoice$(1 TO 11)
  
  Mchoice$(1) = "Customer Inquiry"
  Mchoice$(2) = "Customer Balance Listing"
  Mchoice$(3) = "Transaction Journal"
  Mchoice$(4) = "Detailed Catagory Listing "
  Mchoice$(5) = "Detailed Customer Listing"
  Mchoice$(6) = "Decal Listing"
  Mchoice$(7) = "Expired Decal Listing"
  Mchoice$(8) = "Resident Report"
  Mchoice$(9) = "Owner Report"
  Mchoice$(10) = "Non-Owner Report"
  Mchoice$(11) = "Exit to OS"
  
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

  Row = ((25 - (UBOUND(Mchoice$))) \ 2) + 1

  Col = ((80 - MaxLen) \ 2) - 1
  Help$ = "Main Reports Menu"

  
  DO

    LOCATE Row, Col, 0

    DisplayDCScrn "MENUBAK"

    TitleBox 3, Col, MaxLen + 3, "Decal Report Menu", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    '--Set upper left corner of menu, turn off the cursor
    
     PrintTitle user$
     PrintHelp Help$
    
    ShowCursor
    
    VertMenu Mchoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      IF GotDCCusts THEN
        ClearBack
        LookUp RecNo&, "Customer", False, True, False
        IF RecNo& > 0 THEN
          CustomerInquiry RecNo&
        END IF
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
      ResidentReport
    CASE 9
      OwnerReport
    CASE 10
      NonOwnerReport
    CASE 11
      HideCursor
      CLS
      END
    END SELECT
  LOOP

'END
  RUN "dcmenu"

SUB CatListing
  SHARED Choice$()
  
  
  ReportFile$ = "DCCODLST.PRN"  'Report File Name
  FF$ = CHR$(12)
  MaxLines = 53
  LineCnt = 0

  size = 2500

  Start = 1     'start at array element 1
  Dir = 0       'sort direction - use anything else for descending
  SSize = 16    'total size of each TYPE element
  MOff = 0      'offset into the TYPE for the key element
  MSize = 16    'size of the key element - coded as follows:
  
  REDIM Array(1 TO size) AS Struct
  
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
  
  
  PRINT #RptHandle, CHR$(27); CHR$(58);         ' oki 320 12 cpi
  
  GOSUB PrintRptHeader
  
  FOR Cnt = 1 TO Count
    GET TrHandle, Array(Cnt).RecNum, DCCatCodeRec(1)
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
  PRINT #RptHandle, TAB(20); "Va. Decal System : Decal Catagory Code Listing "
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; page
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
    Array(Cnt!).who = LEFT$(DCCatCodeRec(1).CATCODE, 3) + STRING$(11, " ")
    Array(Cnt!).RecNum = Cnt!
  NEXT Cnt!
  
  Count = TrNumRecs
  
  SortT Array(Start), Count, Dir, SSize, MOff, MSize
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

SUB CustomerBalance
  
  SHARED Choice$()
  
  ReportFile$ = "DCCusBal.PRN"  'Report File Name
  FF$ = CHR$(12)
  MaxLines = 58
  LineCnt = 0
  CustCnt = 0
  LPTPort% = 1

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
  
  REDIM Choice$(3, 2)         ' assume maximum of 99 codes

'  NumOFDCCatRecs = NumOFDCCatRecs + 1

  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Name"
  Choice$(2, 0) = "Number"

  Choice$(0, 1) = "2"
  Choice$(1, 1) = "SCREEN"
  Choice$(2, 1) = "PRINTER"

  Action = 1
  ShowCursor
  DisplayDCScrn ScrnName$

  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      Order$ = Form$(1, 0)
      DevSpec$ = LEFT$(Form$(2, 0), 1)
      EXIT DO
    CASE EscKey
      AbortFlag = True
      EXIT DO
    END SELECT
  LOOP

  IF AbortFlag THEN
    GOTO DoneHere1
  END IF
  
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
  
  FOR Cnt& = 1 TO NumOfDcRecs
    IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN
      GET IdxFile, Cnt&, DCIdxRec(1)
      GET DCFile, DCIdxRec(1).IdxRecord, DCCustRec(1)
      CustomerNumber = DCIdxRec(1).IdxRecord
    ELSE
      GET DCFile, Cnt&, DCCustRec(1)
      CustomerNumber = Cnt&
    END IF
    IF DCCustRec(1).Deleted <> "Y" THEN
      IF DCCustRec(1).AcctBal <> 0 THEN
        IF LineCnt >= MaxLines THEN
          PRINT #RptHandle, FF$
          GOSUB PrintCustBalRptHeader
        END IF
        PRINT #RptHandle, USING "#####"; CustomerNumber;
        PRINT #RptHandle, TAB(10); DCCustRec(1).BILLNAME;
        PRINT #RptHandle, TAB(50); USING "$$####,#.##"; DCCustRec(1).AcctBal
        CustCnt = CustCnt + 1
        TotalBal# = Round#(TotalBal# + DCCustRec(1).AcctBal)
        LineCnt = LineCnt + 1
      END IF
    END IF
  NEXT Cnt&
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
  
DoneHere1:
  EXIT SUB
  
  
PrintCustBalRptHeader:
  page = page + 1
  PRINT #RptHandle, TAB(21); "Va. Decal System : Detailed Customer Listing"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; page
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
END SUB

SUB CustomerInquiry (RecNo&)

  LPTPort% = 1

  LibName$ = "DC"
  ScrnName$ = "DCCUSTIQ"
  Help$ = "Customer Inquiry"
  'LOCATE 1, 1, 0
  
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

  FOR Cnt = 1 TO 17
    Fld(Cnt).Protected = True
  NEXT

  Action = 1
  FirstTime = True

  ShowCursor
  DisplayDCScrn ScrnName$
  PrintHelp Help$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF FirstTime THEN
      FirstTime = False
      GOSUB LoadCustInfo
      Action = 1
      Frm(1).FldNo = 19
    END IF
    
    SELECT CASE Frm(1).KeyCode
    CASE F4KEY
      GOSUB PrintHistory
      DisplayDCScrn ScrnName$
      Action = 1
    CASE F5KEY
      GOSUB PRINTCars
      DisplayDCScrn ScrnName$
      Action = 1
    CASE EscKey
      CLOSE
      EXIT DO
    END SELECT
  LOOP

EXIT SUB

LoadCustInfo:
  OpenDCCustFile NumOfDcRecs, DCFile
  AccountRecord = RecNo&
  GET DCFile, AccountRecord, DCCustRec(1)
  CLOSE

  LSET Form$(1, 0) = DCCustRec(1).CUSTNUMB
  LSET Form$(2, 0) = DCCustRec(1).SORTNAME
  LSET Form$(3, 0) = DCCustRec(1).BILLNAME
  LSET Form$(4, 0) = DCCustRec(1).ADDRESS1
  LSET Form$(5, 0) = DCCustRec(1).ADDRESS2
  LSET Form$(6, 0) = DCCustRec(1).CITY
  LSET Form$(7, 0) = DCCustRec(1).STATE
  LSET Form$(8, 0) = DCCustRec(1).ZIPCODE
  LSET Form$(9, 0) = STR$(DCCustRec(1).AcctBal)
  LSET Form$(10, 0) = DCCustRec(1).SoSec
  LSET Form$(12, 0) = DCCustRec(1).DRVLIC
  LSET Form$(14, 0) = Num2Date(DCCustRec(1).DATEOPED)
  LSET Form$(16, 0) = DCCustRec(1).CashOnly
  LSET Form$(11, 0) = DCCustRec(1).Resident
  LSET Form$(13, 0) = DCCustRec(1).Owner
  LSET Form$(15, 0) = DCCustRec(1).HPHONE
  LSET Form$(17, 0) = DCCustRec(1).WPHONE
  CustomerGrabed = 1
  Action = 1
RETURN
    
PrintHistory:
  RunBalance# = 0
  page = 0
  REDIM DCTransRec(1) AS DCTransRecType
  DCTransRecLen = LEN(DCTransRec(1))
  
  MaxLines = 58
  ReportFile$ = "DCHIST.PRN"

  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  DCTransFile = FREEFILE
  OPEN "DCTrans.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCTransFile LEN = DCTransRecLen
  
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
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi
  PRINT #RptHandle, CHR$(12);
  CLOSE         'Close all open files now
  
  PrintRptFile "Customer Account History", ReportFile$, LPTPort%, RetCode%, EntryPoint
  KILL ReportFile$
RETURN
  
PrintCustInqRptHeader:
  page = page + 1
  PRINT #RptHandle, TAB(18); "Va. Vehicle Decals : Detailed Customer History"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; page
  PRINT #RptHandle, "Customer Name: "; DCCustRec(1).BILLNAME
  PRINT #RptHandle, "  Date"; TAB(10); "Description "; TAB(50); "Type"; TAB(60); "  Amount"; TAB(73); "Balance"
  PRINT #RptHandle, STRING$(80, "=")
  LineCnt = 3
RETURN
  
  
PRINTCars:
  page = 0
  
  'Open Vehicle File
  REDIM DCVRec(1) AS DCVehType
  DCVehRecLen = LEN(DCVRec(1))
  DCVFile = FREEFILE
  OPEN "DCVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCVFile LEN = DCVehRecLen
  NumOfVrecs = LOF(DCVFile) \ DCVehRecLen
  
  MaxLines = 58
  ReportFile$ = "DCCARS.PRN"
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  GOSUB PrintCarsRptHeader
  
  IF LineCnt >= MaxLines THEN
    PRINT #RptHandle, FF$
    GOSUB PrintCustInqRptHeader
  END IF

  VehRec& = DCCustRec(1).FirstCar
  WHILE VehRec& <> 0
    GET DCVFile, VehRec&, DCVRec(1)
    
    IF LineCnt >= MaxLines THEN
      PRINT #RptHandle, FF$
      GOSUB PrintCustInqRptHeader
    END IF
    
    IF DCVRec(1).Active = "Y" THEN
      Desc$ = RTRIM$(DCVRec(1).MakeModel) + " " + (RTRIM$(DCVRec(1).Desc))
      Desc$ = LEFT$(Desc$, 53)
      PRINT #RptHandle, Desc$;
      PRINT #RptHandle, TAB(55); LTRIM$(LEFT$(DCVRec(1).Sticker, 13));
      PRINT #RptHandle, TAB(70); Num2Date$(DCVRec(1).ExpireDate)
      LineCnt = LineCnt + 1
    END IF
    
    VehRec& = DCVRec(1).NextRec

  WEND
  
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi
  PRINT #RptHandle, CHR$(12);
  CLOSE         'Close all open files now
  
  PrintRptFile "Customer Vehicles Listing", ReportFile$, LPTPort%, RetCode%, EntryPoint
  KILL ReportFile$
  
RETURN
  
PrintCarsRptHeader:
  page = page + 1
  PRINT #RptHandle, TAB(18); "Va. Vehicle Decals : Detailed Vehicle Listing"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; page
  PRINT #RptHandle, "Customer Name: "; DCCustRec(1).BILLNAME
  PRINT #RptHandle, "Vehicle Description"; TAB(55); "Sticker #"; TAB(70); " Expires"
  PRINT #RptHandle, STRING$(80, "=")
  LineCnt = 5
RETURN
  
END SUB

SUB CustomerListing
  
  
  SHARED Choice$()
  
  
  ReportFile$ = "DCDetCus.PRN"  'Report File Name
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
      GET DCFile, DCIdxRec(1).IdxRecord, DCCustRec(1)
      CustomerNumber = DCIdxRec(1).IdxRecord
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
      PRINT #RptHandle, DCCustRec(1).BILLNAME
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
  page = page + 1
  PRINT #RptHandle, TAB(22); "Va. Decal System : Detailed Customer Listing"
  PRINT #RptHandle, "Catagory: "; Catagory$
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; page
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
  
  
  REDIM Choice$(100, 3)         ' assume maximum of 99 codes
  
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
    CASE EscKey
      Canceled$ = "Y"
      RETURN
    END SELECT
  LOOP
  
END SUB

SUB ExpiredLicenseListing
  SHARED Choice$()
  
  
  ReportFile$ = "DCEXPLIC.PRN"  'Report File Name
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
      GET DCFile, DCIdxRec(1).IdxRecord, DCCustRec(1)
      CustomerNumber = DCIdxRec(1).IdxRecord
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
            PRINT #RptHandle, CustomerNumber; TAB(10); DCCustRec(1).BILLNAME;
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
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi
  CLOSE         'Close all open files now
  
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
  page = page + 1
  PRINT #RptHandle, TAB(29); "Expired Decal Listing"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; page
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
  
  
  REDIM Choice$(100, 3)         ' assume maximum of 99 codes
  
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
    CASE EscKey
      Canceled$ = "Y"
      RETURN
    END SELECT
  LOOP
  
END SUB

SUB LicenseListing
  
  SHARED Choice$()
  
  ReportFile$ = "DCLICLST.PRN"  'Report File Name
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
      GET DCFile, DCIdxRec(1).IdxRecord, DCCustRec(1)
      CustomerNumber = DCIdxRec(1).IdxRecord
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
            PRINT #RptHandle, CustomerNumber; TAB(10); DCCustRec(1).BILLNAME;
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
  
  EXIT SUB
  
  
PrintLicenseListRptHeader:
  page = page + 1
  PRINT #RptHandle, TAB(20); "Va. Vehicle Decal Current Sticker Listing"
  PRINT #RptHandle, "Catagory: "; Catagory$
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; page
  PRINT #RptHandle, "Cust #"; TAB(10); "Customer Name"; TAB(50); "Sticker #"; TAB(70); "Valid Thru"
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
  
  
  REDIM Choice$(100, 3)         ' assume maximum of 99 codes
  
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
    CASE EscKey
      Canceled$ = "Y"
      RETURN
    END SELECT
  LOOP
  
END SUB

SUB NonOwnerReport
  
  SHARED Choice$()
  
  
  ReportFile$ = "DCNoOwn.PRN"   'Report File Name
  FF$ = CHR$(12)
  MaxLines = 58
  LineCnt = 0
  CustCnt = 0
  LPTPort% = 1
  
  GOSUB SelectNoOwnOutput: IF Canceled$ = "Y" THEN EXIT SUB
  
  
  
  
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  
  'PRINT #RptHandle, CHR$(27); CHR$(58); ' oki 320 12 cpi
  
  GOSUB PrintNoOwnRptHeader
  
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
  
  'Open Non-Owned Vehicle File
  REDIM DCOVRec(1) AS DCOldVehType
  DCOVRecLen = LEN(DCOVRec(1))
  DCOVFile = FREEFILE
  OPEN "DCOLDVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCOVFile LEN = DCOVRecLen
  NumOfOVRecs! = LOF(DCOVFile) \ DCOVRecLen
  
  
  IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN NumOfDcRecs = LOF(IdxFile) \ LEN(DCIdxRec(1))
  
  FOR Cnt! = 1 TO NumOfDcRecs
    IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN
      GET IdxFile, Cnt!, DCIdxRec(1)
      GET DCFile, DCIdxRec(1).IdxRecord, DCCustRec(1)
      CustomerNumber = DCIdxRec(1).IdxRecord
    ELSE
      GET DCFile, Cnt!, DCCustRec(1)
      CustomerNumber = Cnt!
    END IF
    
    
    
    IF DCCustRec(1).Deleted <> "Y" THEN
      IF LineCnt >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintNoOwnRptHeader
      END IF
      FOR Cnt1! = 1 TO NumOfOVRecs!
        GET DCOVFile, Cnt1!, DCOVRec(1)
        IF DCOVRec(1).CustRec = CustomerNumber THEN
          PRINT #RptHandle, USING "#####"; CustomerNumber;
          PRINT #RptHandle, TAB(10); DCCustRec(1).BILLNAME; TAB(50); DCCustRec(1).SoSec
          PRINT #RptHandle, DCOVRec(1).Make; TAB(40); "Year: "; DCOVRec(1).Year
          CustCnt = CustCnt + 1
          PRINT #RptHandle, STRING$(80, "-"): LineCnt = LineCnt + 3
        END IF
      NEXT Cnt1!
    END IF
  NEXT Cnt!
  GOSUB PrintNoOwnRptEnding
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi
  CLOSE         'Close all open files now
  
  IF DevSpec$ = "S" THEN
    EntryPoint = 2
  ELSE
    EntryPoint = 5
  END IF
  Header$ = "Va. Owner Report"
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  KILL ReportFile$
  
  EXIT SUB
  
  
PrintNoOwnRptHeader:
  page = page + 1
  PRINT #RptHandle, TAB(12); "Va. Decal System : Old Vehicles Owned as of January 1"
  PRINT #RptHandle, "Tax Year: "; ResYear$
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; page
  PRINT #RptHandle, "Acct #"; TAB(10); "Customer Name"; TAB(50); "Soc Sec #"
  PRINT #RptHandle, STRING$(80, "=")
  LineCnt = 5
  RETURN
  
PrintNoOwnRptEnding:
  PRINT #RptHandle, STRING$(80, "-")
  PRINT #RptHandle, "Total Customers Printed: "; USING "#####"; CustCnt
  PRINT #RptHandle,
  PRINT #RptHandle, FF$
  RETURN
  
SelectNoOwnOutput:
  LibName$ = "DC"
  ScrnName$ = "DCNOOWN"
  
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
  
  
  REDIM Choice$(100, 3)         ' assume maximum of 99 codes
  
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
      ResYear$ = Form$(2, 0)
      DevSpec$ = LEFT$(Form$(3, 0), 1)
      RETURN
    CASE EscKey
      Canceled$ = "Y"
      RETURN
    END SELECT
  LOOP
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
  
  SHARED Choice$()
  
  
  ReportFile$ = "DCOWNREP.PRN"  'Report File Name
  FF$ = CHR$(12)
  MaxLines = 58
  LineCnt = 0
  CustCnt = 0
  LPTPort% = 1
  
  GOSUB SelectOwnRepOutput: IF Canceled$ = "Y" THEN EXIT SUB
  
  
  
  
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  
  'PRINT #RptHandle, CHR$(27); CHR$(58); ' oki 320 12 cpi
  
  GOSUB PrintOwnRepRptHeader
  
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
      GET DCFile, DCIdxRec(1).IdxRecord, DCCustRec(1)
      CustomerNumber = DCIdxRec(1).IdxRecord
    ELSE
      GET DCFile, Cnt!, DCCustRec(1)
      CustomerNumber = Cnt!
    END IF
    IF DCCustRec(1).Deleted <> "Y" THEN
      IF LineCnt >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintOwnRepRptHeader
      END IF
      IF DCCustRec(1).Owner = "Y" THEN
        PRINT #RptHandle, USING "#####"; CustomerNumber;
        PRINT #RptHandle, TAB(10); DCCustRec(1).BILLNAME; TAB(50); DCCustRec(1).SoSec
        CustCnt = CustCnt + 1
        
        VehRecord! = DCCustRec(1).FirstCar
        IF VehRecord! > 0 THEN PRINT #RptHandle, "Vehicles Owned as of January 1"
        WHILE VehRecord! <> 0
          GET DCVFile, VehRecord!, DCVRec(1)
          PRINT #RptHandle, RTRIM$(DCVRec(1).MakeModel) + " " + RTRIM$(DCVRec(1).Desc); TAB(60); "Tag #"; DCVRec(1).StateTag
          LineCnt = LineCnt + 1
          VehRecord! = DCVRec(1).NextRec
        WEND
        PRINT #RptHandle, STRING$(80, "-"): LineCnt = LineCnt + 1
      END IF
    END IF
  NEXT Cnt!
  GOSUB PrintOwnRepRptEnding
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi
  CLOSE         'Close all open files now
  
  IF DevSpec$ = "S" THEN
    EntryPoint = 2
  ELSE
    EntryPoint = 5
  END IF
  Header$ = "Va. Owner Report"
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  KILL ReportFile$
  
  EXIT SUB
  
  
PrintOwnRepRptHeader:
  page = page + 1
  PRINT #RptHandle, TAB(22); "Va. Decal System : Customer Owner Listing"
  PRINT #RptHandle, "Tax Year: "; ResYear$
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; page
  PRINT #RptHandle, "Acct #"; TAB(10); "Customer Name"; TAB(50); "Soc Sec #"
  PRINT #RptHandle, STRING$(80, "=")
  LineCnt = 5
  RETURN
  
PrintOwnRepRptEnding:
  PRINT #RptHandle, STRING$(80, "-")
  PRINT #RptHandle, "Total Customers Printed: "; USING "#####"; CustCnt
  PRINT #RptHandle,
  PRINT #RptHandle, FF$
  RETURN
  
SelectOwnRepOutput:
  LibName$ = "DC"
  ScrnName$ = "DCOWNREP"
  
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
  
  
  REDIM Choice$(100, 3)         ' assume maximum of 99 codes
  
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
      ResYear$ = Form$(2, 0)
      DevSpec$ = LEFT$(Form$(3, 0), 1)
      RETURN
    CASE EscKey
      Canceled$ = "Y"
      RETURN
    END SELECT
  LOOP
  
  
END SUB

SUB ResidentReport
  
  
  SHARED Choice$()
  
  
  ReportFile$ = "DCRESREP.PRN"  'Report File Name
  FF$ = CHR$(12)
  MaxLines = 58
  LineCnt = 0
  CustCnt = 0
  LPTPort% = 1
  
  GOSUB SelectResRepOutput: IF Canceled$ = "Y" THEN EXIT SUB
  
  
  
  
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  
  'PRINT #RptHandle, CHR$(27); CHR$(58); ' oki 320 12 cpi
  
  GOSUB PrintResRepRptHeader
  
  ' Print Main Body
  OpenDCCustFile NumOfDcRecs, DCFile
  REDIM DCIdxRec(1) AS DCCustIdxRecType
  IdxFile = FREEFILE
  OPEN "DCCust.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS #IdxFile LEN = LEN(DCIdxRec(1))
  
  
  
  
  IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN NumOfDcRecs = LOF(IdxFile) \ LEN(DCIdxRec(1))
  
  FOR Cnt! = 1 TO NumOfDcRecs
    IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN
      GET IdxFile, Cnt!, DCIdxRec(1)
      GET DCFile, DCIdxRec(1).IdxRecord, DCCustRec(1)
      CustomerNumber = DCIdxRec(1).IdxRecord
    ELSE
      GET DCFile, Cnt!, DCCustRec(1)
      CustomerNumber = Cnt!
    END IF
    IF DCCustRec(1).Deleted <> "Y" THEN
      IF LineCnt >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintResRepRptHeader
      END IF
      IF DCCustRec(1).Resident = "Y" THEN
        PRINT #RptHandle, USING "#####"; CustomerNumber;
        PRINT #RptHandle, TAB(10); DCCustRec(1).BILLNAME; TAB(50); DCCustRec(1).SoSec
        CustCnt = CustCnt + 1
        LineCnt = LineCnt + 1
      END IF
    END IF
  NEXT Cnt!
  GOSUB PrintResRepRptEnding
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi
  CLOSE         'Close all open files now
  
  IF DevSpec$ = "S" THEN
    EntryPoint = 2
  ELSE
    EntryPoint = 5
  END IF
  Header$ = "Va. Resident Report"
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  KILL ReportFile$
  
  EXIT SUB
  
  
PrintResRepRptHeader:
  page = page + 1
  PRINT #RptHandle, TAB(22); "Va. Decal System : Customer Resident Listing"
  PRINT #RptHandle, "Tax Year: "; ResYear$
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; page
  PRINT #RptHandle, "Acct #"; TAB(10); "Customer Name"; TAB(50); "Soc Sec #"
  PRINT #RptHandle, STRING$(80, "=")
  LineCnt = 5
  RETURN
  
PrintResRepRptEnding:
  PRINT #RptHandle, STRING$(80, "-")
  PRINT #RptHandle, "Total Customers Printed: "; USING "#####"; CustCnt
  PRINT #RptHandle,
  PRINT #RptHandle, FF$
  RETURN
  
  
  
  
  
  
  
SelectResRepOutput:
  LibName$ = "DC"
  ScrnName$ = "DCRESREP"
  
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
  
  
  REDIM Choice$(100, 3)         ' assume maximum of 99 codes
  
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
      ResYear$ = Form$(2, 0)
      DevSpec$ = LEFT$(Form$(3, 0), 1)
      RETURN
    CASE EscKey
      Canceled$ = "Y"
      RETURN
    END SELECT
  LOOP
  
END SUB

SUB ShowNoCode
  LibName$ = "DC"
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
  LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
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
  
  DIM Cat$(250), Cat1$(250), CatAmt#(250), CatAmt1#(250)
  
  ReportFile$ = "DCTrans.PRN"   'Report File Name
  FF$ = CHR$(12)
  MaxLines = 53
  LPTPort% = 1
  LineCnt = 0

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
  
  REDIM Array(1 TO size) AS Struct

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
    GET DCTRFile, Array(Cnt).RecNum, DCTransRec(1)
    
    IF LineCnt >= MaxLines THEN
      PRINT #RptHandle, FF$
      GOSUB PrintRptHeader2
    END IF
    
    'Get Customer
    OpenDCCustFile NumOfDcRecs, DCFile
    GET DCFile, VAL(DCTransRec(1).CustomerNumber), DCCustRec(1)
    
    PRINT #RptHandle, Num2Date$(DCTransRec(1).TransDate);
    PRINT #RptHandle, TAB(13); LEFT$(DCCustRec(1).BILLNAME, 25);
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
    'Now Subtotal by Decal Type
    GOSUB ChargeSub
    GOSUB PaymentSub
    
  NEXT Cnt
  GOSUB PrintRptEnding2
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi
  CLOSE         'Close all open files now
  
  IF DevSpec$ = "S" THEN
    EntryPoint = 2
  ELSE
    EntryPoint = 5
  END IF
  Header$ = "Transaction Journal"
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  KILL ReportFile$
  
  EXIT SUB
  
ChargeSub:
  
  IF DCTransRec(1).TransType = 1 THEN
    IF CatCnt = 0 THEN
      CatCnt = 1
      Cat$(1) = LTRIM$(DCTransRec(1).DecalCat)
      CatAmt#(1) = DCTransRec(1).TransAmount
      RETURN
    END IF
    FOR CatLoop = 1 TO CatCnt
      IF Cat$(CatLoop) = LTRIM$(DCTransRec(1).DecalCat) THEN
        CatAmt#(CatLoop) = CatAmt#(CatLoop) + DCTransRec(1).TransAmount
        RETURN
      END IF
    NEXT CatLoop
    CatCnt = CatCnt + 1
    Cat$(CatCnt) = LTRIM$(DCTransRec(1).DecalCat)
    CatAmt#(CatCnt) = DCTransRec(1).TransAmount
  END IF
RETURN
  
PaymentSub:
  IF DCTransRec(1).TransType = 2 THEN
    IF CatCnt1 = 0 THEN
      CatCnt1 = 1
      Cat1$(1) = LTRIM$(DCTransRec(1).DecalCat)
      CatAmt1#(1) = DCTransRec(1).TransAmount
      RETURN
    END IF
    FOR CatLoop = 1 TO CatCnt1
      IF Cat$(CatLoop) = LTRIM$(DCTransRec(1).DecalCat) THEN
        CatAmt1#(CatLoop) = CatAmt1#(CatLoop) + DCTransRec(1).TransAmount
        RETURN
      END IF
    NEXT CatLoop
    CatCnt1 = CatCnt1 + 1
    Cat$(CatCnt1) = LTRIM$(DCTransRec(1).DecalCat)
    CatAmt1#(CatCnt1) = DCTransRec(1).TransAmount
  END IF
RETURN
  
PrintRptHeader2:
  page = page + 1
  PRINT #RptHandle, TAB(20); "Va. Vehicle Decals : Transactions Journal "
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; page
  PRINT #RptHandle, ""
  PRINT #RptHandle, "  Date"; TAB(13); "Customer Name"; TAB(40); "Type"; TAB(50); "Description"; TAB(70); "Amount"
  PRINT #RptHandle, STRING$(80, "=")
  LineCnt = 5
RETURN
  
PrintRptEnding2:
  PRINT #RptHandle, STRING$(80, "-")
  PRINT #RptHandle, TAB(69); USING "$$#####,#.##"; TotalAmt#
  FOR Cnt& = 1 TO 2
    IF TotalAmt#(Cnt&) <> 0 THEN
      PRINT #RptHandle, "Trans Type : ";
      IF Cnt& = 1 THEN PRINT #RptHandle, "Charges  ";
      IF Cnt& = 2 THEN PRINT #RptHandle, "Payments ";
      PRINT #RptHandle, "     Total Amount: "; USING "$$######,#.##"; TotalAmt#(Cnt&)
    END IF
  NEXT Cnt&
  PRINT #RptHandle, STRING$(79, "-")
  PRINT #RptHandle, "Catagory Totals : Charges"
  PRINT #RptHandle, " "
  PRINT #RptHandle, "Catagory"; TAB(20); "       Amount"
  FOR Lp = 1 TO CatCnt
    PRINT #RptHandle, Cat$(Lp); TAB(20); USING "$$######,#.##"; CatAmt#(Lp)
  NEXT
  PRINT #RptHandle, " "
  PRINT #RptHandle, " "
  PRINT #RptHandle, "Catagory Totals : Payments"
  PRINT #RptHandle, " "
  PRINT #RptHandle, "Catagory"; TAB(20); "       Amount"
  FOR Lp = 1 TO CatCnt1
    PRINT #RptHandle, Cat$(Lp); TAB(20); USING "$$######,#.##"; CatAmt1#(Lp)
  NEXT
  PRINT #RptHandle, FF$
RETURN
  
GetReportInformation2:
  
  DCTRFile = FREEFILE
  OPEN "DCTrans.Dat" FOR RANDOM ACCESS READ WRITE SHARED AS #DCTRFile LEN = LEN(DCTransRec(1))
  TrNumRecs = LOF(DCTRFile) / LEN(DCTransRec(1))
  
  FOR Cnt& = 1 TO TrNumRecs
    GET DCTRFile, Cnt&, DCTransRec(1)
    IF LEFT$(Type$, 1) = "A" THEN
      IF DCTransRec(1).TransDate >= BegDate AND DCTransRec(1).TransDate <= EndDate THEN
        Count = Count + 1
        Array(Count).who = LTRIM$(STR$(DCTransRec(1).TransDate)) + DCTransRec(1).CustomerNumber + STRING$(4, " ")
        Array(Count).RecNum = Cnt&
      END IF
    END IF
    
    IF LEFT$(Type$, 1) = "C" THEN
      IF DCTransRec(1).TransType = 1 THEN
        IF DCTransRec(1).TransDate >= BegDate AND DCTransRec(1).TransDate <= EndDate THEN
          Count = Count + 1
          Array(Count).who = LTRIM$(STR$(DCTransRec(1).TransDate)) + DCTransRec(1).CustomerNumber + STRING$(4, " ")
          Array(Count).RecNum = Cnt&
        END IF
      END IF
    END IF
    
    IF LEFT$(Type$, 1) = "P" THEN
      IF DCTransRec(1).TransType = 2 THEN
        IF DCTransRec(1).TransDate >= BegDate AND DCTransRec(1).TransDate <= EndDate THEN
          Count = Count + 1
          Array(Count).who = LTRIM$(STR$(DCTransRec(1).TransDate)) + DCTransRec(1).CustomerNumber + STRING$(4, " ")
          Array(Count).RecNum = Cnt&
        END IF
      END IF
    END IF
  NEXT Cnt&
  
  SortT Array(Start), Count, Dir, SSize, MOff, MSize

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
    CASE EscKey
      Canceled$ = "Y"
      RETURN
    END SELECT
  LOOP
  RETURN
  
  
END SUB

