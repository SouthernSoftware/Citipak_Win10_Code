DECLARE SUB DisplayTaxScrn (ScrnName$)
DEFINT A-Z
DECLARE SUB OpenTaxCustFile (NumOfTaxRecs%, TaxFile%)
DECLARE SUB CustomerListing ()
DECLARE SUB TAXCustomerMenu ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB ClearBack ()
DECLARE SUB SendDist2GL ()
DECLARE SUB UBMiscMenu ()
DECLARE SUB UBBillMenu ()
DECLARE SUB UBCustomerMenu ()
DECLARE SUB ClearScrn ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION WEnvTest% ()
  
  '$INCLUDE: 'DefCnf.BI'
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)
DECLARE SUB HideCursor ()
DECLARE SUB CursorOff ()
DECLARE SUB TextCursor (MouseFg%, MouseBg%)
  
  CONST False = 0, True = NOT False
  
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'Taxfiles.BI'
  '$INCLUDE: 'PROPAbst.BI'
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  
  TYPE Struct
    who AS STRING * 14
    RecNum AS INTEGER
  END TYPE
  
  STACK 5000
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 12)
  
  MChoice$(1) = "Customer Inquiry"
  MChoice$(2) = "Master Customer Listing"
  MChoice$(3) = "Master Abstract Listing"
  MChoice$(4) = "Master Balance Listing"
  MChoice$(5) = "Master Mortage Code Listing"
  MChoice$(6) = "Master Valuation Listing"
  MChoice$(7) = "Transaction Journal"
  MChoice$(8) = "Late Listing"
  MChoice$(9) = "Sr. Citizens Listing"
  MChoice$(10) = "Advertising Listing"
  MChoice$(11) = "Mailing Labels"
  MChoice$(12) = "Exit to DOS"
  
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 18   'limit the box length to go no lower than line 18
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
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    
    ClearBack
    
    TitleBox 2, Col, MaxLen + 3, "Tax Billing Setup Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      'TAXCustomerMenu
      
    CASE 2
      CustomerListing
    CASE 3
    CASE 4
    CASE 5
    CASE 6
    CASE 7
    CASE 8
    CASE 9
    CASE 10
    CASE 11
    CASE 12
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP
  
  IF WEnvTest THEN
    Ext$ = ".bas"
  ELSE
    Ext$ = ".exe"
  END IF

  IF Exist("Taxmenu" + Ext$) THEN
    RUN "TaxMenu"
  ELSE
    HideCursor
    ClearScrn
  END IF
  
  END

SUB CustomerListing
  
  SHARED Choice$()

  ReportFile$ = "TaxCust.PRN"   'Report File Name

'CommaFmt$ = "########,.##"    'format takes 13 chars
'TotalFmt$ = "#########,.##"   'format takes 14 chars
'SumLine$ = STRING$(13, "-")   'column summary line
'DivLine$ = STRING$(77, "-")   'dashed line
'DivLine2$ = STRING$(77, "=")  'Double Line

  Dash80$ = STRING$(80, "=")

  size = 5500
  Start = 1     'start at array element 1
  Dir = 0       'sort direction - use anything else for descending
  SSize = 16    'total size of each TYPE element
  MOff = 0      'offset into the TYPE for the key element
  MSize = 16    'size of the key element - coded as follows:

  REDIM array(1 TO size) AS Struct
  
  FF$ = CHR$(12)

  MaxLines = 60
  LineCnt = 0
  CustCnt = 0
  
  GOSUB SelectDetailCustomerOutput

'IF Canceled$ = "Y" THEN EXIT SUB
  IF AbortFlag THEN EXIT SUB
  
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  GOSUB PrintDetailCustomerRptHeader
  
  ' Print Main Body

  OpenTaxCustFile NumOfTaxRecs, TaxFile
  
  IF UsingIndex THEN
    GOSUB GetNameIndex
  END IF
  
  FOR Cnt = 1 TO NumOfTaxRecs

    IF UsingIndex THEN
      CustRecNo = array(Cnt).RecNum
    ELSE
      CustRecNo = Cnt
    END IF

    GET TaxFile, CustRecNo, TaxCustRec(1)
    
'    IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN
'      GET TaxFile, array(Cnt!).RecNum, TaxCustRec(1)
'      CustRecNo = array(Cnt!).RecNum
'    ELSE
'      GET TaxFile, Cnt!, TaxCustRec(1)
'      CustRecNo = Cnt!
'    END IF

    'Help$ = "Processing Record # " + STR$(Cnt!)
    ' Main Print Processing Here
    
'IF NOT (TaxCustRec(1).Deleted) THEN

    IF NOT TaxCustRec(1).Deleted THEN

      IF LineCnt >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintDetailCustomerRptHeader
      END IF
      
      ' Print Line Here
      ' Get Name First

      Nme$ = QPTrim$(TaxCustRec(1).FNAME) + " " + QPTrim$(TaxCustRec(1).LName)

'Nme$ = RTRIM$(TaxCustRec(1).FNAME) + " " + RTRIM$(TaxCustRec(1).LName)
'Nme$ = LTRIM$(Nme$)
      
'IF Detail$ = "Summary" THEN
      IF NOT DetailFlag THEN
        PRINT #RptHandle, CustRecNo; TAB(10); Nme$; TAB(60); TaxCustRec(1).Active
        LineCnt = LineCnt + 1
      ELSE
        PRINT #RptHandle, "Cust #: "; CustRecNo; TAB(15); Nme$
        PRINT #RptHandle, "Active: "; TaxCustRec(1).Active; TAB(15); TaxCustRec(1).Addr1
        PRINT #RptHandle, "Int'st: "; TaxCustRec(1).Interest; TAB(15); TaxCustRec(1).Addr2
        PRINT #RptHandle, "Exempt: "; TaxCustRec(1).TaxExempt; TAB(15); RTRIM$(TaxCustRec(1).City) + ", "; RTRIM$(TaxCustRec(1).State) + "  " + RTRIM$(TaxCustRec(1).Zip)
        PRINT #RptHandle, ""
        LineCnt = LineCnt + 5
      END IF
      CustCnt = CustCnt + 1
    END IF
  NEXT
'NEXT Cnt

  GOSUB PrintDetailCustomerRptEnding
  
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi

  CLOSE         'Close all open files now
  
  IF DevSpec$ = "P" THEN
    EntryPoint = 4
'**
  ELSEIF DevSpec$ = "S" THEN
    EntryPoint = 2
'**
  ELSE
    EntryPoint = 1
  END IF
  
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  KILL ReportFile$
  
  EXIT SUB
  
  
PrintDetailCustomerRptHeader:
  Page = Page + 1
  PRINT #RptHandle, TAB(20); "Property Tax Detailed Customer Listing"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
'IF Detail$ = "Summary" THEN
  IF NOT DetailFlag THEN
    PRINT #RptHandle, "Summary Format"
    PRINT #RptHandle, "Acct #"; TAB(10); "Name"; TAB(55); "Active"
    PRINT #RptHandle, Dash80$
    LineCnt = 5
  ELSE
    PRINT #RptHandle, "Detail Format"
    PRINT #RptHandle, STRING$(132, "=")
    LineCnt = 4
  END IF
RETURN
  
PrintDetailCustomerRptEnding:
  PRINT #RptHandle, Dash80$
  PRINT #RptHandle, "Total Customers Printed: "; USING "#####"; CustCnt
  PRINT #RptHandle,
  PRINT #RptHandle, FF$
RETURN
  
SelectDetailCustomerOutput:
  LibName$ = "TAX"
  ScrnName$ = "CUSTRPT"
  
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
  
  REDIM Choice$(0 TO 2, 0 TO 2)
  
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Name"
  Choice$(2, 0) = "Number"
  Choice$(0, 1) = "2"
  Choice$(1, 1) = "Summary"
  Choice$(2, 1) = "Detail"
  Choice$(0, 2) = "3"
  Choice$(1, 2) = "Screen"
  Choice$(2, 2) = "Printer"
  
  Action = 1
  ShowCursor

  DisplayTaxScrn ScrnName$
'LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%
  
  'printhelp help$
  Action = 1
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      
      IF LEFT$(Form$(1, 0), 1) = "N" THEN
        UsingIndex = True
      ELSE
        UsingIndex = False
      END IF

'Order$ = LEFT$(Form$(1, 0), 1)

      IF LEFT$(Form$(2, 0), 1) = "S" THEN
        DetailFlag = False
      ELSE
        DetailFlag = True
      END IF
'      Detail$ = Form$(2, 0)

      DevSpec$ = LEFT$(Form$(3, 0), 1)
      ExitFlag = True
'RETURN
    CASE EscKey
      AbortFlag = True
'Canceled$ = "Y"
      ExitFlag = True 'EXIT DO
'RETURN
    END SELECT
  LOOP UNTIL ExitFlag

RETURN
  
GetNameIndex:
  FOR Cnt = 1 TO NumOfTaxRecs
    GET TaxFile, Cnt, TaxCustRec(1)
    array(Cnt).who = UCASE$(TaxCustRec(1).SNAME) + " "
    array(Cnt).RecNum = Cnt
    Count = NumOfTaxRecs
  NEXT
'NEXT Cnt

'Sort Them Here
  SortT array(Start), Count, Dir, SSize, MOff, MSize
RETURN
  
  
END SUB

SUB OpenTaxCustFile (NumOfTaxRecs, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "TAXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
  
END SUB

