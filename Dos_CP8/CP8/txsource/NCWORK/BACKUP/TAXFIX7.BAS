DEFINT A-Z
DECLARE SUB Abstracts ()
DECLARE SUB ShowCustHistory (CustRec&)
DECLARE SUB CustHistoryRpt ()
DECLARE SUB GetPropRecList (PropRecs() AS LONG, CustRec&)
DECLARE SUB GetPersRecList (PersRecs() AS LONG, CustRec&)
DECLARE SUB LookUp (RecNo&, Text$, ChkBalFlag%, CLSFlag%, SSNFlag%)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB OpenTaxCustFile (NumOfTaxRecs%, TaxFile%)
DECLARE SUB OpenTaxPropFile (NumOfPropRecs%, PropTaxFile%)
DECLARE SUB OpenTaxPersFile (NumOfPersRecs%, PersTaxFile%)
DECLARE SUB AbtractListing ()
DECLARE SUB BalanceListing ()
DECLARE SUB MortgageCodeList ()
DECLARE SUB MasterValueList ()
DECLARE SUB TransactionJournal ()
DECLARE SUB LateListing ()
DECLARE SUB CustomerInquiry ()
DECLARE SUB SrCitizensList ()
DECLARE SUB Labels ()
DECLARE SUB AdListing ()
DECLARE SUB DisplayTaxScrn (ScrnName$)
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
DECLARE FUNCTION Round# (B#)
  '$INCLUDE: 'DefCnf.BI'
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)
DECLARE SUB HideCursor ()
DECLARE SUB CursorOff ()
DECLARE SUB TextCursor (MouseFg%, MouseBg%)
DECLARE SUB WaitForAction ()
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION GetCustBalance# (RecNo&)
  
  
  CONST False = 0, True = NOT False
  
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'Taxfiles.BI'
  '$INCLUDE: 'PROPAbst.BI'
  '$INCLUDE: 'TAXRPTTY.BI'
  
  DIM SHARED TaxSetupRec AS TaxMasterType
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED PersRec(1) AS PersonalRecType
  DIM SHARED TaxTrans(1) AS TaxTransactionType
  
  TYPE Struct
    who AS STRING * 14
    RecNum AS INTEGER
  END TYPE
  
  STACK 5000
  
  ClearBack
  CustHistoryRpt
  
  
  END

SUB oCustomerListing
  
  SHARED Choice$()
  ReportFile$ = "TaxCust.PRN"   'Report File Name
  CommaFmt$ = "########,.##"    'format takes 13 chars
  TotalFmt$ = "#########,.##"   'format takes 14 chars
  SumLine$ = STRING$(13, "-")   'column summary line
  DivLine$ = STRING$(77, "-")   'dashed line
  DivLine2$ = STRING$(77, "=")  'Double Line
  
  size = 5500
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
  
  FF$ = CHR$(12)
  MaxLines = 60
  LineCnt = 0
  CustCnt = 0
  
  GOSUB SelectDetailCustomerOutput
  IF Canceled$ = "Y" THEN EXIT SUB
  
  
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  GOSUB oPrintDetailCustomerRptHeader
  
  ' Print Main Body
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  
  
  IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN
    GOSUB oGetNameIndex
  END IF
  
  FOR Cnt! = 1 TO NumOfTaxRecs
    
    IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN
      GET TaxFile, Array(Cnt!).RecNum, TaxCustRec(1)
      CustomerNumber = Array(Cnt!).RecNum
    ELSE
      GET TaxFile, Cnt!, TaxCustRec(1)
      CustomerNumber = Cnt!
    END IF
    Help$ = "Processing Record # " + STR$(Cnt!)
    
    ' Main Print Processing Here
    
    IF NOT (TaxCustRec(1).Deleted) THEN
      
      IF LineCnt >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB oPrintDetailCustomerRptHeader
      END IF
      
      ' Print Line Here
      ' Get Name First
      
      nme$ = RTRIM$(TaxCustRec(1).FName) + " " + RTRIM$(TaxCustRec(1).LName)
      nme$ = LTRIM$(nme$)
      
      IF Detail$ = "Summary" THEN
        PRINT #RptHandle, CustomerNumber; TAB(10); nme$; TAB(60); TaxCustRec(1).ACTIVE
        LineCnt = LineCnt + 1
      ELSE
        PRINT #RptHandle, "Cust #: "; CustomerNumber; TAB(15); nme$
        PRINT #RptHandle, "Active: "; TaxCustRec(1).ACTIVE; TAB(15); TaxCustRec(1).ADDR1
        PRINT #RptHandle, "Int'st: "; TaxCustRec(1).INTEREST; TAB(15); TaxCustRec(1).ADDR2
        PRINT #RptHandle, "Exempt: "; TaxCustRec(1).TaxExempt; TAB(15); RTRIM$(TaxCustRec(1).CITY) + ", "; RTRIM$(TaxCustRec(1).STATE) + "  " + RTRIM$(TaxCustRec(1).ZIP)
        PRINT #RptHandle, ""
        LineCnt = LineCnt + 5
        
        
        
      END IF
      CustCnt = CustCnt + 1
      
    END IF
  NEXT Cnt!
  GOSUB oPrintDetailCustomerRptEnding
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi
  
  CLOSE         'Close all open files now
  
  IF DevSpec$ = "P" THEN
    EntryPoint = 4
  ELSE
    EntryPoint = 1
  END IF
  
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  KILL ReportFile$
  
  EXIT SUB
  
  
oPrintDetailCustomerRptHeader:
  Page = Page + 1
  PRINT #RptHandle, TAB(20); "Property Tax Detailed Customer Listing"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
  IF Detail$ = "Summary" THEN
    PRINT #RptHandle, "Summary Format"
    PRINT #RptHandle, "Acct #"; TAB(10); "Name"; TAB(55); "Active"
    PRINT #RptHandle, STRING$(80, "=")
    LineCnt = 5
  ELSE
    PRINT #RptHandle, "Detail Format"
    PRINT #RptHandle, STRING$(132, "=")
    LineCnt = 4
  END IF
  RETURN
  
oPrintDetailCustomerRptEnding:
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
      Detail$ = Form$(2, 0)
      DevSpec$ = LEFT$(Form$(3, 0), 1)
      RETURN
    CASE EscKey
      Canceled$ = "Y"
      RETURN
    END SELECT
  LOOP
  
  
oGetNameIndex:
  FOR Cnt! = 1 TO NumOfTaxRecs
    GET TaxFile, Cnt!, TaxCustRec(1)
    Array(Cnt!).who = UCASE$(TaxCustRec(1).SName) + " "
    Array(Cnt!).RecNum = Cnt!
    Count = NumOfTaxRecs
  NEXT Cnt!
  
  'Sort Them Here
  SortT Array(Start), Count, Dir, SSize, MOff, MSize
  RETURN
  
  
END SUB

SUB OpenTaxCustFile (NumOfTaxRecs, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "TAXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
  
END SUB

SUB OpenTaxPersFile (NumOfPersRecs, PersTaxFile)
  PersTaxFile = FREEFILE
  OPEN "TAXPERS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PersTaxFile LEN = LEN(PersRec(1))
  NumOfPersRecs = LOF(PersTaxFile) / LEN(PersRec(1))
END SUB

SUB OpenTaxPropFile (NumOfPropRecs, PropTaxFile)
  PropTaxFile = FREEFILE
  OPEN "TAXPROP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PropTaxFile LEN = LEN(PropertyRec(1))
  NumOfPropRecs = LOF(PropTaxFile) / LEN(PropertyRec(1))
END SUB

SUB SrCitizensList
  SHARED Choice$()
  
  REDIM Array(1 TO 1) AS Struct 'Template for the sort array
  ReportFile$ = "TaxSC.PRN"     'Report File Name
  Dash80$ = STRING$(80, "=")
  FF$ = CHR$(12)
  
  MaxLines = 56
  LineCnt = 0
  CustCnt = 0
  
  LibName$ = "TAX"
  ScrnName$ = "SCRPT"
  
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
  Choice$(1, 0) = "Name Order"
  Choice$(2, 0) = "Account Number"
  Choice$(0, 2) = "3"
  Choice$(1, 2) = "Screen"
  Choice$(2, 2) = "Printer"
  
  Action = 1
  ClearBack
  
  ShowCursor
  
  DisplayTaxScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      IF LEFT$(Form$(1, 0), 1) = "N" THEN
        UsingIndex = True
      ELSE
        UsingIndex = False
      END IF
      TaxRate! = VAL(Form$(2, 0))
      DevSpec$ = LEFT$(Form$(3, 0), 1)
      ExitFlag = True
    CASE EscKey
      AbortFlag = True
      ExitFlag = True           'EXIT DO
    END SELECT
  LOOP UNTIL ExitFlag
  
  IF AbortFlag THEN EXIT SUB
  
  RptHandle = FREEFILE
  
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  GOSUB PrintSCRptHeader
  
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile
  
  ClearBack
  ShowProcessingScrn "Senior Citizen Listing"
  
  IF UsingIndex AND NumOfTaxRecs > 0 THEN
    GOSUB GetNameIndexSC
  END IF
  
  FOR Cnt = 1 TO NumOfTaxRecs
    IF UsingIndex THEN
      CustRecNo = Array(Cnt).RecNum
    ELSE
      CustRecNo = Cnt
    END IF
    
    GET TaxFile, CustRecNo, TaxCustRec(1)
    'Set SC Amt to Zero
    SCAmt# = 0
    IF NOT TaxCustRec(1).Deleted THEN
      IF LineCnt >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintSCRptHeader
      END IF
      
      nme$ = QPTrim$(TaxCustRec(1).FName) + " " + QPTrim$(TaxCustRec(1).LName)
      nme$ = QPTrim$(nme$)      'this one cleans up those with only last name
      
      
      'Now Show Property Records Next
      
      IF TaxCustRec(1).FirstPropRec > 0 THEN
        PropertyRecord! = TaxCustRec(1).FirstPropRec
        WHILE PropertyRecord! <> 0
          GET #PropTaxFile, PropertyRecord!, PropertyRec(1)
          SCAmt# = SCAmt# + PropertyRec(1).EXMPSENI
          PropertyRecord! = PropertyRec(1).NextRec
        WEND
      END IF
      
      'NOW CHECK PERSONAL PROPERTY
      IF TaxCustRec(1).FirstPersRec > 0 THEN
        PropertyRecord! = TaxCustRec(1).FirstPersRec
        WHILE PropertyRecord! <> 0
          GET #PersTaxFile, PropertyRecord!, PersRec(1)
          SCAmt# = SCAmt# + PersRec(1).EXMPSENI
          PropertyRecord! = PersRec(1).NextRec
        WEND
      END IF
      
      IF SCAmt# > 0 THEN
        TaxLoss# = (SCAmt# * TaxRate!) / 100
        PRINT #RptHandle, TaxCustRec(1).CSSN; TAB(15); nme$; TAB(57); USING "$$#######,#"; SCAmt#;
        PRINT #RptHandle, TAB(71); USING "$####,#.##"; TaxLoss#
        PRINT #RptHandle, TAB(15); RTRIM$(TaxCustRec(1).ADDR1) + " " + RTRIM$(TaxCustRec(1).CITY) + " " + TaxCustRec(1).STATE + " " + TaxCustRec(1).ZIP
        PRINT #RptHandle, ""
        LineCnt = LineCnt + 3
        CustCnt = CustCnt + 1
        TotalSCAmt# = TotalSCAmt# + SCAmt#
        TotalTaxLoss# = TotalTaxLoss# + TaxLoss#
      END IF
    END IF
    ShowPctComp Cnt, NumOfTaxRecs
  NEXT
  
  GOSUB PrintSCRptEnding
  
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi
  
  CLOSE         'Close all open files now
  
  IF DevSpec$ = "P" THEN
    EntryPoint = 4
  ELSEIF DevSpec$ = "S" THEN
    EntryPoint = 2
  ELSE
    EntryPoint = 1
  END IF
  
  ERASE Array, Frm, Form$, Fld, TaxCustRec
  
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  KILL ReportFile$
  
  EXIT SUB
  
PrintSCRptHeader:
  Page = Page + 1
  PRINT #RptHandle, TAB(20); "Senior Citizen Discount Report    :   Form AV-22C"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
  PRINT #RptHandle, ""
  PRINT #RptHandle, "Soc Sec #"; TAB(15); "Name/Address"; TAB(57); "Exempt Amt"; TAB(71); "Tax Loss"
  PRINT #RptHandle, Dash80$
  LineCnt = 5
  RETURN
  
PrintSCRptEnding:
  PRINT #RptHandle, Dash80$
  PRINT #RptHandle, "Total Customers Printed: "; USING "#####"; CustCnt
  PRINT #RptHandle, "Total Value of Discount: "; USING "$$#######,#"; TotalSCAmt#
  PRINT #RptHandle, "  Total Tax Loss Amount: "; USING "$$######,#.##"; TotalTaxLoss#
  PRINT #RptHandle, FF$
  RETURN
  
GetNameIndexSC:
  REDIM Array(1 TO NumOfTaxRecs) AS Struct
  FOR Cnt = 1 TO NumOfTaxRecs
    GET TaxFile, Cnt, TaxCustRec(1)
    Array(Cnt).who = UCASE$(TaxCustRec(1).SName) + " "
    Array(Cnt).RecNum = Cnt
  NEXT
  
  'Sort Them Here
  SortT Array(1), NumOfTaxRecs, 0, LEN(Array(1)), 0, 14
  RETURN
  
END SUB

