DEFINT A-Z

DEFSNG A-Z
DEFINT A-Z
SUB PrintEditList
  SHARED Choice$()
  DIM Cat$(250), CatAmt#(250)   'Set Maximum Catagories at 250
  ReportFile$ = "DCPAYED.PRN"   'Report File Name
  CommaFmt$ = "########,.##"    'format takes 13 chars
  TotalFmt$ = "#########,.##"   'format takes 14 chars
  SumLine$ = STRING$(13, "-")   'column summary line
  DivLine$ = STRING$(77, "-")   'dashed line
  DivLine2$ = STRING$(77, "=")  'Double Line
  FF$ = CHR$(12)
  MaxLines = 53
  LPTPORT% = 1
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
  
  '  REDIM array(1 TO size) AS Struct
  
  GOSUB SelectOutput
  
  
  CustRecLen = LEN(DCCustRec(1))
  TrHandle = FREEFILE
  OPEN "DCCust.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS TrHandle LEN = CustRecLen
  TrNumRecs = LOF(TrHandle) \ CustRecLen
  
  DCEditRecLen = LEN(EditPaymentRec(1))
  DCFile = FREEFILE
  OPEN "DCEDPYT.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCFile LEN = DCEditRecLen
  NumOfDcRecs = LOF(DCFile) \ DCEditRecLen
  
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  GOSUB PrintRptHeader
  
  FOR Cnt = 1 TO NumOfDcRecs
    GET DCFile, Cnt, EditPaymentRec(1)
    GET TrHandle, VAL(EditPaymentRec(1).CustNumber), DCCustRec(1)
    
    IF LineCnt >= MaxLines THEN
      PRINT #RptHandle, FF$
      GOSUB PrintRptHeader
    END IF
    IF EditPaymentRec(1).Amount <> 0 THEN
      PRINT #RptHandle, VAL(DCCustRec(1).CUSTNUMB); TAB(10); DCCustRec(1).BILLNAME; TAB(55); USING "$$#####,#.##"; EditPaymentRec(1).Amount;
      PRINT #RptHandle, TAB(73); EditPaymentRec(1).DecalCat
      PRINT #RptHandle, "Veh Desc: "; EditPaymentRec(1).MakeModel; TAB(50); "State Tag# "; RTRIM$(EditPaymentRec(1).StateTag)
      PRINT #RptHandle, STRING$(80, "-")
      TotalCust = TotalCust + 1
      TotalValue# = TotalValue# + EditPaymentRec(1).Amount
      TotalValue# = INT((TotalValue# * 100) + .5) / 100
      GOSUB CatagoryTotal
      LineCnt = LineCnt + 3
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
  Header$ = "Payment Edit Listing"
  PrintRptFile Header$, ReportFile$, LPTPORT%, RetCode%, EntryPoint
  
  KILL ReportFile$
  EXIT SUB
CatagoryTotal:
  IF CatCnt = 0 THEN
    CatCnt = 1
    Cat$(1) = LTRIM$(EditPaymentRec(1).DecalCat)
    CatAmt#(1) = EditPaymentRec(1).Amount
    RETURN
  END IF
  FOR CatLoop = 1 TO CatCnt
    IF Cat$(CatLoop) = LTRIM$(EditPaymentRec(1).DecalCat) THEN
      CatAmt#(CatLoop) = CatAmt#(CatLoop) + EditPaymentRec(1).Amount
      RETURN
    END IF
  NEXT CatLoop
  CatCnt = CatCnt + 1
  Cat$(CatCnt) = LTRIM$(EditPaymentRec(1).DecalCat)
  CatAmt#(CatCnt) = EditPaymentRec(1).Amount
  RETURN
  
  
  
  RETURN
  
PrintRptHeader:
  page = page + 1
  PRINT #RptHandle, TAB(15); "Decal System : Payment (Decal Purchase's) Edit Listing"
  PRINT #RptHandle, TAB(21); "      Report Date: "; DATE$; TAB(68); "Page #"; page
  PRINT #RptHandle, ""
  PRINT #RptHandle, "Cust #"; TAB(10); "Billing Name"; TAB(56); "Payment Amt"; TAB(73); "Type"
  PRINT #RptHandle, STRING$(80, "=")
  LineCnt = 5
  RETURN
  
PrintRptEnding:
  PRINT #RptHandle, "Number of Entries .. "; USING "####,#"; TotalCust;
  PRINT #RptHandle, TAB(55); USING "$$#####,#.##"; TotalValue#
  PRINT #RptHandle, FF$
  PRINT #RptHandle, TAB(15); "Decal System : Payment (Decal Purchase's) Edit Listing"
  PRINT #RptHandle, TAB(21); "      Report Date: "; DATE$; TAB(68); "Page #"; page
  PRINT #RptHandle, "Catagory Totals"
  PRINT #RptHandle, "Catagory"; TAB(20); "       Amount"
  FOR Lp = 1 TO CatCnt
    PRINT #RptHandle, Cat$(Lp); TAB(20); USING "$$######,#.##"; CatAmt#(Lp)
  NEXT Lp
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
  PrintHelp help$
  Action = 1
  COLOR 14: LOCATE 9, 23: PRINT "Payment Edit List"; : COLOR 10: PRINT " ]";
  
  
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

