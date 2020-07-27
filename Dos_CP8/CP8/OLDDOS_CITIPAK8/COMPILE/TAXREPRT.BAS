DEFINT A-Z
DECLARE SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE SUB PrintEditReport ()
DECLARE SUB MakeAdCostReport ()
DECLARE SUB MakeCustIndex (IdxType%)
DECLARE SUB MortCodeReport ()
DECLARE SUB TaxPaymentRpt ()
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
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB HideCursor ()
DECLARE SUB CursorOff ()
DECLARE SUB TextCursor (MouseFg%, MouseBg%)
DECLARE SUB WaitForAction ()
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION GetCustBalance# (RecNo&)
DECLARE SUB QPrintRC (Text$, Row, Col, Attrib)
DECLARE FUNCTION ParseBillNum$ (Text$)

  CONST False = 0, True = NOT False
  
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'Taxfiles.BI'
  '$INCLUDE: 'PROPAbst.BI'
  '$INCLUDE: 'TAXRPTTY.BI'
  ''$INCLUDE: 'TAXCONST.BI'

  DIM SHARED TaxSetUpRec AS TaxMasterType
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED PersRec(1) AS PersonalRecType
  DIM SHARED TaxTrans(1) AS TaxTransactionType
  
  TYPE PaySumType
    Year        AS INTEGER
    Principle   AS DOUBLE
    Interest    AS DOUBLE
    PENALTY     AS DOUBLE
    ColAdd      AS DOUBLE
  END TYPE
  
  
  TYPE Struct
    who AS STRING * 14
    RecNum AS INTEGER
  END TYPE
  
  TYPE ParcelSortType
    Parcel    AS STRING * 12
    CustRec   AS INTEGER
    ParcelRec AS INTEGER
  END TYPE

  STACK 5000
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 14)
  
  MChoice$(1) = "Customer Inquiry"
  MChoice$(2) = "Customer Transaction History "
  MChoice$(3) = "Master Customer Listing"
  MChoice$(4) = "Master Balance Listing"
  MChoice$(5) = "Master Valuation Listing"
  MChoice$(6) = "Tax Payment Report"
  MChoice$(7) = "Transaction Journal"
  MChoice$(8) = "Mailing Labels"
  MChoice$(9) = "Late Listing"
  MChoice$(10) = "Mortage Code Report"
  MChoice$(11) = "Sr. Citizens Listing"
  MChoice$(12) = "Print Abstracts"
  MChoice$(13) = "Print Advertising Report"
  MChoice$(14) = "Exit to DOS"
  
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
  Row = ((25 - (UBOUND(MChoice$))) \ 2) + 2
  Col = ((80 - MaxLen) \ 2) - 1
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    
    ClearBack
    
    TitleBox 2, Col, MaxLen + 3, "Tax Billing Reports Menu", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      CustomerInquiry
    CASE 2
      'unrem
      CustHistoryRpt
    CASE 3
      CustomerListing
    CASE 4
      BalanceListing
    CASE 5
      MasterValueList
    CASE 6
      TaxPaymentRpt
    CASE 7
      TransactionJournal
    CASE 8
      Labels
    CASE 9
      LateListing
    CASE 10
      MortCodeReport
    CASE 11
      SrCitizensList
    CASE 12
      Abstracts
    CASE 13
      MakeAdCostReport
    CASE 14
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

SUB Abstracts
  
  SHARED Choice$()
  
  REDIM Array(1 TO 1) AS Struct 'Template for the sort array
  ReportFile$ = "TaxAbst.PRN"   'Report File Name
  Dash80$ = STRING$(80, "=")
  FF$ = CHR$(12)
  
  MaxLines = 56
  LineCnt = 0
  CustCnt = 0
  
  LibName$ = "TAX"
  ScrnName$ = "ABLIST"
  
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
      DetailFlag = True
      TaxYear$ = Form$(2, 0)
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
  
  
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  
  IF UsingIndex AND NumOfTaxRecs > 0 THEN
    GOSUB GetNameIndex2
  END IF
  
  ClearBack
  ShowProcessingScrn "ABSTRACTS"
  
  FOR Cnt = 1 TO NumOfTaxRecs
    IF UsingIndex THEN
      CustRecNo = Array(Cnt).RecNum
    ELSE
      CustRecNo = Cnt
    END IF
    
    GET TaxFile, CustRecNo, TaxCustRec(1)
    
    IF NOT TaxCustRec(1).Deleted THEN
      
      NME$ = QPTrim$(TaxCustRec(1).FName) + " " + QPTrim$(TaxCustRec(1).LName)
      NME$ = QPTrim$(NME$)      'this one cleans up those with only last name
      
      
      
      'Now Show Property Records Next
      
      IF TaxCustRec(1).FirstPropRec > 0 THEN
        
        PropertyRecord! = TaxCustRec(1).FirstPropRec
        FirstProperty = 1
        
        WHILE PropertyRecord! <> 0
          
          GOSUB AbHeading
          
          PRINT #RptHandle, "Property Owned..."
          GET #PropTaxFile, PropertyRecord!, PropertyRec(1)
          PRINT #RptHandle, "            ---------------------------------------------------------------"
          PRINT #RptHandle, "      PIN#: "; QPTrim$(PropertyRec(1).RealPin)
          PRINT #RptHandle, "            ---------------------------------------------"
          PRINT #RptHandle, "      MAP#:"; QPTrim$(PropertyRec(1).MAP); "/"; QPTrim$(PropertyRec(1).BLOCK); "/"; QPTrim$(PropertyRec(1).LOTNUMB); TAB(50); "Real Value: "; "_____________"
          PRINT #RptHandle, "            ---------------------------------------------"
          PRINT #RptHandle, "      Desc: "; QPTrim$(PropertyRec(1).PROPNOT1)
          PRINT #RptHandle, "      Desc: "; QPTrim$(PropertyRec(1).PROPNOT2)
          PRINT #RptHandle, "      Desc: "; QPTrim$(PropertyRec(1).PROPNOT3)
          PRINT #RptHandle, "            ---------------------------------------------------------------"
          PRINT #RptHandle, "Acres/Lots:"
          PRINT #RptHandle, "            ---------------------------------------------------------------"
          PRINT #RptHandle, STRING$(79, "=")
          
          IF FirstProperty = 1 THEN
            PRINT #RptHandle, STRING$(79, "=")
            PRINT #RptHandle, "**** PERSONAL PROPERTY ****"
            PRINT #RptHandle, ""
            PRINT #RptHandle, "   Year        Make              Model            Vehicle ID Number"
            PRINT #RptHandle, "|        |               |                   |                             | "
            PRINT #RptHandle, "|        |               |                   |                             | "
            PRINT #RptHandle, "|        |               |                   |                             | "
            PRINT #RptHandle, "|        |               |                   |                             | "
            PRINT #RptHandle, "|        |               |                   |                             | "
            PRINT #RptHandle, "---------------------------------------------------------------------------"
            PRINT #RptHandle, ""
            PRINT #RptHandle, "Total Personal Value "; TAB(50); "            "; "_________________"
            PRINT #RptHandle, STRING$(79, "=")
            PRINT #RptHandle, ""
            PRINT #RptHandle, "  Customer SS# : _______________________   Spouse SS# ___________________"
            PRINT #RptHandle, ""
            PRINT #RptHandle, "Spouse Employer: _______________________   Res. Phone # _________________"
            PRINT #RptHandle, ""
            PRINT #RptHandle, "Customer Signature: ________________________________________"
            PRINT #RptHandle, "==============================================================================="
            PRINT #RptHandle, "Machinery & Fixtures:______________________"
            PRINT #RptHandle, "            Supplies:______________________"
            PRINT #RptHandle, "      Farm Equipment:______________________"
            PRINT #RptHandle, "      Boats & Motors:______________________"
            PRINT #RptHandle, "        Mobile Homes:______________________"
            PRINT #RptHandle, "             Campers:______________________"
            PRINT #RptHandle, ""
            PRINT #RptHandle, "     TOTAL PERS PROP:______________________"
            PRINT #RptHandle, ""
            PRINT #RptHandle, "     TOTAL REAL PROP:______________________"
            PRINT #RptHandle, ""
            PRINT #RptHandle, " Less Age/Disb. Excl:______________________"
            PRINT #RptHandle, ""
            PRINT #RptHandle, " TOTAL TAXABLE VALUE:______________________"
            PRINT #RptHandle, CHR$(12);
            FirstProperty = 0
          ELSE
            PRINT #RptHandle, CHR$(12);
          END IF
          
          PropertyRecord! = PropertyRec(1).NextRec
        WEND
        
      ELSE
        GOSUB AbHeading
        'Print Blank Form
        PRINT #RptHandle, "            ---------------------------------------------------------------"
        PRINT #RptHandle, "      PIN#: "
        PRINT #RptHandle, "            ---------------------------------------------"
        PRINT #RptHandle, "      MAP#:"; TAB(50); "Real Value: "; "_____________"
        PRINT #RptHandle, "            ---------------------------------------------"
        PRINT #RptHandle, "      Desc: "
        PRINT #RptHandle, "      Desc: "
        PRINT #RptHandle, "      Desc: "
        PRINT #RptHandle, "            ---------------------------------------------------------------"
        PRINT #RptHandle, "Acres/Lots:"
        PRINT #RptHandle, "            ---------------------------------------------------------------"
        PRINT #RptHandle, STRING$(79, "=")
        PRINT #RptHandle, "**** PERSONAL PROPERTY ****"
        PRINT #RptHandle, ""
        PRINT #RptHandle, "   Year        Make              Model            Vehicle ID Number"
        PRINT #RptHandle, "|        |               |                   |                             | "
        PRINT #RptHandle, "|        |               |                   |                             | "
        PRINT #RptHandle, "|        |               |                   |                             | "
        PRINT #RptHandle, "|        |               |                   |                             | "
        PRINT #RptHandle, "|        |               |                   |                             | "
        PRINT #RptHandle, "---------------------------------------------------------------------------"
        PRINT #RptHandle, ""
        PRINT #RptHandle, "Total Personal Value "; TAB(50); "            "; "_________________"
        PRINT #RptHandle, STRING$(79, "=")
        PRINT #RptHandle, ""
        PRINT #RptHandle, "  Customer SS# : _______________________   Spouse SS# ___________________"
        PRINT #RptHandle, ""
        PRINT #RptHandle, "Spouse Employer: _______________________   Res. Phone # _________________"
        PRINT #RptHandle, ""
        PRINT #RptHandle, "Customer Signature: ________________________________________"
        PRINT #RptHandle, "==============================================================================="
        PRINT #RptHandle, "Machinery & Fixtures:______________________"
        PRINT #RptHandle, "            Supplies:______________________"
        PRINT #RptHandle, "      Farm Equipment:______________________"
        PRINT #RptHandle, "      Boats & Motors:______________________"
        PRINT #RptHandle, "        Mobile Homes:______________________"
        PRINT #RptHandle, "             Campers:______________________"
        PRINT #RptHandle, ""
        PRINT #RptHandle, "     TOTAL PERS PROP:______________________"
        PRINT #RptHandle, ""
        PRINT #RptHandle, "     TOTAL REAL PROP:______________________"
        PRINT #RptHandle, ""
        PRINT #RptHandle, " Less Age/Disb. Excl:______________________"
        PRINT #RptHandle, ""
        PRINT #RptHandle, " TOTAL TAXABLE VALUE:______________________"
        PRINT #RptHandle, CHR$(12);
      END IF
      
    END IF
    ShowPctComp Cnt, NumOfTaxRecs
  NEXT
  
  PRINT #RptHandle, FF$
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
  
  
  
  IF parcelrecord = 0 THEN
    RETURN
  END IF
  
  GET 1, parcelrecord
  PRINT #10, "            ---------------------------------------------------------------"
  PRINT #10, "      PIN#: "; PIN$
  PRINT #10, "            ---------------------------------------------"
  PRINT #10, "      MAP#: "; pmap$;
  IF showparcel$ = "Y" THEN
    PRINT #10, TAB(50); "Real Value: "; USING "$$######,#"; CVD(pland$) + CVD(pbldg$)
  ELSE
    PRINT #10, TAB(50); "Real Value: "; "_________________"
  END IF
  PRINT #10, "            ---------------------------------------------"
  PRINT #10, "      Desc: "; pdesc$
  PRINT #10, "            ---------------------------------------------------------------"
  PRINT #10, "      Subd: "; psub$
  PRINT #10, "            ---------------------------------------------------------------"
  PRINT #10, "Acres/Lots:"
  PRINT #10, "            ---------------------------------------------------------------"
  PRINT #10, STRING$(79, "=")
  
AbHeading:
  TxFile = FREEFILE
  OPEN "TAXSETUP.DAT" FOR RANDOM AS #TxFile LEN = LEN(TaxSetUpRec)
  IF LOF(TxFile) > 0 THEN
    GET TxFile, 1, TaxSetUpRec
    TownName$ = TaxSetUpRec.Name
  ELSE
    TownName$ = ""
  END IF
  CLOSE TxFile
  PRINT #RptHandle, "Abstract Listing : "; TownName$; TAB(60); "Tax Year "; TaxYear$
  PRINT #RptHandle, ""
  PRINT #RptHandle, "Account of ..."
  PRINT #RptHandle, LEFT$(NME$, 40); TAB(55); "SS# "; TaxCustRec(1).CSSN
  PRINT #RptHandle, TaxCustRec(1).Addr1
  PRINT #RptHandle, TaxCustRec(1).Addr2
  PRINT #RptHandle, RTRIM$(TaxCustRec(1).City); " "; TaxCustRec(1).State; " "; TaxCustRec(1).Zip;
  PRINT #RptHandle, TAB(60); "Acct #: "; CustRecNo
  PRINT #RptHandle, STRING$(79, "=")
  RETURN
  
GetNameIndex2:
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

SUB AbtractListing
  
  REDIM MChoice$(1 TO 4)
  MChoice$(1) = "Personal Abstract Listing"
  MChoice$(2) = "Property Abstract Listing"
  MChoice$(3) = "Exit to OS"
  
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
  Row = ((25 - (UBOUND(MChoice$))) \ 2)
  Col = ((80 - MaxLen) \ 2) - 1
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    
    ClearBack
    
    TitleBox 2, Col, MaxLen + 3, "Tax Abstract Reports Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      
    CASE 2
    CASE 3
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP
  
END SUB

SUB BalanceListing
  SHARED Choice$()
  DIM Balance#(100), Year(100), YR(100), GBalance#(100), IBalance#(100), ABalance#(100)
  REDIM PropRec(1)  AS PropertyRecType

  REDIM Array(1 TO 1) AS Struct 'Template for the sort array
  ReportFile$ = "TaxBal.PRN"    'Report File Name
  Dash80$ = STRING$(80, "=")
  FF$ = CHR$(12)
  
  MaxLines = 53
  LineCnt = 0
  CustCnt = 0
  
  LibName$ = "TAX"
  ScrnName$ = "BALRPT"
  
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
  Choice$(0, 1) = "2"
  Choice$(1, 1) = "Summary"
  Choice$(2, 1) = "Detail"
  Choice$(0, 2) = "4"
  Choice$(1, 2) = "Screen"
  Choice$(2, 2) = "Printer"
  Form$(3, 0) = "ALL"
  
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
      IF LEFT$(Form$(2, 0), 1) = "D" THEN
        DetailFlag = True
      ELSE
        DetailFlag = False
      END IF
      TaxYear$ = RTRIM$(Form$(3, 0))
      DevSpec$ = LEFT$(Form$(4, 0), 1)
      ExitFlag = True
    CASE EscKey
      AbortFlag = True
      ExitFlag = True           'EXIT DO
    END SELECT
  LOOP UNTIL ExitFlag
  
  IF AbortFlag THEN EXIT SUB
  
  RptHandle = FREEFILE
  
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  GOSUB PrintBalanceRptHeader
  
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile
  
  IF UsingIndex AND NumOfTaxRecs > 0 THEN
    GOSUB GetNameIndexBal
  END IF
  
  ClearBack
  ShowProcessingScrn "Master Balance Listing"
  
  FOR Cnt& = 1 TO NumOfTaxRecs
    IF UsingIndex THEN
      CustRecNo = Array(Cnt&).RecNum
    ELSE
      CustRecNo = Cnt&
    END IF
    
    GET TaxFile, CustRecNo, TaxCustRec(1)
    
    IF NOT TaxCustRec(1).Deleted THEN
      IF LineCnt >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintBalanceRptHeader
      END IF
      'Check to See if Balance on File
      GOSUB CheckBalance
      
      IF Balance# <> 0 THEN
        IF DetailFlag THEN
          'Detail Format Print Each Line
          'Get Name First
          NME$ = QPTrim$(TaxCustRec(1).FName) + " " + QPTrim$(TaxCustRec(1).LName)
          NME$ = QPTrim$(NME$)  'this one cleans up those with only last name
          Balance# = 0
          'Open the Trans File
          TransFile = FREEFILE
          OPEN "TaxTrans.dat" FOR RANDOM AS TransFile LEN = LEN(TaxTrans(1))
          TransRecord& = TaxCustRec(1).LastTrans
          WHILE TransRecord& <> 0
            GET TransFile, TransRecord&, TaxTrans(1)
            IF TaxTrans(1).TranType = 1 AND (TaxYear$ = "ALL" OR VAL(TaxYear$) = TaxTrans(1).TaxYear) THEN
              Balance# = TaxTrans(1).Revenue.Principle1 + TaxTrans(1).Revenue.Principle2 + TaxTrans(1).Revenue.Principle3 + TaxTrans(1).Revenue.Principle4 + TaxTrans(1).Revenue.Principle5
              Balance# = Balance# + TaxTrans(1).Revenue.Interest + TaxTrans(1).Revenue.PENALTY + TaxTrans(1).Revenue.Collection
              Balance# = Balance# - (TaxTrans(1).Revenue.Principle1Pd + TaxTrans(1).Revenue.Principle2Pd + TaxTrans(1).Revenue.Principle3Pd + TaxTrans(1).Revenue.Principle4Pd + TaxTrans(1).Revenue.Principle5Pd)
              Balance# = Balance# - (TaxTrans(1).Revenue.InterestPD + TaxTrans(1).Revenue.PenaltyPd + TaxTrans(1).Revenue.CollectionPd)
              Balance# = Round#(Balance#)
              IF Balance# > 0 THEN
                PRINT #RptHandle, USING "#####"; CustRecNo;
                PRINT #RptHandle, TAB(10); LEFT$(NME$, 33);
                PRINT #RptHandle, TAB(45); LEFT$(TaxTrans(1).Description, 16);
                PRINT #RptHandle, TAB(61); TaxTrans(1).TaxYear;
                PRINT #RptHandle, TAB(68); USING "$$######,#.##"; Balance#
                PRINT #RptHandle, TAB(10); "Tax: "; USING "######,#.##"; Round#((TaxTrans(1).Revenue.Principle1 - TaxTrans(1).Revenue.Principle1Pd));
                PRINT #RptHandle, "  ";
                PRINT #RptHandle, "Int't: "; USING "######,#.##"; Round#((TaxTrans(1).Revenue.Interest - TaxTrans(1).Revenue.InterestPD));
                PRINT #RptHandle, "  ";
                PRINT #RptHandle, "Ad/Col: "; USING "######,#.##"; Round#((TaxTrans(1).Revenue.Collection - TaxTrans(1).Revenue.CollectionPd))

                PropPin$ = "N/A"
                IF TaxCustRec(1).FirstPropRec > 0 THEN
                  GET #PropTaxFile, TaxCustRec(1).FirstPropRec, PropRec(1)
                  PropPin$ = PropRec(1).RealPin
                END IF
                PRINT #RptHandle, TAB(10); "PIN: "; PropPin$
                TotalBalance# = TotalBalance# + Balance#
                InterestBalance# = InterestBalance# + Round#((TaxTrans(1).Revenue.Interest - TaxTrans(1).Revenue.InterestPD))
                AdBalance# = AdBalance# + Round#((TaxTrans(1).Revenue.Collection - TaxTrans(1).Revenue.CollectionPd))
                LineCnt = LineCnt + 3
                CustCnt = CustCnt + 1
                
                IF NCnt = 0 THEN
                  NCnt = 1
                  YR(NCnt) = TaxTrans(1).TaxYear
                  GBalance#(NCnt) = Balance#
                  IBalance#(NCnt) = Round#((TaxTrans(1).Revenue.Interest - TaxTrans(1).Revenue.InterestPD))
                  ABalance#(NCnt) = Round#((TaxTrans(1).Revenue.Collection - TaxTrans(1).Revenue.CollectionPd))
                  
                  GOTO skipem1
                END IF
                
                FOR LCnt = 1 TO NCnt
                  UpFlag = 0    'Set Flag to Not Updated
                  IF YR(LCnt) = TaxTrans(1).TaxYear THEN
                    GBalance#(LCnt) = GBalance#(LCnt) + Balance#
                    IBalance#(LCnt) = IBalance#(LCnt) + Round#((TaxTrans(1).Revenue.Interest - TaxTrans(1).Revenue.InterestPD))
                    ABalance#(LCnt) = ABalance#(LCnt) + Round#((TaxTrans(1).Revenue.Collection - TaxTrans(1).Revenue.CollectionPd))
                    GOTO skipem1
                  END IF
                NEXT LCnt
                NCnt = NCnt + 1
                YCnt = NCnt
                GBalance#(NCnt) = Balance#
                IBalance#(NCnt) = Round#((TaxTrans(1).Revenue.Interest - TaxTrans(1).Revenue.InterestPD))
                ABalance#(NCnt) = Round#((TaxTrans(1).Revenue.Collection - TaxTrans(1).Revenue.CollectionPd))
                YR(NCnt) = TaxTrans(1).TaxYear
skipem1:
              END IF
            END IF
            TransRecord& = TaxTrans(1).LastTrans
          WEND
          CLOSE TransFile
        ELSE
          'Summary Format
          'Get Name First
          NME$ = QPTrim$(TaxCustRec(1).FName) + " " + QPTrim$(TaxCustRec(1).LName)
          NME$ = QPTrim$(NME$)  'this one cleans up those with only last name
          
          Balance# = 0
          'Open the Trans File
          TransFile = FREEFILE
          OPEN "TaxTrans.dat" FOR RANDOM AS TransFile LEN = LEN(TaxTrans(1))
          TransRecord& = TaxCustRec(1).LastTrans
          
          WHILE TransRecord& <> 0
            GET TransFile, TransRecord&, TaxTrans(1)
            
            IF TaxTrans(1).TranType = 1 AND (TaxYear$ = "ALL" OR VAL(TaxYear$) = TaxTrans(1).TaxYear) THEN
              Balance# = 0
              Balance# = TaxTrans(1).Revenue.Principle1 + TaxTrans(1).Revenue.Principle2 + TaxTrans(1).Revenue.Principle3 + TaxTrans(1).Revenue.Principle4 + TaxTrans(1).Revenue.Principle5
              Balance# = Balance# + TaxTrans(1).Revenue.Interest + TaxTrans(1).Revenue.PENALTY + TaxTrans(1).Revenue.Collection
              Balance# = Balance# - (TaxTrans(1).Revenue.Principle1Pd + TaxTrans(1).Revenue.Principle2Pd + TaxTrans(1).Revenue.Principle3Pd + TaxTrans(1).Revenue.Principle4Pd + TaxTrans(1).Revenue.Principle5Pd)
              Balance# = Balance# - (TaxTrans(1).Revenue.InterestPD + TaxTrans(1).Revenue.PenaltyPd + TaxTrans(1).Revenue.CollectionPd)
              Balance# = Round#(Balance#)
            ELSE
              Balance# = 0
            END IF
            
            'IF Balance# = .01# THEN
            'unrem
            IF Balance# > 0 THEN
              IF NCnt = 0 THEN
                NCnt = 1
                Year(NCnt) = TaxTrans(1).TaxYear
                Balance#(NCnt) = Balance#
                GOTO skipem
              END IF
              
              FOR LCnt = 1 TO NCnt
                UpFlag = 0      'Set Flag to Not Updated
                IF Year(LCnt) = TaxTrans(1).TaxYear THEN
                  Balance#(LCnt) = Balance#(LCnt) + Balance#
                  GOTO skipem
                END IF
              NEXT LCnt
              NCnt = NCnt + 1
              Balance#(NCnt) = Balance#
              Year(NCnt) = TaxTrans(1).TaxYear
            END IF
skipem:
            TransRecord& = TaxTrans(1).LastTrans
          WEND
          CLOSE TransFile
          
          FOR PCnt = 1 TO NCnt
            IF Balance#(PCnt) > 0 THEN
              PRINT #RptHandle, USING "#####"; CustRecNo;
              PRINT #RptHandle, TAB(10); LEFT$(NME$, 33);
              PRINT #RptHandle, TAB(61); Year(PCnt);
              PRINT #RptHandle, TAB(68); USING "$$######,#.##"; Balance#(PCnt)
              TotalBalance# = TotalBalance# + Balance#(PCnt)
              
              'Update Year Balance
              Updated = 0
              Year = Year(PCnt)
              IF YCnt = 0 THEN
                YCnt = 1
                YR(YCnt) = Year
                GBalance#(YCnt) = GBalance#(YCnt) + Balance#(PCnt)
              ELSE
                FOR TCnt = 1 TO YCnt
                  IF Year = YR(TCnt) THEN
                    GBalance#(TCnt) = GBalance#(TCnt) + Balance#(PCnt)
                    Updated = 1
                  END IF
                NEXT TCnt
                IF Updated = 0 THEN
                  YCnt = YCnt + 1
                  YR(YCnt) = Year
                  GBalance#(YCnt) = Balance#(PCnt)
                END IF
              END IF
              LineCnt = LineCnt + 1
              CustCnt = CustCnt + 1
              Year(PCnt) = 0
              NCnt = 0
            END IF
          NEXT PCnt
          FOR ClearCnt = 1 TO 99
            Balance#(ClearCnt) = 0
          NEXT ClearCnt
        END IF  'End for Balance Check
      END IF    'End for Detail Cust
    END IF      'End of Delete Cust
    ShowPctComp Cnt&, NumOfTaxRecs
    
  NEXT Cnt&
  
  GOSUB PrintBalanceRptEnding
  
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
  
PrintBalanceRptHeader:
  PAGE = PAGE + 1
  PRINT #RptHandle, TAB(21); "Property Tax Customer Balance Listing"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; PAGE
  IF NOT DetailFlag THEN
    PRINT #RptHandle, "Summary Format"
  ELSE
    PRINT #RptHandle, "Detail Format"
  END IF
  PRINT #RptHandle, "Acct #"; TAB(10); "Name"; TAB(45); "Description"; TAB(60); "Tax Year"; TAB(71); "Balance"
  PRINT #RptHandle, Dash80$
  LineCnt = 5
  RETURN
  
PrintBalanceRptEnding:
  GOSUB SORTTHEM
  
  PRINT #RptHandle, Dash80$
  PRINT #RptHandle, "Total Customer Lines Printed: "; USING "#####"; CustCnt
  PRINT #RptHandle, "      Total Tax Balance: "; USING "$$########,#.##"; TotalBalance#
  IF DetailFlag THEN
    PRINT #RptHandle, "    Total Tax Principle: "; USING "$$########,#.##"; TotalBalance# - (InterestBalance# + AdBalance#)
    PRINT #RptHandle, "         Total Interest: "; USING "$$########,#.##"; InterestBalance#
    PRINT #RptHandle, "      Total Ad/Col Cost: "; USING "$$########,#.##"; AdBalance#
    PRINT #RptHandle, FF$
    PRINT #RptHandle, "Tax Totals By Year"
    PRINT #RptHandle, "Tax Year"; TAB(15); "Tax Amount"; TAB(33); "Interest"; TAB(50); "Ad/Collection"; TAB(70); "Total Due"
    PRINT #RptHandle, STRING$(79, "-")
    FOR lll = 1 TO YCnt
      PRINT #RptHandle, YR(lll); TAB(12); USING "$$#####,#.##"; GBalance#(lll) - (IBalance#(lll) + ABalance#(lll));
      PRINT #RptHandle, TAB(30); USING "$$####,#.##"; IBalance#(lll);
      PRINT #RptHandle, TAB(50); USING "$$####,#.##"; ABalance#(lll);
      PRINT #RptHandle, TAB(65); USING "$$#######,#.##"; GBalance#(lll)
    NEXT lll
    RETURN
  END IF
  PRINT #RptHandle, FF$
  PRINT #RptHandle, "Tax Totals By Year"
  PRINT #RptHandle, "Tax Year"; TAB(15); "Amount Owed"
  PRINT #RptHandle, STRING$(40, "-")
  FOR lll = 1 TO YCnt
    PRINT #RptHandle, YR(lll); TAB(15); USING "$$#######,#.##"; GBalance#(lll)
  NEXT lll
  PRINT #RptHandle, FF$
  RETURN
  
GetNameIndexBal:
  REDIM Array(1 TO NumOfTaxRecs) AS Struct
  FOR Cnt = 1 TO NumOfTaxRecs
    GET TaxFile, Cnt, TaxCustRec(1)
    Array(Cnt).who = UCASE$(TaxCustRec(1).SName) + " "
    Array(Cnt).RecNum = Cnt
  NEXT
  
  'Sort Them Here
  SortT Array(1), NumOfTaxRecs, 0, LEN(Array(1)), 0, 14
  RETURN
  
CheckBalance:
  IF TaxCustRec(1).LastTrans = 0 THEN Balance# = 0: RETURN
  'Open the File and Look For A Balance
  TransFile = FREEFILE
  OPEN "TaxTrans.dat" FOR RANDOM AS TransFile LEN = LEN(TaxTrans(1))
  TransRecord& = TaxCustRec(1).LastTrans
  WHILE TransRecord& <> 0
    GET TransFile, TransRecord&, TaxTrans(1)
    IF TaxTrans(1).TranType = 1 THEN
      Balance# = TaxTrans(1).Revenue.Principle1 + TaxTrans(1).Revenue.Principle2 + TaxTrans(1).Revenue.Principle3 + TaxTrans(1).Revenue.Principle4 + TaxTrans(1).Revenue.Principle5
      Balance# = Balance# + TaxTrans(1).Revenue.Interest + TaxTrans(1).Revenue.PENALTY + TaxTrans(1).Revenue.Collection
      Balance# = Balance# - (TaxTrans(1).Revenue.Principle1Pd + TaxTrans(1).Revenue.Principle2Pd + TaxTrans(1).Revenue.Principle3Pd + TaxTrans(1).Revenue.Principle4Pd + TaxTrans(1).Revenue.Principle5Pd)
      Balance# = Balance# - (TaxTrans(1).Revenue.InterestPD + TaxTrans(1).Revenue.PenaltyPd + TaxTrans(1).Revenue.CollectionPd)
      Balance# = Round#(Balance#)
      IF Balance# > 0 THEN CLOSE TransFile: RETURN
    END IF
    TransRecord& = TaxTrans(1).LastTrans
  WEND
  Balance# = 0
  CLOSE TransFile
  RETURN
  
SORTTHEM:
26000         REM sort
  Count = YCnt
26020 M = Count
26030 M = INT(M / 2)
26040 IF M = 0 THEN 26190
26050 FOR st = 1 TO M
26060 i = st
26070 j = st + M
26080 SW = 0
26090 IF YR(i) <= YR(j) THEN 26120
26100 SW = 1
26110 SWAP YR(i), YR(j)
  SWAP GBalance#(i), GBalance#(j)
  SWAP IBalance#(i), IBalance#(j)
  SWAP ABalance#(i), ABalance#(j)
26120 i = j
26130 j = j + M
26140 IF j <= Count THEN 26090
26150 IF SW = 0 THEN 26170
26160 GOTO 26060
26170 NEXT st
26180 GOTO 26030
26190 RETURN
  
  
END SUB

SUB CustomerInquiry
  
  SHARED Choice$()
  
  REDIM TaxInq(1) AS TaxCustInqType
  
  ClearBack
  LookUp RecNo&, "Customer Inquiry", False, True, False
  CustAcct& = RecNo&
  IF RecNo& <= 0 THEN
    GOTO CustInqExit
  END IF
  
  FirstTime = True
  
  LibName$ = "TAX"
  ScrnName$ = "TCUSTINQ"
  
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
  ClearBack
  
  ShowCursor
  
  DisplayTaxScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      FirstTime = False
      GOSUB LoadInqData
      Action = 1
    END IF
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F4KEY
      IF CustAcct& > 0 THEN
        ShowCustHistory -CustAcct&              'DON'T CHANGE THIS
        Action = 1
      END IF
    CASE EscKey
      ExitFlag = True           'EXIT DO
    END SELECT
  LOOP UNTIL ExitFlag
  
  
CustInqExit:
  EXIT SUB
  
LoadInqData:
  RealValue# = 0
  SeniExmp# = 0
  OthrExmp# = 0
  PersValue# = 0
  MOBHValue# = 0
  MERHValue# = 0
  FarmValue# = 0
  MachValue# = 0
  
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  GET TaxFile, RecNo&, TaxCustRec(1)
  CLOSE TaxFile
  
  REDIM PersRecs(0) AS LONG
  REDIM PropRecs(0) AS LONG
  
  GetPersRecList PersRecs(), RecNo&
  GetPropRecList PropRecs(), RecNo&
  
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  IF PropRecs(0) > 0 THEN
    FOR Cnt& = 1 TO PropRecs(0)
      GET PropTaxFile, PropRecs(Cnt&), PropertyRec(1)
      RealValue# = Round#(RealValue# + PropertyRec(1).PROPVALU)
      SeniExmp# = Round#(SeniExmp# + PropertyRec(1).EXMPSENI)
      OthrExmp# = Round#(OthrExmp# + PropertyRec(1).EXMPOTHR)
    NEXT
  END IF
  CLOSE PropTaxFile
  
  OpenTaxPersFile NumOfPersRecs, PersTaxFile
  IF PersRecs(0) > 0 THEN
    FOR Cnt& = 1 TO PersRecs(0)
      GET PersTaxFile, PersRecs(Cnt&), PersRec(1)
      PersValue# = Round#(PersValue# + PersRec(1).PERSVAL)
      MOBHValue# = Round#(MOBHValue# + PersRec(1).MHVALUE)
      MERHValue# = Round#(MERHValue# + PersRec(1).MCVALUE)
      FarmValue# = Round#(FarmValue# + PersRec(1).CVALUE)
      MachValue# = Round#(MachValue# + PersRec(1).MTVALUE)
      SeniExmp# = Round#(SeniExmp# + PersRec(1).EXMPSENI)
      OthrExmp# = Round#(OthrExmp# + PersRec(1).EXMPOTHR)
    NEXT
  END IF
  CLOSE PersTaxFile
  
  TaxInq(1).ACCT = RecNo&
  TaxInq(1).OPENDATE = TaxCustRec(1).OPENDATE
  TaxInq(1).CUSTNAME = QPTrim$(TaxCustRec(1).FName) + " " + QPTrim$(TaxCustRec(1).LName)
  TaxInq(1).HPHONE = TaxCustRec(1).HPHONE
  TaxInq(1).CSSN = TaxCustRec(1).CSSN
  TaxInq(1).WPHONE = TaxCustRec(1).WPHONE
  TaxInq(1).Addr1 = TaxCustRec(1).Addr1
  TaxInq(1).Addr2 = TaxCustRec(1).Addr2
  TaxInq(1).City = TaxCustRec(1).City
  TaxInq(1).State = TaxCustRec(1).State
  TaxInq(1).Zip = TaxCustRec(1).Zip
  TaxInq(1).Active = TaxCustRec(1).Active
  TaxInq(1).Interest = TaxCustRec(1).Interest
  TaxInq(1).EXEMPT = TaxCustRec(1).TaxExempt
  TaxInq(1).PENALTY = TaxCustRec(1).PENALTY
  TaxInq(1).SRCITDIS = SeniExmp#
  TaxInq(1).ODISCOUN = OthrExmp#
  TaxInq(1).PROPVALU = RealValue#
  TaxInq(1).PERSVAL = PersValue#
  TaxInq(1).MHVALUE = MOBHValue#
  TaxInq(1).MCVALUE = MERHValue#
  TaxInq(1).CVALUE = FarmValue#
  TaxInq(1).MTVALUE = MachValue#
  TaxInq(1).Balance = GetCustBalance#(CustAcct&)
  
  BCopy VARSEG(TaxInq(1)), VARPTR(TaxInq(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), LEN(Form$(0, 0)), 0
  UnPackBuffer 0, 0, Form$(), Fld()
  Action = 1
  RETURN
END SUB

SUB CustomerListing
  
  SHARED Choice$()
  
  REDIM Array(1 TO 1) AS Struct          'Template for the sort array
  REDIM PArray(1 TO 1) AS ParcelSortType 'Template for the sort array
  ReportFile$ = "TaxCust.PRN"   'Report File Name
  Dash80$ = STRING$(80, "=")
  FF$ = CHR$(12)

  TxFile = FREEFILE
  OPEN "TAXSETUP.DAT" FOR RANDOM AS #TxFile LEN = LEN(TaxSetUpRec)
  IF LOF(TxFile) > 0 THEN
    GET TxFile, 1, TaxSetUpRec
    FirstTime = True
  END IF
  CLOSE
  
  IF INSTR(TaxSetUpRec.Name, "SUNSET") > 0 THEN
    SunBeach = True
  END IF

  LineCnt = 0
  CustCnt = 0
  
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
  
  REDIM Choice$(0 TO 3, 0 TO 3)
  
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Name Order"
  Choice$(2, 0) = "Account Number"
  IF SunBeach THEN
    Choice$(3, 0) = "Parcel Number"
  END IF

  Choice$(0, 1) = "2"
  Choice$(1, 1) = "Summary"
  Choice$(2, 1) = "Detail"
  
  Choice$(0, 2) = "3"
  Choice$(1, 2) = "Both"
  Choice$(2, 2) = "Active"
  Choice$(3, 2) = "Inactive"
  
  Choice$(0, 3) = "4"
  Choice$(1, 3) = "Screen"
  Choice$(2, 3) = "Printer"
  
  Action = 1
  ClearBack
  
  ShowCursor
  
  DisplayTaxScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      SELECT CASE LEFT$(Form$(1, 0), 1)
      CASE "N"
        UsingIndex = True
      CASE "P"
        UsingParcel = True
      CASE ELSE
        UsingIndex = False
      END SELECT

      IF LEFT$(Form$(2, 0), 1) = "D" THEN
        DetailFlag = True
        MaxLines = 45
      ELSE
        DetailFlag = False
        MaxLines = 50
      END IF
      
      AStat$ = QPTrim$(LEFT$(Form$(3, 0), 1))
      SELECT CASE AStat$
      CASE "B"
        StatFlag = -1
      CASE "A"
        StatFlag = 1
      CASE "I"
        StatFlag = 0
      END SELECT
      
      DevSpec$ = LEFT$(Form$(4, 0), 1)
      ExitFlag = True
    CASE EscKey
      AbortFlag = True
      ExitFlag = True           'EXIT DO
    END SELECT
  LOOP UNTIL ExitFlag
  
  IF AbortFlag THEN EXIT SUB
  
  RptHandle = FREEFILE
  
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  GOSUB PrintDetailCustomerRptHeader
  
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile
  
  IF UsingIndex AND NumOfTaxRecs > 0 THEN
    GOSUB GetNameIndex
  ELSEIF UsingParcel THEN
    GOSUB GetParcelIndex
  END IF
  

  ClearBack
  ShowProcessingScrn "Master Customer Listing"

  IF UsingIndex THEN
    MaxRecs = NumOfTaxRecs
  ELSEIF UsingParcel THEN
    MaxRecs = PCnt
    DetailFlag = True
    StatFlag = True
  ELSE
    MaxRecs = NumOfTaxRecs
  END IF
  
  'GOTO Jump2Here

  FOR Cnt = 1 TO MaxRecs
    IF UsingIndex THEN
      CustRecNo = Array(Cnt).RecNum
    ELSEIF UsingParcel THEN
      CustRecNo = PArray(Cnt).CustRec
    ELSE
      CustRecNo = Cnt
    END IF
    
    GET TaxFile, CustRecNo, TaxCustRec(1)
    
    '    IF TaxCustRec(1).Deleted = 0 THEN
    '    ELSE
    '      TaxCustRec(1).Deleted = -1
    '      PUT TaxFile, CustRecNo, TaxCustRec(1)
    '    END IF
    
    IF TaxCustRec(1).Deleted = 0 THEN
      SELECT CASE TaxCustRec(1).Active
      CASE "Y"
        IF StatFlag = 0 THEN
          GOTO SkipCust1
        END IF
      CASE ELSE
        IF StatFlag = 1 THEN
          GOTO SkipCust1
        END IF
      END SELECT
      
      IF LineCnt >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintDetailCustomerRptHeader
      END IF
      
      NME$ = QPTrim$(TaxCustRec(1).FName) + " " + QPTrim$(TaxCustRec(1).LName)
      NME$ = QPTrim$(NME$)      'this one cleans up those with only last name
      
      IF NOT DetailFlag THEN
        PRINT #RptHandle, USING "#####"; CustRecNo;
        PRINT #RptHandle, TAB(10); NME$; TAB(60); TaxCustRec(1).Active; TAB(65); TaxCustRec(1).PIN
        LineCnt = LineCnt + 1
      ELSE
        PRINT #RptHandle, USING "  Acct: ####"; CustRecNo;
        PRINT #RptHandle, TAB(15); NME$
        PRINT #RptHandle, "Active: "; TaxCustRec(1).Active; TAB(15); TaxCustRec(1).Addr1
        PRINT #RptHandle, "Int'st: "; TaxCustRec(1).Interest; TAB(15); TaxCustRec(1).Addr2
        PRINT #RptHandle, "Exempt: "; TaxCustRec(1).TaxExempt; TAB(15); RTRIM$(TaxCustRec(1).City) + ", "; RTRIM$(TaxCustRec(1).State) + "  " + RTRIM$(TaxCustRec(1).Zip)
        
        PRINT #RptHandle, ""
        LineCnt = LineCnt + 5
        
        'Now Show Property Records Next
        
        IF UsingParcel THEN
          IF LineCnt >= MaxLines THEN
            PRINT #RptHandle, FF$
            GOSUB PrintDetailCustomerRptHeader
            PRINT #RptHandle, USING "  Acct: ####"; CustRecNo;
            PRINT #RptHandle, TAB(15); NME$
            LineCnt = LineCnt + 2
          END IF
          PRINT #RptHandle, "Property..."
          GET #PropTaxFile, PArray(Cnt).ParcelRec, PropertyRec(1)
          ThisPin$ = QPTrim$(PropertyRec(1).RealPin)
          IF LEN(ThisPin$) = 0 THEN
            ThisPin$ = "  *BLANK!"
          END IF
          PRINT #RptHandle, "Pin # "; ThisPin$; TAB(50); " Value: "; USING "$$########,#.##"; PropertyRec(1).PROPVALU
          PRINT #RptHandle, "Desc  "; QPTrim$(PropertyRec(1).PROPNOT1); TAB(50); "Exempt: "; USING "$$########,#.##"; PropertyRec(1).EXMPSENI
          PRINT #RptHandle, "Desc  "; QPTrim$(PropertyRec(1).PROPNOT2); TAB(50); "Oth Ex: "; ; USING "$$########,#.##"; PropertyRec(1).EXMPOTHR
          PRINT #RptHandle, "Desc  "; QPTrim$(PropertyRec(1).PROPNOT3); TAB(50); "Mortgage Code: "; QPTrim$(PropertyRec(1).MortCode)
          PRINT #RptHandle, "Map/Blk/Lot - "; QPTrim$(PropertyRec(1).MAP); "/"; QPTrim$(PropertyRec(1).BLOCK); "/"; QPTrim$(PropertyRec(1).LOTNUMB); TAB(40); "Late (Y/N): "; PropertyRec(1).LATELIST
          PRINT #RptHandle, STRING$(79, "-")
          LineCnt = LineCnt + 6

        ELSE
          IF TaxCustRec(1).FirstPropRec > 0 THEN
            PropertyRecord! = TaxCustRec(1).FirstPropRec
            WHILE PropertyRecord! <> 0
              IF LineCnt >= MaxLines THEN
                PRINT #RptHandle, FF$
                GOSUB PrintDetailCustomerRptHeader
                PRINT #RptHandle, USING "  Acct: ####"; CustRecNo;
                PRINT #RptHandle, TAB(15); NME$
                LineCnt = LineCnt + 2
              END IF
              PRINT #RptHandle, "Property Owned..."
              GET #PropTaxFile, PropertyRecord!, PropertyRec(1)
              'Reset Property Link
              'IF PropertyRec(1).CustPin <> CustRecNo THEN STOP
              PropertyRec(1).CustPin = CustRecNo
              PUT #PropTaxFile, PropertyRecord!, PropertyRec(1)
              PRINT #RptHandle, "Pin # "; QPTrim$(PropertyRec(1).RealPin); TAB(50); " Value: "; USING "$$########,#.##"; PropertyRec(1).PROPVALU
              PRINT #RptHandle, "Desc  "; QPTrim$(PropertyRec(1).PROPNOT1); TAB(50); "Exempt: "; USING "$$########,#.##"; PropertyRec(1).EXMPSENI
              PRINT #RptHandle, "Desc  "; QPTrim$(PropertyRec(1).PROPNOT2); TAB(50); "Oth Ex: "; ; USING "$$########,#.##"; PropertyRec(1).EXMPOTHR
              PRINT #RptHandle, "Desc  "; QPTrim$(PropertyRec(1).PROPNOT3); TAB(50); "Mortgage Code: "; QPTrim$(PropertyRec(1).MortCode)
              PRINT #RptHandle, "Map/Blk/Lot - "; QPTrim$(PropertyRec(1).MAP); "/"; QPTrim$(PropertyRec(1).BLOCK); "/"; QPTrim$(PropertyRec(1).LOTNUMB); TAB(40); "Late (Y/N): "; PropertyRec(1).LATELIST
              PRINT #RptHandle, STRING$(79, "-")
              LineCnt = LineCnt + 6
              OldRecord! = PropertyRecord!
              PropertyRecord! = PropertyRec(1).NextRec
              IF OldRecord! = PropertyRecord! THEN PropertyRecord! = 0
            WEND
          END IF
        END IF
        'IF UsingParcel = 0 THEN
          'NOW CHECK PERSONAL PROPERTY
          IF TaxCustRec(1).FirstPersRec > 0 THEN
            PropertyRecord! = TaxCustRec(1).FirstPersRec
            WHILE PropertyRecord! <> 0
              IF LineCnt >= MaxLines THEN
                PRINT #RptHandle, FF$
                GOSUB PrintDetailCustomerRptHeader
                PRINT #RptHandle, USING "  Acct: ####"; CustRecNo;
                PRINT #RptHandle, TAB(15); NME$
                LineCnt = LineCnt + 2
              END IF
              PRINT #RptHandle, "Personal Property Owned..."
              GET #PersTaxFile, PropertyRecord!, PersRec(1)
              'IF PersRec(1).CustPin <> CustRecNo THEN STOP
              PersRec(1).CustPin = CustRecNo
              PUT #PersTaxFile, PropertyRecord!, PersRec(1)
              PRINT #RptHandle, "Pin # "; QPTrim$(PersRec(1).PropPin); TAB(50); " PP Value: "; USING "$$########,#.##"; PersRec(1).PERSVAL
              PRINT #RptHandle, "Desc  "; QPTrim$(PersRec(1).DESC1); TAB(50); " MH Value: "; USING "$$########,#.##"; PersRec(1).MHVALUE
              PRINT #RptHandle, "Desc  "; QPTrim$(PersRec(1).DESC2); TAB(50); " MC Value: "; USING "$$########,#.##"; PersRec(1).MCVALUE
              PRINT #RptHandle, "Desc  "; QPTrim$(PersRec(1).DESC3); TAB(50); " CV Value: "; USING "$$########,#.##"; PersRec(1).CVALUE
              PRINT #RptHandle, "Desc  "; QPTrim$(PersRec(1).Desc4); TAB(50); " MT Value: "; USING "$$########,#.##"; PersRec(1).MTVALUE
              PRINT #RptHandle, "Desc  "; QPTrim$(PersRec(1).Desc5); TAB(50); "Exempt: "; USING "$$########,#.##"; PersRec(1).EXMPSENI
              PRINT #RptHandle, STRING$(79, "-")
              LineCnt = LineCnt + 7
              OldRecord! = PropertyRecord!
              PropertyRecord! = PersRec(1).NextRec
              IF OldRecord! = PropertyRecord! THEN PropertyRecord! = 0
            WEND
          END IF
        'END IF
      END IF
      CustCnt = CustCnt + 1
    END IF
SkipCust1:
    ShowPctComp Cnt, NumOfTaxRecs
  NEXT
  
  GOSUB PrintDetailCustomerRptEnding
  
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
  
  
  EXIT SUB
  
PrintDetailCustomerRptHeader:
  PAGE = PAGE + 1
  PRINT #RptHandle, TAB(20); "Property Tax Detailed Customer Listing"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; PAGE
  IF NOT DetailFlag THEN
    PRINT #RptHandle, "Summary Format"
    PRINT #RptHandle, "Acct #"; TAB(10); "Name"; TAB(55); "Active"
    PRINT #RptHandle, Dash80$
    LineCnt = 3
  ELSE
    PRINT #RptHandle, "Detail Format"
    PRINT #RptHandle, Dash80$
    LineCnt = 2
  END IF
  RETURN
  
PrintDetailCustomerRptEnding:
  PRINT #RptHandle, Dash80$
  PRINT #RptHandle, "Total Customers Printed: "; USING "#####"; CustCnt
  PRINT #RptHandle,
  PRINT #RptHandle, FF$
  RETURN
  
GetNameIndex:
  REDIM Array(1 TO NumOfTaxRecs) AS Struct
  FOR Cnt = 1 TO NumOfTaxRecs
    GET TaxFile, Cnt, TaxCustRec(1)
    Array(Cnt).who = UCASE$(TaxCustRec(1).SName) + " "
    Array(Cnt).RecNum = Cnt
  NEXT
  
  'Sort Them Here
  SortT Array(1), NumOfTaxRecs, 0, LEN(Array(1)), 0, 14
  RETURN
  
  'Jump2Here:
  '  FOR Cnt = 1 TO NumOfTaxRecs
  '    CustRecNo = Cnt
  '    GET TaxFile, CustRecNo, TaxCustRec(1)
  '    TaxCustRec(1).Acct = Cnt
  '    TaxCustRec(1).Pin = Cnt
  '    PUT TaxFile, CustRecNo, TaxCustRec(1)
  '  NEXT
  'CLOSE
  END

GetParcelIndex:
  REDIM PArray(1 TO 1) AS ParcelSortType
  FOR Cnt = 1 TO NumOfPropRecs
    GET #PropTaxFile, Cnt, PropertyRec(1)
    IF (PropertyRec(1).Deleted = 0) AND (PropertyRec(1).CustPin > 0) THEN
      PCnt = PCnt + 1
      REDIM PRESERVE PArray(1 TO PCnt) AS ParcelSortType
      LSET PArray(PCnt).Parcel = PropertyRec(1).RealPin
      PArray(PCnt).ParcelRec = Cnt
      PArray(PCnt).CustRec = PropertyRec(1).CustPin
    END IF
  NEXT
  'SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
  SortT PArray(1), PCnt, 0, LEN(PArray(1)), 0, 12
RETURN


'
END SUB

SUB Labels
  
  SHARED Choice$()
  
  REDIM Array(1 TO 1) AS Struct 'Template for the sort array
  ReportFile$ = "TAXLABEL.PRN"  'Report File Name
  Dash80$ = STRING$(80, "=")
  FF$ = CHR$(12)
  
  MaxLines = 56
  LineCnt = 0
  CustCnt = 0
  
  LibName$ = "TAX"
  ScrnName$ = "VCUSTLAB"
  
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
  Choice$(0, 2) = "4"
  Choice$(1, 2) = "Screen"
  Choice$(2, 2) = "Printer"
  Form$(2, 0) = "R"             'Default to Real
  Form$(3, 0) = "N"             'Default to No
  Action = 1
  ClearBack
  
  ShowCursor
  
  DisplayTaxScrn ScrnName$
  
  DO
'
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      IF LEFT$(Form$(1, 0), 1) = "N" THEN
        UsingIndex = True
      ELSE
        UsingIndex = False
      END IF
      TaxType$ = Form$(2, 0)
      IF LEFT$(Form$(3, 0), 1) = "Y" THEN
        DetailFlag = True
      ELSE
        DetailFlag = False
      END IF
      DevSpec$ = LEFT$(Form$(4, 0), 1)
      ExitFlag = True
    CASE F5KEY
      GOSUB PrintAlign
    CASE EscKey
      AbortFlag = True
      ExitFlag = True           'EXIT DO
    END SELECT
  LOOP UNTIL ExitFlag
  
  IF AbortFlag THEN EXIT SUB
  
  RptHandle = FREEFILE
  
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile
  
  IF UsingIndex AND NumOfTaxRecs > 0 THEN
    GOSUB GetNameIndexML
  END IF
  
  ClearBack
  ShowProcessingScrn "Mailing Labels"
  
  FOR Cnt = 1 TO NumOfTaxRecs
    IF UsingIndex THEN
      CustRecNo = Array(Cnt).RecNum
    ELSE
      CustRecNo = Cnt
    END IF
    
    GET TaxFile, CustRecNo, TaxCustRec(1)
    
    IF NOT TaxCustRec(1).Deleted THEN
      'Mortcode Test Here
      IF DetailFlag AND TaxType$ = "R" THEN
        PropRec& = TaxCustRec(1).FirstPropRec
        WHILE PropRec& <> 0
          GET PropTaxFile, PropRec&, PropertyRec(1)
          MC$ = PropertyRec(1).MortCode
          MC$ = RTRIM$(MC$)
          PropRec& = PropertyRec(1).NextRec
        WEND
      END IF
      
      IF DetailFlag AND TaxType$ = "R" AND LEN(MC$) > 0 THEN
        NME$ = QPTrim$(TaxCustRec(1).FName) + " " + QPTrim$(TaxCustRec(1).LName)
        NME$ = QPTrim$(NME$)    'this one cleans up those with only last name
        PRINT #RptHandle, USING "#####"; CustRecNo;
        PRINT #RptHandle, TAB(20); "MC="; MC$
        PRINT #RptHandle, NME$
        PRINT #RptHandle, TaxCustRec(1).Addr1
        PRINT #RptHandle, TaxCustRec(1).Addr2
        PRINT #RptHandle, TaxCustRec(1).City; " "; TaxCustRec(1).State; " "; TaxCustRec(1).Zip
        PRINT #RptHandle,
      END IF
      IF NOT (DetailFlag) THEN
        NME$ = QPTrim$(TaxCustRec(1).FName) + " " + QPTrim$(TaxCustRec(1).LName)
        NME$ = QPTrim$(NME$)    'this one cleans up those with only last name
        PRINT #RptHandle, USING "#####"; CustRecNo
        PRINT #RptHandle, NME$
        PRINT #RptHandle, TaxCustRec(1).Addr1
        PRINT #RptHandle, TaxCustRec(1).Addr2
        PRINT #RptHandle, TaxCustRec(1).City; " "; TaxCustRec(1).State; " "; TaxCustRec(1).Zip
        PRINT #RptHandle,
      END IF
      CustCnt = CustCnt + 1
    END IF
    ShowPctComp Cnt, NumOfTaxRecs
  NEXT
  
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
  
  'KILL ReportFile$
  
  EXIT SUB
  
PrintAlign:
  LPRINT STRING$(35, "X")
  LPRINT STRING$(35, "X")
  LPRINT STRING$(35, "X")
  LPRINT STRING$(35, "X")
  LPRINT STRING$(35, "X")
  LPRINT
RETURN
  
  
GetNameIndexML:
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

SUB LateListing
  SHARED Choice$()
  
  REDIM Array(1 TO 1) AS Struct 'Template for the sort array
  ReportFile$ = "TaxLATE.PRN"   'Report File Name
  Dash80$ = STRING$(80, "=")
  FF$ = CHR$(12)
  
  MaxLines = 56
  LineCnt = 0
  CustCnt = 0
  
  LibName$ = "TAX"
  ScrnName$ = "LATERPT"
  
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
  Form$(2, 0) = "Detail"
  Fld(2).Protected = -1
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
      DetailFlag$ = LEFT$(Form$(2, 0), 1)
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
  
  GOSUB PrintLateRptHeader
  
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile
  
  IF UsingIndex AND NumOfTaxRecs > 0 THEN
    GOSUB GetNameIndexLate
  END IF
  
  ClearBack
  ShowProcessingScrn "Late Listing Report"
  
  FOR Cnt = 1 TO NumOfTaxRecs
    IF UsingIndex THEN
      CustRecNo = Array(Cnt).RecNum
    ELSE
      CustRecNo = Cnt
    END IF
    
    GET TaxFile, CustRecNo, TaxCustRec(1)
    
    IF NOT TaxCustRec(1).Deleted THEN
      
      ' Check Line on Page and Form Feed if Necessary
      IF LineCnt >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintLateRptHeader
      END IF
      
      NME$ = QPTrim$(TaxCustRec(1).FName) + " " + QPTrim$(TaxCustRec(1).LName)
      NME$ = QPTrim$(NME$)      'this one cleans up those with only last name
      
      
      
      
      'First Show Property Records
      
      IF TaxCustRec(1).FirstPropRec > 0 THEN
        PropertyRecord! = TaxCustRec(1).FirstPropRec
        WHILE PropertyRecord! <> 0
          GET #PropTaxFile, PropertyRecord!, PropertyRec(1)
          IF PropertyRec(1).LATELIST = "Y" THEN
            PRINT #RptHandle, USING "######"; CustRecNo;
            PRINT #RptHandle, TAB(10); NME$; TAB(57); USING "$$#######,#"; PropertyRec(1).PROPVALU
            TotalLateAmt# = TotalLateAmt# + PropertyRec(1).PROPVALU
            LineCnt = LineCnt + 1
            CustCnt = CustCnt + 1
          END IF
          PropertyRecord! = PropertyRec(1).NextRec
        WEND
      END IF
      
      'NOW CHECK PERSONAL PROPERTY
      IF TaxCustRec(1).FirstPersRec > 0 THEN
        PropertyRecord! = TaxCustRec(1).FirstPersRec
        WHILE PropertyRecord! <> 0
          GET #PersTaxFile, PropertyRecord!, PersRec(1)
          
          IF PersRec(1).LATELIST = "Y" THEN
            PValue# = PersRec(1).PERSVAL# + PersRec(1).MHVALUE + PersRec(1).MCVALUE + PersRec(1).CVALUE + PersRec(1).MTVALUE
            PRINT #RptHandle, USING "######"; CustRecNo;
            PRINT #RptHandle, TAB(10); NME$; TAB(57); USING "$$#######,#"; PValue#
            TotalLateAmt# = TotalLateAmt# + PValue#
            PValue# = 0
            LineCnt = LineCnt + 1
            CustCnt = CustCnt + 1
          END IF
          PropertyRecord! = PersRec(1).NextRec
        WEND
      END IF
      
    END IF
    ShowPctComp Cnt, NumOfTaxRecs
  NEXT
  
  GOSUB PrintLateRptEnding
  
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
  
PrintLateRptHeader:
  PAGE = PAGE + 1
  PRINT #RptHandle, TAB(24); "Tax Customer Late Listing Report"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; PAGE
  PRINT #RptHandle, ""
  PRINT #RptHandle, "Acct #"; TAB(10); "Name"; TAB(55); "Property Value"
  PRINT #RptHandle, Dash80$
  LineCnt = 5
  RETURN
  
PrintLateRptEnding:
  PRINT #RptHandle, Dash80$
  PRINT #RptHandle, "        Total Late Listings Printed: "; USING "#####"; CustCnt
  PRINT #RptHandle, "Total Value of Late Listed Property: "; USING "$$#######,#"; TotalLateAmt#
  PRINT #RptHandle, FF$
  RETURN
  
GetNameIndexLate:
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

SUB MakeAdCostReport

  SHARED Choice$()

  REDIM Array(1 TO 1) AS Struct 'Template for the sort array
  'REDIM PropDesc$(250)

  REDIM PropDesc$(0 TO 1)

  Dash80$ = STRING$(80, "=")
  
  REDIM AdvTrans(1) AS InterestRecType

  TaxTranLen = LEN(TaxTrans(1))
  TaxCustLen = LEN(TaxCustRec(1))
  TAdvLen = LEN(AdvTrans(1))
  
BeginOver:
  CustCnt = 0
  LibName$ = "TAX"
  ScrnName$ = "ADRPT"
  
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
  'Form$(2, 0) = "0"
  'Form$(4, 0) = DATE$
  Action = 1
  ClearBack

  FirstTime = True

  ShowCursor
  
  DisplayTaxScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF FirstTime THEN
      FirstTime = False
      Fld(2).Protected = True
      Fld(4).Protected = True
      QPrintRC SPACE$(30), 11, 20, 0
      QPrintRC SPACE$(30), 15, 20, 0
    END IF

    SELECT CASE Frm(1).KeyCode
    
    CASE F10Key
      IF LEFT$(Form$(1, 0), 1) = "N" THEN
        UsingIndex = True
      ELSE
        UsingIndex = False
      END IF
      AmountCharged# = Value#(Form$(2, 0), ecode%)
      TaxYear = Value#(Form$(3, 0), ecode%)
      PostDate$ = Form$(4, 0)
      ExitFlag = True
    CASE EscKey
      AbortFlag = True
      ExitFlag = True           'EXIT DO
    END SELECT
  LOOP UNTIL ExitFlag
  
  IF AbortFlag THEN EXIT SUB

  IF Exist%("TAXADV4P.DAT") THEN
    KILL "TAXADV4P.DAT"             'kill any old work file
  END IF
  
  PropTaxFile = FREEFILE
  OPEN "TAXPROP.DAT" FOR RANDOM SHARED AS #PropTaxFile LEN = LEN(PropertyRec(1))
  NumOfPropRecs = LOF(PropTaxFile) / LEN(PropertyRec(1))

  TransFile = FREEFILE
  OPEN "TaxTrans.dat" FOR RANDOM SHARED AS TransFile LEN = TaxTranLen
  
  TAdvFile = FREEFILE
  OPEN "TAXADV4P.DAT" FOR RANDOM SHARED AS TAdvFile LEN = TAdvLen
  
  TaxFile = FREEFILE
  OPEN TaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxCustLen

  NumOfTaxRecs = LOF(TaxFile) / TaxCustLen
  
  IF UsingIndex AND NumOfTaxRecs > 0 THEN
    GOSUB GetAdNameIndex
  END IF
  
  ClearBack
  ShowProcessingScrn "Ad/Collection Cost Listing"
  
  FOR Cnt = 1 TO NumOfTaxRecs
    IF UsingIndex THEN
      CustRecNo& = Array(Cnt).RecNum
    ELSE
      CustRecNo& = Cnt
    END IF
    
    GET TaxFile, CustRecNo&, TaxCustRec(1)

    'IF CustRecNo& = 1853 THEN STOP

    IF NOT TaxCustRec(1).Deleted THEN

      
      TransRList = 0
      TransRecord& = TaxCustRec(1).LastTrans
      WHILE TransRecord& <> 0
        GET TransFile, TransRecord&, TaxTrans(1)
        CurTransRecord& = TransRecord&
        IF TaxTrans(1).TranType = 1 AND TaxTrans(1).TaxYear = TaxYear AND TaxCustRec(1).FirstPropRec > 0 THEN
          Principle# = Round#(TaxTrans(1).Revenue.Principle1 + TaxTrans(1).Revenue.Principle2 + TaxTrans(1).Revenue.Principle3 + TaxTrans(1).Revenue.Principle4 + TaxTrans(1).Revenue.Principle5)

          'leland wants only actual principle.......
          'Principle# = Round#(Principle# + TaxTrans(1).Revenue.PENALTY + TaxTrans(1).Revenue.Collection + TaxTrans(1).Revenue.Interest)

          Paid# = Round#(TaxTrans(1).Revenue.Principle1Pd + TaxTrans(1).Revenue.Principle2Pd + TaxTrans(1).Revenue.Principle3Pd + TaxTrans(1).Revenue.Principle4Pd + TaxTrans(1).Revenue.Principle5Pd)

          'leland wants only actual principle.......
          'Paid# = Round#(Paid# + TaxTrans(1).Revenue.PenaltyPd + TaxTrans(1).Revenue.CollectionPd + TaxTrans(1).Revenue.InterestPD)

          Balance# = Round#(Principle# - Paid#)
          
          TransRList = TransRList + 1
          IF Balance# > 0 THEN
            BillNumber$ = TaxTrans(1).Description
            BillNumber$ = ParseBillNum$(BillNumber$)
            REDIM PropDesc$(0 TO 1)
            'Get Prop Desc Here
            PropRec& = TaxCustRec(1).FirstPropRec
            HowMany = 0
            z1 = 0
            WHILE PropRec& <> 0
              GET PropTaxFile, PropRec&, PropertyRec(1)
              z1 = z1 + 1
              REDIM PRESERVE PropDesc$(0 TO z1)
              PropDesc$(z1) = QPTrim$(PropertyRec(1).RealPin)
              'PropDesc$(z1) = QPTrim$(PropertyRec(1).PROPNOT1)
              'PropDesc$(z1) = QPTrim$(PropertyRec(1).MAP) + "\" + QPTrim$(PropertyRec(1).BLOCK) + "\" + QPTrim$(PropertyRec(1).LOTNUMB) + "  " + QPTrim$(PropertyRec(1).PROPNOT1)
              PropRec& = PropertyRec(1).NextRec
            WEND
            PropDesc$ = PropDesc$(z1)
            NME$ = QPTrim$(TaxCustRec(1).FName) + " " + QPTrim$(TaxCustRec(1).LName)
            NME$ = QPTrim$(NME$)                'this one cleans up those with only last name
            IF TransRList = 1 THEN
              PropDesc$ = PropDesc$(z1)
            END IF

            IF (TransRList > 1) AND (TransRList <= z1) THEN
              'z2 = (z1 - TransRList) + 1
              'IF z2 < 0 THEN z2 = 1
              'PropDesc$ = PropDesc$(z2)
              PropDesc$ = PropDesc$(TransRList)
            END IF

            TotalCharged# = TotalCharged# + AmountCharged#

            AdvTrans(1).CustRec = CustRecNo&
            AdvTrans(1).CUSTNAME = NME$
            AdvTrans(1).TaxYear = TaxYear
            AdvTrans(1).Amount = Balance#
            AdvTrans(1).BillNumber = BillNumber$
            AdvTrans(1).BillRec = TransRecord&
            AdvTrans(1).InfoTxt = PropDesc$
            AdvTrans(1).DelFlag = 0
            PUT #TAdvFile, , AdvTrans(1)
            Balance# = 0
          END IF
        END IF
        TransRecord& = TaxTrans(1).LastTrans
      WEND
      CustCnt = CustCnt + 1
    END IF
    
    ShowPctComp Cnt, NumOfTaxRecs
  NEXT
  
  CLOSE         'Close all open files now

  ERASE Array, Frm, Form$, Fld, TaxCustRec

  REDIM AdcTrans(1) AS InterestRecType

  TAdvLen = LEN(AdcTrans(1))
  Dash$ = STRING$(79, "-")
  TheDate$ = DATE$
  MaxLines = 54
  LineCnt = 0

  TAdcFile = FREEFILE
  OPEN "TAXADV4P.DAT" FOR RANDOM SHARED AS TAdcFile LEN = TAdvLen
  AdcRecord& = LOF(TAdcFile) / TAdvLen
  IF AdcRecord& = 0 THEN
    CLOSE
    CursorOff
    ClearBack
    DisplayTaxScrn "ERRSCRN1"
    QPrintRC "No Active Ad-Cost Transactions!", 10, 25, -1
    QPrintRC "Press any key to continue.", 12, 28, -1
    WaitForAction

    EXIT SUB
  END IF
  'Get Current Year
  GET #TAdcFile, 1, AdcTrans(1)
  WhatYear = AdcTrans(1).TaxYear

  ClearBack
  ShowProcessingScrn "Creating Ad-Cost Report"
  RptFile = FREEFILE

  OPEN "TAXADV4P.RPT" FOR OUTPUT AS #RptFile

  GOSUB AdCostRptHeading

  FOR Cnt& = 1 TO AdcRecord&
    GET #TAdcFile, Cnt&, AdcTrans(1)
    IF AdcTrans(1).DelFlag <> 0 THEN
      GOTO SkipThisAdcTrans
    END IF
    PRINT #RptFile, USING "#####"; AdcTrans(1).CustRec;
    PRINT #RptFile, " "; LEFT$(AdcTrans(1).CUSTNAME, 29);
    PRINT #RptFile, TAB(40); AdcTrans(1).InfoTxt;
    PRINT #RptFile, TAB(70); USING "######.##"; AdcTrans(1).Amount
    CustCnt = CustCnt + 1

    LineCnt = LineCnt + 1
    TotalADCost# = TotalADCost# + AdcTrans(1).Amount

    IF AdcTrans(1).TaxYear = WhatYear THEN
      CurADCost# = CurADCost# + AdcTrans(1).Amount
    ELSE
      PastADCost# = PastADCost# + AdcTrans(1).Amount
    END IF

    IF LineCnt >= MaxLines THEN
      PRINT #RptFile, CHR$(12);
      GOSUB AdCostRptHeading
    END IF
SkipThisAdcTrans:
    ShowPctComp Cnt&, AdcRecord&
  NEXT

  PRINT #RptFile, Dash$
  PRINT #RptFile, CHR$(12);

  CLOSE

  PrintRptFile "Tax Ad-Cost Report.", "TAXADV4P.RPT", 1, RetCode%, 0

  EXIT SUB

AdCostRptHeading:
  PAGE = PAGE + 1
  PRINT #RptFile, TAB(25); "Tax Billing : Advertising Report"
  PRINT #RptFile,
  PRINT #RptFile, "Date: "; TheDate$; TAB(65); "Page #"; PAGE
  PRINT #RptFile, "Tax Year: "; WhatYear
  PRINT #RptFile, "Customer Name                         Property Description        Principle"
  PRINT #RptFile, Dash$
  LineCnt = 5
RETURN

'  PrintEditReport
  
  EXIT SUB
  

GetAdNameIndex:
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

SUB MasterValueList
  
  SHARED Choice$()

  DIM TSetUp(1) AS TaxMasterType
  TaxSetupLen = LEN(TSetUp(1))

  TxFile = FREEFILE
  OPEN "TAXSETUP.DAT" FOR RANDOM AS #TxFile LEN = TaxSetupLen
  IF LOF(TxFile) > 0 THEN
    GET TxFile, 1, TSetUp(1)
    FirstTime = True
  END IF
  CLOSE

  IF INSTR(TSetUp(1).Name, "SUNSET") > 0 THEN
    SunBeech = True
  END IF
  
  REDIM Array(1 TO 1) AS Struct 'Template for the sort array
  ReportFile$ = "TaxValu.PRN"   'Report File Name
  Dash80$ = STRING$(80, "=")
  FF$ = CHR$(12)
  
  LineCnt = 0
  CustCnt = 0
  
  LibName$ = "TAX"
  ScrnName$ = "VALRPT"
  
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
  
  REDIM Choice$(0 TO 3, 0 TO 4)
  
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Name Order"
  Choice$(2, 0) = "Account Number"
  
  Choice$(0, 1) = "2"
  Choice$(1, 1) = "Summary"
  Choice$(2, 1) = "Detail"
  
  Choice$(0, 2) = "3"
  Choice$(1, 2) = "Both"
  Choice$(2, 2) = "Personal"
  Choice$(3, 2) = "Real"
  
  Choice$(0, 3) = "4"
  Choice$(1, 3) = "Both"
  Choice$(2, 3) = "Active"
  Choice$(3, 3) = "Inactive"
  
  Choice$(0, 4) = "5"
  Choice$(1, 4) = "Screen"
  Choice$(2, 4) = "Printer"
  
  Action = 1
  FirstTime = True

  ClearBack
  ShowCursor
  DisplayTaxScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
 '   IF FirstTime THEN
 '     LSET Form$(1, 0) = Choice$(1, 0)
 '     FirstTime = False
 '     Action = 1
 '     Fld(1).Protected = True
 '   END IF

    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      IF LEFT$(Form$(1, 0), 1) = "N" THEN
        UsingIndex = True
      ELSE
        UsingIndex = False
      END IF

      IF LEFT$(Form$(2, 0), 1) = "D" THEN
        DetailFlag = True
      ELSE
        DetailFlag = False
      END IF
      IF DetailFlag THEN
        MaxLines = 40
      ELSE
        MaxLines = 50
      END IF

      PType$ = Form$(3, 0)

      IF (LEFT$(PType$, 1) = "R") AND SunBeech THEN
        UsingIndex = True
        GOSUB MakePinIndex
      END IF

      AStat$ = QPTrim$(LEFT$(Form$(4, 0), 1))
      SELECT CASE AStat$
      CASE "B"
        StatFlag = -1
      CASE "A"
        StatFlag = 1
      CASE "I"
        StatFlag = 0
      END SELECT
      DevSpec$ = LEFT$(Form$(5, 0), 1)
      ExitFlag = True
    CASE EscKey
      AbortFlag = True
      ExitFlag = True           'EXIT DO
    END SELECT
  LOOP UNTIL ExitFlag
  
  IF AbortFlag THEN EXIT SUB
  
  IF UsingIndex THEN
    SubByName = True
    FirstCust = True
  END IF
  
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  GOSUB PrintMasterValueHeader
  
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile

  IF SunBeech THEN
    IF (LEFT$(PType$, 1) <> "R") THEN
      SunBeech = False
    END IF
  END IF

  IF (LEFT$(PType$, 1) = "R") AND SunBeech THEN
    NumOfTaxRecs = Found&
    SubByName = False
  END IF

  IF (UsingIndex AND NumOfTaxRecs > 0) AND NOT SunBeech THEN
    GOSUB GetNameIndex1
  END IF
  
  ClearBack
  ShowProcessingScrn "Master Valuation Report"


  FOR Cnt = 1 TO NumOfTaxRecs
    IF UsingIndex THEN
      CustRecNo = Array(Cnt).RecNum
    ELSE
      CustRecNo = Cnt
    END IF

'dale
'IF CustRecNo < 1076 OR CustRecNo > 1077 THEN
'  GOTO SkipEm10
'END IF

    GET TaxFile, CustRecNo, TaxCustRec(1)
    
    PValue# = 0
    IF NOT TaxCustRec(1).Deleted THEN
      
      SELECT CASE TaxCustRec(1).Active
      CASE "Y"
        IF StatFlag = 0 THEN
          GOTO SkipEm10
        END IF
      CASE ELSE
        IF StatFlag = 1 THEN
          GOTO SkipEm10
        END IF
      END SELECT
      
      IF LineCnt >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintMasterValueHeader
      END IF
      LstName$ = QPTrim$(TaxCustRec(1).LName)
      FstName$ = QPTrim$(TaxCustRec(1).FName)
      NME$ = LstName$ + " " + FstName$
      NME$ = QPTrim$(NME$)      'this one cleans up those with only last name
      TChar$ = LEFT$(LstName$, 1)
      
      IF FirstCust THEN         'First cust set lastname char
        FirstCust = False
        FChar$ = TChar$
      END IF
      
      IF (FChar$ <> TChar$) AND SubByName THEN  'if this a new last name first char
        GOSUB PrintLNameSubTotals               'print subtotals for last char
        FChar$ = TChar$         'set new last char
      END IF
      
      IF NOT DetailFlag THEN
        'Figure Values
        'Real Value First
        RealValue# = 0
        Discnt# = 0
        IF TaxCustRec(1).FirstPropRec > 0 THEN
          PropertyRecord! = TaxCustRec(1).FirstPropRec
          WHILE PropertyRecord! <> 0
            GET PropTaxFile, PropertyRecord!, PropertyRec(1)
            RealValue# = Round#(RealValue# + PropertyRec(1).PROPVALU)
            IF LEFT$(PType$, 1) <> "P" THEN
              Discnt# = Round#(Discnt# + PropertyRec(1).EXMPSENI)
              Discnt# = Round#(Discnt# + PropertyRec(1).EXMPOTHR)
            END IF
            PropertyRecord! = PropertyRec(1).NextRec
          WEND
        END IF
        
        'Personal Property Here
        PersValue# = 0
        IF TaxCustRec(1).FirstPersRec > 0 THEN
          PropertyRecord! = TaxCustRec(1).FirstPersRec
          WHILE PropertyRecord! <> 0
            GET PersTaxFile, PropertyRecord!, PersRec(1)
            PersValue# = Round#(PersValue# + PersRec(1).PERSVAL)
            PersValue# = Round#(PersValue# + PersRec(1).MHVALUE)
            PersValue# = Round#(PersValue# + PersRec(1).MCVALUE)
            PersValue# = Round#(PersValue# + PersRec(1).CVALUE)
            PersValue# = Round#(PersValue# + PersRec(1).MTVALUE)
            IF LEFT$(PType$, 1) <> "R" THEN
              Discnt# = Round#(Discnt# + PersRec(1).EXMPSENI)
            END IF
            PropertyRecord! = PersRec(1).NextRec
          WEND
        END IF
        
        IF LEFT$(PType$, 1) = "R" THEN
          PersValue# = 0
        END IF
        IF LEFT$(PType$, 1) = "P" THEN
          RealValue# = 0
        END IF
        IF LEN(NME$) = 0 AND RealValue# = 0 AND PersValue# = 0 THEN
          TaxCustRec(1).Deleted = -1
          PUT TaxFile, CustRecNo, TaxCustRec(1)
          GOTO SkipEm10
        END IF
        
        PRINT #RptHandle, USING "#####"; CustRecNo;
        PRINT #RptHandle, TAB(8); LEFT$(NME$, 27);
        PRINT #RptHandle, TAB(37); USING "########,#"; RealValue#;
        PRINT #RptHandle, TAB(48); USING "########,#"; PersValue#;
        PRINT #RptHandle, TAB(59); USING "########,#"; Discnt#;
        PRINT #RptHandle, TAB(70); USING "########,#"; Round#(RealValue# + PersValue# - Discnt#)
        LineCnt = LineCnt + 1
        CReal# = Round#(CReal# + RealValue#)
        CPers# = Round#(CPers# + PersValue#)
        CDisc# = Round#(CDisc# + Discnt#)
        TotalReal# = Round#(TotalReal# + RealValue#)
        TotalPers# = Round#(TotalPers# + PersValue#)
        TotalDisc# = Round#(TotalDisc# + Discnt#)
      ELSE
        'Figure Values
        'Real Value First
        RealValue# = 0
        Discnt# = 0
        IF TaxCustRec(1).FirstPropRec > 0 THEN
          PropertyRecord! = TaxCustRec(1).FirstPropRec
          WHILE PropertyRecord! <> 0
            GET PropTaxFile, PropertyRecord!, PropertyRec(1)
            RealValue# = Round#(RealValue# + PropertyRec(1).PROPVALU)
            IF LEFT$(PType$, 1) <> "P" THEN
              Discnt# = Round#(Discnt# + PropertyRec(1).EXMPSENI)
              Discnt# = Round#(Discnt# + PropertyRec(1).EXMPOTHR)
            END IF
            PropertyRecord! = PropertyRec(1).NextRec
          WEND
        END IF
        'Personal Property Here
        PersValue# = 0
        IF TaxCustRec(1).FirstPersRec > 0 THEN
          PropertyRecord! = TaxCustRec(1).FirstPersRec
          WHILE PropertyRecord! <> 0
            GET PersTaxFile, PropertyRecord!, PersRec(1)
'            IF PersRec(1).PERSVAL < 0 OR PersRec(1).MHVALUE < 0 OR PersRec(1).MCVALUE < 0 OR PersRec(1).CVALUE < 0 OR PersRec(1).MTVALUE < 0 THEN
'              LPRINT CustRecNo
'              GOTO NextPropRec
'            END IF
            
            PersValue# = Round#(PersValue# + PersRec(1).PERSVAL)
            PersValue# = Round#(PersValue# + PersRec(1).MHVALUE)
            PersValue# = Round#(PersValue# + PersRec(1).MCVALUE)
            PersValue# = Round#(PersValue# + PersRec(1).CVALUE)
            PersValue# = Round#(PersValue# + PersRec(1).MTVALUE)
            IF LEFT$(PType$, 1) <> "R" THEN
              Discnt# = Round#(Discnt# + PersRec(1).EXMPSENI)
            END IF
NextPropRec:
            PropertyRecord! = PersRec(1).NextRec
          WEND
        END IF
        
        IF LEFT$(PType$, 1) = "R" THEN
          PersValue# = 0
        END IF
        IF LEFT$(PType$, 1) = "P" THEN
          RealValue# = 0
        END IF
        IF LEN(NME$) = 0 AND RealValue# = 0 AND PersValue# = 0 THEN
          TaxCustRec(1).Deleted = -1
          PUT TaxFile, CustRecNo, TaxCustRec(1)
          GOTO SkipEm10
        END IF
        
        PRINT #RptHandle, USING "#####"; CustRecNo;
        PRINT #RptHandle, TAB(8); LEFT$(NME$, 27);
        PRINT #RptHandle, TAB(37); USING "########,#"; RealValue#;
        PRINT #RptHandle, TAB(48); USING "########,#"; PersValue#;
        PRINT #RptHandle, TAB(59); USING "########,#"; Discnt#;
        PRINT #RptHandle, TAB(70); USING "########,#"; Round#(RealValue# + PersValue# - Discnt#)
        CReal# = Round#(CReal# + RealValue#)
        CPers# = Round#(CPers# + PersValue#)
        CDisc# = Round#(CDisc# + Discnt#)
        
        TotalReal# = Round#(TotalReal# + RealValue#)
        TotalPers# = Round#(TotalPers# + PersValue#)
        TotalDisc# = Round#(TotalDisc# + Discnt#)
        
        LineCnt = LineCnt + 1
        
        'Now Show Detail Support Here
        IF LEFT$(PType$, 1) = "P" THEN
        ELSE
          IF TaxCustRec(1).FirstPropRec > 0 THEN
            PropertyRecord! = TaxCustRec(1).FirstPropRec
            PFlag = 0
            WHILE PropertyRecord! <> 0
              GET PropTaxFile, PropertyRecord!, PropertyRec(1)
              PRINT #RptHandle, TAB(15); "Property Pin# "; QPTrim$(PropertyRec(1).RealPin);
              PRINT #RptHandle, TAB(52); "Value: "; USING "#######,#"; PropertyRec(1).PROPVALU
              
              IF LEN(QPTrim$(PropertyRec(1).MAP)) > 0 THEN
                MapCnt = MapCnt + 1
              END IF
              'STOP
              
              'PRINT #RptHandle, TAB(15); "MAP "; PropertyRec(1).MAP;
              'PRINT #RptHandle, TAB(40); "BLK "; PropertyRec(1).BLOCK;
              'PRINT #RptHandle, TAB(60); "LOT "; PropertyRec(1).LOTNUMB
              
              LineCnt = LineCnt + 1
              PFlag = 1
              PropertyRecord! = PropertyRec(1).NextRec
            WEND
          END IF
        END IF
        IF LEFT$(PType$, 1) = "R" THEN
        ELSE
          IF TaxCustRec(1).FirstPersRec > 0 THEN
            PropertyRecord! = TaxCustRec(1).FirstPersRec
            WHILE PropertyRecord! <> 0
              GET PersTaxFile, PropertyRecord!, PersRec(1)

'              IF PersRec(1).PERSVAL < 0 OR PersRec(1).MHVALUE < 0 OR PersRec(1).MCVALUE < 0 OR PersRec(1).CVALUE < 0 OR PersRec(1).MTVALUE < 0 THEN
'                LPRINT CustRecNo
'                GOTO NextPropRec2
'              END IF

              PValue# = Round#(PValue# + PersRec(1).PERSVAL)
              PValue# = Round#(PValue# + PersRec(1).MHVALUE)
              PValue# = Round#(PValue# + PersRec(1).MCVALUE)
              PValue# = Round#(PValue# + PersRec(1).CVALUE)
              PValue# = Round#(PValue# + PersRec(1).MTVALUE)
              PRINT #RptHandle, TAB(15); "Pers Abstract# "; PersRec(1).PropPin;
              PRINT #RptHandle, TAB(52); "Value: "; USING "#######,#"; PValue#
              LineCnt = LineCnt + 1
              PFlag = 1
NextPropRec2:
              PropertyRecord! = PersRec(1).NextRec
            WEND
          END IF
        END IF
      END IF
      CustCnt = CustCnt + 1
      NCustCnt = NCustCnt + 1
    END IF
    
    IF PFlag = 1 THEN
      PRINT #RptHandle, ""
      PFlag = 0
    END IF
    
    
SkipEm10:
    ShowPctComp Cnt, NumOfTaxRecs
  NEXT Cnt
  IF SubByName THEN
    LastOneFlag = True
    GOSUB PrintLNameSubTotals
  END IF
  
  GOSUB PrintMasterValueEnding
  
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
  
  'KILL ReportFile$
  
  EXIT SUB
  
PrintMasterValueHeader:
  PAGE = PAGE + 1
  PRINT #RptHandle, TAB(25); "Property Tax Valuation Listing"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; PAGE
  IF NOT DetailFlag THEN
    PRINT #RptHandle, "Summary Format"
    PRINT #RptHandle, TAB(40); "[------------- Valuations --------------]"
    PRINT #RptHandle, "Acct #"; TAB(8); "Name"; TAB(43); "Real"; TAB(53); "Pers"; TAB(62); "Discnt"; TAB(77); "Net"
    PRINT #RptHandle, Dash80$
    LineCnt = 6
  ELSE
    PRINT #RptHandle, "Detail Format"
    PRINT #RptHandle, TAB(40); "[------------- Valuations -------------]"
    PRINT #RptHandle, "Acct #"; TAB(8); "Name"; TAB(46); "Real"; TAB(55); "Pers"; TAB(63); "Discnt"; TAB(76); "Net"
    PRINT #RptHandle, Dash80$
    LineCnt = 6
  END IF
RETURN
  
PrintMasterValueEnding:
  PRINT #RptHandle, Dash80$
  PRINT #RptHandle, "Total Customers Printed: "; USING "#####"; CustCnt
  PRINT #RptHandle, "Totals ..."
  PRINT #RptHandle, "Real Value: "; USING "$$##########,#"; TotalReal#
  PRINT #RptHandle, "Pers Value: "; USING "$$##########,#"; TotalPers#
  PRINT #RptHandle, "Discount  : "; USING "$$##########,#"; TotalDisc#
  PRINT #RptHandle, "Net Value : "; USING "$$##########,#"; Round#(TotalReal# + TotalPers# - TotalDisc#)
  PRINT #RptHandle, FF$
RETURN
  
GetNameIndex1:
  REDIM Array(1 TO NumOfTaxRecs) AS Struct
  FOR Cnt = 1 TO NumOfTaxRecs
    GET TaxFile, Cnt, TaxCustRec(1)
    Array(Cnt).who = UCASE$(TaxCustRec(1).SName) + " "
    Array(Cnt).RecNum = Cnt
  NEXT
  
  'Sort Them Here
  SortT Array(1), NumOfTaxRecs, 0, LEN(Array(1)), 0, 14
RETURN

MakePinIndex:
'dale
  Found& = 0
  OpenTaxCustFile NumOfTaxRecs, TaxFile  'get num of cust recs
  CLOSE TaxFile

  REDIM TempPropRec(1) AS PropertyRecType
  PropRecLen = LEN(TempPropRec(1))

  PropFile = FREEFILE
  OPEN "TAXPROP.DAT" FOR RANDOM SHARED AS PropFile LEN = PropRecLen
  NumProp& = LOF(PropFile) / PropRecLen

  REDIM Array(1 TO NumProp&) AS Struct
  FOR Cnt& = 1 TO NumProp&
    GET PropFile, Cnt&, TempPropRec(1)
    TempPin& = TempPropRec(1).CustPin
    IF TempPin& <= 0 OR TempPin& > NumOfTaxRecs THEN
      GOTO BadPinSkip
    END IF
    Found& = Found& + 1
    Array(Found&).who = LEFT$(QPTrim$(TempPropRec(1).RealPin), 3)
    Array(Found&).RecNum = TempPin&
BadPinSkip:
  NEXT
  CLOSE PropFile

  SortT Array(1), CINT(Found&), 0, LEN(Array(1)), 0, 3
'SortT (Element, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)


RETURN
  
PrintLNameSubTotals:
  PRINT #RptHandle, "["; STRING$(78, "-"); "]"
  PRINT #RptHandle, "  Sub Totals for '"; FChar$; "'      "; USING "#####"; NCustCnt;
  PRINT #RptHandle, TAB(46); "Real"; TAB(55); "Pers"; TAB(63); "Discnt"; TAB(76); "Net"
  PRINT #RptHandle, TAB(37); USING "########,#"; CReal#;
  PRINT #RptHandle, TAB(48); USING "########,#"; CPers#;
  PRINT #RptHandle, TAB(59); USING "########,#"; CDisc#;
  PRINT #RptHandle, TAB(70); USING "########,#"; Round#(CReal# + CPers# - CDisc#)
  PRINT #RptHandle, FF$
  IF NOT LastOneFlag THEN
    GOSUB PrintMasterValueHeader
  END IF
  NCustCnt = 0
  CReal# = 0
  CPers# = 0
  CDisc# = 0
  
  RETURN
  
END SUB

SUB MortCodeReport
  SHARED Choice$()

  REDIM MCode$(250), MDesc$(250)
  
  REDIM MortCodeRec(1) AS MortCodeRecType
  MortCodeRecLen = LEN(MortCodeRec(1))
  REDIM Array(1 TO 1) AS Struct 'Template for the sort array
  ReportFile$ = "TaxMORT.PRN"   'Report File Name
  Dash80$ = STRING$(80, "=")
  FF$ = CHR$(12)
  
  MaxLines = 56
  LineCnt = 0
  CustCnt = 0
  
  LibName$ = "TAX"
  ScrnName$ = "LATERPT"
  
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

  FirstTime = True

  ShowCursor
  
  DisplayTaxScrn ScrnName$
  QPrintRC "Mortgage Code Report ", 8, 24, 14
  QPrintRC "]", 8, 45, 10
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      LSET Form$(2, 0) = "N/A"
      Fld(2).Protected = -1
      Action = 1
    END IF

    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      IF LEFT$(Form$(1, 0), 1) = "N" THEN
        UsingIndex = True
      ELSE
        UsingIndex = False
      END IF
      DetailFlag$ = LEFT$(Form$(2, 0), 1)
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
  
  GOSUB PrintMCHeader
  
  GOSUB MCodes
  
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile
  
  IF UsingIndex AND NumOfTaxRecs > 0 THEN
    GOSUB GetNameIndexMC
  END IF
  
  ClearBack
  ShowProcessingScrn "Mortgage Code Report"
  
  FOR MCodes = 1 TO NumOfCatRecs
    QPrintRC "Processing Mortgage Code: " + MCode$(MCodes), 11, 27, -1
    PRINT #RptHandle, "Mortgage Code "; MCode$(MCodes); "  - "; MDesc$(MCodes)
    LineCnt = LineCnt + 1
    
    FOR Cnt = 1 TO NumOfTaxRecs
      IF UsingIndex THEN
        CustRecNo = Array(Cnt).RecNum
      ELSE
        CustRecNo = Cnt
      END IF
      
      GET TaxFile, CustRecNo, TaxCustRec(1)
      
      IF NOT TaxCustRec(1).Deleted THEN
        
        ' Check Line on Page and Form Feed if Necessary
        IF LineCnt >= MaxLines THEN
          PRINT #RptHandle, FF$
          GOSUB PrintMCHeader
          PRINT #RptHandle, "Mortgage Code "; MCode$(MCodes); "  - "; MDesc$(MCodes)
        END IF
        
        NME$ = QPTrim$(TaxCustRec(1).FName) + " " + QPTrim$(TaxCustRec(1).LName)
        NME$ = QPTrim$(NME$)    'this one cleans up those with only last name
        
        'First Show Property Records
        
        IF TaxCustRec(1).FirstPropRec > 0 THEN
          PropertyRecord! = TaxCustRec(1).FirstPropRec
          WHILE PropertyRecord! <> 0
            GET #PropTaxFile, PropertyRecord!, PropertyRec(1)
            IF QPTrim$(PropertyRec(1).MortCode) = MCode$(MCodes) THEN
              PRINT #RptHandle, "Acct # "; USING "######"; CustRecNo;
              PRINT #RptHandle, TAB(20); "PIN # "; PropertyRec(1).RealPin
              PRINT #RptHandle, NME$; TAB(60); " MC: "; PropertyRec(1).MortCode
              PRINT #RptHandle, TaxCustRec(1).Addr1; TAB(60); "VAL:"; USING "$$#######,#"; PropertyRec(1).PROPVALU
              PRINT #RptHandle, TaxCustRec(1).Addr2
              PRINT #RptHandle, RTRIM$(TaxCustRec(1).City); " ";
              PRINT #RptHandle, RTRIM$(TaxCustRec(1).State); " ";
              PRINT #RptHandle, RTRIM$(TaxCustRec(1).Zip)
              PRINT #RptHandle, STRING$(79, "-")
              
              LineCnt = LineCnt + 6
              CustCnt = CustCnt + 1
            END IF
            PropertyRecord! = PropertyRec(1).NextRec
          WEND
        END IF
        
      END IF
      ShowPctComp MCodes, NumOfCatRecs
    NEXT Cnt
  NEXT MCodes
  
  GOSUB PrintMCEnding
  
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
  
PrintMCHeader:
  PAGE = PAGE + 1
  PRINT #RptHandle, TAB(24); "Tax Customer Mortgage Code Report"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; PAGE
  PRINT #RptHandle, Dash80$
  LineCnt = 5
  RETURN
  
PrintMCEnding:
  PRINT #RptHandle, Dash80$
  PRINT #RptHandle, "        Total Listings Printed: "; USING "#####"; CustCnt
  PRINT #RptHandle, FF$
  RETURN
  
GetNameIndexMC:
  REDIM Array(1 TO NumOfTaxRecs) AS Struct
  FOR Cnt = 1 TO NumOfTaxRecs
    GET TaxFile, Cnt, TaxCustRec(1)
    Array(Cnt).who = UCASE$(TaxCustRec(1).SName) + " "
    Array(Cnt).RecNum = Cnt
  NEXT
  
  'Sort Them Here
  SortT Array(1), NumOfTaxRecs, 0, LEN(Array(1)), 0, 14
  RETURN

MCodes:
  MortFile = FREEFILE
  OPEN "TAXMORT.DAT" FOR RANDOM AS MortFile LEN = MortCodeRecLen
  NumOfCatRecs = LOF(MortFile) \ MortCodeRecLen
  IF NumOfCatRecs > 0 THEN
    GOSUB SortMortCodes
    REDIM MChoice$(1 TO NumOfCatRecs)
    FOR Cnt = 1 TO NumOfCatRecs
      GET MortFile, Array(Cnt).RecNum, MortCodeRec(1)
      MCode$(Cnt) = RTRIM$(MortCodeRec(1).MortCode)
      MDesc$(Cnt) = RTRIM$(MortCodeRec(1).BName)
      
    NEXT Cnt
  END IF
  
  
SortMortCodes:
  REDIM Array(1 TO NumOfCatRecs) AS Struct
  FOR Cnt = 1 TO NumOfCatRecs
    GET MortFile, Cnt, MortCodeRec(1)
    Array(Cnt).who = MortCodeRec(1).MortCode + STRING$(11, " ")
    Array(Cnt).RecNum = Cnt
  NEXT Cnt
  SortT Array(1), NumOfTaxRecs, 0, LEN(Array(1)), 0, 14
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
      
      NME$ = QPTrim$(TaxCustRec(1).FName) + " " + QPTrim$(TaxCustRec(1).LName)
      NME$ = QPTrim$(NME$)      'this one cleans up those with only last name
      
      
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
        PRINT #RptHandle, TaxCustRec(1).CSSN; TAB(15); NME$; TAB(57); USING "$$#######,#"; SCAmt#;
        PRINT #RptHandle, TAB(71); USING "$####,#.##"; TaxLoss#
        PRINT #RptHandle, TAB(15); RTRIM$(TaxCustRec(1).Addr1) + " " + RTRIM$(TaxCustRec(1).City) + " " + TaxCustRec(1).State + " " + TaxCustRec(1).Zip
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
  
  'KILL ReportFile$
  
  EXIT SUB
  
PrintSCRptHeader:
  PAGE = PAGE + 1
  PRINT #RptHandle, TAB(20); "Senior Citizen Discount Report    :   Form AV-22C"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; PAGE
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

SUB TaxPaymentRpt
  
  SHARED Choice$()
  ReportFile$ = "TRANJOUR.PRN"  'Report File Name
  Dash80$ = STRING$(80, "=")
  FF$ = CHR$(12)
  
  FirstYear = True
  FirstTime = True
  REDIM TranRec(1 TO 2) AS TaxTransactionType
  TaxTranRecLen = LEN(TranRec(1))
  
  REDIM YearInfo(1 TO 1) AS PaySumType
  
  MaxLines = 56
  LineCnt = 0
  CustCnt = 0
  
  LibName$ = "TAX"
  ScrnName$ = "TAXPAYRP"
  
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
  
  Choice$(0, 1) = "3"
  Choice$(1, 1) = "Screen"
  Choice$(2, 1) = "Printer"
  
  Action = 1
  ClearBack
  
  ShowCursor
  
  DisplayTaxScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    'IF FirstTime THEN
    '  FirstTime = False
    '  Form$(1, 0) = "01-01-1900"
    '  Form$(2, 0) = "12-31-2000"
    '  Action = 1
    'END IF
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      BDate$ = Form$(1, 0)
      BDate = Date2Num%(BDate$)
      EDate$ = Form$(2, 0)
      EDate = Date2Num%(EDate$)
      DevSpec$ = LEFT$(Form$(3, 0), 1)
      TrType = 2
      IF TrType >= 0 AND TrType <= 8 AND EDate >= BDate THEN
        ExitFlag = True
      END IF
    CASE EscKey
      AbortFlag = True
      ExitFlag = True           'EXIT DO
    END SELECT
  LOOP UNTIL ExitFlag
  
  IF AbortFlag THEN EXIT SUB

  ClearBack
  MakeCustIndex 1
  
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  GOSUB PrintTransJourRptHeader
  
  OpenTaxCustFile 0&, TaxFile
  CustRecLen = LEN(TaxCustRec(1))
  
  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
  NumOfTranRecs& = LOF(TaxTran) / TaxTranRecLen
  
  ClearBack
  ShowProcessingScrn "Tax Payment Report"
  
  FOR TCnt& = NumOfTranRecs& TO 1 STEP -1
    GET TaxTran, TCnt&, TranRec(1)
    CustBeenPrinted = False
    IF LineCnt >= MaxLines THEN
      PRINT #RptHandle, FF$
      GOSUB PrintTransJourRptHeader
    END IF
    IF TranRec(1).TranType = 2 THEN
      IF TranRec(1).TransDate >= BDate AND TranRec(1).TransDate <= EDate THEN
        IF TranRec(1).CustomerRec > 0 THEN
          GET TaxFile, TranRec(1).CustomerRec, TaxCustRec(1)
          PRINT #RptHandle, QPTrim$(TaxCustRec(1).FName); " "; QPTrim$(TaxCustRec(1).LName)
          LineCnt = LineCnt + 2
          CustBeenPrinted = True
        END IF
        IF TranRec(1).BelongTo > 0 THEN
          GET TaxTran, TranRec(1).BelongTo, TranRec(2)
          TaxYear = TranRec(2).TaxYear
          TaxYear$ = QPTrim$(STR$(TranRec(2).TaxYear))
        ELSE
          TaxYear = 0
          TaxYear$ = "????"
        END IF
        GOSUB ParsetoYear
        PRINT #RptHandle, Num2Date$(TranRec(1).TransDate);
        PRINT #RptHandle, TAB(12); TranRec(1).Description;
        PRINT #RptHandle, TAB(50); TaxYear$;    '" "; TranRec(1).CustomerRec;
        'PRINT #RptHandle, TAB(50); TranRec(1).CustomerRec;
        PRINT #RptHandle, TAB(68); USING "$$######,#.##"; TranRec(1).Amount
        LineCnt = LineCnt + 1
        TotalAmt# = Round#(TotalAmt# + TranRec(1).Amount)
        Principle# = Round#(Principle# + TranRec(1).Revenue.Principle1Pd)
        Interest# = Round#(Interest# + TranRec(1).Revenue.InterestPD)
        PENALTY# = Round#(PENALTY# + TranRec(1).Revenue.PenaltyPd)
        Collect# = Round#(Collect# + TranRec(1).Revenue.CollectionPd)
      END IF
    END IF
    OffSet& = NumOfTranRecs& - TCnt&
    IF OffSet& < 1 THEN OffSet& = 1
    ShowPctCompL OffSet&, NumOfTranRecs&
  NEXT
  
  SortT YearInfo(1), YCnt, 0, LEN(YearInfo(1)), 0, 2
  'SortT (Element, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
  
  GOSUB PrintTransJourRptEnding
  
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi
  
  CLOSE         'Close all open files now
  
  IF DevSpec$ = "P" THEN
    EntryPoint = 4
  ELSEIF DevSpec$ = "S" THEN
    EntryPoint = 2
  ELSE
    EntryPoint = 1
  END IF
  
  ERASE Frm, Form$, Fld, TaxCustRec
  
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  KILL ReportFile$
  
  EXIT SUB
  
PrintTransJourRptHeader:
  PAGE = PAGE + 1
  PRINT #RptHandle, TAB(22); "Property Tax Payment Report Listing"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; PAGE
  PRINT #RptHandle, "Date Range: "; BDate$; "   Thru: "; EDate$
  PRINT #RptHandle, "Tr Date"; TAB(12); "Description"; TAB(50); "Tax Year"; TAB(70); "  Amount"
  PRINT #RptHandle, Dash80$
  LineCnt = 7
RETURN
  
PrintTransJourRptEnding:
  PRINT #RptHandle, Dash80$
  PRINT #RptHandle,
  PRINT #RptHandle, "Summary"
  PRINT #RptHandle, " Bill Year     Principle         Interest          Penalty       Collect/Ad"
  PRINT #RptHandle, Dash80$
  FOR zz = 1 TO YCnt
    PRINT #RptHandle, TAB(5); YearInfo(zz).Year;
    PRINT #RptHandle, USING "$$#######,#.##"; YearInfo(zz).Principle; TAB(28); YearInfo(zz).Interest; TAB(45); YearInfo(zz).PENALTY; TAB(62); YearInfo(zz).ColAdd
  NEXT
  PRINT #RptHandle, Dash80$
  PRINT #RptHandle, "   Totals:";
  PRINT #RptHandle, USING "$$#######,#.##"; Principle#; TAB(28); Interest#; TAB(45); PENALTY#; TAB(62); Collect#
  
  'PRINT #RptHandle, "   Principle: "; USING "$$#######,#.##"; Principle#
  'PRINT #RptHandle, "    Interest: "; USING "$$#######,#.##"; Interest#
  'PRINT #RptHandle, "     Penalty: "; USING "$$#######,#.##"; Penalty#
  'PRINT #RptHandle, "  Collect/Ad: "; USING "$$#######,#.##"; Collect#
  'PRINT #RptHandle, "---------------------------------------"
  'PRINT #RptHandle, "Total Amount: "; USING "$$#######,#.##"; TotalAmt#
  
  PRINT #RptHandle, FF$
RETURN
  
ParsetoYear:
  
  IF FirstYear THEN
    FirstYear = False
    YCnt = 1
    YearInfo(1).Year = TaxYear
    YearInfo(1).Principle = TranRec(1).Revenue.Principle1Pd
    YearInfo(1).Interest = TranRec(1).Revenue.InterestPD
    YearInfo(1).PENALTY = TranRec(1).Revenue.PenaltyPd
    YearInfo(1).ColAdd = TranRec(1).Revenue.CollectionPd
  ELSE
    FOR ZZCnt = 1 TO YCnt
      IF YearInfo(ZZCnt).Year = TaxYear THEN
        YearInfo(ZZCnt).Principle = Round#(YearInfo(ZZCnt).Principle + TranRec(1).Revenue.Principle1Pd)
        YearInfo(ZZCnt).Interest = Round#(YearInfo(ZZCnt).Interest + TranRec(1).Revenue.InterestPD)
        YearInfo(ZZCnt).PENALTY = Round#(YearInfo(ZZCnt).PENALTY + TranRec(1).Revenue.PenaltyPd)
        YearInfo(ZZCnt).ColAdd = Round#(YearInfo(ZZCnt).ColAdd + TranRec(1).Revenue.CollectionPd)
        GOTO GotoYear
      END IF
    NEXT
    'if we fell thru to here then we don't have this year started
    'do it
    YCnt = YCnt + 1
    REDIM PRESERVE YearInfo(1 TO YCnt) AS PaySumType
    YearInfo(YCnt).Year = TaxYear
    YearInfo(YCnt).Principle = TranRec(1).Revenue.Principle1Pd
    YearInfo(YCnt).Interest = TranRec(1).Revenue.InterestPD
    YearInfo(YCnt).PENALTY = TranRec(1).Revenue.PenaltyPd
    YearInfo(YCnt).ColAdd = TranRec(1).Revenue.CollectionPd
  END IF
  
GotoYear:
  
  RETURN
  
  
  'IF TrType <> 1 THEN
  '  IF TranRec(1).BelongTo > 0 THEN
  '    GET TaxTran, TranRec(1).BelongTo, TranRec(2)
  '    'IF TaxType$ = TranRec(2).BillType THEN
  '      IF NOT CustBeenPrinted THEN
  '        IF TranRec(2).CustomerRec > 0 THEN
  '          GET TaxFile, TranRec(2).CustomerRec, TaxCustRec(1)
  '          PRINT #RptHandle, QPTrim$(TaxCustRec(1).FName); " "; QPTrim$(TaxCustRec(1).LName)
  '          CustBeenPrinted = True
  '          LineCnt = LineCnt + 1
  '        END IF
  '      END IF
  '
  '      GET TaxTran, TCnt&, TranRec(1)
  '      PRINT #RptHandle, Num2Date$(TranRec(1).TransDate);
  '      PRINT #RptHandle, TAB(12); TranRec(1).Description;
  '
  '      IF TranRec(2).TaxYear > 0 THEN
  '        PRINT #RptHandle, TAB(50); TranRec(2).TaxYear;
  '        'ELSEIF TrType = 2 THEN
  '        '  PRINT #RptHandle, TAB(50); TranRec(1).CustomerRec;
  '      ELSE
  '        PRINT #RptHandle, TAB(50); "N/A";
  '      END IF
  '      PRINT #RptHandle, TAB(68); USING "$$######,#.##"; TranRec(1).Amount
  '      LineCnt = LineCnt + 1
  '      TotalAmt# = TotalAmt# + TranRec(1).Amount
  '    'END IF
  '  END IF
  'END IF
  
  
END SUB

SUB TransactionJournal
  SHARED Choice$()
  
  REDIM Array(1 TO 1) AS Struct 'Template for the sort array
  ReportFile$ = "TRANJOUR.PRN"  'Report File Name
  Dash80$ = STRING$(80, "=")
  FF$ = CHR$(12)
  
  REDIM TranRec(1 TO 2) AS TaxTransactionType
  TaxTranRecLen = LEN(TranRec(1))
  
  MaxLines = 56
  LineCnt = 0
  CustCnt = 0
  
  LibName$ = "TAX"
  ScrnName$ = "TRANSREP"
  
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
  
  REDIM Choice$(0 TO 9, 0 TO 2)
  
  Choice$(0, 0) = "1"
  'Choice$(1, 0) = "0-All Transactions"
  Choice$(1, 0) = "1-Billing"
  Choice$(2, 0) = "2-Payments"
  Choice$(3, 0) = "3-Release/Abatements"
  Choice$(4, 0) = "4-Interest"
  Choice$(5, 0) = "5-Penalty"
  Choice$(6, 0) = "6-Collection Cost"
  Choice$(7, 0) = "7-Adjustments"
  Choice$(8, 0) = "8-Misc. Costs"
  
  Choice$(0, 1) = "4"
  Choice$(1, 1) = "Real Esate"
  Choice$(2, 1) = "Personal"
  Choice$(3, 1) = "Combined"
  
  Choice$(0, 2) = "5"
  Choice$(1, 2) = "Screen"
  Choice$(2, 2) = "Printer"
  
  Action = 1
  ClearBack
  
  ShowCursor
  
  DisplayTaxScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    'if FirstTime
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      TrType = VAL(Form$(1, 0))
      BDate$ = Form$(2, 0)
      BDate = Date2Num%(BDate$)

      EDate$ = Form$(3, 0)
      EDate = Date2Num%(EDate$)

      'EDate$ = BDate$

      TaxType$ = QPTrim$(LEFT$(Form$(4, 0), 1))
      IF LEN(TaxType$) = 0 THEN
        NoBillType = True
      END IF
      DevSpec$ = LEFT$(Form$(5, 0), 1)
      IF TrType >= 0 AND TrType <= 8 AND EDate >= BDate THEN
        ExitFlag = True
      END IF
    CASE EscKey
      AbortFlag = True
      ExitFlag = True           'EXIT DO
    END SELECT
  LOOP UNTIL ExitFlag
  
  IF AbortFlag THEN EXIT SUB
  
  RptHandle = FREEFILE
  
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  GOSUB PrintPayRptHeader
  
  OpenTaxCustFile 0&, TaxFile
  CustRecLen = LEN(TaxCustRec(1))
  
  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
  NumOfTranRecs& = LOF(TaxTran) / TaxTranRecLen

  'OPEN "TR.LST" FOR RANDOM AS #12 LEN = 4

  ClearBack
  ShowProcessingScrn "Transaction Journal Listing"
  FOR TCnt& = NumOfTranRecs& TO 1 STEP -1
    GET TaxTran, TCnt&, TranRec(1)

    IF TranRec(1).TransDate < 0 THEN
      'TranRec(1).CustomerRec
      BadTrCnt& = BadTrCnt& + 1
      'STOP
    END IF

    CustBeenPrinted = False
    IF LineCnt >= MaxLines THEN
      PRINT #RptHandle, FF$
      GOSUB PrintPayRptHeader
    END IF
    IF TranRec(1).TransDate >= BDate AND TranRec(1).TransDate <= EDate THEN

      'IF TranRec(1).CustomerRec = 1776 THEN STOP

      IF NoBillType THEN
        GOTO SkipIntoHere:
      END IF
      IF TrType = 0 OR TranRec(1).TranType = TrType THEN
        IF ((TaxType$ = TranRec(1).BillType) OR NoBillType) AND (TrType = 1 OR TrType = 0) THEN
SkipIntoHere:

          'PUT #12, , TCnt&

          IF NOT CustBeenPrinted THEN
            IF TranRec(1).CustomerRec > 0 THEN
              GET TaxFile, TranRec(1).CustomerRec, TaxCustRec(1)
              PRINT #RptHandle, QPTrim$(TaxCustRec(1).FName); " "; QPTrim$(TaxCustRec(1).LName);
              'PRINT #RptHandle, "!"; TranRec(1).Posted2GL; "!"
              PRINT #RptHandle, TAB(50); USING "##########"; TranRec(1).CustomerRec
              'PRINT #RptHandle, "!TR"; TCnt&; "!"
              LineCnt = LineCnt + 2
              CustBeenPrinted = True
            END IF
          END IF
          
          PRINT #RptHandle, Num2Date$(TranRec(1).TransDate);
          PRINT #RptHandle, TAB(12); TranRec(1).Description;
          PRINT #RptHandle, TAB(50); TranRec(1).TaxYear;
          PRINT #RptHandle, TAB(68); USING "$$######,#.##"; TranRec(1).Amount
          LineCnt = LineCnt + 1
          TotalAmt# = TotalAmt# + TranRec(1).Amount
        END IF
        
        IF TrType <> 1 THEN
          IF TranRec(1).BelongTo > 0 THEN
            GET TaxTran, TranRec(1).BelongTo, TranRec(2)
            IF TaxType$ = TranRec(2).BillType THEN
              IF NOT CustBeenPrinted THEN
                IF TranRec(2).CustomerRec > 0 THEN
                  GET TaxFile, TranRec(2).CustomerRec, TaxCustRec(1)
                  PRINT #RptHandle, QPTrim$(TaxCustRec(1).FName); " "; QPTrim$(TaxCustRec(1).LName)
                  CustBeenPrinted = True
                  LineCnt = LineCnt + 1
                END IF
              END IF
              
              GET TaxTran, TCnt&, TranRec(1)
              PRINT #RptHandle, Num2Date$(TranRec(1).TransDate);
              PRINT #RptHandle, TAB(12); TranRec(1).Description;
              
              IF TranRec(2).TaxYear > 0 THEN
                PRINT #RptHandle, TAB(50); TranRec(2).TaxYear;
                'ELSEIF TrType = 2 THEN
                '  PRINT #RptHandle, TAB(50); TranRec(1).CustomerRec;
              ELSE
                PRINT #RptHandle, TAB(50); "N/A";
              END IF
              PRINT #RptHandle, TAB(68); USING "$$######,#.##"; TranRec(1).Amount
              LineCnt = LineCnt + 1
              TotalAmt# = TotalAmt# + TranRec(1).Amount
            END IF
          END IF
        END IF
      END IF
    END IF
    ShowPctCompL TCnt&, NumOfTranRecs&
  NEXT
  
  GOSUB PrintPayRptFooter
  
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
  
  'BadTrCnt&
  'KILL ReportFile$
  
  EXIT SUB
  
PrintPayRptHeader:
  PAGE = PAGE + 1
  PRINT #RptHandle, TAB(20); "Property Tax Detailed Journal Listing"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; PAGE
  PRINT #RptHandle, "Tax Type: ";
  IF TaxType$ = "R" THEN PRINT #RptHandle, "Real Estate"
  IF TaxType$ = "P" THEN PRINT #RptHandle, "Personal"
  IF TaxType$ = "C" THEN PRINT #RptHandle, "Combined"
  PRINT #RptHandle, "Transaction Journal for ";
  IF TrType = 0 THEN PRINT #RptHandle, "All Transactions"
  IF TrType = 1 THEN PRINT #RptHandle, "Billing Transactions"
  IF TrType = 2 THEN PRINT #RptHandle, "Payment Transactions"
  IF TrType = 3 THEN PRINT #RptHandle, "Release/Abatement Transactions"
  IF TrType = 4 THEN PRINT #RptHandle, "Interest Transactions"
  IF TrType = 5 THEN PRINT #RptHandle, "Penalty Transactions"
  IF TrType = 6 THEN PRINT #RptHandle, "Collection Cost Transactions"
  IF TrType = 7 THEN PRINT #RptHandle, "Adjustments Transactions"
  IF TrType = 8 THEN PRINT #RptHandle, "Collection/Add Cost Adjustments"
  PRINT #RptHandle, "Date Range: Beg on "; BDate$; " Ending on "; EDate$
  PRINT #RptHandle, "Tr Date"; TAB(12); "Description"; TAB(50); "Tax Year"; TAB(70); "  Amount"
  PRINT #RptHandle, Dash80$
  LineCnt = 7
  RETURN
  
PrintPayRptFooter:
  PRINT #RptHandle, Dash80$
  PRINT #RptHandle, "Total Amount: "; USING "$$#######,#.##"; TotalAmt#
  PRINT #RptHandle,
  PRINT #RptHandle, FF$
  RETURN
  
  
  
END SUB

