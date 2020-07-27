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
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 13)
  
  MChoice$(1) = "Customer Inquiry"
  MChoice$(2) = "Customer Transaction History"
  MChoice$(3) = "Master Customer Listing"
  MChoice$(4) = "Master Balance Listing"
  MChoice$(5) = "Master Valuation Listing"
  MChoice$(6) = "Transaction Journal"
  MChoice$(7) = "Late Listing"
  MChoice$(8) = "Sr. Citizens Listing"
  MChoice$(9) = "Advertising Listing"
  MChoice$(10) = "Print Abstracts"
  MChoice$(11) = "Exit to DOS"
  
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
    
    TitleBox 2, Col, MaxLen + 3, "Tax Billing Reports Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      CustomerInquiry
    CASE 2
      'CustHistoryRpt
    CASE 3
      CustomerListing
    CASE 4
      BalanceListing
    CASE 5
      MasterValueList
    CASE 6
      TransactionJournal
    CASE 7
      LateListing
    CASE 8
      SrCitizensList
    CASE 9
      AdListing
    CASE 10
      Abstracts
    CASE 11
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
      
      nme$ = QPTrim$(TaxCustRec(1).FName) + " " + QPTrim$(TaxCustRec(1).LName)
      nme$ = QPTrim$(nme$)      'this one cleans up those with only last name
      
      
      
      'Now Show Property Records Next
      
      IF TaxCustRec(1).FirstPropRec > 0 THEN
        
        PropertyRecord! = TaxCustRec(1).FirstPropRec
        FirstProperty = 1
        
        WHILE PropertyRecord! <> 0
          
          GOSUB AbHeading
          
          PRINT #RptHandle, "Property Owned..."
          GET #PropTaxFile, PropertyRecord!, PropertyRec(1)
          PRINT #RptHandle, "            ---------------------------------------------------------------"
          PRINT #RptHandle, "      PIN#: "; QPTrim$(PropertyRec(1).REALPIN)
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
  PRINT #10, "      PIN#: "; pin$
  PRINT #10, "            ---------------------------------------------"
  PRINT #10, "      MAP#: "; pmap$;
  IF showparcel$ = "Y" THEN
    PRINT #10, TAB(50); "Real Value: "; USING "$$######,#"; CVD(PLAND$) + CVD(PBLDG$)
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
  OPEN "TAXSETUP.DAT" FOR RANDOM AS #TxFile LEN = LEN(TaxSetupRec)
  IF LOF(TxFile) > 0 THEN
    GET TxFile, 1, TaxSetupRec
    TownName$ = TaxSetupRec.Name
  ELSE
    TownName$ = ""
  END IF
  CLOSE TxFile
  PRINT #RptHandle, "Abstract Listing : "; TownName$; TAB(60); "Tax Year "; TaxYear$
  PRINT #RptHandle, ""
  PRINT #RptHandle, "Account of ..."
  PRINT #RptHandle, LEFT$(nme$, 40); TAB(55); "SS# "; TaxCustRec(1).CSSN
  PRINT #RptHandle, TaxCustRec(1).ADDR1
  PRINT #RptHandle, TaxCustRec(1).ADDR2
  PRINT #RptHandle, RTRIM$(TaxCustRec(1).CITY); " "; TaxCustRec(1).STATE; " "; TaxCustRec(1).ZIP;
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

SUB AdListing
  SHARED Choice$()
  
  REDIM Array(1 TO 1) AS Struct 'Template for the sort array
  REDIM PropDesc$(250)
  ReportFile$ = "TxAdList.PRN"  'Report File Name
  Dash80$ = STRING$(80, "=")
  FF$ = CHR$(12)
  
BeginOver:
  MaxLines = 56
  LineCnt = 0
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
  Choice$(0, 1) = "5"
  Choice$(1, 1) = "Screen"
  Choice$(2, 1) = "Printer"
  Form$(2, 0) = "0"
  Form$(4, 0) = DATE$
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
      AmountCharged# = Value#(Form$(2, 0), ecode%)
      TaxYear = Value#(Form$(3, 0), ecode%)
      PostDate$ = Form$(4, 0)
      DevSpec$ = LEFT$(Form$(5, 0), 1)
      IF AmountToCharge# <> 0 AND LEN(QPTrim$(PostDate$)) <> 10 THEN
        Action = 1
        GOTO BeginOver
      END IF
      ExitFlag = True
    CASE EscKey
      AbortFlag = True
      ExitFlag = True           'EXIT DO
    END SELECT
  LOOP UNTIL ExitFlag
  
  IF AbortFlag THEN EXIT SUB
  
  RptHandle = FREEFILE
  
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  GOSUB PrintAdRptHeader
  
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  
  TransFile = FREEFILE
  OPEN "TaxTrans.dat" FOR RANDOM AS TransFile LEN = LEN(TaxTrans(1))
  
  IF UsingIndex AND NumOfTaxRecs > 0 THEN
    GOSUB GetAdNameIndex
  END IF
  
  ClearBack
  ShowProcessingScrn "Ad/Collection Cost Listing"
  
  FOR Cnt = 1 TO NumOfTaxRecs
    IF UsingIndex THEN
      CustRecNo = Array(Cnt).RecNum
    ELSE
      CustRecNo = Cnt
    END IF
    
    
    
    GET TaxFile, CustRecNo, TaxCustRec(1)
    
    IF NOT TaxCustRec(1).Deleted THEN
      IF LineCnt >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintAdRptHeader
      END IF
      
      
      TransRList = 0
      TransRecord& = TaxCustRec(1).LastTrans
      WHILE TransRecord& <> 0
        
        GET TransFile, TransRecord&, TaxTrans(1)
        CurTransRecord& = TransRecord&
        
        IF TaxTrans(1).TranType = 1 AND TaxTrans(1).TaxYear = TaxYear AND TaxCustRec(1).FirstPropRec > 0 THEN
          Balance# = TaxTrans(1).Revenue.Principle1 + TaxTrans(1).Revenue.Principle2 + TaxTrans(1).Revenue.Principle3 + TaxTrans(1).Revenue.Principle4 + TaxTrans(1).Revenue.Principle5
          Balance# = Balance# + TaxTrans(1).Revenue.INTEREST + TaxTrans(1).Revenue.PENALTY + TaxTrans(1).Revenue.Collection
          Balance# = Balance# - (TaxTrans(1).Revenue.Principle1Pd + TaxTrans(1).Revenue.Principle2Pd + TaxTrans(1).Revenue.Principle3Pd + TaxTrans(1).Revenue.Principle4Pd + TaxTrans(1).Revenue.Principle5Pd)
          Balance# = Balance# - (TaxTrans(1).Revenue.InterestPd + TaxTrans(1).Revenue.PenaltyPd + TaxTrans(1).Revenue.CollectionPd)
          Balance# = Round#(Balance#)
          TransRList = TransRList + 1
          
          IF Balance# > 0 THEN
            
            'Get Prop Desc Here
            PropRec& = TaxCustRec(1).FirstPropRec
            OpenTaxPropFile NumOfPropRecs, PropTaxFile
            HowMany = 0
            Z1 = 0
            WHILE PropRec& <> 0
              GET PropTaxFile, PropRec&, PropertyRec(1)
              Z1 = Z1 + 1
              PropDesc$(Z1) = QPTrim$(PropertyRec(1).PROPNOT1)
              PropRec& = PropertyRec(1).NextRec
            WEND
            IF TransRList = 1 THEN
              PropDesc$ = PropDesc$(Z1)
            END IF
            IF TransRList > 1 THEN
              Z2 = (Z1 - TransRList) + 1
              IF Z2 < 0 THEN Z2 = 1
              PropDesc$ = PropDesc$(Z2)
            END IF
            
            nme$ = QPTrim$(TaxCustRec(1).FName) + " " + QPTrim$(TaxCustRec(1).LName)
            nme$ = QPTrim$(nme$)                'this one cleans up those with only last name
            
            PRINT #RptHandle, LEFT$(nme$, 38);
            PRINT #RptHandle, TAB(40); LEFT$(PropDesc$, 29);
            PRINT #RptHandle, TAB(71); USING "######.##"; Balance#
            LineCnt = LineCnt + 1
            
            
            
            
            
            
            TotalCharged# = TotalCharged# + AmountCharged#
            
            CLOSE PropTaxFile
            
            
            IF AmountCharged# <> 0 THEN
              ' Now Add Amt to Bill and Put Back
              TaxTrans(1).Revenue.Collection = TaxTrans(1).Revenue.Collection + AmountCharged#
              PUT TransFile, TransRecord&, TaxTrans(1)
              
              
              
              TaxTrans(1).TransDate = Date2Num%(PostDate$)
              TaxTrans(1).TaxYear = TaxYear
              TaxTrans(1).TranType = 6          '1=Bill 2=Payment 3=Release 4=Interest 5=Penalty 6=Collection/Ad Cost Billing
              TaxTrans(1).BillType = "C"        'R=Real P=Personal Property C=Combined (NC/GA)
              TaxTrans(1).Amount = AmountCharged#               'Total Transaction Amount
              TaxTrans(1).Revenue.Principle1 = 0
              TaxTrans(1).Revenue.Principle2 = 0
              TaxTrans(1).Revenue.Principle3 = 0
              TaxTrans(1).Revenue.Principle4 = 0
              TaxTrans(1).Revenue.Principle5 = 0
              TaxTrans(1).Revenue.INTEREST = 0
              TaxTrans(1).Revenue.PENALTY = 0
              TaxTrans(1).Revenue.Collection = AmountCharged#
              TaxTrans(1).Revenue.Future1 = 0
              TaxTrans(1).Revenue.Future2 = 0
              TaxTrans(1).Revenue.Principle1Pd = 0
              TaxTrans(1).Revenue.Principle2Pd = 0
              TaxTrans(1).Revenue.Principle3Pd = 0
              TaxTrans(1).Revenue.Principle4Pd = 0
              TaxTrans(1).Revenue.Principle5Pd = 0
              TaxTrans(1).Revenue.InterestPd = 0
              TaxTrans(1).Revenue.PenaltyPd = 0
              TaxTrans(1).Revenue.CollectionPd = 0
              TaxTrans(1).Revenue.Future1Pd = 0
              TaxTrans(1).Revenue.Future2Pd = 0
              
              
              TaxTrans(1).Description = "Collection/Ad Charge"
              TaxTrans(1).Posted2GL = "N"
              TaxTrans(1).CustomerRec = CustRecNo
              TaxTrans(1).LastTrans = 0
              TaxTrans(1).BelongTo = CurTransRecord&
              TaxTrans(1).Padding = ""
              
              'Increment Transaction File Record Count
              NextRecord& = (LOF(TransFile) / LEN(TaxTrans(1))) + 1
              PUT #TransFile, NextRecord&, TaxTrans(1)
              
              'Update the Customer Pointers Now
              GET TaxFile, CustRecNo, TaxCustRec(1)
              
              IF TaxCustRec(1).LastTrans = 0 THEN
                TaxCustRec(1).LastTrans = NextRecord&
                PUT #TaxFile, CustRecNo, TaxCustRec(1)
              ELSE
                Previous& = TaxCustRec(1).LastTrans
                TaxCustRec(1).LastTrans = NextRecord&
                PUT #TaxFile, CustRecNo, TaxCustRec(1)
                
                GET #TransFile, NextRecord&, TaxTrans(1)
                TaxTrans(1).LastTrans = Previous&
                PUT #TransFile, NextRecord&, TaxTrans(1)
              END IF
            END IF
            'Now Grab Back the Original Transaction Holding the Bill to Get the Next One
            GET TransFile, CurTransRecord&, TaxTrans(1)
            Balance# = 0
          END IF
        END IF
        TransRecord& = TaxTrans(1).LastTrans
      WEND
      
      
      
      
      CustCnt = CustCnt + 1
    END IF
    
    ShowPctComp Cnt, NumOfTaxRecs
  NEXT
  
  GOSUB PrintAdRptEnding
  
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
  
PrintAdRptHeader:
  Page = Page + 1
  PRINT #RptHandle, TAB(20); "Property Tax Ad/Collection Listing"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
  PRINT #RptHandle, "Name"; TAB(40); "Prop Desc"; TAB(68); "Tax Balance"
  PRINT #RptHandle, Dash80$
  LineCnt = 3
  RETURN
  
  
PrintAdRptEnding:
  PRINT #RptHandle, Dash80$
  PRINT #RptHandle, "Total Ad/Collection Amount Charged: "; USING "$$######,#.##"; TotalCharged#
  
  PRINT #RptHandle,
  PRINT #RptHandle, FF$
  RETURN
  
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

SUB BalanceListing
  SHARED Choice$()
  DIM Balance#(100), Year(100), Yr(100), GBalance#(100), IBalance#(100), ABalance#(100)
  
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
          nme$ = QPTrim$(TaxCustRec(1).FName) + " " + QPTrim$(TaxCustRec(1).LName)
          nme$ = QPTrim$(nme$)  'this one cleans up those with only last name
          Balance# = 0
          'Open the Trans File
          TransFile = FREEFILE
          OPEN "TaxTrans.dat" FOR RANDOM AS TransFile LEN = LEN(TaxTrans(1))
          TransRecord& = TaxCustRec(1).LastTrans
          
          WHILE TransRecord& <> 0
            GET TransFile, TransRecord&, TaxTrans(1)
            IF TaxTrans(1).TranType = 1 AND (TaxYear$ = "ALL" OR VAL(TaxYear$) = TaxTrans(1).TaxYear) THEN
              Balance# = TaxTrans(1).Revenue.Principle1 + TaxTrans(1).Revenue.Principle2 + TaxTrans(1).Revenue.Principle3 + TaxTrans(1).Revenue.Principle4 + TaxTrans(1).Revenue.Principle5
              Balance# = Balance# + TaxTrans(1).Revenue.INTEREST + TaxTrans(1).Revenue.PENALTY + TaxTrans(1).Revenue.Collection
              Balance# = Balance# - (TaxTrans(1).Revenue.Principle1Pd + TaxTrans(1).Revenue.Principle2Pd + TaxTrans(1).Revenue.Principle3Pd + TaxTrans(1).Revenue.Principle4Pd + TaxTrans(1).Revenue.Principle5Pd)
              Balance# = Balance# - (TaxTrans(1).Revenue.InterestPd + TaxTrans(1).Revenue.PenaltyPd + TaxTrans(1).Revenue.CollectionPd)
              Balance# = Round#(Balance#)
              IF Balance# > 0 THEN
                PRINT #RptHandle, USING "#####"; CustRecNo;
                PRINT #RptHandle, TAB(10); LEFT$(nme$, 33);
                PRINT #RptHandle, TAB(45); LEFT$(TaxTrans(1).Description, 16);
                PRINT #RptHandle, TAB(61); TaxTrans(1).TaxYear;
                PRINT #RptHandle, TAB(68); USING "$$######,#.##"; Balance#
                PRINT #RptHandle, TAB(10); "Tax: "; USING "######,#.##"; Round#((TaxTrans(1).Revenue.Principle1 - TaxTrans(1).Revenue.Principle1Pd));
                PRINT #RptHandle, "  ";
                PRINT #RptHandle, "Int't: "; USING "######,#.##"; Round#((TaxTrans(1).Revenue.INTEREST - TaxTrans(1).Revenue.InterestPd));
                PRINT #RptHandle, "  ";
                PRINT #RptHandle, "Ad/Col: "; USING "######,#.##"; Round#((TaxTrans(1).Revenue.Collection - TaxTrans(1).Revenue.CollectionPd))
                TotalBalance# = TotalBalance# + Balance#
                InterestBalance# = InterestBalance# + Round#((TaxTrans(1).Revenue.INTEREST - TaxTrans(1).Revenue.InterestPd))
                AdBalance# = AdBalance# + Round#((TaxTrans(1).Revenue.Collection - TaxTrans(1).Revenue.CollectionPd))
                LineCnt = LineCnt + 2
                CustCnt = CustCnt + 1
                
                IF NCnt = 0 THEN
                  NCnt = 1
                  Yr(NCnt) = TaxTrans(1).TaxYear
                  GBalance#(NCnt) = Balance#
                  IBalance#(NCnt) = Round#((TaxTrans(1).Revenue.INTEREST - TaxTrans(1).Revenue.InterestPd))
                  ABalance#(NCnt) = Round#((TaxTrans(1).Revenue.Collection - TaxTrans(1).Revenue.CollectionPd))
                  
                  GOTO skipem1
                END IF
                
                FOR LCnt = 1 TO NCnt
                  UpFlag = 0    'Set Flag to Not Updated
                  IF Yr(LCnt) = TaxTrans(1).TaxYear THEN
                    GBalance#(LCnt) = GBalance#(LCnt) + Balance#
                    IBalance#(LCnt) = IBalance#(LCnt) + Round#((TaxTrans(1).Revenue.INTEREST - TaxTrans(1).Revenue.InterestPd))
                    ABalance#(LCnt) = ABalance#(LCnt) + Round#((TaxTrans(1).Revenue.Collection - TaxTrans(1).Revenue.CollectionPd))
                    GOTO skipem1
                  END IF
                NEXT LCnt
                NCnt = NCnt + 1
                YCnt = NCnt
                GBalance#(NCnt) = Balance#
                IBalance#(NCnt) = Round#((TaxTrans(1).Revenue.INTEREST - TaxTrans(1).Revenue.InterestPd))
                ABalance#(NCnt) = Round#((TaxTrans(1).Revenue.Collection - TaxTrans(1).Revenue.CollectionPd))
                Yr(NCnt) = TaxTrans(1).TaxYear
skipem1:
              END IF
            END IF
            TransRecord& = TaxTrans(1).LastTrans
          WEND
          CLOSE TransFile
          
        ELSE
          
          'Summary Format
          'Get Name First
          nme$ = QPTrim$(TaxCustRec(1).FName) + " " + QPTrim$(TaxCustRec(1).LName)
          nme$ = QPTrim$(nme$)  'this one cleans up those with only last name
          
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
              Balance# = Balance# + TaxTrans(1).Revenue.INTEREST + TaxTrans(1).Revenue.PENALTY + TaxTrans(1).Revenue.Collection
              Balance# = Balance# - (TaxTrans(1).Revenue.Principle1Pd + TaxTrans(1).Revenue.Principle2Pd + TaxTrans(1).Revenue.Principle3Pd + TaxTrans(1).Revenue.Principle4Pd + TaxTrans(1).Revenue.Principle5Pd)
              Balance# = Balance# - (TaxTrans(1).Revenue.InterestPd + TaxTrans(1).Revenue.PenaltyPd + TaxTrans(1).Revenue.CollectionPd)
              Balance# = Round#(Balance#)
            ELSE
              Balance# = 0
            END IF
            
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
              PRINT #RptHandle, TAB(10); LEFT$(nme$, 33);
              PRINT #RptHandle, TAB(61); Year(PCnt);
              PRINT #RptHandle, TAB(68); USING "$$######,#.##"; Balance#(PCnt)
              TotalBalance# = TotalBalance# + Balance#(PCnt)
              
              'Update Year Balance
              Updated = 0
              Year = Year(PCnt)
              IF YCnt = 0 THEN
                YCnt = 1
                Yr(YCnt) = Year
                GBalance#(YCnt) = GBalance#(YCnt) + Balance#(PCnt)
              ELSE
                FOR TCnt = 1 TO YCnt
                  IF Year = Yr(TCnt) THEN
                    GBalance#(TCnt) = GBalance#(TCnt) + Balance#(PCnt)
                    Updated = 1
                  END IF
                NEXT TCnt
                IF Updated = 0 THEN
                  YCnt = YCnt + 1
                  Yr(YCnt) = Year
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
  Page = Page + 1
  PRINT #RptHandle, TAB(21); "Property Tax Customer Balance Listing"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
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
      PRINT #RptHandle, Yr(lll); TAB(12); USING "$$#####,#.##"; GBalance#(lll) - (IBalance#(lll) + ABalance#(lll));
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
    PRINT #RptHandle, Yr(lll); TAB(15); USING "$$#######,#.##"; GBalance#(lll)
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
      Balance# = Balance# + TaxTrans(1).Revenue.INTEREST + TaxTrans(1).Revenue.PENALTY + TaxTrans(1).Revenue.Collection
      Balance# = Balance# - (TaxTrans(1).Revenue.Principle1Pd + TaxTrans(1).Revenue.Principle2Pd + TaxTrans(1).Revenue.Principle3Pd + TaxTrans(1).Revenue.Principle4Pd + TaxTrans(1).Revenue.Principle5Pd)
      Balance# = Balance# - (TaxTrans(1).Revenue.InterestPd + TaxTrans(1).Revenue.PenaltyPd + TaxTrans(1).Revenue.CollectionPd)
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
26090 IF Yr(i) <= Yr(j) THEN 26120
26100 SW = 1
26110 SWAP Yr(i), Yr(j)
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
  MACHValue# = 0
  
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
      MACHValue# = Round#(MACHValue# + PersRec(1).MTVALUE)
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
  TaxInq(1).ADDR1 = TaxCustRec(1).ADDR1
  TaxInq(1).ADDR2 = TaxCustRec(1).ADDR2
  TaxInq(1).CITY = TaxCustRec(1).CITY
  TaxInq(1).STATE = TaxCustRec(1).STATE
  TaxInq(1).ZIP = TaxCustRec(1).ZIP
  TaxInq(1).ACTIVE = TaxCustRec(1).ACTIVE
  TaxInq(1).INTEREST = TaxCustRec(1).INTEREST
  TaxInq(1).EXEMPT = TaxCustRec(1).TaxExempt
  TaxInq(1).PENALTY = TaxCustRec(1).PENALTY
  TaxInq(1).SRCITDIS = SeniExmp#
  TaxInq(1).ODISCOUN = OthrExmp#
  TaxInq(1).PROPVALU = RealValue#
  TaxInq(1).PERSVAL = PersValue#
  TaxInq(1).MHVALUE = MOBHValue#
  TaxInq(1).MCVALUE = MERHValue#
  TaxInq(1).CVALUE = FarmValue#
  TaxInq(1).MTVALUE = MACHValue#
  TaxInq(1).Balance = GetCustBalance#(CustAcct&)
  
  BCopy VARSEG(TaxInq(1)), VARPTR(TaxInq(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), LEN(Form$(0, 0)), 0
  UnPackBuffer 0, 0, Form$(), Fld()
  Action = 1
  RETURN
END SUB

SUB CustomerListing
  
  SHARED Choice$()
  
  REDIM Array(1 TO 1) AS Struct 'Template for the sort array
  ReportFile$ = "TaxCust.PRN"   'Report File Name
  Dash80$ = STRING$(80, "=")
  FF$ = CHR$(12)
  
  
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
  
  REDIM Choice$(0 TO 2, 0 TO 2)
  
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Name Order"
  Choice$(2, 0) = "Account Number"
  Choice$(0, 1) = "2"
  Choice$(1, 1) = "Summary"
  Choice$(2, 1) = "Detail"
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
      IF LEFT$(Form$(2, 0), 1) = "D" THEN
        DetailFlag = True
        MaxLines = 45
      ELSE
        DetailFlag = False
        MaxLines = 50
      END IF
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
  
  GOSUB PrintDetailCustomerRptHeader
  
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile
  
  IF UsingIndex AND NumOfTaxRecs > 0 THEN
    GOSUB GetNameIndex
  END IF
  
  ClearBack
  ShowProcessingScrn "Master Customer Listing"
  
  FOR Cnt = 1 TO NumOfTaxRecs
    IF UsingIndex THEN
      CustRecNo = Array(Cnt).RecNum
    ELSE
      CustRecNo = Cnt
    END IF
    
    GET TaxFile, CustRecNo, TaxCustRec(1)
    
    IF TaxCustRec(1).Deleted = 0 THEN
      
    ELSE
      TaxCustRec(1).Deleted = -1
      PUT TaxFile, CustRecNo, TaxCustRec(1)
    END IF
    
    IF NOT TaxCustRec(1).Deleted THEN
      IF LineCnt >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintDetailCustomerRptHeader
      END IF



      nme$ = QPTrim$(TaxCustRec(1).FName) + " " + QPTrim$(TaxCustRec(1).LName)
      nme$ = QPTrim$(nme$)      'this one cleans up those with only last name
      
      IF NOT DetailFlag THEN
        PRINT #RptHandle, USING "#####"; CustRecNo;
        PRINT #RptHandle, TAB(10); nme$; TAB(60); TaxCustRec(1).ACTIVE; TAB(65); TaxCustRec(1).pin
        LineCnt = LineCnt + 1
      ELSE
        PRINT #RptHandle, USING "  Acct: ####"; CustRecNo;
        PRINT #RptHandle, TAB(15); nme$
        PRINT #RptHandle, "Active: "; TaxCustRec(1).ACTIVE; TAB(15); TaxCustRec(1).ADDR1
        PRINT #RptHandle, "Int'st: "; TaxCustRec(1).INTEREST; TAB(15); TaxCustRec(1).ADDR2
        PRINT #RptHandle, "Exempt: "; TaxCustRec(1).TaxExempt; TAB(15); RTRIM$(TaxCustRec(1).CITY) + ", "; RTRIM$(TaxCustRec(1).STATE) + "  " + RTRIM$(TaxCustRec(1).ZIP)
        
        PRINT #RptHandle, ""
        LineCnt = LineCnt + 5
        
        'Now Show Property Records Next
        
        IF TaxCustRec(1).FirstPropRec > 0 THEN
          
          PropertyRecord! = TaxCustRec(1).FirstPropRec
          
          WHILE PropertyRecord! <> 0
            
            IF LineCnt >= MaxLines THEN
              PRINT #RptHandle, FF$
              GOSUB PrintDetailCustomerRptHeader
              PRINT #RptHandle, USING "  Acct: ####"; CustRecNo;
              PRINT #RptHandle, TAB(15); nme$
              LineCnt = LineCnt + 2
            END IF
            
            PRINT #RptHandle, "Property Owned..."
            GET #PropTaxFile, PropertyRecord!, PropertyRec(1)
             'Reset Property Link
             PropertyRec(1).CustPin = CustRecNo
             PUT #PropTaxFile, PropertyRecord!, PropertyRec(1)

            PRINT #RptHandle, "Pin # "; QPTrim$(PropertyRec(1).REALPIN); TAB(50); " Value: "; USING "$$########,#.##"; PropertyRec(1).PROPVALU
            PRINT #RptHandle, "Desc  "; QPTrim$(PropertyRec(1).PROPNOT1); TAB(50); "Exempt: "; USING "$$########,#.##"; PropertyRec(1).EXMPSENI
            PRINT #RptHandle, "Desc  "; QPTrim$(PropertyRec(1).PROPNOT2); TAB(50); "Oth Ex: "; ; USING "$$########,#.##"; PropertyRec(1).EXMPOTHR
            PRINT #RptHandle, "Desc  "; QPTrim$(PropertyRec(1).PROPNOT3); TAB(50); "Mortgage Code: "; QPTrim$(PropertyRec(1).MORTCODE)
            PRINT #RptHandle, "Map/Blk/Lot - "; QPTrim$(PropertyRec(1).MAP); "/"; QPTrim$(PropertyRec(1).BLOCK); "/"; QPTrim$(PropertyRec(1).LOTNUMB); TAB(40); "Late (Y/N): "; PropertyRec(1).LATELIST
            PRINT #RptHandle, STRING$(79, "-")
            LineCnt = LineCnt + 6
            OldRecord! = PropertyRecord!
            PropertyRecord! = PropertyRec(1).NextRec
            IF OldRecord! = PropertyRecord! THEN PropertyRecord! = 0
            
          WEND
        END IF
        
        
        'NOW CHECK PERSONAL PROPERTY
        IF TaxCustRec(1).FirstPersRec > 0 THEN
          
          PropertyRecord! = TaxCustRec(1).FirstPersRec
          
          WHILE PropertyRecord! <> 0
            
            IF LineCnt >= MaxLines THEN
              PRINT #RptHandle, FF$
              GOSUB PrintDetailCustomerRptHeader
              PRINT #RptHandle, USING "  Acct: ####"; CustRecNo;
              PRINT #RptHandle, TAB(15); nme$
              LineCnt = LineCnt + 2
            END IF
            
            PRINT #RptHandle, "Personal Property Owned..."
            GET #PersTaxFile, PropertyRecord!, PersRec(1)
            PersRec(1).CustPin = CustRecNo
            PUT #PersTaxFile, PropertyRecord!, PersRec(1)
            PRINT #RptHandle, "Pin # "; QPTrim$(PersRec(1).PROPPIN); TAB(50); " PP Value: "; USING "$$########,#.##"; PersRec(1).PERSVAL
            PRINT #RptHandle, "Desc  "; QPTrim$(PersRec(1).DESC1); TAB(50); " MH Value: "; USING "$$########,#.##"; PersRec(1).MHVALUE
            PRINT #RptHandle, "Desc  "; QPTrim$(PersRec(1).DESC2); TAB(50); " MC Value: "; USING "$$########,#.##"; PersRec(1).MCVALUE
            PRINT #RptHandle, "Desc  "; QPTrim$(PersRec(1).DESC3); TAB(50); " CV Value: "; USING "$$########,#.##"; PersRec(1).CVALUE
            PRINT #RptHandle, "Desc  "; QPTrim$(PersRec(1).DESC4); TAB(50); " MT Value: "; USING "$$########,#.##"; PersRec(1).MTVALUE
            PRINT #RptHandle, "Desc  "; QPTrim$(PersRec(1).DESC5); TAB(50); "Exempt: "; USING "$$########,#.##"; PersRec(1).EXMPSENI
            PRINT #RptHandle, STRING$(79, "-")
            LineCnt = LineCnt + 7
            OldRecord! = PropertyRecord!
            PropertyRecord! = PersRec(1).NextRec
            IF OldRecord! = PropertyRecord! THEN PropertyRecord! = 0
          WEND
        END IF
      END IF
      CustCnt = CustCnt + 1
    END IF
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
  
  KILL ReportFile$
  
  EXIT SUB
  
PrintDetailCustomerRptHeader:
  Page = Page + 1
  PRINT #RptHandle, TAB(20); "Property Tax Detailed Customer Listing"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
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
  
END SUB

SUB Labels
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
      
      nme$ = QPTrim$(TaxCustRec(1).FName) + " " + QPTrim$(TaxCustRec(1).LName)
      nme$ = QPTrim$(nme$)      'this one cleans up those with only last name
      
      
      
      
      'First Show Property Records
      
      IF TaxCustRec(1).FirstPropRec > 0 THEN
        PropertyRecord! = TaxCustRec(1).FirstPropRec
        WHILE PropertyRecord! <> 0
          GET #PropTaxFile, PropertyRecord!, PropertyRec(1)
          IF PropertyRec(1).LATELIST = "Y" THEN
            PRINT #RptHandle, USING "######"; CustRecNo;
            PRINT #RptHandle, TAB(10); nme$; TAB(57); USING "$$#######,#"; PropertyRec(1).PROPVALU
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
            PRINT #RptHandle, TAB(10); nme$; TAB(57); USING "$$#######,#"; PValue#
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
  Page = Page + 1
  PRINT #RptHandle, TAB(24); "Tax Customer Late Listing Report"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
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

SUB MasterValueList
  SHARED Choice$()
  
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
  
  REDIM Choice$(0 TO 3, 0 TO 3)
  
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
      Ptype$ = Form$(3, 0)
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
  
  GOSUB PrintMasterValueHeader
  
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile
  
  IF UsingIndex AND NumOfTaxRecs > 0 THEN
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
    
    GET TaxFile, CustRecNo, TaxCustRec(1)
    
    PValue# = 0
    IF NOT TaxCustRec(1).Deleted THEN
      IF LineCnt >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintMasterValueHeader
      END IF
      nme$ = QPTrim$(TaxCustRec(1).FName) + " " + QPTrim$(TaxCustRec(1).LName)
      nme$ = QPTrim$(nme$)      'this one cleans up those with only last name
      IF NOT DetailFlag THEN
        'Figure Values
        'Real Value First
        RealValue# = 0
        Discnt# = 0
        IF TaxCustRec(1).FirstPropRec > 0 THEN
          PropertyRecord! = TaxCustRec(1).FirstPropRec
          WHILE PropertyRecord! <> 0
            GET PropTaxFile, PropertyRecord!, PropertyRec(1)
            RealValue# = RealValue# + PropertyRec(1).PROPVALU
            IF LEFT$(Ptype$, 1) <> "P" THEN
              Discnt# = Discnt# + PropertyRec(1).EXMPSENI
              Discnt# = Discnt# + PropertyRec(1).EXMPOTHR
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
            PersValue# = PersValue# + PersRec(1).PERSVAL
            PersValue# = PersValue# + PersRec(1).MHVALUE
            PersValue# = PersValue# + PersRec(1).MCVALUE
            PersValue# = PersValue# + PersRec(1).CVALUE
            PersValue# = PersValue# + PersRec(1).MTVALUE
            IF LEFT$(Ptype$, 1) <> "R" THEN
              Discnt# = Discnt# + PersRec(1).EXMPSENI
            END IF
            PropertyRecord! = PersRec(1).NextRec
          WEND
        END IF
        
        IF LEFT$(Ptype$, 1) = "R" THEN
          PersValue# = 0
        END IF
        IF LEFT$(Ptype$, 1) = "P" THEN
          RealValue# = 0
        END IF
        IF LEN(nme$) = 0 AND RealValue# = 0 AND PersValue# = 0 THEN
         TaxCustRec(1).Deleted = -1
         PUT TaxFile, CustRecNo, TaxCustRec(1)
         GOTO skipem10
        END IF

        PRINT #RptHandle, USING "#####"; CustRecNo;
        PRINT #RptHandle, TAB(8); LEFT$(nme$, 27);
        PRINT #RptHandle, TAB(37); USING "########,#"; RealValue#;
        PRINT #RptHandle, TAB(48); USING "########,#"; PersValue#;
        PRINT #RptHandle, TAB(59); USING "########,#"; Discnt#;
        PRINT #RptHandle, TAB(70); USING "########,#"; RealValue# + PersValue# - Discnt#
        LineCnt = LineCnt + 1
        TotalReal# = TotalReal# + RealValue#
        TotalPers# = TotalPers# + PersValue#
        TotalDisc# = TotalDisc# + Discnt#
      ELSE
        'Figure Values
        'Real Value First
        RealValue# = 0
        Discnt# = 0
        IF TaxCustRec(1).FirstPropRec > 0 THEN
          PropertyRecord! = TaxCustRec(1).FirstPropRec
          WHILE PropertyRecord! <> 0
            GET PropTaxFile, PropertyRecord!, PropertyRec(1)
            RealValue# = RealValue# + PropertyRec(1).PROPVALU
            IF LEFT$(Ptype$, 1) <> "P" THEN
              Discnt# = Discnt# + PropertyRec(1).EXMPSENI
              Discnt# = Discnt# + PropertyRec(1).EXMPOTHR
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
            PersValue# = PersValue# + PersRec(1).PERSVAL
            PersValue# = PersValue# + PersRec(1).MHVALUE
            PersValue# = PersValue# + PersRec(1).MCVALUE
            PersValue# = PersValue# + PersRec(1).CVALUE
            PersValue# = PersValue# + PersRec(1).MTVALUE
            IF LEFT$(Ptype$, 1) <> "R" THEN
              Discnt# = Discnt# + PersRec(1).EXMPSENI
            END IF
            PropertyRecord! = PersRec(1).NextRec
          WEND
        END IF
        
        IF LEFT$(Ptype$, 1) = "R" THEN
          PersValue# = 0
        END IF
        IF LEFT$(Ptype$, 1) = "P" THEN
          RealValue# = 0
        END IF
        IF LEN(nme$) = 0 AND RealValue# = 0 AND PersValue# = 0 THEN
         TaxCustRec(1).Deleted = -1
         PUT TaxFile, CustRecNo, TaxCustRec(1)
         GOTO skipem10
        END IF

        PRINT #RptHandle, USING "#####"; CustRecNo;
        PRINT #RptHandle, TAB(8); LEFT$(nme$, 27);
        PRINT #RptHandle, TAB(37); USING "########,#"; RealValue#;
        PRINT #RptHandle, TAB(48); USING "########,#"; PersValue#;
        PRINT #RptHandle, TAB(59); USING "########,#"; Discnt#;
        PRINT #RptHandle, TAB(70); USING "########,#"; RealValue# + PersValue# - Discnt#
        TotalReal# = TotalReal# + RealValue#
        TotalPers# = TotalPers# + PersValue#
        TotalDisc# = TotalDisc# + Discnt#
        
        LineCnt = LineCnt + 1
        
        'Now Show Detail Support Here
        IF LEFT$(Ptype$, 1) = "P" THEN
        ELSE
          IF TaxCustRec(1).FirstPropRec > 0 THEN
            PropertyRecord! = TaxCustRec(1).FirstPropRec
            PFlag = 0
            WHILE PropertyRecord! <> 0
              GET PropTaxFile, PropertyRecord!, PropertyRec(1)
              PRINT #RptHandle, TAB(15); "Property Pin# "; QPTrim$(PropertyRec(1).REALPIN);
              PRINT #RptHandle, TAB(52); "Value: "; USING "#######,#"; PropertyRec(1).PROPVALU
              LineCnt = LineCnt + 1
              PFlag = 1
              PropertyRecord! = PropertyRec(1).NextRec
            WEND
          END IF
        END IF
        IF LEFT$(Ptype$, 1) = "R" THEN
        ELSE
          IF TaxCustRec(1).FirstPersRec > 0 THEN
            PropertyRecord! = TaxCustRec(1).FirstPersRec
            WHILE PropertyRecord! <> 0
              GET PersTaxFile, PropertyRecord!, PersRec(1)
              PValue# = PValue# + PersRec(1).PERSVAL
              PValue# = PValue# + PersRec(1).MHVALUE
              PValue# = PValue# + PersRec(1).MCVALUE
              PValue# = PValue# + PersRec(1).CVALUE
              PValue# = PValue# + PersRec(1).MTVALUE
              PRINT #RptHandle, TAB(15); "Pers Abstract# "; PersRec(1).PROPPIN;
              PRINT #RptHandle, TAB(52); "Value: "; USING "#######,#"; PValue#
              LineCnt = LineCnt + 1
              PFlag = 1
              PropertyRecord! = PersRec(1).NextRec
            WEND
          END IF
        END IF
      END IF
          CustCnt = CustCnt + 1
    END IF
    
    IF PFlag = 1 THEN
      PRINT #RptHandle, ""
      PFlag = 0
    END IF


skipem10:
    ShowPctComp Cnt, NumOfTaxRecs
  NEXT Cnt
  
  
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
  
  KILL ReportFile$
  
  EXIT SUB
  
PrintMasterValueHeader:
  Page = Page + 1
  PRINT #RptHandle, TAB(25); "Property Tax Valuation Listing"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
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
  PRINT #RptHandle, "Net Value : "; USING "$$##########,#"; TotalReal# + TotalPers# - TotalDisc#
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
  
END SUB

SUB MortgageCodeList
END SUB

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

SUB TransactionJournal
  SHARED Choice$()
  
  REDIM Array(1 TO 1) AS Struct 'Template for the sort array
  ReportFile$ = "TRANJOUR.PRN"  'Report File Name
  Dash80$ = STRING$(80, "=")
  FF$ = CHR$(12)
  
  REDIM TranRec(1) AS TaxTransactionType
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
      BDate$ = Form$(2, 0): BDate = Date2Num%(BDate$)
      EDate$ = Form$(3, 0): EDate = Date2Num%(EDate$)
      TaxType$ = LEFT$(Form$(4, 0), 1)
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
  
  GOSUB PrintTransJourRptHeader

  OpenTaxCustFile 0&, TaxFile
  CustRecLen = LEN(TaxCustRec(1))

  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
  NumOfTranRecs& = LOF(TaxTran) / TaxTranRecLen
  
  ClearBack
  ShowProcessingScrn "Transaction Journal Listing"
  
  FOR TCnt& = NumOfTranRecs& TO 1 STEP -1
    GET TaxTran, TCnt&, TranRec(1)
    CustBeenPrinted = False
    IF LineCnt >= MaxLines THEN
      PRINT #RptHandle, FF$
      GOSUB PrintTransJourRptHeader
    END IF
    IF TranRec(1).TransDate >= BDate AND TranRec(1).TransDate <= EDate THEN
      IF TrType = 0 OR TranRec(1).TranType = TrType THEN
        IF TaxType$ = TranRec(1).BillType AND (TrType = 1 OR TrType = 0) THEN

          IF NOT CustBeenPrinted THEN
            IF TranRec(1).CustomerRec > 0 THEN
              GET TaxFile, TranRec(1).CustomerRec, TaxCustRec(1)
              PRINT #RptHandle, QPTrim$(TaxCustRec(1).FName); " "; QPTrim$(TaxCustRec(1).LName)
              LineCnt = LineCnt + 1
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
            GET TaxTran, TranRec(1).BelongTo, TranRec(1)
            IF TaxType$ = TranRec(1).BillType THEN
            IF NOT CustBeenPrinted THEN
              IF TranRec(1).CustomerRec > 0 THEN
                GET TaxFile, TranRec(1).CustomerRec, TaxCustRec(1)
                PRINT #RptHandle, QPTrim$(TaxCustRec(1).FName); " "; QPTrim$(TaxCustRec(1).LName)
                CustBeenPrinted = True
                LineCnt = LineCnt + 1
              END IF
            END IF

              GET TaxTran, TCnt&, TranRec(1)
              PRINT #RptHandle, Num2Date$(TranRec(1).TransDate);
              PRINT #RptHandle, TAB(12); TranRec(1).Description;
              IF TranRec(1).TaxYear > 0 THEN
                PRINT #RptHandle, TAB(50); TranRec(1).TaxYear;
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
    ShowPctComp TCnt&, NumOfTranRecs&
  NEXT
  
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
  
  ERASE Array, Frm, Form$, Fld, TaxCustRec
  
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  KILL ReportFile$
  
  EXIT SUB
  
PrintTransJourRptHeader:
  Page = Page + 1
  PRINT #RptHandle, TAB(20); "Property Tax Detailed Journal Listing"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
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
  
PrintTransJourRptEnding:
  PRINT #RptHandle, Dash80$
  PRINT #RptHandle, "Total Amount: "; USING "$$#######,#.##"; TotalAmt#
  PRINT #RptHandle,
  PRINT #RptHandle, FF$
  RETURN
  
  
  
END SUB

