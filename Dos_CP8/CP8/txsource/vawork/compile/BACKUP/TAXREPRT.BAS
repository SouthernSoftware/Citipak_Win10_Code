DEFINT A-Z
DECLARE SUB CustHistoryRpt1 ()
DECLARE SUB CustHistoryRpt2 ()
DECLARE SUB CustomerInquiry2 ()
DECLARE SUB CustomerInquiry1 ()
DECLARE SUB OpenPPTaxCustFile (NumOfTaxRecs%, TaxFile%)
DECLARE SUB OpenRETaxCustFile (NumOfTaxRecs%, TaxFile%)
DECLARE SUB CustHistoryRpt ()
DECLARE SUB GetPropRecList (PropRecs() AS LONG, CustRec&)
DECLARE SUB GetPersRecList (PersRecs() AS LONG, CustRec&)
DECLARE SUB LookUp (RecNo&, Text$, ChkBalFlag%, CLSFlag%, SSNFlag%, TaxType%)
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


  CONST False = 0, True = NOT False
  
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'Taxfiles.BI'
  '$INCLUDE: 'PROPAbst.BI'

  '$INCLUDE: 'TAXRPTTY.BI'
  
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
  
  MChoice$(1) = "Customer Inquiry - Real Estate"
  MChoice$(2) = "Customer Inquiry - Personal Property"
  MChoice$(3) = "Customer Transaction History - Real"
  MChoice$(4) = "Customer Transaction History - Pers"
  MChoice$(5) = "Master Customer Listing"
  MChoice$(6) = "Master Balance Listing"
  MChoice$(7) = "Master Valuation Listing"
  MChoice$(8) = "Transaction Journal"
  MChoice$(9) = "Mailing Labels"
  MChoice$(10) = "Exit to DOS"
  
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
      CustomerInquiry1
    CASE 2
      CustomerInquiry2
    CASE 3
      CustHistoryRpt1
    CASE 4
      CustHistoryRpt2
    CASE 5
      CustomerListing
    CASE 6
      BalanceListing
    CASE 7
      MasterValueList
    CASE 8
      TransactionJournal
    CASE 9
      Labels
    CASE 10
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

SUB BalanceListing
  SHARED Choice$()
  DIM Balance#(100), Year(100)
  REDIM Array(1 TO 1) AS Struct 'Template for the sort array
  ReportFile$ = "TaxBal.PRN"    'Report File Name
  Dash80$ = STRING$(80, "=")
  FF$ = CHR$(12)
  
  MaxLines = 53
  LineCnt = 0
  CustCnt = 0
  
  LibName$ = "TAX"
  ScrnName$ = "VABALRPT"
  
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
  Choice$(0, 1) = "3"
  Choice$(1, 1) = "Summary"
  Choice$(2, 1) = "Detail"
  Choice$(0, 2) = "4"
  Choice$(1, 2) = "Screen"
  Choice$(2, 2) = "Printer"

  Form$(2, 0) = "R"
  
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
      TaxType$ = Form$(2, 0)
      IF LEFT$(Form$(3, 0), 1) = "D" THEN
        DetailFlag = True
      ELSE
        DetailFlag = False
      END IF
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
  IF TaxType$ = "R" THEN
    OpenRETaxCustFile NumOfTaxRecs, TaxFile
   ELSE
    OpenPPTaxCustFile NumOfTaxRecs, TaxFile
  END IF
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
          Nme$ = QPTrim$(TaxCustRec(1).FNAME) + " " + QPTrim$(TaxCustRec(1).LName)
          Nme$ = QPTrim$(Nme$)  'this one cleans up those with only last name
          
          'Open the Trans File
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
              IF Balance# > 0 THEN
                PRINT #RptHandle, USING "#####"; CustRecNo;
                PRINT #RptHandle, TAB(10); LEFT$(Nme$, 33);
                PRINT #RptHandle, TAB(45); LEFT$(TaxTrans(1).Description, 16);
                PRINT #RptHandle, TAB(61); TaxTrans(1).TaxYear;
                PRINT #RptHandle, TAB(68); USING "$$######,#.##"; Balance#


               IF TaxType$ = "P" THEN
                PRINT #RptHandle, TAB(10); "PP Tax: "; USING "#####.##"; Round#((TaxTrans(1).Revenue.Principle1 - TaxTrans(1).Revenue.Principle1Pd));
                PRINT #RptHandle, "  ";
                PRINT #RptHandle, "MT Tax: "; USING "#####.##"; Round#((TaxTrans(1).Revenue.Principle2 - TaxTrans(1).Revenue.Principle2Pd));
                PRINT #RptHandle, "  ";
                PRINT #RptHandle, "MC Tax: "; USING "#####.##"; Round#((TaxTrans(1).Revenue.Principle3 - TaxTrans(1).Revenue.Principle3Pd));
                PRINT #RptHandle, "  ";
                PRINT #RptHandle, "FE Tax: "; USING "#####.##"; Round#((TaxTrans(1).Revenue.Principle4 - TaxTrans(1).Revenue.Principle4Pd))
                
                PRINT #RptHandle, TAB(10); "MH Tax: "; USING "#####.##"; Round#((TaxTrans(1).Revenue.Principle5 - TaxTrans(1).Revenue.Principle5Pd));
                PRINT #RptHandle, "  ";
                PRINT #RptHandle, " Int't: "; USING "#####.##"; Round#((TaxTrans(1).Revenue.INTEREST - TaxTrans(1).Revenue.InterestPd));
                PRINT #RptHandle, "  ";
                PRINT #RptHandle, "    Pen: "; USING "#####.##"; Round#((TaxTrans(1).Revenue.PENALTY - TaxTrans(1).Revenue.PenaltyPd))
                TotalBalance# = TotalBalance# + Balance#
                PRINT #RptHandle, STRING$(80, "-")
                LineCnt = LineCnt + 4
                CustCnt = CustCnt + 1
                ELSE
                PRINT #RptHandle, TAB(10); "Real Tax: "; USING "#####.##"; Round#((TaxTrans(1).Revenue.Principle1 - TaxTrans(1).Revenue.Principle1Pd));
                PRINT #RptHandle, "  ";
                PRINT #RptHandle, " Int't: "; USING "#####.##"; Round#((TaxTrans(1).Revenue.INTEREST - TaxTrans(1).Revenue.InterestPd));
                PRINT #RptHandle, "  ";
                PRINT #RptHandle, "    Pen: "; USING "#####.##"; Round#((TaxTrans(1).Revenue.PENALTY - TaxTrans(1).Revenue.PenaltyPd))
                TotalBalance# = TotalBalance# + Balance#
                PRINT #RptHandle, STRING$(80, "-")
                LineCnt = LineCnt + 3
                CustCnt = CustCnt + 1
               END IF
                
              END IF
            END IF
            TransRecord& = TaxTrans(1).LastTrans
          WEND
          CLOSE TransFile
          
        ELSE
          
          'Summary Format
          'Get Name First
          Nme$ = QPTrim$(TaxCustRec(1).FNAME) + " " + QPTrim$(TaxCustRec(1).LName)
          Nme$ = QPTrim$(Nme$)  'this one cleans up those with only last name
          
          
          'Open the Trans File
          TransFile = FREEFILE
          OPEN "TaxTrans.dat" FOR RANDOM AS TransFile LEN = LEN(TaxTrans(1))
          TransRecord& = TaxCustRec(1).LastTrans
          
          WHILE TransRecord& <> 0
            GET TransFile, TransRecord&, TaxTrans(1)
            
            IF TaxTrans(1).TranType = 1 THEN
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
                GOTO SkipEm
              END IF
              
              FOR LCnt = 1 TO NCnt
                UpFlag = 0      'Set Flag to Not Updated
                IF Year(LCnt) = TaxTrans(1).TaxYear THEN
                  Balance#(LCnt) = Balance#(LCnt) + Balance#
                  GOTO SkipEm
                END IF
              NEXT LCnt
              
              
              NCnt = NCnt + 1
              Balance#(NCnt) = Balance#
              Year(NCnt) = TaxTrans(1).TaxYear
            END IF
            
            
SkipEm:
            
            TransRecord& = TaxTrans(1).LastTrans
          WEND
          CLOSE TransFile
          
          FOR PCnt = 1 TO NCnt
            IF Balance#(PCnt) > 0 THEN
              PRINT #RptHandle, USING "#####"; CustRecNo;
              PRINT #RptHandle, TAB(10); LEFT$(Nme$, 33);
              PRINT #RptHandle, TAB(61); Year(PCnt);
              PRINT #RptHandle, TAB(68); USING "$$######,#.##"; Balance#(PCnt)
              TotalBalance# = TotalBalance# + Balance#(PCnt)
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
  PRINT #RptHandle, Dash80$
  PRINT #RptHandle, "Total Customer Lines Printed: "; USING "#####"; CustCnt
  PRINT #RptHandle, "               Total Balance: "; USING "$$########,#.##"; TotalBalance#
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
  
END SUB

SUB CustomerInquiry1

  SHARED Choice$()

  REDIM TaxInq(1) AS RETaxCustInqType

  TaxType% = 1
  ClearBack
  LookUp RecNo&, "Customer Inquiry", False, True, False, TaxType%
  IF RecNo& <= 0 THEN
    GOTO CustInqExit
  END IF

  FirstTime = True

  LibName$ = "TAX"
  ScrnName$ = "RCUSTINQ"

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

  OpenRETaxCustFile NumOfTaxRecs, TaxFile
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
      BldgValue# = Round#(SeniExmp# + PropertyRec(1).EXMPSENI)
      OthrExmp# = Round#(OthrExmp# + PropertyRec(1).EXMPOTHR)
    NEXT
  END IF
  CLOSE PropTaxFile


  TaxInq(1).ACCT = RecNo&
  TaxInq(1).OPENDATE = TaxCustRec(1).OPENDATE
  TaxInq(1).CUSTNAME = QPTrim$(TaxCustRec(1).FNAME) + " " + QPTrim$(TaxCustRec(1).LName)
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
  TaxInq(1).ODISCOUN = OthrExmp#
  TaxInq(1).PROPVALU = RealValue#
  TaxInq(1).PERSVAL = BldgValue#

  BCopy VARSEG(TaxInq(1)), VARPTR(TaxInq(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), LEN(Form$(0, 0)), 0
  UnPackBuffer 0, 0, Form$(), Fld()
  Action = 1
RETURN
END SUB

SUB CustomerInquiry2

  SHARED Choice$()

  REDIM TaxInq(1) AS TaxCustInqType

  TaxType% = 2 'Personal Property

  ClearBack
  LookUp RecNo&, "Customer Inquiry", False, True, False, TaxType%
  IF RecNo& <= 0 THEN
    GOTO CustInqExit2
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
      GOSUB LoadInqData2
      Action = 1
    END IF

    SELECT CASE Frm(1).KeyCode
    CASE EscKey
      ExitFlag = True           'EXIT DO
    END SELECT
  LOOP UNTIL ExitFlag


CustInqExit2:
EXIT SUB

LoadInqData2:
  RealValue# = 0
  SeniExmp# = 0
  OthrExmp# = 0
  PersValue# = 0
  MOBHValue# = 0
  MERHValue# = 0
  FarmValue# = 0
  MACHValue# = 0

  OpenPPTaxCustFile NumOfTaxRecs, TaxFile
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
  TaxInq(1).CUSTNAME = QPTrim$(TaxCustRec(1).FNAME) + " " + QPTrim$(TaxCustRec(1).LName)
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
  TaxInq(1).SRCITDIS = 0
  TaxInq(1).ODISCOUN = 0
  TaxInq(1).PROPVALU = 0
  TaxInq(1).PERSVAL = PersValue#
  TaxInq(1).MHVALUE = MOBHValue#
  TaxInq(1).MCVALUE = MERHValue#
  TaxInq(1).CVALUE = FarmValue#
  TaxInq(1).MTVALUE = MACHValue#

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
  
  MaxLines = 56
  LineCnt = 0
  CustCnt = 0
  
  LibName$ = "TAX"
  ScrnName$ = "VCUSTRPT"
  
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
  Choice$(0, 1) = "3"
  Choice$(1, 1) = "Summary"
  Choice$(2, 1) = "Detail"
  Choice$(0, 2) = "4"
  Choice$(1, 2) = "Screen"
  Choice$(2, 2) = "Printer"
  Form$(2, 0) = "R"                     'Always Default Real
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
      TaxType$ = Form$(2, 0)
      IF LEFT$(Form$(3, 0), 1) = "D" THEN
        DetailFlag = True
      ELSE
        DetailFlag = False
      END IF
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
  IF TaxType$ = "R" THEN
   OpenRETaxCustFile NumOfTaxRecs, TaxFile
    ELSE
   OpenPPTaxCustFile NumOfTaxRecs, TaxFile
  END IF
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
    
    IF NOT TaxCustRec(1).Deleted THEN
      IF LineCnt >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintDetailCustomerRptHeader
      END IF
      
      Nme$ = QPTrim$(TaxCustRec(1).FNAME) + " " + QPTrim$(TaxCustRec(1).LName)
      Nme$ = QPTrim$(Nme$)      'this one cleans up those with only last name
      
      IF NOT DetailFlag THEN
        PRINT #RptHandle, USING "#####"; CustRecNo;
        PRINT #RptHandle, TAB(10); Nme$; TAB(60); TaxCustRec(1).ACTIVE
        LineCnt = LineCnt + 1
      ELSE
        PRINT #RptHandle, USING "  Acct: ####"; CustRecNo;
        PRINT #RptHandle, TAB(15); Nme$
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
              PRINT #RptHandle, TAB(15); Nme$
              LineCnt = LineCnt + 2
            END IF
            
            PRINT #RptHandle, "Property Owned..."
            GET #PropTaxFile, PropertyRecord!, PropertyRec(1)
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
              PRINT #RptHandle, TAB(15); Nme$
              LineCnt = LineCnt + 2
            END IF
            
            PRINT #RptHandle, "Personal Property Owned..."
            GET #PersTaxFile, PropertyRecord!, PersRec(1)
            PRINT #RptHandle, "Pin # "; QPTrim$(PersRec(1).PROPPIN); TAB(50); " PP Value: "; USING "$$########,#.##"; PersRec(1).PERSVAL
            PRINT #RptHandle, "Desc  "; QPTrim$(PersRec(1).DESC1); TAB(50); " MH Value: "; USING "$$########,#.##"; PersRec(1).MHVALUE
            PRINT #RptHandle, "Desc  "; QPTrim$(PersRec(1).DESC2); TAB(50); " MC Value: "; USING "$$########,#.##"; PersRec(1).MCVALUE
            PRINT #RptHandle, "Desc  "; QPTrim$(PersRec(1).DESC3); TAB(50); " FE Value: "; USING "$$########,#.##"; PersRec(1).CVALUE
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
      
      Nme$ = QPTrim$(TaxCustRec(1).FNAME) + " " + QPTrim$(TaxCustRec(1).LName)
      Nme$ = QPTrim$(Nme$)      'this one cleans up those with only last name
      
      
      
      
      'First Show Property Records
      
      IF TaxCustRec(1).FirstPropRec > 0 THEN
        PropertyRecord! = TaxCustRec(1).FirstPropRec
        WHILE PropertyRecord! <> 0
          GET #PropTaxFile, PropertyRecord!, PropertyRec(1)
          IF PropertyRec(1).LATELIST = "Y" THEN
            PRINT #RptHandle, USING "######"; CustRecNo;
            PRINT #RptHandle, TAB(10); Nme$; TAB(57); USING "$$#######,#"; PropertyRec(1).PROPVALU
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
            PRINT #RptHandle, TAB(10); Nme$; TAB(57); USING "$$#######,#"; PValue#
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
  
  MaxLines = 50
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
  Choice$(1, 2) = "Real"
  Choice$(2, 2) = "Personal"
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
  IF Ptype$ = "R" THEN
   OpenRETaxCustFile NumOfTaxRecs, TaxFile
  ELSE
   OpenPPTaxCustFile NumOfTaxRecs, TaxFile
  END IF
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
      
      Nme$ = QPTrim$(TaxCustRec(1).FNAME) + " " + QPTrim$(TaxCustRec(1).LName)
      Nme$ = QPTrim$(Nme$)      'this one cleans up those with only last name
      
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
            BldgValue# = BldgValue# + PropertyRec(1).EXMPSENI
            IF LEFT$(Ptype$, 1) <> "P" THEN
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
        PRINT #RptHandle, USING "#####"; CustRecNo;
        PRINT #RptHandle, TAB(8); LEFT$(Nme$, 27);
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
            RealValue# = RealValue# + PropertyRec(1).PROPVALU + PropertyRec(1).EXMPSENI
            
            IF LEFT$(Ptype$, 1) <> "P" THEN
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
        PRINT #RptHandle, USING "#####"; CustRecNo;
        PRINT #RptHandle, TAB(8); LEFT$(Nme$, 27);
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
    END IF
    
    IF PFlag = 1 THEN PRINT #RptHandle, "": PFlag = 0
    CustCnt = CustCnt + 1
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
      
      Nme$ = RTRIM$(TaxCustRec(1).FNAME) + " " + RTRIM$(TaxCustRec(1).LName)
      Nme$ = LTRIM$(Nme$)
      
      IF Detail$ = "Summary" THEN
        PRINT #RptHandle, CustomerNumber; TAB(10); Nme$; TAB(60); TaxCustRec(1).ACTIVE
        LineCnt = LineCnt + 1
      ELSE
        PRINT #RptHandle, "Cust #: "; CustomerNumber; TAB(15); Nme$
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

SUB OpenPPTaxCustFile (NumOfTaxRecs, TaxFile)

  TaxFile = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))

END SUB

SUB OpenRETaxCustFile (NumOfTaxRecs, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "RETXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
  
END SUB

SUB OpenTaxCustFile (NumOfTaxRecs%, TaxFile%)
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
      
      Nme$ = QPTrim$(TaxCustRec(1).FNAME) + " " + QPTrim$(TaxCustRec(1).LName)
      Nme$ = QPTrim$(Nme$)      'this one cleans up those with only last name
      
      
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
          SCAmt# = SCAmt# + PropertyRec(1).EXMPSENI
          PropertyRecord! = PersRec(1).NextRec
        WEND
      END IF
      
      IF SCAmt# > 0 THEN
        TaxLoss# = (SCAmt# * TaxRate!) / 100
        PRINT #RptHandle, TaxCustRec(1).CSSN; TAB(15); Nme$; TAB(57); USING "$$#######,#"; SCAmt#;
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
END SUB

