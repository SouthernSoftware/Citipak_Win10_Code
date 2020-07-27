DEFINT A-Z
DECLARE SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE SUB ShowCustPersList (CustRec&, TaxType%)
DECLARE SUB CustHistoryRpt ()
DECLARE SUB CustHistoryRpt1 ()
DECLARE SUB CustHistoryRpt2 ()
DECLARE SUB CustomerInquiry2 ()
DECLARE SUB CustomerInquiry1 ()
DECLARE SUB OpenPPTaxCustFile (NumOfTaxRecs%, TaxFile%)
DECLARE SUB OpenRETaxCustFile (NumOfTaxRecs%, TaxFile%)
DECLARE SUB ShowCustPropList (CustRec&, TaxType%)
DECLARE SUB ShowCustHistory (CustRec&, TaxType%)
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
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION WEnvTest% ()
DECLARE FUNCTION Round# (B#)

DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
'DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB HideCursor ()
DECLARE SUB CursorOff ()
DECLARE SUB TextCursor (MouseFg%, MouseBg%)
DECLARE SUB WaitForAction ()
  
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
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED PersRec(1) AS PersonalRecType
  DIM SHARED TaxTrans(1) AS TaxTransactionType
  
  
  TYPE Struct
    who AS STRING * 14
    RecNum AS INTEGER
  END TYPE
  
  TYPE BalDetailType
    Year          AS INTEGER
    CustCnt       AS INTEGER
    Principle1    AS DOUBLE     'Va Personal Prop
    Principle2    AS DOUBLE     'For Va Only     Mach/Tools
    Principle3    AS DOUBLE     'For Va Only     Merch Cap
    Principle4    AS DOUBLE     'For Va Only     Farm Equip
    Principle5    AS DOUBLE     'For Va Only     Mobile Homes
    Interest      AS DOUBLE
    Penalty       AS DOUBLE
    Collection    AS DOUBLE
  END TYPE
  
  STACK 5000
  
  '--Dim the choice array to the number of menu items
  REDIM Mchoice$(1 TO 13)
  
  Mchoice$(1) = "Customer Inquiry - Real Estate"
  Mchoice$(2) = "Customer Inquiry - Personal Property"
  Mchoice$(3) = "Customer Transaction History - Real"
  Mchoice$(4) = "Customer Transaction History - Pers"
  Mchoice$(5) = "Master Customer Listing"
  Mchoice$(6) = "Master Balance Listing"
  Mchoice$(7) = "Master Valuation Listing"
  Mchoice$(8) = "Transaction Journal"
  Mchoice$(9) = "Mailing Labels"
  Mchoice$(10) = "Exit to DOS"
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 18   'limit the box length to go no lower than line 18
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
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    
    ClearBack
    
    TitleBox 2, Col, MaxLen + 3, "Tax Billing Reports Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu Mchoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      CustomerInquiry1
    CASE 2
      CustomerInquiry2
    CASE 3
      'unrem
      'CustHistoryRpt1
    CASE 4
      'unrem
      'CustHistoryRpt2
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
      CLEAR
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
    CLEAR
    RUN "TaxMenu"
  ELSE
    HideCursor
    ClearScrn
  END IF
  
  END

SUB BalanceListing
  
  SHARED Choice$()
  
  REDIM BalDet(0 TO 100) AS BalDetailType
  
  REDIM Array(1 TO 1) AS Struct 'Template for the sort array
  ReportFile$ = "TaxBal.PRN"    'Report File Name
  Dash80$ = STRING$(80, "=")
  FF$ = CHR$(12)
  
  Fmt1$ = "#####.##"
  
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
  
  TransFile = FREEFILE
  OPEN "TaxTrans.dat" FOR RANDOM AS TransFile LEN = LEN(TaxTrans(1))
  
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
      
      GOSUB CheckBalance        'Check to See if Balance on File
      
      IF Balance# <> 0 THEN
        Nme$ = QPTrim$(TaxCustRec(1).FNAME) + " " + QPTrim$(TaxCustRec(1).LName)
        Nme$ = QPTrim$(Nme$)    'this one cleans up those with only last name
        
        TransRecord& = TaxCustRec(1).LastTrans
        WHILE TransRecord& <> 0
          GET TransFile, TransRecord&, TaxTrans(1)
          IF TaxTrans(1).TranType = 1 THEN
            Balance# = TaxTrans(1).Revenue.Principle1 + TaxTrans(1).Revenue.Principle2 + TaxTrans(1).Revenue.Principle3 + TaxTrans(1).Revenue.Principle4 + TaxTrans(1).Revenue.Principle5
            Balance# = Balance# + TaxTrans(1).Revenue.Interest + TaxTrans(1).Revenue.Penalty + TaxTrans(1).Revenue.Collection
            Balance# = Round(Balance# - (TaxTrans(1).Revenue.Principle1Pd + TaxTrans(1).Revenue.Principle2Pd + TaxTrans(1).Revenue.Principle3Pd + TaxTrans(1).Revenue.Principle4Pd + TaxTrans(1).Revenue.Principle5Pd))
            Balance# = Balance# - (TaxTrans(1).Revenue.InterestPd + TaxTrans(1).Revenue.PenaltyPd + TaxTrans(1).Revenue.CollectionPd)
            Balance# = Round#(Balance#)
'NOTE: DALE
            IF Balance# < 0 THEN
            'IF Balance# > 0 THEN
              PRINT #RptHandle, USING "#####"; CustRecNo;
              PRINT #RptHandle, TAB(10); LEFT$(Nme$, 33);
              PRINT #RptHandle, TAB(45); LEFT$(TaxTrans(1).Description, 16);
              PRINT #RptHandle, TAB(61); TaxTrans(1).TaxYear;
              PRINT #RptHandle, TAB(67); USING "$$######,#.##"; Balance#
              LineCnt = LineCnt + 1
              CustCnt = CustCnt + 1
              TotalBalance# = TotalBalance# + Balance#
              IF DetailFlag THEN                'Detail Format Print Each Line
                IF TaxType$ = "P" THEN          'Personal Property
                  PRINT #RptHandle, TAB(10); "PP Tax: "; USING Fmt1$; Round#(TaxTrans(1).Revenue.Principle1 - TaxTrans(1).Revenue.Principle1Pd);
                  PRINT #RptHandle, "  ";
                  PRINT #RptHandle, "MT Tax: "; USING Fmt1$; Round#(TaxTrans(1).Revenue.Principle2 - TaxTrans(1).Revenue.Principle2Pd);
                  PRINT #RptHandle, "  ";
                  PRINT #RptHandle, "MC Tax: "; USING Fmt1$; Round#(TaxTrans(1).Revenue.Principle3 - TaxTrans(1).Revenue.Principle3Pd);
                  PRINT #RptHandle, "  ";
                  PRINT #RptHandle, "FE Tax: "; USING Fmt1$; Round#(TaxTrans(1).Revenue.Principle4 - TaxTrans(1).Revenue.Principle4Pd)
                  PRINT #RptHandle, TAB(10); "MH Tax: "; USING Fmt1$; Round#(TaxTrans(1).Revenue.Principle5 - TaxTrans(1).Revenue.Principle5Pd);
                  PRINT #RptHandle, "  ";
                  PRINT #RptHandle, " Int't: "; USING Fmt1$; Round#(TaxTrans(1).Revenue.Interest - TaxTrans(1).Revenue.InterestPd);
                  PRINT #RptHandle, "  ";
                  PRINT #RptHandle, "    Pen: "; USING Fmt1$; Round#(TaxTrans(1).Revenue.Penalty - TaxTrans(1).Revenue.PenaltyPd)
                  PRINT #RptHandle, STRING$(80, "-")
                  LineCnt = LineCnt + 4
                ELSE            'Real Estate
                  PRINT #RptHandle, TAB(10); "Real Tax: "; USING Fmt1$; Round#(TaxTrans(1).Revenue.Principle1 - TaxTrans(1).Revenue.Principle1Pd);
                  PRINT #RptHandle, "  ";
                  PRINT #RptHandle, " Int't: "; USING Fmt1$; Round#(TaxTrans(1).Revenue.Interest - TaxTrans(1).Revenue.InterestPd);
                  PRINT #RptHandle, "  ";
                  PRINT #RptHandle, "    Pen: "; USING Fmt1$; Round#(TaxTrans(1).Revenue.Penalty - TaxTrans(1).Revenue.PenaltyPd)
                  PRINT #RptHandle, STRING$(80, "-")
                  LineCnt = LineCnt + 3
                END IF
              END IF
              GOSUB SumYearInfo
            END IF
          END IF
          TransRecord& = TaxTrans(1).LastTrans
        WEND
      END IF    'End for Detail Cust
    END IF      'End of Delete Cust
    ShowPctComp Cnt&, NumOfTaxRecs
    
  NEXT
  
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
  
  '  KILL ReportFile$
  
  EXIT SUB
  
PrintBalanceRptHeader:
  Page = Page + 1
  PRINT #RptHandle, TAB(17); "Property Tax Customer Balance Listing";
  IF TaxType$ = "R" THEN
    PRINT #RptHandle, " (Real)"
  ELSE
    PRINT #RptHandle, " (Pers)"
  END IF
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
  IF NOT DetailFlag THEN
    PRINT #RptHandle, "               Total Balance: "; USING "$$########,#.##"; TotalBalance#
  END IF
  PRINT #RptHandle, FF$
  GOSUB SortThem
  PRINT #RptHandle, "Tax Totals By Year"
  IF DetailFlag THEN
    PRINT #RptHandle, " Year     Count"
    PRINT #RptHandle, "    PP Tax    MT Tax     MC Tax     FE Tax     MH Tax    Int/Pen    |    TOTALS"
    PRINT #RptHandle, STRING$(79, "-")
  ELSE
    PRINT #RptHandle, "Tax Year";
    PRINT #RptHandle, TAB(12); "Tax Amount"
    PRINT #RptHandle, STRING$(40, "-")
  END IF
  
  
  FOR Cnt = 1 TO NCnt
    Balance# = 0
    Balance# = Round(BalDet(Cnt).Principle1 + BalDet(Cnt).Principle2 + BalDet(Cnt).Principle3 + BalDet(Cnt).Principle4 + BalDet(Cnt).Principle5)
    Balance# = Round(Balance# + BalDet(Cnt).Interest + BalDet(Cnt).Penalty)
    GBalance# = Round#(GBalance# + Balance#)
    BalDet(0).Principle1 = Round#(BalDet(0).Principle1 + BalDet(Cnt).Principle1)
    BalDet(0).Principle2 = Round#(BalDet(0).Principle2 + BalDet(Cnt).Principle2)
    BalDet(0).Principle3 = Round#(BalDet(0).Principle3 + BalDet(Cnt).Principle3)
    BalDet(0).Principle4 = Round#(BalDet(0).Principle4 + BalDet(Cnt).Principle4)
    BalDet(0).Principle5 = Round#(BalDet(0).Principle5 + BalDet(Cnt).Principle5)
    BalDet(0).Interest = Round#(BalDet(0).Interest + BalDet(Cnt).Interest + BalDet(Cnt).Penalty)
    IF NOT DetailFlag THEN
      PRINT #RptHandle, BalDet(Cnt).Year; TAB(10);
      PRINT #RptHandle, USING "########.##"; Balance#
    ELSE
      PRINT #RptHandle, BalDet(Cnt).Year; TAB(12); USING "####"; BalDet(Cnt).CustCnt;
      PRINT #RptHandle, TAB(69); "|"
      PRINT #RptHandle, USING "#######.##"; BalDet(Cnt).Principle1; TAB(11); BalDet(Cnt).Principle2; TAB(22); BalDet(Cnt).Principle3; TAB(33); BalDet(Cnt).Principle4; TAB(44); BalDet(Cnt).Principle5; TAB(55); Round#(BalDet(Cnt).Interest + BalDet(Cnt _
).Penalty);
      PRINT #RptHandle, TAB(69); "|"; USING "#######.##"; TAB(70); Balance#
    END IF
  NEXT
  IF DetailFlag THEN
    PRINT #RptHandle, STRING$(79, "-")
    PRINT #RptHandle, USING "#######.##"; BalDet(0).Principle1; TAB(11); BalDet(0).Principle2; TAB(22); BalDet(0).Principle3; TAB(33); BalDet(0).Principle4; TAB(44); BalDet(0).Principle5; TAB(55); BalDet(0).Interest;
    PRINT #RptHandle, TAB(69); "|"; USING "#######.##"; TAB(70); GBalance#
  END IF
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
  TransRecord& = TaxCustRec(1).LastTrans
  WHILE TransRecord& <> 0
    GET TransFile, TransRecord&, TaxTrans(1)
    IF TaxTrans(1).TranType = 1 THEN
      Balance# = TaxTrans(1).Revenue.Principle1 + TaxTrans(1).Revenue.Principle2 + TaxTrans(1).Revenue.Principle3 + TaxTrans(1).Revenue.Principle4 + TaxTrans(1).Revenue.Principle5
      Balance# = Balance# + TaxTrans(1).Revenue.Interest + TaxTrans(1).Revenue.Penalty + TaxTrans(1).Revenue.Collection
      Balance# = Balance# - (TaxTrans(1).Revenue.Principle1Pd + TaxTrans(1).Revenue.Principle2Pd + TaxTrans(1).Revenue.Principle3Pd + TaxTrans(1).Revenue.Principle4Pd + TaxTrans(1).Revenue.Principle5Pd)
      Balance# = Balance# - (TaxTrans(1).Revenue.InterestPd + TaxTrans(1).Revenue.PenaltyPd + TaxTrans(1).Revenue.CollectionPd)
      Balance# = Round#(Balance#)
      'Routine Start to Fix Posting Bug  ******************
      'GOSUB FixBug
      'Routine End to Fix Posting Bug    ******************
'NOTE: DALE
      IF Balance# < 0 THEN
        GOTO CheckBalReturn
      END IF
    '  IF Balance# > 0 THEN
    '    GOTO CheckBalReturn
    '  END IF

    END IF
    CTransRecord& = TransRecord&
    TransRecord& = TaxTrans(1).LastTrans
    IF TransRecord& = CTransRecord& THEN TransRecord& = 0
  WEND
  Balance# = 0
  
CheckBalReturn:
  RETURN
  
SortThem:
26000         REM sort
  Count = NCnt
26020 M = Count
26030 M = INT(M / 2)
26040 IF M = 0 THEN 26190
26050 FOR st = 1 TO M
26060 I = st
26070 J = st + M
26080 SW = 0
26090 IF BalDet(I).Year <= BalDet(J).Year THEN 26120
26100 SW = 1
26110 SWAP BalDet(I), BalDet(J)
26120 I = J
26130 J = J + M
26140 IF J <= Count THEN 26090
26150 IF SW = 0 THEN 26170
26160 GOTO 26060
26170 NEXT st
26180 GOTO 26030
26190 RETURN
  
SumYearInfo:
  NewFlag = True                'Set Flag to Not Updated
  
  FOR LCnt = 1 TO NCnt
    IF BalDet(LCnt).Year = TaxTrans(1).TaxYear THEN
      BalDet(LCnt).CustCnt = BalDet(LCnt).CustCnt + 1
      BalDet(LCnt).Principle1 = Round(BalDet(LCnt).Principle1 + Round#(TaxTrans(1).Revenue.Principle1 - TaxTrans(1).Revenue.Principle1Pd))
      BalDet(LCnt).Principle2 = Round(BalDet(LCnt).Principle2 + Round#(TaxTrans(1).Revenue.Principle2 - TaxTrans(1).Revenue.Principle2Pd))
      BalDet(LCnt).Principle3 = Round(BalDet(LCnt).Principle3 + Round#(TaxTrans(1).Revenue.Principle3 - TaxTrans(1).Revenue.Principle3Pd))
      BalDet(LCnt).Principle4 = Round(BalDet(LCnt).Principle4 + Round#(TaxTrans(1).Revenue.Principle4 - TaxTrans(1).Revenue.Principle4Pd))
      BalDet(LCnt).Principle5 = Round(BalDet(LCnt).Principle5 + Round#(TaxTrans(1).Revenue.Principle5 - TaxTrans(1).Revenue.Principle5Pd))
      BalDet(LCnt).Interest = Round(BalDet(LCnt).Interest + Round#(TaxTrans(1).Revenue.Interest - TaxTrans(1).Revenue.InterestPd))
      BalDet(LCnt).Penalty = Round(BalDet(LCnt).Penalty + Round#(TaxTrans(1).Revenue.Penalty - TaxTrans(1).Revenue.PenaltyPd))
      NewFlag = False
      EXIT FOR
    END IF
  NEXT
  IF NewFlag THEN               'if we didn't find this year
    NCnt = NCnt + 1
    BalDet(NCnt).Year = TaxTrans(1).TaxYear
    BalDet(NCnt).CustCnt = 1
    BalDet(NCnt).Principle1 = Round#(TaxTrans(1).Revenue.Principle1 - TaxTrans(1).Revenue.Principle1Pd)
    BalDet(NCnt).Principle2 = Round#(TaxTrans(1).Revenue.Principle2 - TaxTrans(1).Revenue.Principle2Pd)
    BalDet(NCnt).Principle3 = Round#(TaxTrans(1).Revenue.Principle3 - TaxTrans(1).Revenue.Principle3Pd)
    BalDet(NCnt).Principle4 = Round#(TaxTrans(1).Revenue.Principle4 - TaxTrans(1).Revenue.Principle4Pd)
    BalDet(NCnt).Principle5 = Round#(TaxTrans(1).Revenue.Principle5 - TaxTrans(1).Revenue.Principle5Pd)
    BalDet(NCnt).Interest = Round#(TaxTrans(1).Revenue.Interest - TaxTrans(1).Revenue.InterestPd)
    BalDet(NCnt).Penalty = Round#(TaxTrans(1).Revenue.Penalty - TaxTrans(1).Revenue.PenaltyPd)
  END IF
  
GotYearInfo:
  RETURN
  
FixBug:
  'IF TaxTrans(1).Revenue.Principle1Pd > TaxTrans(1).Revenue.Principle1 THEN
  '  TaxTrans(1).Revenue.Principle1Pd = TaxTrans(1).Revenue.Principle1
  'END IF
  'IF TaxTrans(1).Revenue.Principle2Pd > TaxTrans(1).Revenue.Principle2 THEN
  '  TaxTrans(1).Revenue.Principle2Pd = TaxTrans(1).Revenue.Principle2
  'END IF
  'IF TaxTrans(1).Revenue.Principle3Pd > TaxTrans(1).Revenue.Principle3 THEN
  '  TaxTrans(1).Revenue.Principle3Pd = TaxTrans(1).Revenue.Principle3
  'END IF
  'IF TaxTrans(1).Revenue.Principle4Pd > TaxTrans(1).Revenue.Principle4 THEN
  '  TaxTrans(1).Revenue.Principle4Pd = TaxTrans(1).Revenue.Principle4
  'END IF
  'IF TaxTrans(1).Revenue.Principle5Pd > TaxTrans(1).Revenue.Principle5 THEN
  '  TaxTrans(1).Revenue.Principle5Pd = TaxTrans(1).Revenue.Principle5
  'END IF
  'IF TaxTrans(1).Revenue.InterestPd > TaxTrans(1).Revenue.Interest THEN
  '  TaxTrans(1).Revenue.InterestPd = TaxTrans(1).Revenue.Interest
  'END IF
  'IF TaxTrans(1).Revenue.PenaltyPd > TaxTrans(1).Revenue.Penalty THEN
  '  TaxTrans(1).Revenue.PenaltyPd = TaxTrans(1).Revenue.Penalty
  'END IF
  'PUT TransFile, TransRecord&, TaxTrans(1)
  
  RETURN
END SUB

SUB CustomerInquiry1
  
  SHARED Choice$()
  
  REDIM TaxInq(1) AS RETaxCustInqType
  REDIM TempScrn(0)
  
  TaxType% = 1
  ClearBack
  LookUp RecNo&, "Customer Inquiry", False, True, False, TaxType%
  
  CustAcct& = RecNo&
  
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
      
    CASE F4Key
      IF CustAcct& > 0 THEN
        ShowCustHistory CustAcct&, 1            'DON'T CHANGE THIS
        Action = 1
      END IF
      
    CASE F7Key
      IF CustAcct& > 0 THEN
        ShowCustPropList CustAcct&, 1           'DON'T CHANGE THIS
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
  
  
  TaxInq(1).Acct = RecNo&
  TaxInq(1).OPENDATE = TaxCustRec(1).OPENDATE
  TaxInq(1).CUSTNAME = QPTrim$(TaxCustRec(1).FNAME) + " " + QPTrim$(TaxCustRec(1).LName)
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
  TaxInq(1).Penalty = TaxCustRec(1).Penalty
  TaxInq(1).ODISCOUN = OthrExmp#
  TaxInq(1).PROPVALU = RealValue#
  TaxInq(1).PersVal = BldgValue#
  
  BCopy VARSEG(TaxInq(1)), VARPTR(TaxInq(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), LEN(Form$(0, 0)), 0
  UnPackBuffer 0, 0, Form$(), Fld()
  Action = 1
  RETURN
END SUB

SUB CustomerInquiry2
  
  SHARED Choice$()
  
  REDIM TaxInq(1) AS TaxCustInqType
  
  TaxType% = 2  'Personal Property
  
  ClearBack
  LookUp RecNo&, "Customer Inquiry", False, True, False, TaxType%
  IF RecNo& <= 0 THEN
    GOTO CustInqExit2
  END IF
  CustAcct& = RecNo&
  FirstTime = True
  
  LibName$ = "TAX"
  ScrnName$ = "PCUSTINQ"
  
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
      
    CASE F4Key
      IF CustAcct& > 0 THEN
        ShowCustHistory -CustAcct&, 2           'DON'T CHANGE THIS
        Action = 1
      END IF
    CASE F7Key
      IF CustAcct& > 0 THEN
        ShowCustPersList CustAcct&, 1           'DON'T CHANGE THIS
        Action = 1
        
      END IF
      
      
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
      PersValue# = Round#(PersValue# + PersRec(1).PersVal)
      MOBHValue# = Round#(MOBHValue# + PersRec(1).MHVALUE)
      MERHValue# = Round#(MERHValue# + PersRec(1).MCVALUE)
      FarmValue# = Round#(FarmValue# + PersRec(1).CVALUE)
      MACHValue# = Round#(MACHValue# + PersRec(1).MTVALUE)
      SeniExmp# = Round#(SeniExmp# + PersRec(1).EXMPSENI)
      OthrExmp# = Round#(OthrExmp# + PersRec(1).EXMPOTHR)
    NEXT
  END IF
  CLOSE PersTaxFile
  
  TaxInq(1).Acct = RecNo&
  TaxInq(1).OPENDATE = TaxCustRec(1).OPENDATE
  TaxInq(1).CUSTNAME = QPTrim$(TaxCustRec(1).FNAME) + " " + QPTrim$(TaxCustRec(1).LName)
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
  TaxInq(1).Penalty = TaxCustRec(1).Penalty
  TaxInq(1).SRCITDIS = 0
  TaxInq(1).ODISCOUN = 0
  TaxInq(1).PROPVALU = PersValue#
  TaxInq(1).PersVal = MOBHValue#
  TaxInq(1).MHVALUE = MERHValue#
  TaxInq(1).MCVALUE = FarmValue#
  TaxInq(1).CVALUE = MACHValue#
  TaxInq(1).MTVALUE = 0
  
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
  
  MaxLines = 52
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
  Form$(2, 0) = "R"             'Always Default Real
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
    
    '    'IF CustRecNo = 34 THEN STOP
    
    IF NOT TaxCustRec(1).Deleted THEN
      
      GOSUB CheckLineCnt
      
      Nme$ = QPTrim$(TaxCustRec(1).FNAME) + " " + QPTrim$(TaxCustRec(1).LName)
      Nme$ = QPTrim$(Nme$)      'this one cleans up those with only last name
      
      IF NOT DetailFlag THEN
        'IF NOT INSTR(TaxCustRec(1).City, ",") > 0 THEN
        PRINT #RptHandle, USING "#####"; CustRecNo;
        PRINT #RptHandle, TAB(10); Nme$; TAB(60); TaxCustRec(1).CountyAcct
        'PRINT #RptHandle, TAB(10); Nme$; TAB(60); TaxCustRec(1).Active
        'PRINT #RptHandle, TAB(10); Nme$; TAB(60); RTRIM$(TaxCustRec(1).City)
        LineCnt = LineCnt + 1
        'END IF
      ELSE
        PRINT #RptHandle, USING "  Acct: ####"; CustRecNo;
        PRINT #RptHandle, TAB(15); Nme$
        PRINT #RptHandle, "Active: "; TaxCustRec(1).Active; TAB(15); TaxCustRec(1).Addr1
        PRINT #RptHandle, "Int'st: "; TaxCustRec(1).Interest; TAB(15); TaxCustRec(1).Addr2
        PRINT #RptHandle, "Exempt: "; TaxCustRec(1).TaxExempt; TAB(15); RTRIM$(TaxCustRec(1).City) + ", "; RTRIM$(TaxCustRec(1).State) + "  " + RTRIM$(TaxCustRec(1).Zip)
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
            PRINT #RptHandle, "Pin # "; QPTrim$(PropertyRec(1).REALPIN); TAB(50); "  Land: "; USING "$$########,#.##"; PropertyRec(1).PROPVALU
            PRINT #RptHandle, "Desc  "; QPTrim$(PropertyRec(1).PROPNOT1); TAB(50); "  Bldg: "; USING "$$########,#.##"; PropertyRec(1).EXMPSENI
            PRINT #RptHandle, "Desc  "; QPTrim$(PropertyRec(1).PROPNOT2); TAB(50); "Exempt: "; ; USING "$$########,#.##"; PropertyRec(1).EXMPOTHR
            PRINT #RptHandle, "Desc  "; QPTrim$(PropertyRec(1).PROPNOT3); TAB(50); "Mortgage Code: "; QPTrim$(PropertyRec(1).MORTCODE)
            PRINT #RptHandle, "Map/Blk/Lot - "; QPTrim$(PropertyRec(1).MAP); "/"; QPTrim$(PropertyRec(1).BLOCK); "/"; QPTrim$(PropertyRec(1).LOTNUMB); TAB(40); "Late (Y/N): "; PropertyRec(1).LATELIST
            PRINT #RptHandle, STRING$(79, "-")
            LineCnt = LineCnt + 6
            OldRecord! = PropertyRecord!
            PropertyRecord! = PropertyRec(1).NextRec
            IF OldRecord! = PropertyRecord! THEN PropertyRecord! = 0
            'GOSUB CheckLineCnt
          WEND
        END IF
        GOSUB CheckLineCnt
        
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
            PRINT #RptHandle, "Pin # "; QPTrim$(PersRec(1).PROPPIN); TAB(50); " PP Value: "; USING "$$########,#.##"; PersRec(1).PersVal
            PRINT #RptHandle, "Desc  "; QPTrim$(PersRec(1).DESC1); TAB(50); " MH Value: "; USING "$$########,#.##"; PersRec(1).MHVALUE
            PRINT #RptHandle, "Desc  "; QPTrim$(PersRec(1).DESC2); TAB(50); " MC Value: "; USING "$$########,#.##"; PersRec(1).MCVALUE
            PRINT #RptHandle, "Desc  "; QPTrim$(PersRec(1).DESC3); TAB(50); " FE Value: "; USING "$$########,#.##"; PersRec(1).CVALUE
            PRINT #RptHandle, "Desc  "; QPTrim$(PersRec(1).Desc4); TAB(50); " MT Value: "; USING "$$########,#.##"; PersRec(1).MTVALUE
            PRINT #RptHandle, "Desc  "; QPTrim$(PersRec(1).Desc5); TAB(50); "Exempt: "; USING "$$########,#.##"; PersRec(1).EXMPSENI
            PRINT #RptHandle, STRING$(79, "-")
            LineCnt = LineCnt + 7
            OldRecord! = PropertyRecord!
            PropertyRecord! = PersRec(1).NextRec
            IF OldRecord! = PropertyRecord! THEN PropertyRecord! = 0
            '    GOSUB CheckLineCnt
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
  
  'KILL ReportFile$
  
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
  
CheckLineCnt:
  IF LineCnt >= MaxLines THEN
    PRINT #RptHandle, FF$
    GOSUB PrintDetailCustomerRptHeader
  END IF
  RETURN
  
  
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
    CASE F5Key
      GOSUB PrintAlign
    CASE EscKey
      AbortFlag = True
      ExitFlag = True           'EXIT DO
    END SELECT
  LOOP UNTIL ExitFlag
  
  IF AbortFlag THEN EXIT SUB
  
  RptHandle = FREEFILE
  
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  
  IF TaxType$ = "R" THEN
    OpenRETaxCustFile NumOfTaxRecs, TaxFile
  ELSE
    OpenPPTaxCustFile NumOfTaxRecs, TaxFile
  END IF
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
          MC$ = PropertyRec(1).MORTCODE
          MC$ = RTRIM$(MC$)
          PropRec& = PropertyRec(1).NextRec
        WEND
      END IF
      
      IF DetailFlag AND TaxType$ = "R" AND LEN(MC$) > 0 THEN
        Nme$ = QPTrim$(TaxCustRec(1).FNAME) + " " + QPTrim$(TaxCustRec(1).LName)
        Nme$ = QPTrim$(Nme$)    'this one cleans up those with only last name
        PRINT #RptHandle, USING "#####"; CustRecNo;
        PRINT #RptHandle, TAB(20); "MC="; MC$
        PRINT #RptHandle, Nme$
        PRINT #RptHandle, TaxCustRec(1).Addr1
        PRINT #RptHandle, TaxCustRec(1).Addr2
        PRINT #RptHandle, TaxCustRec(1).City; " "; TaxCustRec(1).State; " "; TaxCustRec(1).Zip
        PRINT #RptHandle,
      END IF
      IF NOT (DetailFlag) THEN
        Nme$ = QPTrim$(TaxCustRec(1).FNAME) + " " + QPTrim$(TaxCustRec(1).LName)
        Nme$ = QPTrim$(Nme$)    'this one cleans up those with only last name
        PRINT #RptHandle, USING "#####"; CustRecNo
        PRINT #RptHandle, Nme$
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
  
  KILL ReportFile$
  
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
            PValue# = PersRec(1).PersVal# + PersRec(1).MHVALUE + PersRec(1).MCVALUE + PersRec(1).CVALUE + PersRec(1).MTVALUE
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
  Choice$(0, 3) = "5"
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
      PType$ = LEFT$(Form$(3, 0), 1)
      DevSpec$ = LEFT$(Form$(5, 0), 1)
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
  IF PType$ = "R" THEN
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
    
    'IF CustRecNo = 34 THEN STOP
    
    PValue# = 0
    BldgValue# = 0
    RealValue# = 0
    IF NOT TaxCustRec(1).Deleted THEN
      IF LineCnt >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintMasterValueHeader
      END IF
      
      Nme$ = QPTrim$(TaxCustRec(1).FNAME) + " " + QPTrim$(TaxCustRec(1).LName)
      Nme$ = QPTrim$(Nme$)      'this one cleans up those with only last name
      Addr$ = QPTrim$(TaxCustRec(1).Addr1)
      City$ = QPTrim$(TaxCustRec(1).City) + ", " + QPTrim$(TaxCustRec(1).State) + " " + QPTrim$(TaxCustRec(1).Zip)
      
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
            IF LEFT$(PType$, 1) <> "P" THEN
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
            PersValue# = PersValue# + PersRec(1).PersVal
            PersValue# = PersValue# + PersRec(1).MHVALUE
            PersValue# = PersValue# + PersRec(1).MCVALUE
            PersValue# = PersValue# + PersRec(1).CVALUE
            PersValue# = PersValue# + PersRec(1).MTVALUE
            IF LEFT$(PType$, 1) <> "R" THEN
              Discnt# = Discnt# + PersRec(1).EXMPSENI
            END IF
            PropertyRecord! = PersRec(1).NextRec
          WEND
        END IF

'IF PersValue# = 0 THEN
'  GOTO NotThisCust
'END IF

        IF LEFT$(PType$, 1) = "R" THEN
          PersValue# = 0
        END IF
        IF LEFT$(PType$, 1) = "P" THEN
          RealValue# = 0
        END IF
        PRINT #RptHandle, USING "#####"; CustRecNo;
        PRINT #RptHandle, TAB(8); LEFT$(Nme$, 27);
        IF PType$ = "R" THEN
          PRINT #RptHandle, TAB(37); USING "########,#"; RealValue#;
          PRINT #RptHandle, TAB(48); USING "########,#"; BldgValue#;
          PRINT #RptHandle, TAB(59); USING "########,#"; Discnt#;
          PRINT #RptHandle, TAB(70); USING "########,#"; RealValue# + BldgValue# - Discnt#
        ELSE
          PRINT #RptHandle, TAB(37); USING "########,#"; PersValue#;
          PRINT #RptHandle, TAB(59); USING "########,#"; Discnt#;
          PRINT #RptHandle, TAB(70); USING "########,#"; RealValue# + PersValue# - Discnt#
        END IF
        LineCnt = LineCnt + 1
        TotalReal# = TotalReal# + RealValue#
        TotalBldg# = TotalBldg# + BldgValue#
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
            BldgValue# = BldgValue# + PropertyRec(1).EXMPSENI
            IF LEFT$(PType$, 1) <> "P" THEN
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
            PersValue# = PersValue# + PersRec(1).PersVal
            PersValue# = PersValue# + PersRec(1).MHVALUE
            PersValue# = PersValue# + PersRec(1).MCVALUE
            PersValue# = PersValue# + PersRec(1).CVALUE
            PersValue# = PersValue# + PersRec(1).MTVALUE
            IF LEFT$(PType$, 1) <> "R" THEN
              Discnt# = Discnt# + PersRec(1).EXMPSENI
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
        PRINT #RptHandle, USING "#####"; CustRecNo;
        PRINT #RptHandle, TAB(8); LEFT$(Nme$, 27);
        IF PType$ = "R" THEN
          PRINT #RptHandle, TAB(37); USING "########,#"; RealValue#;
          PRINT #RptHandle, TAB(48); USING "########,#"; BldgValue#;
          PRINT #RptHandle, TAB(59); USING "########,#"; Discnt#;
          PRINT #RptHandle, TAB(70); USING "########,#"; RealValue# + BldgValue# - Discnt#
        ELSE
          PRINT #RptHandle, TAB(37); USING "########,#"; PersValue#;
          PRINT #RptHandle, TAB(59); USING "########,#"; Discnt#;
          PRINT #RptHandle, TAB(70); USING "########,#"; RealValue# + PersValue# - Discnt#
        END IF
        
        PRINT #RptHandle, TAB(8); Addr$
        PRINT #RptHandle, TAB(8); City$
        PRINT #RptHandle, TAB(8); "SSN: "; TaxCustRec(1).CSSN
        
        TotalReal# = TotalReal# + RealValue#
        TotalBldg# = TotalBldg# + BldgValue#
        TotalPers# = TotalPers# + PersValue#
        TotalDisc# = TotalDisc# + Discnt#
        
        LineCnt = LineCnt + 4
        
        'Now Show Detail Support Here
        IF LEFT$(PType$, 1) = "P" THEN
        ELSE
          IF TaxCustRec(1).FirstPropRec > 0 THEN
            PropertyRecord! = TaxCustRec(1).FirstPropRec
            PFlag = 0
            WHILE PropertyRecord! <> 0
              GET PropTaxFile, PropertyRecord!, PropertyRec(1)
              PRINT #RptHandle, TAB(15); "Property Pin# "; QPTrim$(PropertyRec(1).REALPIN);
              PRINT #RptHandle, TAB(52); "Value: "; USING "#######,#"; PropertyRec(1).PROPVALU + PropertyRec(1).EXMPSENI
              LineCnt = LineCnt + 1
              PFlag = 1
              PropertyRecord! = PropertyRec(1).NextRec
            WEND
            
          END IF
        END IF
        IF LEFT$(PType$, 1) = "R" THEN
        ELSE
          CarCount = 0
          IF TaxCustRec(1).FirstPersRec > 0 THEN
            PropertyRecord! = TaxCustRec(1).FirstPersRec
            WHILE PropertyRecord! <> 0
              GET PersTaxFile, PropertyRecord!, PersRec(1)
              PValue# = 0
              PValue# = PValue# + PersRec(1).PersVal
              PValue# = PValue# + PersRec(1).MHVALUE
              PValue# = PValue# + PersRec(1).MCVALUE
              PValue# = PValue# + PersRec(1).CVALUE
              PValue# = PValue# + PersRec(1).MTVALUE
              PRINT #RptHandle, TAB(5); "Pers Abstract# "; QPTrim$(PersRec(1).PROPPIN);
              PRINT #RptHandle, TAB(45); "Value: "; USING "#######,#"; PValue#
              'PRINT #RptHandle, TAB(45); "Value: "; USING "#######,#"; PersRec(1).PERSVAL
              PRINT #RptHandle, TAB(12); "VIN: "; QPTrim$(PersRec(1).DESC1)
              PRINT #RptHandle, TAB(5); "Make/Model: "; QPTrim$(PersRec(1).DESC2)
              LineCnt = LineCnt + 3
              PFlag = 1
              'CarCount = CarCount + 1
              'IF MaxCount < CarCount THEN
              '  MaxCount = CarCount
              '  IF MaxCount = 21 THEN STOP
              'END IF
              PropertyRecord! = PersRec(1).NextRec
            WEND
          END IF
        END IF
      END IF
    END IF
    
    IF PFlag = 1 THEN PRINT #RptHandle, "": PFlag = 0
    CustCnt = CustCnt + 1
NotThisCust:
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
  
  'KILL ReportFile$
  
  EXIT SUB
  
PrintMasterValueHeader:
  Page = Page + 1
  PRINT #RptHandle, TAB(25); "Property Tax Valuation Listing"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
  IF PType$ = "R" THEN
    IF NOT DetailFlag THEN
      PRINT #RptHandle, "Summary Format"
      PRINT #RptHandle, TAB(40); "[------------- Valuations --------------]"
      PRINT #RptHandle, "Acct #"; TAB(8); "Name"; TAB(43); "Real"; TAB(53); "Bldg"; TAB(62); "Discnt"; TAB(77); "Net"
      PRINT #RptHandle, Dash80$
      LineCnt = 6
    ELSE
      PRINT #RptHandle, "Detail Format"
      PRINT #RptHandle, TAB(40); "[------------- Valuations -------------]"
      PRINT #RptHandle, "Acct #"; TAB(8); "Name"; TAB(46); "Real"; TAB(55); "Bldg"; TAB(63); "Discnt"; TAB(76); "Net"
      PRINT #RptHandle, Dash80$
      LineCnt = 6
    END IF
  ELSE
    IF NOT DetailFlag THEN
      PRINT #RptHandle, "Summary Format"
      PRINT #RptHandle, TAB(40); "[------------- Valuations --------------]"
      PRINT #RptHandle, "Acct #"; TAB(8); "Name"; TAB(43); "Pers"; TAB(62); "Discnt"; TAB(77); "Net"
      PRINT #RptHandle, Dash80$
      LineCnt = 6
    ELSE
      PRINT #RptHandle, "Detail Format"
      PRINT #RptHandle, TAB(40); "[------------- Valuations -------------]"
      PRINT #RptHandle, "Acct #"; TAB(8); "Name"; TAB(46); "Pers"; TAB(63); "Discnt"; TAB(76); "Net"
      PRINT #RptHandle, Dash80$
      LineCnt = 6
    END IF
    
    
  END IF
  RETURN
  
PrintMasterValueEnding:
  PRINT #RptHandle, Dash80$
  PRINT #RptHandle, "Total Customers Printed: "; USING "#####"; CustCnt
  PRINT #RptHandle, "Totals ..."
  IF PType$ = "R" THEN
    PRINT #RptHandle, "Real Value: "; USING "$$##########,#"; TotalReal#
    PRINT #RptHandle, "Bldg Value: "; USING "$$##########,#"; TotalBldg#
    PRINT #RptHandle, "Discount  : "; USING "$$##########,#"; TotalDisc#
    PRINT #RptHandle, "Net Value : "; USING "$$##########,#"; TotalReal# + TotalBldg# - TotalDisc#
  ELSE
    PRINT #RptHandle, "Pers Value: "; USING "$$##########,#"; TotalPers#
    PRINT #RptHandle, "Discount  : "; USING "$$##########,#"; TotalDisc#
    PRINT #RptHandle, "Net Value : "; USING "$$##########,#"; TotalReal# + TotalPers# - TotalDisc#
  END IF
  
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
  
  REDIM Choice$(0 TO 8, 0 TO 2)
  
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "0-All Transactions"
  Choice$(2, 0) = "1-Billing"
  Choice$(3, 0) = "2-Payments"
  Choice$(4, 0) = "3-Release/Abatements"
  Choice$(5, 0) = "4-Interest"
  Choice$(6, 0) = "5-Penalty"
  Choice$(7, 0) = "6-Collection Cost"
  Choice$(8, 0) = "7-Adjustments"
  
  Choice$(0, 1) = "4"
  Choice$(1, 1) = "Real Esate"
  Choice$(2, 1) = "Personal"
  
  Choice$(0, 2) = "5"
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
      TrType = VAL(Form$(1, 0))
      BDate$ = Form$(2, 0)
      BDate = Date2Num%(BDate$)
      EDate$ = Form$(3, 0)
      EDate = Date2Num%(EDate$)
      
      IF LEFT$(Form$(4, 0), 1) = "R" THEN
        TaxType$ = "R"
      ELSEIF LEFT$(Form$(4, 0), 1) = "P" THEN
        TaxType$ = "P"
      ELSE
        TaxType$ = ""
      END IF
      DevSpec$ = LEFT$(Form$(5, 0), 1)
      IF TrType >= 0 AND TrType <= 7 AND EDate >= BDate THEN
        ExitFlag = True
      END IF
    CASE EscKey
      AbortFlag = True
      ExitFlag = True           'EXIT DO
    END SELECT
  LOOP UNTIL ExitFlag
  
  IF AbortFlag THEN EXIT SUB

  OpenPPTaxCustFile NumOfTaxRecs, PPFile
  

  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  GOSUB PrintTransJourRptHeader


  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
  NumOfTranRecs& = LOF(TaxTran) / TaxTranRecLen
  
  ClearBack
  ShowProcessingScrn "Master Customer Listing"
  
  FOR TCnt& = NumOfTranRecs& TO 1 STEP -1
    GET TaxTran, TCnt&, TranRec(1)
    IF LineCnt >= MaxLines THEN
      PRINT #RptHandle, FF$
      GOSUB PrintTransJourRptHeader
    END IF
    IF TranRec(1).TransDate >= BDate AND TranRec(1).TransDate <= EDate THEN
      
    ELSE
      GOTO SkipEm1
    END IF
    
    IF TrType = 0 OR TranRec(1).TranType = TrType THEN
    ELSE
      GOTO SkipEm1
    END IF
    IF TrType = 1 THEN
      IF TaxType$ = TranRec(1).BillType THEN
      ELSEIF TaxType$ = "" THEN
      ELSE
        GOTO SkipEm1
      END IF
    END IF
    IF TrType = 2 OR TrType = 3 THEN
      GET TaxTran, TranRec(1).BelongTo, TranRec(1)
      IF (TaxType$ = TranRec(1).BillType) OR (TaxType$ = "") THEN
        TaxYear$ = LTRIM$(STR$(TranRec(1).TaxYear))
        GET TaxTran, TCnt&, TranRec(1)
      ELSE
        GOTO SkipEm1
      END IF
    END IF
    
    GOSUB PrintLineType0or1
    
SkipEm1:
    ShowPctCompL TCnt&, NumOfTranRecs&
  NEXT
  
  GOSUB PrintTransJourRptEnding
  
  PRINT #RptHandle, MT&; CHR$(18);  ' oki 320 10 cpi
  
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
  
PrintLineType0or1:
  MT& = MT& + 1
  PRINT #RptHandle, Num2Date$(TranRec(1).TransDate);
  'PRINT #RptHandle, TranRec(1).BelongTo;
  PRINT #RptHandle, TAB(12); TranRec(1).Description;
  'PRINT #RptHandle, TAB(12); TCnt&;

  PRINT #RptHandle, TAB(45); "";
  IF TranRec(1).TranType = 1 THEN
    PRINT #RptHandle, TAB(45); "BILL";

'    TranRec(1).Amount = 0
'    TranRec(1).Revenue.Principle1 = 0
'    TranRec(1).Revenue.Principle2 = 0
'    TranRec(1).Revenue.Principle3 = 0
'    TranRec(1).Revenue.Principle4 = 0
'    TranRec(1).Revenue.Principle5 = 0
'    TranRec(1).Revenue.Interest = 0
'    TranRec(1).Revenue.Penalty = 0
'    TranRec(1).Revenue.Collection = 0
'    TranRec(1).Revenue.Future1 = 0
'    TranRec(1).Revenue.Future2 = 0
'    TranRec(1).Revenue.Principle1Pd = 0
'    TranRec(1).Revenue.Principle2Pd = 0
'    TranRec(1).Revenue.Principle3Pd = 0
'    TranRec(1).Revenue.Principle4Pd = 0
'    TranRec(1).Revenue.Principle5Pd = 0
'    TranRec(1).Revenue.InterestPd = 0
'    TranRec(1).Revenue.PenaltyPd = 0
'    TranRec(1).Revenue.CollectionPd = 0
'    TranRec(1).Revenue.Future1Pd = 0
'    TranRec(1).Revenue.Future2Pd = 0
'    PUT TaxTran, TCnt&, TranRec(1)
'    GET PPFile, TranRec(1).CustomerRec, TaxCustRec(1)
'    TaxCustRec(1).LastTrans = TranRec(1).LastTrans
'    PUT PPFile, TranRec(1).CustomerRec, TaxCustRec(1)

'    Interest# = TranRec(1).Amount

'    WhatTr& = TranRec(1).BelongTo
'    GET TaxTran, WhatTr&, TranRec(1)
'    TranRec(1).Revenue.Interest = Round#(TranRec(1).Revenue.Interest - Interest#)
'    PUT TaxTran, WhatTr&, TranRec(1)

  END IF
  IF TranRec(1).TranType = 2 THEN
    IF TranRec(1).BillType = "R" THEN
      PRINT #RptHandle, TAB(45); "PMT R";
    ELSEIF TranRec(1).BillType = "P" THEN
      PRINT #RptHandle, TAB(45); "PMT P";
    ELSE
      PRINT #RptHandle, TAB(45); "PYMT";
    END IF
  END IF
  IF TranRec(1).TranType = 3 THEN
    PRINT #RptHandle, TAB(45); "ABATE";
  END IF

  IF TranRec(1).TranType = 4 THEN
    PRINT #RptHandle, TAB(45); "INT";

'temp
'    GET PPFile, TranRec(1).CustomerRec, TaxCustRec(1)
'    TaxCustRec(1).LastTrans = TranRec(1).LastTrans
'    PUT PPFile, TranRec(1).CustomerRec, TaxCustRec(1)
'    Interest# = TranRec(1).Amount
'    WhatTr& = TranRec(1).BelongTo
'    GET TaxTran, WhatTr&, TranRec(1)
'    TranRec(1).Revenue.Interest = Round#(TranRec(1).Revenue.Interest - Interest#)
'    PUT TaxTran, WhatTr&, TranRec(1)
'temp

  END IF

  IF TranRec(1).TranType = 5 THEN
    PRINT #RptHandle, TAB(45); "PEN";
  END IF
  IF TranRec(1).TranType = 6 THEN
    PRINT #RptHandle, TAB(45); "COLL";
  END IF
  IF TranRec(1).TranType = 7 THEN
    PRINT #RptHandle, TAB(45); "ADJ";
  END IF
  PRINT #RptHandle, TAB(52); USING "#####"; TranRec(1).CustomerRec;

  IF TrType = 2 OR TrType = 3 THEN
    PRINT #RptHandle, TAB(60); TCnt&;
'    PRINT #RptHandle, TAB(60); TaxYear$;
  ELSE
    PRINT #RptHandle, TAB(60); TranRec(1).TaxYear;
  END IF
  PRINT #RptHandle, TAB(68); USING "$$######,#.##"; TranRec(1).Amount
  LineCnt = LineCnt + 1
  TotalAmt# = TotalAmt# + TranRec(1).Amount
  TotalAmt# = Round#(TotalAmt#)
  RETURN
  
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
  PRINT #RptHandle, "Date Range: Beg on "; BDate$; " Ending on "; EDate$
  PRINT #RptHandle, "Tran Date"; TAB(12); "Description"; TAB(45); "Type"; TAB(52); "Acct #"; TAB(61); "Year"; TAB(70); "  Amount"
  PRINT #RptHandle, Dash80$
  LineCnt = 7
  RETURN
  
PrintTransJourRptEnding:
  PRINT #RptHandle, Dash80$
  PRINT #RptHandle, "Total Transaction Amount "; USING "$$######,#.##"; TotalAmt#
  PRINT #RptHandle,
  PRINT #RptHandle, FF$
  RETURN
  
  
  
END SUB

