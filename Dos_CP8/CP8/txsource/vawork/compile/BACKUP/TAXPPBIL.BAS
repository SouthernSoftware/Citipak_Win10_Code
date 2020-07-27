DEFINT A-Z
DECLARE SUB PostTaxBills ()
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB RestScrn (Array%())
DECLARE SUB SaveScrn (Array%())
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB MakeCustIndex (IdxType)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB DisplayTaxScrn (ScrnName$)
DECLARE SUB TAXPrebilling ()
DECLARE SUB TAXCustomerMenu ()
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
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION Date2Num% (TheDate$)
  
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
  
DECLARE FUNCTION Round# (DblNum#)
  
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
DECLARE SUB GetPersRecList (PersRecs() AS LONG, CustRec&)
DECLARE SUB GetPropRecList (PropRecs() AS LONG, CustRec&)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE FUNCTION FUsing$ (Number$, Image$)
  
  CONST False = 0, True = NOT False
  
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TaxBill.BI'
  '$INCLUDE: 'TAXCONST.BI'
  '$INCLUDE: 'PROPAbst.BI'
  '$INCLUDE: 'TAXBINFO.BI'
  
  STACK 5000
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 6)
  
  MChoice$(1) = "Print Pre-Billing Register         "
  MChoice$(2) = "Print/Reprint Tax Bills "
  MChoice$(3) = "Post Tax Bills"
  MChoice$(4) = "Apply Interest"
  MChoice$(5) = "Apply Penalty"
  MChoice$(6) = "Exit to OS"
  
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
  Row = ((24 - (UBOUND(MChoice$))) \ 2)
  Col = ((80 - MaxLen) \ 2) - 1
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    
    ClearBack
    
    TitleBox 2, Col, MaxLen + 3, "Personal Property Tax Billing Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      TAXPrebilling
    CASE 2
      RUN "TXPRNTPB"
    CASE 3
      PostTaxBills
    CASE 4
      RUN "TAXINT"
    CASE 5
    CASE 6
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

SUB PostTaxBills
  
  REDIM TaxBill(1)  AS VAPPTaxBillType
  REDIM TaxCust(1)  AS TaxCustType
  REDIM TaxTrans(1) AS TaxTransactionType
  REDIM PersRec(1)  AS PersonalRecType
  REDIM BillInfo(1) AS VAPPTaxBillInfoType
  REDIM Revenue(1) AS RevSourceType
  
  
  BillLen = LEN(TaxBill(1))
  BillInfoLen = LEN(BillInfo(1))
  
  PersRecLen = LEN(PersRec(1))
  
  
  
  FGetAH "TAXBINF2.DAT", BillInfo(1), BillInfoLen, 1
  
  REDIM TempScrn(0)
  
  ClearBack
  ShowProcessingScrn "Verifying Billing Information"
  BillFile = FREEFILE
  OPEN PPTaxBillFile FOR RANDOM AS BillFile LEN = BillLen
  NumOfRec& = LOF(BillFile) / BillLen
  
  FOR Cnt& = 1 TO NumOfRec&
    GET #BillFile, Cnt&, TaxBill(1)
    IF TaxBill(1).BillPrinted THEN
      Printed = Printed + 1
    END IF
    ShowPctComp Cnt&, NumOfRec&
  NEXT
  CLOSE
  
  IF Printed < 1 THEN
    GOSUB ShowPostError
    GOTO PostExit
  END IF
  
  PostFlag = MsgBox%("TAX", "POSTBILL")
  
  SELECT CASE PostFlag
  CASE 1
    GOTO Postem
  CASE 2
    CursorOff
    ClearBack
    DisplayTaxScrn "POSTCAN"
    WaitForAction
    GOTO PostExit
  END SELECT
  
Postem:
  ClearBack
  ShowProcessingScrn "Posting Tax Billing Transactions"
  
  'Open All Needed Files Here
  PersFile = FREEFILE
  OPEN TaxPersFile FOR RANDOM SHARED AS PersFile LEN = PersRecLen
  
  
  BillFile = FREEFILE
  OPEN PPTaxBillFile FOR RANDOM AS BillFile LEN = BillLen
  NumOfRec& = LOF(BillFile) / BillLen
  
  CustFile = FREEFILE
  OPEN PPTaxCustFile FOR RANDOM AS CustFile LEN = LEN(TaxCust(1))
  
  TransFile = FREEFILE
  OPEN TaxTransFile FOR RANDOM AS TransFile LEN = LEN(TaxTrans(1))
  
  FOR Cnt& = 1 TO NumOfRec&
    GET #BillFile, Cnt&, TaxBill(1)
    IF TaxBill(1).BillPrinted THEN
      
      'Update the Transaction File First
      TaxTrans(1).TransDate = Date2Num%(DATE$)
      TaxTrans(1).TAXYEAR = TaxBill(1).TAXYEAR
      TaxTrans(1).TranType = 1  '1=Bill 2=Payment 3=Release 4=Interest 5=Penalty 6=Collection/Ad Cost Billing
      TaxTrans(1).BillType = "P"                'R=Real P=Personal Property C=Combined (NC/GA)
      TaxTrans(1).Amount = TaxBill(1).TotalBillDue              'Total Transaction Amount
      TaxTrans(1).Revenue.Principle1 = TaxBill(1).PersTaxDue + TaxBill(1).FETaxDue
      TaxTrans(1).Revenue.Principle2 = TaxBill(1).MTTaxDue
      TaxTrans(1).Revenue.Principle3 = TaxBill(1).MCTaxDue
      TaxTrans(1).Revenue.Principle4 = 0 'TaxBill(1).FETaxDue   for ashland this keeps qual volunteer taxed amt have no farm equip
      TaxTrans(1).Revenue.Principle5 = TaxBill(1).MHTaxDue
      TaxTrans(1).Revenue.Interest = 0
      TaxTrans(1).Revenue.Penalty = 0
      TaxTrans(1).Revenue.Collection = 0
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
      
      
      TaxTrans(1).Description = "Tax Bill #" + STR$(TaxBill(1).BillNumber)
      TaxTrans(1).Posted2GL = "N"
      TaxTrans(1).CustomerRec = TaxBill(1).CustRec
      TaxTrans(1).LastTrans = 0
      TaxTrans(1).BelongTo = 0
      TaxTrans(1).Padding = ""
      
      'Increment Transaction File Record Count
      NextRecord& = (LOF(TransFile) / LEN(TaxTrans(1))) + 1
      PUT #TransFile, NextRecord&, TaxTrans(1)
      
      'Update the Customer Pointers Now
      GET #CustFile, TaxBill(1).CustRec, TaxCust(1)
      
      IF TaxCust(1).LastTrans = 0 THEN
        TaxCust(1).LastTrans = NextRecord&
        PUT #CustFile, TaxBill(1).CustRec, TaxCust(1)
      ELSE
        Previous& = TaxCust(1).LastTrans
        TaxCust(1).LastTrans = NextRecord&
        PUT #CustFile, TaxBill(1).CustRec, TaxCust(1)
        
        GET #TransFile, NextRecord&, TaxTrans(1)
        TaxTrans(1).LastTrans = Previous&
        PUT #TransFile, NextRecord&, TaxTrans(1)
      END IF
      
      'Now Update the Property Records with the Tax Year to prevent duplicate billing per year
      
      PropRecord& = TaxBill(1).PersPropRecord
        WHILE PropRecord& <> 0
         GET PersFile, PropRecord&, PersRec(1)
         PersRec(1).LastYrPrinted = TaxBill(1).TAXYEAR
         PUT PersFile, PropRecord&, PersRec(1)
         PropRecord& = PersRec(1).NextRec
        WEND
      
      
    END IF
    ShowPctComp Cnt&, NumOfRec&
  NEXT
  
  CLOSE
  
  DisplayTaxScrn "UPDATEOK"
  
  WaitForAction
  
  'Now Delete the Tax Bill File so Duplicate's Cannot Be Reproduced
  S$ = "DEL TAXPBILL.DAT"
  SHELL S$
  
PostExit:
  EXIT SUB
  
  
ShowPostError:
  ClearBack
  DisplayTaxScrn "ERRSCRN1"
  QPrintRC "THERE ARE NO TAX BILLS TO POST!", 10, 25, -1
  QPrintRC "Please correct and try again.", 12, 26, -1
  WaitForAction
  RETURN
END SUB

SUB TAXPrebilling
  
  SHARED Choice$()
  
  REDIM TaxCust(1)  AS TaxCustType
  REDIM TaxTrans(1) AS TaxTransactionType
  REDIM PersRec(1)  AS PersonalRecType
  REDIM TBillRec(1) AS VAPPTaxBillType
  REDIM BillInfo(1) AS VAPPTaxBillInfoType
  
  REDIM Fmt$(1 TO 3)
  
  TaxCustLen = LEN(TaxCust(1))  'Length of Cust Record Structure
  PersRecLen = LEN(PersRec(1))
  TBillLen = LEN(TBillRec(1))
  BillInfoLen = LEN(BillInfo(1))
  
  REDIM TempScrn(0)
  
  REDIM Choice$(0 TO 4, 0)
  Choice$(0, 0) = "9"
  Choice$(1, 0) = "1) Account Number Order"
  Choice$(2, 0) = "2) Customer Name Order"
  Choice$(3, 0) = "3) Search Name Order"
  Choice$(4, 0) = "4) Social Security Order"
  
  Dash$ = STRING$(79, "-")
  TheDate$ = DATE$
  Fmt$(1) = "$$########,#"
  Fmt$(2) = "#####"
  Fmt$(3) = "#######,#"
  FF$ = CHR$(12)
  MaxLines = 56
  LineCnt = 0
  
  LibName$ = "TAX"
  ScrnName$ = "VPREBILL"
  
  '--Initialize the form name array
  '--Get the total number of fields from all pages
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  '--Clear all fields
  ClearBack
  
  DisplayTaxScrn ScrnName$
  
  ShowCursor
  Form$(3, 0) = "0"
  Form$(4, 0) = "0"
  Form$(5, 0) = "0"
  Form$(6, 0) = "0"
  Form$(7, 0) = "0"

  
  Action = 1
  FirstTime = True
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF (Frm(1).KeyCode = -72) OR (Frm(1).KeyCode = -15) THEN
      FOR zz = Frm(1).FldNo TO 1 STEP -1        'loop backwards
        IF Fld(zz).Protected = False THEN
          Frm(1).FldNo = zz
          Action = 1
          EXIT FOR
        END IF
      NEXT
    END IF
    
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE -68    'F10Key    Proceed with look up
      GOSUB CheckPreInfo
      IF PreInfoOK THEN
        OKFlag = True
      END IF
    CASE 27
      ExitFlag = True
    END SELECT
    
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 20
        SELECT CASE Frm(1).MCol
        CASE 29 TO 40           'ESC Cancel button
          PressButton 27, 20, 29, 40
        CASE 42 TO 54           'F10 Proceed Button
          PressButton -68, 20, 42, 54
        END SELECT
      END SELECT
    END IF
    
  LOOP UNTIL ExitFlag OR OKFlag
  
  IF ExitFlag THEN
    GOTO ExitPreBill
  END IF
  
  ClearBack
  
  IF Order$ = "2" THEN
    UsingIdx = True
    MakeCustIndex 1
  ELSEIF Order$ = "3" THEN
    UsingIdx = True
    MakeCustIndex 2
  ELSEIF Order$ = "4" THEN
    UsingIdx = True
    MakeCustIndex 3
  END IF
  IF UsingIdx THEN
    NumIdxRecs = FileSize&("TAXTEMP.IDX") / 2
    REDIM IdxBuff(1 TO NumIdxRecs) AS INTEGER
    FGetAH "TAXTEMP.IDX", IdxBuff(1), 2, NumIdxRecs
  END IF
  
  ShowProcessingScrn "Calculating Tax Prebilling"
  
  IF Exist%(PPTaxBillFile) THEN
     KILL PPTaxBillFile            'kill any old work file
  END IF
  
  TBillFile = FREEFILE
  OPEN PPTaxBillFile FOR RANDOM SHARED AS TBillFile LEN = TBillLen

  PersFile = FREEFILE
  OPEN TaxPersFile FOR RANDOM SHARED AS PersFile LEN = PersRecLen

  TaxFile = FREEFILE
  OPEN PPTaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxCustLen
  NumOfCRecs& = LOF(TaxFile) \ TaxCustLen
  
  'Prebill calc start *******************************
  FOR Cnt& = 1 TO NumOfCRecs&
    IF UsingIdx THEN
      CustAcct& = IdxBuff(Cnt&)
    ELSE
      CustAcct& = Cnt&
    END IF
    REDIM TBillRec(1) AS VAPPTaxBillType            'make a new empty record each time
    GET #TaxFile, CustAcct&, TaxCust(1)         'get cust record
    LateFlag$ = "N"             'Initialize Late Flag
    PastFlagSet = 0             'Initialize Past Balance Flag
    
    IF TaxCust(1).FirstPersRec <= 0 AND TaxCust(1).FirstPropRec <= 0 THEN
      NoProp = 1
      GOSUB SetCustInfo
      GOSUB WriteIt2Disk
      GOTO PreBillSkip
    END IF
    NoProp = 0
    GOSUB SetCustInfo
    GOSUB GetPersInfo
    GOSUB WriteIt2Disk
    ShowPctComp Cnt&, NumOfCRecs&
PreBillSkip:
    NEXT
    CLOSE
  'Prebill Calc END   *******************************

  'Prebill Print Start #############
  TotalPers# = 0
  TotalMH# = 0
  TotalMC# = 0
  TotalFE# = 0
  TotalMT# = 0
  TotalEx# = 0
  NumBills& = 0
  
  ShowProcessingScrn "Creating Pers Prop Tax Prebill Report"
  RptFile = FREEFILE
  OPEN PPTaxPreRptFile FOR OUTPUT AS #RptFile
  
  TBillFile = FREEFILE
  OPEN PPTaxBillFile FOR RANDOM SHARED AS TBillFile LEN = TBillLen
  NumOfRec& = LOF(TBillFile) / TBillLen

  GOSUB PreBillHeading

  FOR Cnt& = 1 TO NumOfRec&
  GET #TBillFile, Cnt&, TBillRec(1)
  IF YearFlag > 50 AND TBillRec(1).BillNumber = -1 THEN
   ELSE
      PRINT #RptFile, LTRIM$(LEFT$(TBillRec(1).CustName, 30));
  IF TBillRec(1).BillNumber = -1 THEN
      PRINT #RptFile, TAB(35); "No Bill";
   ELSE

       PRINT #RptFile, TAB(35); USING Fmt$(2); TBillRec(1).BillNumber;
   END IF
       PRINT #RptFile, TAB(45); USING Fmt$(3); TBillRec(1).PersValue;
       PRINT #RptFile, TAB(56); USING Fmt$(3); TBillRec(1).FEValue;
       PRINT #RptFile, TAB(67); USING Fmt$(3); TBillRec(1).MTVALUE;
       PRINT #RptFile, TAB(78); USING Fmt$(3); TBillRec(1).MHVALUE;
       PRINT #RptFile, TAB(98); "VALUE"
      PRINT #RptFile, TBillRec(1).CustAdd1
      PRINT #RptFile, TBillRec(1).CustAdd2;
       PRINT #RptFile, TAB(36); "Acct # "
      PRINT #RptFile, RTRIM$(TBillRec(1).CustAdd3); " ";
       PRINT #RptFile, TBillRec(1).CustZip;
       PRINT #RptFile, TAB(36); USING "####"; TBillRec(1).CustRec;
       PRINT #RptFile, TAB(44); USING "#####,#.##"; TBillRec(1).PersTaxDue;
       PRINT #RptFile, TAB(55); USING "#####,#.##"; TBillRec(1).FETaxDue;
       PRINT #RptFile, TAB(66); USING "#####,#.##"; TBillRec(1).MTTaxDue;
       PRINT #RptFile, TAB(77); USING "#####,#.##"; TBillRec(1).MHTaxDue;
       PRINT #RptFile, TAB(88); USING "#####,#.##"; TBillRec(1).TotalBillDue;
       PRINT #RptFile, TAB(99); "TAX"

      PRINT #RptFile, STRING$(102, "-")
      
      
      LineCnt = LineCnt + 5
   END IF                               'end test for sup billing
      IF TBillRec(1).BillNumber = -1 THEN
      ELSE
      TotalPers# = Round#(TotalPers# + TBillRec(1).PersValue)
      TotalMH# = Round#(TotalMH# + TBillRec(1).MHVALUE)
      TotalMC# = Round#(TotalMC# + TBillRec(1).MCVALUE)
      TotalFE# = Round#(TotalFE# + TBillRec(1).FEValue)
      TotalMT# = Round#(TotalMT# + TBillRec(1).MTVALUE)
      TotalPersTax# = Round#(TotalPersTax# + TBillRec(1).PersTaxDue)
      TotalMHTax# = Round#(TotalMHTax# + TBillRec(1).MHTaxDue)
      TotalMCTax# = Round#(TotalMCTax# + TBillRec(1).MCTaxDue)
      TotalFETax# = Round#(TotalFETax# + TBillRec(1).FETaxDue)
      TotalMTTax# = Round#(TotalMTTax# + TBillRec(1).MTTaxDue)


      TotalBills# = Round#(TotalBills# + TBillRec(1).TotalBillDue)
      TotalLate# = Round#(TotalLate# + TBillRec(1).LateTaxDue)
      TotalPast# = Round#(TotalPast# + TBillRec(1).PriorYrBalance)
      END IF

      IF TBillRec(1).TotalBillDue > 0 THEN
        NumBills& = NumBills& + 1
      END IF
      
      IF LineCnt >= MaxLines THEN
        PRINT #RptFile, CHR$(12);
        GOSUB PreBillHeading
      END IF
    
    ShowPctComp Cnt&, NumOfRec&
  
  NEXT
  
  PRINT #RptFile, FF$;
  PRINT #RptFile, "Personal Property Summary Page"
  PRINT #RptFile, "Date: "; TheDate$; TAB(65); "Page #"; page
  PRINT #RptFile, ""
  PRINT #RptFile, "Number of Bills to Process: "; USING "#####,#"; NumBills&
  PRINT #RptFile, "      Total Pers Valuation: "; USING Fmt$(1); TotalPers#
  PRINT #RptFile, " Total Qual Vol. Valuation: "; USING Fmt$(1); TotalFE#
  PRINT #RptFile, "      Total MH   Valuation: "; USING Fmt$(1); TotalMH#
  PRINT #RptFile, "      Total MC   Valuation: "; USING Fmt$(1); TotalMC#
  PRINT #RptFile, "      Total MT   Valuation: "; USING Fmt$(1); TotalMT#
  PRINT #RptFile,
  PRINT #RptFile, "     Total Pers Tax Amount: "; USING "$$########,#.##"; TotalPersTax#
  PRINT #RptFile, "Total Qual Vol. Tax Amount: "; USING "$$########,#.##"; TotalFETax#
  PRINT #RptFile, "Total Mobl Home Tax Amount: "; USING "$$########,#.##"; TotalMHTax#
  PRINT #RptFile, "Total Merch Cap Tax Amount: "; USING "$$########,#.##"; TotalMCTax#
  PRINT #RptFile, "Total Mach/Tool Tax Amount: "; USING "$$########,#.##"; TotalMTTax#
  PRINT #RptFile, "         Total Late Amount: "; USING "$$########,#.##"; TotalLate#
  PRINT #RptFile, ""
  PRINT #RptFile, "      Total Amount to Bill: "; USING "$$########,#.##"; TotalBills#
  PRINT #RptFile, "    Total Past Amt to Bill: "; USING "$$########,#.##"; TotalPast#
  PRINT #RptFile, "Grand Total Amount to Bill: "; USING "$$########,#.##"; Round#(TotalPast# + TotalBills#)
  PRINT #RptFile, ""
  
  
  CLOSE
  
  
  PrintRptFile "Tax Prebilling Report.", PPTaxPreRptFile, 1, RetCode%, 1
  'PreBill Print END   #############

  ERASE TaxCust, PersRec, TBillRec, BillInfo
  ERASE Fmt$, Frm, Form$, Fld, TempScrn, IdxBuff
  

ExitPreBill:
  
  EXIT SUB
  
PreBillHeading:
  page = page + 1
  PRINT #RptFile, TAB(10); "Persons and Personal Property Assessed For Taxation for the Year "; Whatyear
  PRINT #RptFile, ""
  PRINT #RptFile, "Town of Ashland Virginia by Lois B Chenault, Commissioner of Revenue"
  IF YearFlag > 50 THEN
   PRINT #RptFile, "Supplemental Billing"
  ELSE
   PRINT #RptFile, ""
  END IF
  PRINT #RptFile, "Date: "; TheDate$; TAB(82); "Page #"; page


  PRINT #RptFile, TAB(35); "Bill #"; TAB(45); "Personal"; TAB(56); "  Qual"; TAB(67); "Machinery"; TAB(78); "Mobile"
  PRINT #RptFile, TAB(45); "Property"; TAB(56); "  Vol"; TAB(67); "& Tools"; TAB(78); "Homes"
  PRINT #RptFile, "Customer Name & Address"; TAB(45); "Rate "; USING "#.##"; PERSRATE#;
   PRINT #RptFile, TAB(56); "Rate "; USING " .###"; FERate#;
   PRINT #RptFile, TAB(67); "Rate "; USING "#.##"; MTRate#;
   PRINT #RptFile, TAB(78); "Rate "; USING "#.##"; MHRate#

  
  PRINT #RptFile, STRING$(102, "-")
  LineCnt = 7
  RETURN
  
GetPersInfo:
  PersExmp# = 0
  PersValue# = 0
  MHVALUE# = 0
  MCVALUE# = 0
  FEValue# = 0
  MTVALUE# = 0
   
  PropertyRec! = TaxCust(1).FirstPersRec
   PersValue# = 0
   MHVALUE# = 0
   MCVALUE# = 0
   FEValue# = 0
   MTVALUE# = 0

  WHILE PropertyRec! <> 0
    GET #PersFile, PropertyRec!, PersRec(1)
    IF PersRec(1).LastYrPrinted = Whatyear THEN
      YearFlag = YearFlag + 1   'Count Number of Redo if greater than 50 then sup billing
      ELSE
      IF MID$(PersRec(1).DESC5, 3, 1) = "5" THEN
       FEValue# = FEValue# + Round#(PersRec(1).PERSVAL)
       ELSE
       PersValue# = PersValue# + Round#(PersRec(1).PERSVAL)
      END IF
      
      MHVALUE# = MHVALUE# + Round#(PersRec(1).MHVALUE)
      MCVALUE# = MCVALUE# + Round#(PersRec(1).MCVALUE)
      MTVALUE# = MTVALUE# + Round#(PersRec(1).MTVALUE)
    END IF
    PropertyRec! = PersRec(1).NextRec
   
   WEND


    PersTaxDue# = Round#((PersValue# / 100) * PERSRATE#)
    MHTaxDue# = Round#((MHVALUE# / 100) * MHRate#)
    MCTaxDue# = Round#((MCVALUE# / 100) * MCRate#)
    FETaxDue# = Round#((FEValue# / 100) * FERate#)
    MTTaxDue# = Round#((MTVALUE# / 100) * MTRate#)


    TBillRec(1).ExptValue = 0
    TBillRec(1).PersValue = PersValue#
    TBillRec(1).MHVALUE = MHVALUE#
    TBillRec(1).MCVALUE = MCVALUE#
    TBillRec(1).FEValue = FEValue#
    TBillRec(1).MTVALUE = MTVALUE#

    TBillRec(1).PersTaxDue = PersTaxDue#
    TBillRec(1).MHTaxDue = MHTaxDue#
    TBillRec(1).MCTaxDue = MCTaxDue#
    TBillRec(1).FETaxDue = FETaxDue#
    TBillRec(1).MTTaxDue = MTTaxDue#
    
    TBillRec(1).PersPropRecord = TaxCust(1).FirstPersRec
    TBillRec(1).PersTaxRate = PERSRATE#
    TBillRec(1).TotalBillDue = Round#(TBillRec(1).PersTaxDue + TBillRec(1).MHTaxDue + TBillRec(1).MCTaxDue + TBillRec(1).FETaxDue + TBillRec(1).MTTaxDue)
    IF PersRec(1).LATELIST = "Y" THEN
      LateFlag$ = "Y"
      LateAmt# = Round#(TBillRec(1).TotalBillDue * (LATELIST# / 100))
      TBillRec(1).LateTaxDue = LateAmt#
      TBillRec(1).TotalBillDue = Round#(TBillRec(1).TotalBillDue + LateAmt#)
    END IF
    RETURN
  
WriteIt2Disk:
  'write the info out to disk here.
  TBillRec(1).BillPrinted = False
  IF TBillRec(1).TotalBillDue > 0 THEN
    TBillRec(1).BillNumber = BILLNUM&
  ELSE
    TBillRec(1).BillNumber = -1
  END IF
  
  PUT #TBillFile, , TBillRec(1)
  IF TBillRec(1).TotalBillDue > 0 THEN
    BILLNUM& = BILLNUM& + 1
  END IF
RETURN
  
SetCustInfo:
  TBillRec(1).CustRec = CustAcct&               'cust acct rec
  CustName$ = QPTrim$(TaxCust(1).LName) + "  " + QPTrim$(TaxCust(1).FNAME)
  TBillRec(1).CustName = CustName$
  TBillRec(1).CustAdd1 = TaxCust(1).Addr1
  TBillRec(1).CustAdd2 = TaxCust(1).Addr2
  CitySt$ = QPTrim$(TaxCust(1).City) + " " + TaxCust(1).State
  TBillRec(1).CustAdd3 = CitySt$
  TBillRec(1).CustZip = TaxCust(1).Zip
  TBillRec(1).CustPin = CustAcct&
  TBillRec(1).TAXYEAR = Whatyear
  
  'Set Prior Balance if any
  
  GOSUB GetPastBalance
  IF ABalance# <> 0 THEN
    IF PastFlagSet = 0 THEN
      TBillRec(1).PriorYrBalance = ABalance#
    END IF
    PastFlagSet = 1
  END IF
RETURN
  
CheckPreInfo:
  Whatyear = Value#(Form$(1, 0), ECode)
  BILLNUM& = Value#(Form$(2, 0), ECode)
  PERSRATE# = Value#(Form$(3, 0), ECode)
  MHRate# = Value#(Form$(4, 0), ECode)
  MCRate# = Value#(Form$(5, 0), ECode)
  FERate# = Value#(Form$(6, 0), ECode)
  MTRate# = Value#(Form$(7, 0), ECode)
  LATELIST# = Value#(Form$(8, 0), ECode)
  Order$ = LEFT$(Form$(9, 0), 1)
  DueDate$ = Form$(10, 0)

  IF Whatyear < 1997 OR Whatyear > 2010 THEN
    GOSUB ShowInfoErr
    Frm(1).FldNo = 1
    GOTO InfoErrorExit
  END IF
  
  IF BILLNUM& = 0 THEN
    GOSUB ShowInfoErr
    Frm(1).FldNo = 2
    GOTO InfoErrorExit
  END IF
  
  IF LEN(QPTrim$(Order$)) = 0 THEN
    GOSUB ShowInfoErr
    Frm(1).FldNo = 6
    GOTO InfoErrorExit
  END IF
  
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(BillInfo(1)), VARPTR(BillInfo(1)), BillInfoLen, 0

  FPutAH "TAXBINF2.DAT", BillInfo(1), BillInfoLen, 1    'Pers Prop =2
  
  PreInfoOK = True
  
InfoErrorExit:
  
RETURN
  
ShowInfoErr:
  SaveScrn TempScrn()
  DisplayTaxScrn "ERRSCRN1"
  QPrintRC "Invalid Prebilling Information!", 10, 25, -1
  QPrintRC "Please correct and try again.", 12, 26, -1
  WaitForAction
  RestScrn TempScrn()
RETURN
  
GetPastBalance:
  Balance# = 0
  ABalance# = 0
  IF TaxCust(1).LastTrans > 0 THEN
    TransFile = FREEFILE
    OPEN "TaxTrans.dat" FOR RANDOM SHARED AS TransFile LEN = LEN(TaxTrans(1))
    TransRecord& = TaxCust(1).LastTrans
    DO WHILE TransRecord& <> 0
      GET TransFile, TransRecord&, TaxTrans(1)
      IF TaxTrans(1).TranType = 1 THEN
        Balance# = TaxTrans(1).Revenue.Principle1 + TaxTrans(1).Revenue.Principle2 + TaxTrans(1).Revenue.Principle3 + TaxTrans(1).Revenue.Principle4 + TaxTrans(1).Revenue.Principle5
        Balance# = Balance# + TaxTrans(1).Revenue.Interest + TaxTrans(1).Revenue.Penalty + TaxTrans(1).Revenue.Collection
        Balance# = Balance# - (TaxTrans(1).Revenue.Principle1Pd + TaxTrans(1).Revenue.Principle2Pd + TaxTrans(1).Revenue.Principle3Pd + TaxTrans(1).Revenue.Principle4Pd + TaxTrans(1).Revenue.Principle5Pd)
        Balance# = Balance# - (TaxTrans(1).Revenue.InterestPd + TaxTrans(1).Revenue.PenaltyPd + TaxTrans(1).Revenue.CollectionPd)
        Balance# = Round#(Balance#)
      END IF
      ABalance# = ABalance# + Balance#
      TransRecord& = TaxTrans(1).LastTrans
    LOOP
    CLOSE TransFile
  END IF
RETURN
  
END SUB

