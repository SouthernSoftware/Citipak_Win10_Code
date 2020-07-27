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
  
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
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
  
  TYPE MortRecType
    MORTCODE AS STRING * 8
    MortRec  AS INTEGER
  END TYPE

  '$INCLUDE: 'DefCnf.BI'
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

  DIM SHARED TaxSetUpRec AS TaxMasterType

  TxFile = FREEFILE
  OPEN "TAXSETUP.DAT" FOR RANDOM AS #TxFile LEN = LEN(TaxSetUpRec)
  GET TxFile, 1, TaxSetUpRec
  CLOSE TxFile
  IF INSTR(TaxSetUpRec.Name, "MAGNOLI") > 0 THEN
    MagFlag = True
  END IF

  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 7)
  
  MChoice$(1) = "Print Tax Pre-Billing Register"
  MChoice$(2) = "Print/Reprint Tax Bills "
  MChoice$(3) = "Post Tax Bills"
  MChoice$(4) = "Apply Interest"
  MChoice$(5) = "Print Late Notices"

  IF MagFlag THEN
    MChoice$(6) = "Print Alt Late Notices"
    MChoice$(7) = "Exit to OS"
    TStop = 7
  ELSE
    TStop = 6
    MChoice$(6) = "Exit to OS"
  END IF
  
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
    
    TitleBox 2, Col, MaxLen + 3, "Tax Billing 'Billing' Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      TAXPrebilling
    CASE 2
      RUN "TAXPRNTB"
    CASE 3
      PostTaxBills
    CASE 4
      RUN "TAXINT"
    CASE 5
      RUN "TAXLNOTE"
    CASE 6
      IF TStop = 6 THEN
        HideCursor
        ClearScrn
        END
      ELSE
        RUN "TAXDNOTE"
      END IF
    CASE 7
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
  
  REDIM TaxBill(1)  AS TaxBillType
  REDIM TaxCust(1)  AS TaxCustType
  REDIM TaxTrans(1) AS TaxTransactionType
  REDIM RealRec(1)  AS PropertyRecType
  REDIM PersRec(1)  AS PersonalRecType
  REDIM BillInfo(1) AS TaxBillInfoType
  REDIM Revenue(1) AS RevSourceType
  
  BillLen = LEN(TaxBill(1))
  BillInfoLen = LEN(BillInfo(1))
  RealRecLen = LEN(RealRec(1))
  PersRecLen = LEN(PersRec(1))

  REDIM TaxSetUp(1) AS TaxMasterType
  TxFile = FREEFILE
  OPEN "TAXSETUP.DAT" FOR RANDOM AS #TxFile LEN = LEN(TaxSetUpRec)
  IF LOF(TxFile) > 0 THEN
    GET TxFile, 1, TaxSetUp(1)
  END IF
  CLOSE
  IF INSTR(TaxSetUp(1).Name, "CAROLINA SHORES") > 0 THEN
    ShoresFlag = True
  END IF
  
  FGetAH "TAXBINFO.DAT", BillInfo(1), BillInfoLen, 1
  
  REDIM TempScrn(0)
  
  ClearBack
  ShowProcessingScrn "Verifying Billing Information"

  BillFile = FREEFILE
  OPEN TaxBillFile FOR RANDOM AS BillFile LEN = BillLen
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
  
  RealFile = FREEFILE
  OPEN TaxPropFile FOR RANDOM SHARED AS RealFile LEN = RealRecLen
  
  BillFile = FREEFILE
  OPEN TaxBillFile FOR RANDOM AS BillFile LEN = BillLen
  NumOfRec& = LOF(BillFile) / BillLen
  
  CustFile = FREEFILE
  OPEN TaxCustFile FOR RANDOM AS CustFile LEN = LEN(TaxCust(1))
  
  TransFile = FREEFILE
  OPEN TaxTransFile FOR RANDOM AS TransFile LEN = LEN(TaxTrans(1))
  
  FOR Cnt& = 1 TO NumOfRec&
    GET #BillFile, Cnt&, TaxBill(1)
    IF TaxBill(1).BillPrinted THEN
      'Update the Transaction File First
      TaxTrans(1).TransDate = Date2Num%(DATE$)
      TaxTrans(1).TAXYEAR = TaxBill(1).TAXYEAR
      TaxTrans(1).TranType = 1  '1=Bill 2=Payment 3=Release 4=Interest 5=Penalty 6=Collection/Ad Cost Billing
      TaxTrans(1).BillType = "C"         'R=Real P=Personal Property C=Combined (NC/GA)
      TaxTrans(1).Amount = TaxBill(1).TotalBillDue      'Total Transaction Amount

      IF ShoresFlag THEN
        TaxTrans(1).Revenue.Principle1 = Round(TaxBill(1).TotalBillDue + TaxBill(1).CarShore)
      ELSE
        TaxTrans(1).Revenue.Principle1 = TaxBill(1).TotalBillDue
      END IF

      TaxTrans(1).Revenue.Principle2 = 0
      TaxTrans(1).Revenue.Principle3 = 0
      TaxTrans(1).Revenue.Principle4 = 0
      TaxTrans(1).Revenue.Principle5 = 0
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
      IF TaxBill(1).RealPropRecord > 0 THEN
        GET RealFile, TaxBill(1).RealPropRecord, RealRec(1)
        RealRec(1).LastYrPrinted = TaxBill(1).TAXYEAR
        RealRec(1).PROPDISC = "N"
        PUT RealFile, TaxBill(1).RealPropRecord, RealRec(1)
      END IF
      
      IF TaxBill(1).PersPropRecord > 0 THEN
        GET PersFile, TaxBill(1).PersPropRecord, PersRec(1)
        PersRec(1).LastYrPrinted = TaxBill(1).TAXYEAR
        PersRec(1).DISCOV = "N"
        PUT PersFile, TaxBill(1).PersPropRecord, PersRec(1)
      END IF
      
    END IF
    ShowPctComp Cnt&, NumOfRec&
  NEXT
  
  CLOSE
  
  DisplayTaxScrn "UPDATEOK"
  
  WaitForAction
  
  'Now Delete the Tax Bill File so Duplicate's Cannot Be Reproduced
  S$ = "DEL TAXTBILL.DAT"
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
  REDIM RealRec(1)  AS PropertyRecType
  REDIM PersRec(1)  AS PersonalRecType
  REDIM TBillRec(1) AS TaxBillType
  REDIM BillInfo(1) AS TaxBillInfoType

  REDIM MortCodeRec(1) AS MortCodeRecType
  REDIM MortCodes(1 TO 1) AS MortRecType

  REDIM Fmt$(1 TO 4)
  
  REDIM TaxSetUp(1) AS TaxMasterType
  TxFile = FREEFILE
  OPEN "TAXSETUP.DAT" FOR RANDOM AS #TxFile LEN = LEN(TaxSetUpRec)
  IF LOF(TxFile) > 0 THEN
    GET TxFile, 1, TaxSetUp(1)
  END IF
  CLOSE

  IF INSTR(TaxSetUp(1).Name, "CAROLINA SHORES") > 0 THEN
    ShoresFlag = True
  END IF

  IF INSTR(TaxSetUp(1).Name, "SYLVA") > 0 THEN
    SYLVAFlag = True
  END IF
  
  TaxCustLen = LEN(TaxCust(1))  'Length of Cust Record Structure
  RealRecLen = LEN(RealRec(1))
  PersRecLen = LEN(PersRec(1))
  TBillLen = LEN(TBillRec(1))
  BillInfoLen = LEN(BillInfo(1))
  
  REDIM TempScrn(0)
  
  REDIM Choice$(0 TO 4, 0)
  Choice$(0, 0) = "6"
  Choice$(1, 0) = "1) Account Number Order"
  Choice$(2, 0) = "2) Customer Name Order"
  Choice$(3, 0) = "3) Search Name Order"
  Choice$(4, 0) = "4) Social Security Order"
  
  Dash$ = STRING$(79, "-")
  TheDate$ = DATE$
  Fmt$(1) = "$$########,#"
  Fmt$(2) = "########"
  Fmt$(3) = "########,#"
  Fmt$(4) = "$$########,#.##"

  MaxLines = 60
  LineCnt = 0
  
  LibName$ = "TAX"
  ScrnName$ = "PREBILL"
  
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
  
  Action = 1
  FirstTime = True
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF FirstTime THEN
      FirstTime = False
      LSET Form$(7, 0) = "N"
      Action = 1
    END IF

    IF (Frm(1).KeyCode = -72) OR (Frm(1).KeyCode = -15) THEN
      FOR zz = Frm(1).FldNo TO 1 STEP -1        'loop backwards
        IF Fld(zz).Protected = False THEN
          Frm(1).FldNo = zz
          Action = 1
          EXIT FOR
        END IF
      NEXT
    END IF
    
    IF Frm(1).PrevFld = 4 THEN
      LSET Form$(4, 0) = Form$(3, 0)
      Action = 2
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
  GOSUB LoadMortCodes

  IF Exist%(TaxBillFile) THEN
    KILL TaxBillFile            'kill any old work file
  END IF
  
  TBillFile = FREEFILE
  OPEN TaxBillFile FOR RANDOM SHARED AS TBillFile LEN = TBillLen
  PersFile = FREEFILE
  OPEN TaxPersFile FOR RANDOM SHARED AS PersFile LEN = PersRecLen
  RealFile = FREEFILE

  OPEN TaxPropFile FOR RANDOM SHARED AS RealFile LEN = RealRecLen
  TaxFile = FREEFILE

  OPEN TaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxCustLen
  NumOfCRecs& = LOF(TaxFile) \ TaxCustLen

  'Prebill calc start *******************************
  FOR Cnt& = 1 TO NumOfCRecs&
    IF UsingIdx THEN
      CustAcct& = IdxBuff(Cnt&)
    ELSE
      CustAcct& = Cnt&
    END IF
    REDIM TBillRec(1) AS TaxBillType            'make a new empty record each time

'dale
'IF CustAcct& < 1076 OR CustAcct& > 1077 THEN
'  GOTO PreBillSkip:
'END IF

    GET #TaxFile, CustAcct&, TaxCust(1)         'get cust record

    'IF CustAcct& = 3273 THEN STOP

    Done2Disk = 0
    LateAmt# = 0

    'IF CustAcct& = 17 THEN STOP

    IF TaxCust(1).Deleted <> 0 THEN
      GOTO PreBillSkip:
    END IF
    IF TaxCust(1).Active <> "Y" THEN
      InActive = InActive + 1
      GOTO PreBillSkip:
    END IF
    
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
    GOSUB GetRealInfo

PreBillSkip:
    ShowPctComp Cnt&, NumOfCRecs&
  NEXT
  CLOSE
  'Prebill Calc END   *******************************
  'Prebill Print Start #############

  TotalPers# = 0
  TotalReal# = 0
  TotalEx# = 0
  NumBills& = 0
  
  ShowProcessingScrn "Creating Tax Prebilling Report"
  RptFile = FREEFILE
  OPEN TaxPreRptFile FOR OUTPUT AS #RptFile
  
  TBillFile = FREEFILE
  OPEN TaxBillFile FOR RANDOM SHARED AS TBillFile LEN = TBillLen
  NumOfRec& = LOF(TBillFile) / TBillLen

  GOSUB PreBillHeading
  FOR Cnt& = 1 TO NumOfRec&
    GET #TBillFile, Cnt&, TBillRec(1)

    IF Discovery$ = "Y" AND TBillRec(1).BillNumber = -1 THEN

    ELSE
      PRINT #RptFile, USING "#####"; TBillRec(1).CustRec;
      PRINT #RptFile, TAB(8); LEFT$(TBillRec(1).CustName, 30);
      IF TBillRec(1).PriorYrBalance > 0 THEN
        PRINT #RptFile, TAB(49); USING "#####,#.##"; TBillRec(1).PriorYrBalance;
      END IF
      IF TBillRec(1).BillNumber = -1 THEN
        PRINT #RptFile, TAB(62); "N/A";
      ELSE
        PRINT #RptFile, TAB(60); USING Fmt$(2); TBillRec(1).BillNumber;
      END IF
      PRINT #RptFile, TAB(68); USING "######,#.##"; TBillRec(1).TotalBillDue

'GOTO donewithem:

      PRINT #RptFile, USING Fmt$(3); TBillRec(1).RealValue;
      PRINT #RptFile, TAB(15); USING Fmt$(3); TBillRec(1).PersValue;
      PRINT #RptFile, TAB(30); USING Fmt$(3); TBillRec(1).ExptValue;
      TValue# = Round#(TBillRec(1).RealValue + TBillRec(1).PersValue - TBillRec(1).ExptValue)
      IF TValue# < 0 THEN TValue# = 0
      PRINT #RptFile, TAB(45); USING Fmt$(3); TValue#;
      IF TBillRec(1).LateTaxDue > 0 THEN
        PRINT #RptFile, TAB(58); "Late="; USING "#####.##"; TBillRec(1).LateTaxDue
      ELSE
        PRINT #RptFile, ""
      END IF
      
      LineCnt = LineCnt + 2
      
      TotalPers# = Round#(TotalPers# + TBillRec(1).PersValue)
      TotalReal# = Round#(TotalReal# + TBillRec(1).RealValue)
      TotalEx# = Round#(TotalEx# + TBillRec(1).ExptValue)
      TotalBills# = Round#(TotalBills# + TBillRec(1).TotalBillDue)
      TotalLate# = Round#(TotalLate# + TBillRec(1).LateTaxDue)
      TotalPast# = Round#(TotalPast# + TBillRec(1).PriorYrBalance)
      
      IF TBillRec(1).TotalBillDue > 0 THEN
        NumBills& = NumBills& + 1
      END IF
      
      IF LineCnt >= MaxLines THEN
        PRINT #RptFile, CHR$(12);
        GOSUB PreBillHeading
      END IF
donewithem:
    END IF      'Test for Discovery Bills
    ShowPctComp Cnt&, NumOfRec&
  NEXT

  PRINT #RptFile, CHR$(12);
  PRINT #RptFile, TAB(20); "Property Tax Billing : Pre-Billing Register"
  PRINT #RptFile,
  PRINT #RptFile, "Date: "; TheDate$; TAB(65); "Page #"; page
  PRINT #RptFile, Dash$
  PRINT #RptFile, "Number of Bills to Process: "; USING "#####,#"; NumBills&
  PRINT #RptFile, "      Total Real Valuation: "; USING Fmt$(1); TotalReal#
  PRINT #RptFile, "      Total Pers Valuation: "; USING Fmt$(1); TotalPers#
  PRINT #RptFile, "          Total Exemptions: "; USING Fmt$(1); TotalEx#
  PRINT #RptFile, "      Total Amount to Bill: "; USING Fmt$(4); TotalBills#
  PRINT #RptFile, "    Total Past Amt to Bill: "; USING Fmt$(4); TotalPast#
  PRINT #RptFile, "Grand Total Amount to Bill: "; USING Fmt$(4); Round#(TotalPast# + TotalBills#)
  PRINT #RptFile, ""
  PRINT #RptFile, "         Total Late Amount: "; USING Fmt$(4); TotalLate#
  PRINT #RptFile, "           Inactive Skiped: "; USING "#####,#"; InActive
  IF ShoresFlag THEN
    PRINT #RptFile, "               Storm Water: "; USING "#####,#"; TStorm#
  END IF
  PRINT #RptFile, CHR$(12)
  IF SYLVAFlag THEN
    GOSUB DoSylva5Less
  END IF
  
  CLOSE
  
  ERASE TaxCust, RealRec, PersRec, TBillRec, BillInfo
  ERASE Fmt$, Frm, Form$, Fld, TempScrn, IdxBuff
  
  PrintRptFile "Tax Prebilling Report.", TaxPreRptFile, 1, RetCode%, 0
  'PreBill Print END   #############
  
ExitPreBill:
  
  EXIT SUB
  
DoSylva5Less:
    GOSUB DOSylva5LessHeader
    L5# = 0
    FOR Cnt& = 1 TO NumOfRec&
      GET #TBillFile, Cnt&, TBillRec(1)
      IF TBillRec(1).BillNumber = -100 THEN
        PRINT #RptFile, USING "#####"; TBillRec(1).CustRec;
        PRINT #RptFile, TAB(8); QPTrim$(TBillRec(1).CustName);
        PRINT #RptFile, TAB(63); USING "$$######,#.##"; TBillRec(1).TotalBillDue
        LineCnt = LineCnt + 1
        IF LineCnt >= MaxLines THEN
          PRINT #RptFile, CHR$(12);
          GOSUB DOSylva5LessHeader
        END IF
        L5# = Round#(L5# + TBillRec(1).TotalBillDue)
        TBillRec(1).TotalBillDue = 0
        TBillRec(1).BillNumber = -1
        BCntC = BCntC + 1
        PUT #TBillFile, Cnt&, TBillRec(1)
      END IF
    NEXT
    PRINT #RptFile, Dash$
    PRINT #RptFile, "Totals:"; TAB(23); BCntC;
    PRINT #RptFile, TAB(63); USING "$$######,#.##"; L5#
    PRINT #RptFile, CHR$(12)
RETURN

DOSylva5LessHeader:
  page = page + 1
  PRINT #RptFile, TAB(20); "Property Tax Billing : Pre-Billing Register"
  PRINT #RptFile, ""
  PRINT #RptFile, TAB(5); "Accounts @5.00 or less."
  PRINT #RptFile, "Date: "; TheDate$; TAB(65); "Page #"; page
  PRINT #RptFile, "Acct #"; TAB(8); "Customer Name"; TAB(63); "Amount"
  PRINT #RptFile, Dash$
  LineCnt = 6
RETURN

PreBillHeading:
  page = page + 1
  PRINT #RptFile, TAB(20); "Property Tax Billing : Pre-Billing Register"
  PRINT #RptFile,
  PRINT #RptFile, "Date: "; TheDate$; TAB(65); "Page #"; page
  PRINT #RptFile, "Acct #"; TAB(8); "Customer Name"
  PRINT #RptFile, "Real Value"; TAB(15); "Pers Value"; TAB(30); "Discnt Value"; TAB(45); "Net Valuation"; TAB(62); "Bill #"; TAB(73); "Tax Due"
  PRINT #RptFile, Dash$
  LineCnt = 6
RETURN
  
GetRealInfo:
  ThisProp& = TaxCust(1).FirstPropRec
  IF ThisProp& > 0 THEN
    GET #RealFile, ThisProp&, RealRec(1)
    CustMort$ = QPTrim$(RealRec(1).MORTCODE)
    IF LEN(CustMort$) > 0 THEN
      FOR MortCnt = 1 TO NumMortCodes
        ThisMort$ = QPTrim$(MortCodes(MortCnt).MORTCODE)
        IF ThisMort$ = CustMort$ THEN
          TBillRec(1).MortRec = MortCodes(MortCnt).MortRec
          EXIT FOR
        END IF
      NEXT
    END IF

    IF RealRec(1).LastYrPrinted = WhatYear THEN Discovery$ = "Y"
'here
    'IF RealRec(1).PROPDISC = "Y" THEN STOP
    IF (RealRec(1).LastYrPrinted <> WhatYear) OR (RealRec(1).PROPDISC = "Y") THEN
'020501
      IF SupOnly AND RealRec(1).PROPDISC <> "Y" THEN
        GOTO SupOnlyNGSkip
      END IF
      'IF RealRec(1).LastYrPrinted <> WhatYear THEN
      RealValue# = RealRec(1).PROPVALU
      RealExmp# = Round#(RealRec(1).EXMPSENI + RealRec(1).EXMPOTHR)
      RealCalcVal# = Round#((RealValue# - RealExmp#) / 100)
      RealTaxDue# = Round#(RealCalcVal# * REALRATE#)
      TBillRec(1).ExptValue = Round#(TBillRec(1).ExptValue + RealExmp#)
      TBillRec(1).RealValue = RealValue#
      TBillRec(1).RealTaxDue = RealTaxDue#
      TBillRec(1).RealPropRecord = ThisProp&
      TBillRec(1).RealTaxRate = REALRATE#
      TBillRec(1).RDesc1 = QPTrim$(RealRec(1).PROPNOT1) + " " + QPTrim$(RealRec(1).PROPNOT2)
      'TBillRec(1).RDesc2 = RealRec(1).LOTNUMB
      MAPBLKLOT$ = RealRec(1).MAP + " " + RealRec(1).Block + " " + RealRec(1).LotNumb

      TBillRec(1).RDesc2 = MAPBLKLOT$

      TBillRec(1).REALPIN = RealRec(1).REALPIN
      IF RealRec(1).LATELIST = "Y" THEN
        LateAmt# = Round#((RealTaxDue#) * (LATELIST# / 100))
      ELSE
        LateAmt# = 0
      END IF
      TBillRec(1).LateTaxDue = Round#(TBillRec(1).LateTaxDue + LateAmt#)
      TBillRec(1).TotalBillDue = Round#(TBillRec(1).TotalBillDue + RealTaxDue# + LateAmt#)
      '****** put the first rec to disk
      GOSUB WriteIt2Disk
    END IF      'End of Test For Current Year Tax Bill

'020501

SupOnlyNGSkip:
    NextRealRec& = RealRec(1).NextRec
    IF NextRealRec& > 0 THEN
      DO
        REDIM TBillRec(1) AS TaxBillType        'make a new empty record each time
        GOSUB SetCustInfo
        GET #RealFile, NextRealRec&, RealRec(1)
        IF RealRec(1).LastYrPrinted = WhatYear THEN Discovery$ = "Y"
'here
        IF (RealRec(1).LastYrPrinted <> WhatYear) OR (RealRec(1).PROPDISC = "Y") THEN
'020501
        IF SupOnly AND RealRec(1).PROPDISC <> "Y" THEN
          GOTO SupOnlyNGSkip2
        END IF

        'IF RealRec(1).LastYrPrinted <> WhatYear THEN
          RealValue# = RealRec(1).PROPVALU
          RealExmp# = Round#(RealRec(1).EXMPSENI + RealRec(1).EXMPOTHR)
          RealCalcVal# = Round#((RealValue# - RealExmp#) / 100)
          RealTaxDue# = Round#(RealCalcVal# * REALRATE#)
          TBillRec(1).ExptValue = RealExmp#
          TBillRec(1).RealValue = RealValue#
          TBillRec(1).RealTaxDue = RealTaxDue#
          TBillRec(1).RealPropRecord = NextRealRec&
          TBillRec(1).RealTaxRate = REALRATE#
          TBillRec(1).RDesc1 = RealRec(1).PROPNOT1
          TBillRec(1).RDesc2 = RealRec(1).PROPNOT2
          
          TBillRec(1).REALPIN = RealRec(1).REALPIN
          IF RealRec(1).LATELIST = "Y" THEN
            LateAmt# = Round#(RealTaxDue# * (LATELIST# / 100))
            TBillRec(1).LateTaxDue = LateAmt#
          ELSE
            LateAmt# = 0
          END IF
          TBillRec(1).TotalBillDue = Round#(RealTaxDue# + LateAmt#)
          GOSUB WriteIt2Disk
        END IF

'020501
SupOnlyNGSkip2:
        NextRealRec& = RealRec(1).NextRec
      LOOP WHILE NextRealRec& > 0
    END IF
  ELSE
    GOSUB WriteIt2Disk
  END IF

'testing
  IF GotPersVal AND (Done2Disk = False) THEN
    GOSUB WriteIt2Disk
  END IF
RETURN
  
GetPersInfo:
  GotPersVal = 0
  PersExmp# = 0
  PersValue# = 0

  IF TaxCust(1).FirstPersRec > 0 THEN
    GET #PersFile, TaxCust(1).FirstPersRec, PersRec(1)
    IF PersRec(1).LastYrPrinted = WhatYear THEN Discovery$ = "Y"
'here
    'IF PersRec(1).DISCOV = "Y" THEN STOP
    IF (PersRec(1).LastYrPrinted <> WhatYear) OR (PersRec(1).DISCOV = "Y") THEN
'020501
    IF SupOnly AND PersRec(1).DISCOV <> "Y" THEN
      GOTO SupOnlyNGSkip3
    END IF

    'IF PersRec(1).LastYrPrinted <> WhatYear THEN
      PersValue# = Round#(PersRec(1).PersVal + PersRec(1).MHVALUE + PersRec(1).MCVALUE)
      PersValue# = Round#(PersValue# + PersRec(1).CVALUE + PersRec(1).MTVALUE)
      PersExmp# = Round#(PersRec(1).EXMPSENI + PersRec(1).EXMPOTHR)
      PersCalcVal# = Round#((PersValue# - PersExmp#) / 100)
    ELSE
      PersCalcVal# = 0
    END IF
    IF PersCalcVal# < 0 THEN
      PersCalcVal# = 0
    END IF

    PersTaxDue# = Round#(PersCalcVal# * PERSRATE#)
    TBillRec(1).ExptValue = PersExmp#
    TBillRec(1).PersValue = PersValue#
    TBillRec(1).PersTaxDue = PersTaxDue#
    TBillRec(1).PersPropRecord = TaxCust(1).FirstPersRec
    TBillRec(1).PersTaxRate = PERSRATE#
    TBillRec(1).TotalBillDue = Round#(TBillRec(1).PersTaxDue)
    IF TBillRec(1).TotalBillDue > 0 THEN
      GotPersVal = -1
    END IF
    IF PersRec(1).LATELIST = "Y" THEN
      LateAmt# = Round#((LateAmt# + TBillRec(1).PersTaxDue) * (LATELIST# / 100))
      TBillRec(1).LateTaxDue = LateAmt#
      TBillRec(1).TotalBillDue = Round#(TBillRec(1).PersTaxDue + LateAmt#)
    END IF
  ELSE
    'no need to set the variables all are zero already
  END IF
'020501
SupOnlyNGSkip3:
RETURN

WriteIt2Disk:   'write the info out to disk here.

  TBillRec(1).BillPrinted = False
  IF TBillRec(1).TotalBillDue > 0 THEN
    IF ShoresFlag THEN
      TBillRec(1).CarShore = 60
      TStorm# = Round(TStorm# + TBillRec(1).CarShore)
      StormCnt& = StormCnt& + 1
      'TBillRec(1).TotalBillDue = Round#(TBillRec(1).TotalBillDue + TBillRec(1).CarShore)
    END IF
    TBillRec(1).BillNumber = BILLNUM&
  ELSE
    TBillRec(1).BillNumber = -1
  END IF
  IF SYLVAFlag THEN
    IF (TBillRec(1).TotalBillDue > 0) AND (TBillRec(1).TotalBillDue < 5.01#) THEN
      'L5 = L5 + 1
      'STOP
      'TBillRec(1).TotalBillDue = 0
      TBillRec(1).BillNumber = -100
    END IF
  END IF

  PUT #TBillFile, , TBillRec(1)
  IF TBillRec(1).TotalBillDue > 0 THEN
    BILLNUM& = BILLNUM& + 1
  END IF
  Done2Disk = -1

RETURN
  
SetCustInfo:
  TBillRec(1).CustRec = CustAcct&               'cust acct rec
  CustName$ = QPTrim$(TaxCust(1).FName) + " " + QPTrim$(TaxCust(1).LNAME)
  TBillRec(1).CustName = CustName$
  TBillRec(1).CustAdd1 = TaxCust(1).Addr1
  TBillRec(1).CustAdd2 = TaxCust(1).Addr2
  CitySt$ = QPTrim$(TaxCust(1).City) + " " + TaxCust(1).State
  TBillRec(1).CustAdd3 = CitySt$
  TBillRec(1).CustZip = TaxCust(1).Zip
  TBillRec(1).CustPin = TaxCust(1).PIN
  TBillRec(1).TAXYEAR = WhatYear
  TBillRec(1).RDesc3 = TaxCust(1).CSSN

  'Set Prior Balance if any
  GOSUB GetPastBalance

  IF ABalance# > 0 THEN
    IF PastFlagSet = 0 THEN
      TBillRec(1).PriorYrBalance = ABalance#
    END IF
    PastFlagSet = 1
  END IF
RETURN
  
CheckPreInfo:
  WhatYear = Value#(Form$(1, 0), ECode)
  BILLNUM& = Value#(Form$(2, 0), ECode)
  REALRATE# = Value#(Form$(3, 0), ECode)
  PERSRATE# = Value#(Form$(4, 0), ECode)
  LATELIST# = Value#(Form$(5, 0), ECode)
  Order$ = LEFT$(Form$(6, 0), 1)
  SupOnly = Form$(7, 0) = "Y"

  IF WhatYear < 1997 OR WhatYear > 2010 THEN
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
  FPutAH "TAXBINFO.DAT", BillInfo(1), BillInfoLen, 1
  
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
      Balance# = 0
      TransRecord& = TaxTrans(1).LastTrans
    LOOP
    CLOSE TransFile
  END IF

RETURN

LoadMortCodes:

  MortCodeRecLen = LEN(MortCodeRec(1))
  MortFile = FREEFILE
  OPEN "TAXMORT.DAT" FOR RANDOM AS MortFile LEN = MortCodeRecLen
  NumMortCodes = LOF(MortFile) \ MortCodeRecLen

  IF NumMortCodes > 0 THEN
    REDIM PRESERVE MortCodes(1 TO NumMortCodes) AS MortRecType
    FOR Cnt = 1 TO NumMortCodes
      GET #MortFile, Cnt, MortCodeRec(1)
      MortCodes(Cnt).MORTCODE = MortCodeRec(1).MORTCODE
      MortCodes(Cnt).MortRec = Cnt
    NEXT
  END IF
  CLOSE MortFile

RETURN


END SUB

