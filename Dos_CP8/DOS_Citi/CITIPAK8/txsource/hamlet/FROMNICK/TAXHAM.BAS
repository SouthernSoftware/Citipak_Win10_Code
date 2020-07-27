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
  REDIM MChoice$(1 TO 7)
  
  MChoice$(1) = "Print Tax Pre-Billing Register"
  MChoice$(2) = "Print/Reprint Tax Bills "
  MChoice$(3) = "Post Tax Bills"
  MChoice$(4) = "Apply Interest"
  MChoice$(5) = "Print Late Notices"
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

  TYPE Struct
    who AS STRING * 14
    RecNum AS INTEGER
  END TYPE

  
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
      TaxTrans(1).BillType = "C"                'R=Real P=Personal Property C=Combined (NC/GA)
      TaxTrans(1).Amount = TaxBill(1).TotalBillDue              'Total Transaction Amount
      TaxTrans(1).Revenue.Principle1 = TaxBill(1).TotalBillDue
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
        PUT RealFile, TaxBill(1).RealPropRecord, RealRec(1)
      END IF
      
      IF TaxBill(1).PersPropRecord > 0 THEN
        GET PersFile, TaxBill(1).PersPropRecord, PersRec(1)
        PersRec(1).LastYrPrinted = TaxBill(1).TAXYEAR
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
  DIM MCode$(250), MDesc$(250)
  DIM TotalReal#(1), TotalPers#(1), TotalEx#(1), TotalBills#(1), TotalLate#(1)
  
  REDIM MortCodeRec(1) AS MortCodeRecType
  MortCodeRecLen = LEN(MortCodeRec(1))
  REDIM Array(1 TO 1) AS Struct 'Template for the sort array

  
  REDIM TaxCust(1)  AS TaxCustType
  REDIM TaxTrans(1) AS TaxTransactionType
  REDIM RealRec(1)  AS PropertyRecType
  REDIM PersRec(1)  AS PersonalRecType
  REDIM TBillRec(1) AS TaxBillType
  REDIM BillInfo(1) AS TaxBillInfoType
  
  REDIM Fmt$(1 TO 4)
  
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
  Fmt$(4) = "#########,#"
  
  MaxLines = 56
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


  REM Loop #1 for All Non Mortgate Code Customers

  MasterLoop = 1  'Print All TownShips (No Corps)

  FOR Cnt& = 1 TO NumOfCRecs&
    IF UsingIdx THEN
      CustAcct& = IdxBuff(Cnt&)
    ELSE
      CustAcct& = Cnt&
    END IF

    REDIM TBillRec(1) AS TaxBillType            'make a new empty record each time
    GET #TaxFile, CustAcct&, TaxCust(1)         'get cust record
    
    IF TaxCust(1).Deleted <> 0 THEN
      GOTO PreBillSkip:
    END IF
    
    PastFlagSet = 0             'Initialize Past Balance Flag
    
    IF TaxCust(1).FirstPersRec <= 0 AND TaxCust(1).FirstPropRec <= 0 THEN
      NoProp = 1
      GOSUB SetCustInfo
      GOSUB WriteIt2Disk
      GOTO PreBillSkip
    END IF

    'Fix MortCodes
    
    ThisProp& = TaxCust(1).FirstPropRec
    WHILE ThisProp& <> 0
    GET #RealFile, ThisProp&, RealRec(1)
    IF QPTrim$(RealRec(1).MORTCODE) = "8" THEN
     RealRec(1).MORTCODE = "08"
     PUT #RealFile, ThisProp&, RealRec(1)
    END IF
    IF QPTrim$(RealRec(1).MORTCODE) = "1" THEN
     RealRec(1).MORTCODE = "01"
     PUT #RealFile, ThisProp&, RealRec(1)
    END IF
    IF QPTrim$(RealRec(1).MORTCODE) = "2" THEN
     RealRec(1).MORTCODE = "02"
     PUT #RealFile, ThisProp&, RealRec(1)
    END IF
    ThisProp& = RealRec(1).NextRec
    WEND

  
    'CHECK TOWNSHIP CODE FOR CORP
    KK = INSTR(UCASE$(TaxCust(1).TownShip), "CORP")
    IF KK > 0 THEN
     Corp = 1
     ELSE
     Corp = 0
    END IF
   
    IF Corp = 1 THEN
     GOTO PreBillSkip
    END IF
    
    NoProp = 0
    GOSUB SetCustInfo
    GOSUB GetPersInfo
    GOSUB GetRealInfo
PreBillSkip:
    ShowPctComp Cnt&, NumOfCRecs&

  NEXT Cnt&

  ShowProcessingScrn "Calculating Tax Prebilling"

 'Loop #2 for All  Mortgate Code Customers
  MasterLoop = 2  'Print All Mort Code Customers NExt
  

 'Must Loop For Each MortCode Here

   FOR Cnt& = 1 TO NumOfCRecs&

    IF UsingIdx THEN
      CustAcct& = IdxBuff(Cnt&)
    ELSE
      CustAcct& = Cnt&
    END IF

    REDIM TBillRec(1) AS TaxBillType            'make a new empty record each time
    GET #TaxFile, CustAcct&, TaxCust(1)         'get cust record
    
    'CHECK TOWNSHIP CODE FOR CORP
    KK = INSTR(UCASE$(TaxCust(1).TownShip), "CORP")
    IF KK > 0 THEN
     Corp = 1
     ELSE
     Corp = 0
    END IF

    IF Corp = 0 THEN
     GOTO PreBillSkip2
    END IF

  
    PastFlagSet = 0             'Initialize Past Balance Flag

    IF TaxCust(1).FirstPersRec <= 0 AND TaxCust(1).FirstPropRec <= 0 THEN
      NoProp = 1
      GOSUB SetCustInfo
      GOSUB WriteIt2Disk
      GOTO PreBillSkip2
    END IF

    NoProp = 0
    GOSUB SetCustInfo
    GOSUB GetPersInfo
    GOSUB GetRealInfo
PreBillSkip2:
   ShowPctComp Cnt&, NumOfCRecs&
  NEXT Cnt&

  


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
  

    'CHECK TOWNSHIP CODE FOR CORP
    KK = INSTR(UCASE$(TBillRec(1).TownShip), "CORP")
    IF KK > 0 THEN
     Corp = 1
     ELSE
     Corp = 0
    END IF



    IF Discovery$ = "Y" AND TBillRec(1).BillNumber = -1 THEN GOTO SkipEm
    IF TBillRec(1).BillNumber = -1 AND TBillRec(1).RealValue = 0 AND TBillRec(1).PersValue = 0 THEN GOTO SkipEm


    'Check For End of TownShips
      IF Corp = 1 AND PageTotal = 0 THEN
        PRINT #RptFile, ""
       PRINT #RptFile, "Township Totals: ";
       PRINT #RptFile, TAB(46); USING Fmt$(4); TotalReal#(0);
       PRINT #RptFile, TAB(71); USING Fmt$(4); TotalPers#(0) + TotalReal#(0);
       PRINT #RptFile, TAB(89); USING "######,#.##"; TotalBills#(0) - TotalLate#(0);
       PRINT #RptFile, TAB(103); USING "#####,#.##"; TotalLate#(0);
       PRINT #RptFile, TAB(119); USING "$$######,#.##"; TotalBills#(0)

       PRINT #RptFile, TAB(59); USING Fmt$(3); TotalPers#(0);
       PRINT #RptFile, TAB(81); USING "#######,#"; TotalEx#(0)
       PRINT #RptFile, CHR$(12);
       GOSUB PreBillHeading
       PageTotal = 1
      END IF

       PRINT #RptFile, LEFT$(TBillRec(1).CustName, 33);
       IF TBillRec(1).BillNumber = -1 THEN
         PRINT #RptFile, TAB(36); "N/A";
       ELSE
         PRINT #RptFile, TAB(36); USING Fmt$(2); TBillRec(1).BillNumber;
       END IF
       PRINT #RptFile, TAB(47); USING Fmt$(3); TBillRec(1).RealValue;
       PRINT #RptFile, TAB(59); USING Fmt$(3); TBillRec(1).PersValue;
       PRINT #RptFile, TAB(70); USING Fmt$(3); TBillRec(1).PersValue + TBillRec(1).RealValue;
       PRINT #RptFile, TAB(82); USING "######,#"; TBillRec(1).ExptValue;
       PRINT #RptFile, TAB(90); USING "#####,#.##"; TBillRec(1).TotalBillDue - TBillRec(1).LateTaxDue;
       PRINT #RptFile, TAB(104); USING "#####.##"; TBillRec(1).LateTaxDue;
       PRINT #RptFile, TAB(119); USING "$$#####,#.##"; TBillRec(1).TotalBillDue
      LineCnt = LineCnt + 1
      PRINT #RptFile, TBillRec(1).CustAdd1; TAB(36); "Acct # "; TBillRec(1).CustRec
      LineCnt = LineCnt + 1
      IF LEN(QPTrim$(TBillRec(1).CustAdd2)) > 0 THEN
        PRINT #RptFile, TBillRec(1).CustAdd2
        LineCnt = LineCnt + 1
      END IF
      PRINT #RptFile, RTRIM$(TBillRec(1).CustAdd3); "   "; TBillRec(1).CustZip
      LineCnt = LineCnt + 1
      PRINT #RptFile, TBillRec(1).RealPin; TAB(36); RTRIM$(TBillRec(1).RDesc1); " "; RTRIM$(TBillRec(1).RDesc2)
      LineCnt = LineCnt + 1
      PRINT #RptFile, ""
      LineCnt = LineCnt + 1
     
      TotalPers# = Round#(TotalPers# + TBillRec(1).PersValue)
      TotalReal# = Round#(TotalReal# + TBillRec(1).RealValue)
      TotalEx# = Round#(TotalEx# + TBillRec(1).ExptValue)
      TotalBills# = Round#(TotalBills# + TBillRec(1).TotalBillDue)
      TotalLate# = Round#(TotalLate# + TBillRec(1).LateTaxDue)
     
      TotalPers#(Corp) = Round#(TotalPers#(Corp) + TBillRec(1).PersValue)
      TotalReal#(Corp) = Round#(TotalReal#(Corp) + TBillRec(1).RealValue)
      TotalEx#(Corp) = Round#(TotalEx#(Corp) + TBillRec(1).ExptValue)
      TotalBills#(Corp) = Round#(TotalBills#(Corp) + TBillRec(1).TotalBillDue)
      TotalLate#(Corp) = Round#(TotalLate#(Corp) + TBillRec(1).LateTaxDue)
     

      IF TBillRec(1).TotalBillDue > 0 THEN
        NumBills& = NumBills& + 1
      END IF
      
      IF LineCnt >= MaxLines THEN
        PRINT #RptFile, CHR$(12);
        GOSUB PreBillHeading
      END IF
    
SkipEm:
    ShowPctComp Cnt&, NumOfRec&
  NEXT
       PRINT #RptFile, ""
       PRINT #RptFile, "Corp Totals: ";
       PRINT #RptFile, TAB(46); USING Fmt$(4); TotalReal#(1);
       PRINT #RptFile, TAB(59); USING Fmt$(3); TotalPers#(1);
       PRINT #RptFile, TAB(71); USING Fmt$(4); TotalPers#(1) + TotalReal#(1);
       PRINT #RptFile, TAB(82); USING "######,#"; TotalEx#(1);
       PRINT #RptFile, TAB(90); USING "######,#.##"; TotalBills#(1) - TotalLate#(1);
       PRINT #RptFile, TAB(104); USING "####,#.##"; TotalLate#(1);
       PRINT #RptFile, TAB(119); USING "$$######,#.##"; TotalBills#(1)
       PRINT #RptFile, ""
       PRINT #RptFile, ""
       PRINT #RptFile, "BOOK TOTALS: ";
       PRINT #RptFile, TAB(46); USING Fmt$(4); TotalReal#;
       PRINT #RptFile, TAB(71); USING Fmt$(4); TotalPers# + TotalReal#;
       PRINT #RptFile, TAB(89); USING "#######,#.##"; TotalBills# - TotalLate#;
       PRINT #RptFile, TAB(104); USING "####,#.##"; TotalLate#;
       PRINT #RptFile, TAB(119); USING "$$######,#.##"; TotalBills#

       PRINT #RptFile, TAB(58); USING Fmt$(4); TotalPers#;
       PRINT #RptFile, TAB(81); USING "#######,#"; TotalEx#

      'End of printing totals next **********************************************************
       PRINT #RptFile, CHR$(12);
       CLOSE
       ERASE TaxCust, RealRec, PersRec, TBillRec, BillInfo
       ERASE Fmt$, Frm, Form$, Fld, TempScrn, IdxBuff
       PrintRptFile "Tax Prebilling Report.", TaxPreRptFile, 1, RetCode%, 0
      'PreBill Print END   #############

ExitPreBill:
       EXIT SUB
  
PreBillHeading:
  FOR LL = 1 TO 7
    PRINT #RptFile, ""
  NEXT LL
  'Start Printing at Line # 8
  PRINT #RptFile, TAB(112); USING "#.###"; REALRATE#
  PRINT #RptFile, TAB(112); USING "####"; WhatYear
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  Page = Page + 1
  LineCnt = 11
  RETURN
  
GetRealInfo:
  ThisProp& = TaxCust(1).FirstPropRec
  IF ThisProp& > 0 THEN
    GET #RealFile, ThisProp&, RealRec(1)
    IF RealRec(1).LastYrPrinted = WhatYear THEN Discovery$ = "Y"
    IF RealRec(1).LastYrPrinted <> WhatYear THEN
      
      RealValue# = RealRec(1).PROPVALU
      RealExmp# = Round#(RealRec(1).EXMPSENI + RealRec(1).EXMPOTHR)
      RealCalcVal# = Round#((RealValue# - RealExmp#) / 100)
      RealTaxDue# = Round#(RealCalcVal# * REALRATE#)
      TBillRec(1).ExptValue = Round#(TBillRec(1).ExptValue + RealExmp#)
      TBillRec(1).RealValue = RealValue#
      TBillRec(1).RealTaxDue = RealTaxDue#
      TBillRec(1).RealPropRecord = ThisProp&
      TBillRec(1).RealTaxRate = REALRATE#
      TBillRec(1).RDesc1 = RealRec(1).PROPNOT1
      TBillRec(1).RDesc2 = RealRec(1).PROPNOT2
      TBillRec(1).MORTCODE = RealRec(1).MORTCODE
      TBillRec(1).LotOrAcre = RealRec(1).LOTACRE
      TBillRec(1).LASize = STR$(RealRec(1).PropSize)
      TBillRec(1).RealPin = RealRec(1).RealPin
  '    IF RealRec(1).LATELIST = "Y" THEN
  '      LateAmt# = Round#((RealTaxDue#) * (LATELIST# / 100))
   '   ELSE
        LateAmt# = 0
    '  END IF
      TBillRec(1).LateTaxDue = Round#(TBillRec(1).LateTaxDue + LateAmt#)
      TBillRec(1).TotalBillDue = Round#(TBillRec(1).TotalBillDue + RealTaxDue# + LateAmt#)
      
      '****** put the first rec to disk
      GOSUB WriteIt2Disk
      
    END IF      'End of Test For Current Year Tax Bill
    
    NextRealRec& = RealRec(1).NextRec
    IF NextRealRec& > 0 THEN
      DO
        REDIM TBillRec(1) AS TaxBillType        'make a new empty record each time
        GOSUB SetCustInfo
        GET #RealFile, NextRealRec&, RealRec(1)
        IF RealRec(1).LastYrPrinted = WhatYear THEN Discovery$ = "Y"
        IF RealRec(1).LastYrPrinted <> WhatYear THEN
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
          TBillRec(1).RealPin = RealRec(1).RealPin
'          IF RealRec(1).LATELIST = "Y" THEN
'            LateAmt# = Round#(RealTaxDue# * (LATELIST# / 100))
'            TBillRec(1).LateTaxDue = LateAmt#
 '         ELSE
            LateAmt# = 0
  '        END IF
          
          TBillRec(1).TotalBillDue = Round#(RealTaxDue# + LateAmt#)
          
          GOSUB WriteIt2Disk
        END IF
        NextRealRec& = RealRec(1).NextRec
      LOOP WHILE NextRealRec& > 0
    END IF
  ELSE
    GOSUB WriteIt2Disk
  END IF
  
RETURN
  
GetPersInfo:
  PersExmp# = 0
  PersValue# = 0
  LateAmt# = 0
  IF TaxCust(1).FirstPersRec > 0 THEN
    GET #PersFile, TaxCust(1).FirstPersRec, PersRec(1)
    IF PersRec(1).LastYrPrinted = WhatYear THEN Discovery$ = "Y"
    IF PersRec(1).LastYrPrinted <> WhatYear THEN
      PersValue# = Round#(PersRec(1).PERSVAL + PersRec(1).MHVALUE + PersRec(1).MCVALUE)
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
    IF PersRec(1).LATELIST = "Y" THEN
      LateAmt# = Round#(LateAmt# + TBillRec(1).PersTaxDue * (LATELIST# / 100))
      TBillRec(1).LateTaxDue = LateAmt#
      TBillRec(1).TotalBillDue = Round#(TBillRec(1).PersTaxDue + LateAmt#)
    END IF
  ELSE
    'no need to set the variables all are zero already
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
  CustName$ = QPTrim$(TaxCust(1).FNAME) + " " + QPTrim$(TaxCust(1).LName)
  TBillRec(1).CustName = CustName$
  TBillRec(1).CustAdd1 = TaxCust(1).Addr1
  TBillRec(1).CustAdd2 = TaxCust(1).Addr2
  CitySt$ = QPTrim$(TaxCust(1).City) + " " + TaxCust(1).STATE
  TBillRec(1).CustAdd3 = CitySt$
  TBillRec(1).CustZip = TaxCust(1).Zip
  TBillRec(1).CustPin = TaxCust(1).PIN
  TBillRec(1).TAXYEAR = WhatYear
  TBillRec(1).TownShip = TaxCust(1).TownShip
  
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

END SUB

