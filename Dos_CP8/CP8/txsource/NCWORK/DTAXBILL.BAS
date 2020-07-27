DEFINT A-Z
DECLARE SUB MakeNameIndex ()
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

'  MChoice$(1) = "Print Tax Pre-Billing Register"
'  MChoice$(2) = "Print Tax Bills"
'  MChoice$(3) = "Post Tax Bills"
'  MChoice$(4) = "Reprint Tax Bills"
'  MChoice$(5) = "Apply Interest"
'  MChoice$(6) = "Apply Penalty"
'  MChoice$(7) = "Exit to OS"

  MChoice$(1) = "Print Tax Pre-Billing Register"
  MChoice$(2) = "Print Tax Bills"
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
    CASE 4
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

SUB TAXPrebilling

  SHARED Choice$()
  
  REDIM TaxCust(1)  AS TaxCustType
  REDIM RealRec(1)  AS PropertyRecType
  REDIM PersRec(1)  AS PersonalRecType
  REDIM TBillRec(1) AS TaxBillType
  REDIM BillInfo(1) AS TaxBillInfoType

  TaxCustLen = LEN(TaxCust(1))  'Length of Cust Record Structure
  RealRecLen = LEN(RealRec(1))
  PersRecLen = LEN(PersRec(1))
  TBillLen = LEN(TBillRec(1))
  BillInfoLen = LEN(BillInfo(1))

  REDIM ScrnArray(0)

  REDIM Choice$(0 TO 4, 0)
  Choice$(0, 0) = "6"
  Choice$(1, 0) = "Customer Name Order"
  Choice$(2, 0) = "Account Number Order"

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
      Action = 1
    END IF

    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE -68               'F10Key    Proceed with look up
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

  IF Order$ = "C" THEN
    UsingIdx = True
    MakeNameIndex
    NumIdxRecs = FileSize&("TAXTEMP.IDX") / 2
    REDIM IdxBuff(1 TO NumIdxRecs) AS INTEGER
    FGetAH "TAXTEMP.IDX", IdxBuff(1), 2, NumIdxRecs
  END IF

  ShowProcessingScrn "Calculating Tax Prebilling"

  IF Exist%(TaxBillFile) THEN
    KILL TaxBillFile         'kill any old work file
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
    BillOK = False
    RealCnt = 0

    GET #TaxFile, CustAcct&, TaxCust(1)   'get cust record
    'IF TaxCust(1).FirstPersRec <= 0 AND TaxCust(1).FirstPropRec <= 0 THEN
    '  GOTO PreBillSkip
    'END IF
    REDIM TBillRec(1) AS TaxBillType  'make a new empty record each time
    GOSUB SetCustInfo
    GOSUB GetPersInfo
    GOSUB GetRealInfo
    ShowPctComp Cnt&, NumOfCRecs&
PreBillSkip:
  NEXT
  CLOSE
'Prebill Calc END   *******************************

'Prebill Print Start #############
  ShowProcessingScrn "Creating Tax Prebilling Report"
  RptFile = FREEFILE
  OPEN TaxPreRptFile FOR OUTPUT AS #RptFile
  
  TBillFile = FREEFILE
  OPEN TaxBillFile FOR RANDOM SHARED AS TBillFile LEN = TBillLen
  NumOfRec& = LOF(TBillFile) / TBillLen

  FOR Cnt& = 1 TO NumOfRec&
    GET #TBillFile, Cnt&, TBillRec(1)


    PRINT #RptFile, LEFT$(TBillRec(1).CustName, 30); USING "    ######.##"; TBillRec(1).RealValue; TBillRec(1).PersValue; TBillRec(1).TotalBillDue;
    PRINT #RptFile, TAB(70); TBillRec(1).BillNumber


    ShowPctComp Cnt&, NumOfRec&
  NEXT

  CLOSE

  PrintRptFile "Tax Prebilling Report.", TaxPreRptFile, 1, RetCode%, 0
'PreBill Print END   #############

ExitPreBill:
  EXIT SUB

GetRealInfo:
  ThisProp& = TaxCust(1).FirstPropRec
  IF ThisProp& > 0 THEN
    GET #RealFile, ThisProp&, RealRec(1)
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
    TBillRec(1).REALPIN = RealRec(1).REALPIN
    TBillRec(1).TotalBillDue = Round#(TBillRec(1).PersTaxDue + RealTaxDue#)
'****** put the first rec to disk
    GOSUB WriteIt2Disk
    NextRealRec& = RealRec(1).NextRec
    IF NextRealRec& > 0 THEN
      DO
        REDIM TBillRec(1) AS TaxBillType  'make a new empty record each time
        GOSUB SetCustInfo
        GET #RealFile, NextRealRec&, RealRec(1)
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
        TBillRec(1).TotalBillDue = RealTaxDue#
        GOSUB WriteIt2Disk
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
  IF TaxCust(1).FirstPersRec > 0 THEN
    GET #PersFile, TaxCust(1).FirstPersRec, PersRec(1)
    PersValue# = Round#(PersRec(1).PERSVAL + PersRec(1).MHVALUE + PersRec(1).MCVALUE)
    PersValue# = Round#(PersValue# + PersRec(1).CVALUE + PersRec(1).MTVALUE)
    PersExmp# = Round#(PersRec(1).EXMPSENI + PersRec(1).EXMPOTHR)
    PersCalcVal# = Round#((PersValue# - PersExmp#) / 100)
    PersTaxDue# = Round#(PersCalcVal# * PERSRATE#)
    TBillRec(1).ExptValue = PersExmp#
    TBillRec(1).PersValue = PersValue#
    TBillRec(1).PersTaxDue = PersTaxDue#
    TBillRec(1).PersPropRecord = TaxCust(1).FirstPersRec
    TBillRec(1).PersTaxRate = PERSRATE#
  ELSE
    'no need to set the variables all are zero already
  END IF

RETURN

WriteIt2Disk:
  TBillRec(1).BillPrinted = "N"
  TBillRec(1).BillNumber = BILLNUM&
  PUT #TBillFile, , TBillRec(1)
  BILLNUM& = BILLNUM& + 1
RETURN

SetCustInfo:
  TBillRec(1).CustRec = CustAcct&                'cust acct rec
  CustName$ = QPTrim$(TaxCust(1).FNAME) + " " + QPTrim$(TaxCust(1).LName)
  TBillRec(1).CustName = CustName$
  TBillRec(1).CustAdd1 = TaxCust(1).Addr1
  TBillRec(1).CustAdd2 = TaxCust(1).Addr2
  CitySt$ = QPTrim$(TaxCust(1).City) + " " + TaxCust(1).State
  TBillRec(1).CustAdd3 = CitySt$
  TBillRec(1).CustZip = TaxCust(1).Zip
  TBillRec(1).CustPin = TaxCust(1).PIN
  TBillRec(1).TAXYEAR = WhatYear
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
  SaveScrn ScrnArray()
  DisplayTaxScrn "ERRSCRN1"
  QPrintRC "Invalid Prebilling Information!", 10, 25, -1
  QPrintRC "Please correct and try again.", 12, 26, -1
  WaitForAction
  RestScrn ScrnArray()
RETURN

END SUB

