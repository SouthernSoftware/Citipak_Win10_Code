DECLARE SUB DisplayAPScrn (ScrnName$)
DECLARE SUB OpenAPDistFile (APDistFile%, NumDistRecs&, RecLen%)
DECLARE SUB OpenAPLedgerFile (APLedgerFile%, NumTrans%, RecLen%)
DECLARE SUB OpenVendorFile (VendorFile%, NumVRecs%)
DECLARE SUB OpenAcctFile (AcctFileNum%, NumAccts%)
DECLARE SUB WPickList (Items() AS ANY, Picked%(), NPicked%, DspRow%, UpDateFlag%, Cnf AS ANY)
  REDIM TempScrn(0)
  SaveScrn TempScrn()

  REDIM TVendor(1)       AS VendorRecType
  REDIM TAPLedgerRec(1)  AS APLedger81RecType
  REDIM TAPDistRec(1)    AS APDistRecType

  DistRecLen = LEN(TAPDistRec(1))
  LedgerRecLen = LEN(TAPLedgerRec(1))

  Title$ = SPACE$(68)
  LSET Title$ = " PO Num        TR Date       Req Num            Amount"

  'This array points to the po record number fields
  REDIM POFld(1 TO 6) AS INTEGER

  'This array hold the actual po record numbers
  REDIM OPORecs(1 TO 6) AS LONG
  FOR zz = 1 TO 6
    POFld(zz) = FldNum("POREC" + QPTrim$(STR$(zz)), Fld())
    TNum$ = Form$(POFld(zz), 0)
    IF LEN(QPTrim$(TNum$)) > 0 THEN
      GotOld = True
      OPORecs(zz) = CVL(TNum$)
    ELSE
      OPORecs(zz) = 0
    END IF
  NEXT

  PONumFld = FldNum("PONUM", Fld())

  MaxLen = 59   'Set menu width to zero
  Action = 0    '0 let vertmenu have control until they select or cancel
  BoxBot = 15
  AmtOffSet = 43
  Row = 6
  NumFlds = UBOUND(Form$)
  GotSome = False
  POCnt = 0

  VRecNum = CVI(LEFT$(Form$(FldNum("VRECNUM", Fld()), 0), 2))
  PONumFld = FldNum("PONUM", Fld())

  OpenVendorFile VendorFile, NumVRecs
  GET VendorFile, VRecNum, TVendor(1)
  CLOSE

  OpenAPLedgerFile APLedgerFile, NumLedgerRecs, LedgerRecLen

  LedgerRec& = TVendor(1).FrstTran
  DO WHILE LedgerRec& > 0
    GET APLedgerFile, LedgerRec&, TAPLedgerRec(1)
    IF TAPLedgerRec(1).TrCode = 4 THEN
      GotSome = True
      POCnt = POCnt + 1
      REDIM PRESERVE MTChoice(1 TO POCnt) AS FLen2  'Items to display
      REDIM PRESERVE Picked(1 TO POCnt)             'Selected items list
      REDIM PRESERVE PORecs(1 TO POCnt) AS LONG     'PO's rec numbers
      REDIM PRESERVE PODist(1 TO POCnt) AS LONG     'PO's first distribution
      LSET MTChoice(POCnt).V = TAPLedgerRec(1).PONUM
      MID$(MTChoice(POCnt).V, 14) = Num2Date(TAPLedgerRec(1).TrDate)
      MID$(MTChoice(POCnt).V, 28) = LEFT$(QPTrim$(TAPLedgerRec(1).DOCNum), 10)
      MID$(MTChoice(POCnt).V, AmtOffSet) = FUsing$(STR$(TAPLedgerRec(1).Amt), "#######.##")
      PORecs(POCnt) = LedgerRec&
      PODist(POCnt) = TAPLedgerRec(1).FrstDist
'This retags any previous selected po's
      IF GotOld THEN
        FOR zz = 1 TO 6
          IF OPORecs(zz) = LedgerRec& THEN
            OPCnt = OPCnt + 1
            Picked(OPCnt) = POCnt
            EXIT FOR
          END IF
        NEXT
      END IF
'***************************************
    END IF
    LedgerRec& = TAPLedgerRec(1).NextTrans
  LOOP
  CLOSE

  IF GotSome THEN
    MPaintBox 2, 1, 24, 80, 8   'to add shadow to list
    DisplayAPScrn "PICKOPTI"
    QPrintRC Title$, Row - 1, 7, 112
    WPickList MTChoice(), Picked(), PickedCnt, Row, UpDateFlag, Cnf
  ELSE
    DisplayAPScrn "NOVENDPO"
    WaitForAction
    GOTO ExitPOSelect
  END IF

  IF UpDateFlag THEN   'If the PO Selections were changed then update
    ChangeFlag = True
    TotalAmt# = 0
    TaxActualAdj = 0
    DistCnt = 0
    ActualCnt = 0

    IF PickedCnt = 0 THEN
      FOR CDCnt = AcctFlds(ActualCnt + 1) TO AcctFlds(NumAcctFlds)                'clear all of distrubtions
        LSET Form$(CDCnt, 0) = ""               'fields
        LSET Form$(CDCnt + 1, 0) = ""
        LSET Form$(CDCnt + 2, 0) = ""
        LSET Form$(CDCnt + 3, 0) = ""
        FOR SAVFld = CDCnt TO CDCnt + 3
          SaveField SAVFld, Form$(), Fld(), BadFld
        NEXT
      NEXT
      GOTO Skip2Here
    END IF
    IF GotTaxFile THEN
      IF StaTaxFlag THEN
        DistCnt = DistCnt + 1
        TaxActualAdj = TaxActualAdj + 1
        STAcctRec = FindAcct(QPTrim$(InvTaxRec(1).InvTax(1).AcctNo))
        REDIM PRESERVE DistSum(1 TO DistCnt)  AS DistSumType
        DistSum(DistCnt).DistAcctNum = InvTaxRec(1).InvTax(1).AcctNo
        DistSum(DistCnt).AcctTitle = GetAcctTitle$(STAcctRec)
      END IF

      IF CtyTaxFlag THEN
        DistCnt = DistCnt + 1
        TaxActualAdj = TaxActualAdj + 1
        CTAcctRec = FindAcct(QPTrim$(InvTaxRec(1).InvTax(2).AcctNo))
        REDIM PRESERVE DistSum(1 TO DistCnt)  AS DistSumType
        DistSum(DistCnt).DistAcctNum = InvTaxRec(1).InvTax(2).AcctNo
        DistSum(DistCnt).AcctTitle = GetAcctTitle$(CTAcctRec)
      END IF
    END IF

    OpenAcctFile AcctFileNum, NumAccts
    OpenAPDistFile APDistFile, NumDistRecs&, DistRecLen

    FOR zz = 1 TO PickedCnt
      LSET Form$(POFld(zz), 0) = MKL$(PORecs(Picked(zz))) 'store po rec number
      SaveField POFld(zz), Form$(), Fld(), BadFld
      NextDist& = PODist(zz)
      DO UNTIL NextDist& = 0
        GET APDistFile, NextDist&, TAPDistRec(1)
        DistCnt = DistCnt + 1
        GET AcctFileNum, TAPDistRec(1).DistAcctRec, Acct
        REDIM PRESERVE DistSum(1 TO DistCnt)  AS DistSumType
        DistSum(DistCnt).DistAcctNum = TAPDistRec(1).DistAcctNum
        DistSum(DistCnt).AcctTitle = Acct.Title
        DistSum(DistCnt).DistAmt = TAPDistRec(1).DistAmt
        NextDist& = TAPDistRec(1).NextDist
      LOOP
    NEXT
    CLOSE

    TInvAmt# = 0
    FOR zz = TaxActualAdj + 1 TO DistCnt
      TInvAmt# = Round#(TInvAmt# + DistSum(zz).DistAmt)
    NEXT

    IF GotTaxFile THEN
      IF StaTaxFlag THEN
        TCalc# = InvTaxRec(1).InvTax(1).TaxAmt
        IF TCalc# > 0 THEN
          StaTax# = Round#(TInvAmt# * (TCalc# / 100))
          DistSum(TState).DistAmt = StaTax#
        END IF
      END IF
      IF CtyTaxFlag THEN
        TCalc# = InvTaxRec(1).InvTax(2).TaxAmt
        IF TCalc# > 0 THEN
          CtyTax# = Round#(TInvAmt# * (TCalc# / 100))
          DistSum(TCounty).DistAmt = CtyTax#
        END IF
      END IF
    END IF

Skip2Here:
    IF PickedCnt > 1 THEN
      LSET Form$(PONumFld, 0) = "Multi"
    ELSEIF PickedCnt = 1 THEN
      LSET Form$(PONumFld, 0) = QPTrim$(LEFT$(MTChoice(Picked(1)).V, 13))
    ELSE
      LSET Form$(PONumFld, 0) = ""
    END IF

    FOR zz = PickedCnt + 1 TO 6
      LSET Form$(POFld(zz), 0) = MKL$(0)
      SaveField POFld(zz), Form$(), Fld(), BadFld
    NEXT

    LSET Form$(7, 0) = QPTrim$(STR$(TotalAmt#))
    SaveField 7, Form$(), Fld(), BadFld

    IF ActualCnt < NumAcctFlds THEN
      FOR CDCnt = AcctFlds(ActualCnt + 1) TO AcctFlds(NumAcctFlds)                'clear all of distrubtions
        LSET Form$(CDCnt, 0) = ""               'fields
        LSET Form$(CDCnt + 1, 0) = ""
        LSET Form$(CDCnt + 2, 0) = ""
        LSET Form$(CDCnt + 3, 0) = ""
        FOR SAVFld = CDCnt TO CDCnt + 3
          SaveField SAVFld, Form$(), Fld(), BadFld
        NEXT
      NEXT
    END IF
  END IF

ExitPOSelect:
  RestScrn TempScrn()
  ERASE TempScrn

