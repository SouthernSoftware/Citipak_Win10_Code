DEFINT A-Z
DECLARE SUB ClearBox ()
DECLARE SUB TRGetAcctStruct (GLFundLen%, GLAcctLen%, GLDetLen%)
DECLARE SUB LoadUBSetUpFile (UBSetUpFileNum%, UBSetUpLen%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION QPTrim$ (X$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION FindAcct% (AcctNum$)
DECLARE FUNCTION GetAcctTitle$ (AcctRecNum)
DECLARE FUNCTION QPStrI$ (number%)
DECLARE FUNCTION QPStrL$ (number&)
DECLARE FUNCTION Round# (N#)
DECLARE SUB QPrintRC (T$, R%, c%, Clr%)
DECLARE SUB LibLoadDisplayForm (LibName$, FormName$, Form$(), Fld() AS ANY)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPORT%, RetCode%, EntryPoint%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
  
  '$INCLUDE: 'DefCnf.bi'                    'monitor and color info
  '$INCLUDE: 'FieldInf.bi'                  'field information
  '$INCLUDE: 'FormEdit.bi'                  'form editing information
  '$INCLUDE: 'QSCR.BI'                      'QuickScreen Routines
  '$INCLUDE: 'cmfiles.bi'
  '$INCLUDE: 'GL.BI'                        'GL File Types
  '$INCLUDE: 'ubtrans.bi'
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'TAXGLSET.BI'                  'Tax GL Layout
  '$INCLUDE: 'TaxCust.BI'                   'Tax Trans Layout
  '$INCLUDE: 'AR.bi'                        'A/R FILE LAYOUTS
  '$INCLUDE: 'DC.bi'                        'Decal File Layouts
  
  TYPE TranRecInfoType
    TranDate  AS INTEGER
    TranRecNo AS LONG
  END TYPE

  CONST False = 0, True = NOT False

  DIM SHARED GJRec(1) AS TrEditRecType

SUB ClearBox
  FOR SLine = 11 TO 14
    LOCATE SLine, 20
    PRINT STRING$(40, 32)
  NEXT
  LOCATE 16, 20
  PRINT STRING$(41, 32)

END SUB

SUB ExtractTX (ThruDate%)

  Ref$ = "TX" + LEFT$(DATE$, 2) + MID$(DATE$, 4, 2) + RIGHT$(DATE$, 2)
  
  DIM BillPrinAmt#(31), PayPrinAmt#(31), AdjPrinAmt#(31), BillIntAmt#(31), PayIntAmt#(31), AdjIntAmt#(31), BillAdvAmt#(31), PayAdvAmt#(31), AdjAdvAmt#(31)
  
  REDIM PRESERVE GJInfo(1 TO 3) AS GJXferRecType
  
  REDIM TaxGLAccts(1) AS TaxAcctsType
  TaxAcctRecLen = LEN(TaxGLAccts(1))
  
  REDIM TaxTrans(1) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTrans(1))
  
  REDIM TaxSetUpRec(1) AS TaxMasterType
  TxFile = FREEFILE
  OPEN "TAXSETUP.DAT" FOR RANDOM AS #TxFile LEN = LEN(TaxSetUpRec(1))
  IF LOF(TxFile) > 0 THEN
    GET TxFile, 1, TaxSetUpRec(1)
  ELSE
    Ok = MsgBox%("UBSENDGL", "NOMET1")
    GOTO TaxEnd
  END IF
  
  
  AcctMeth$ = QPTrim$(TaxSetUpRec(1).AcctgMethod)
  IF (LEN(AcctMeth$) = 0) THEN
    Ok = MsgBox%("UBSENDGL", "NOMETH4")
    GOTO TaxEnd
  END IF
  
  SELECT CASE AcctMeth$
  CASE "C"
    InterfaceMethod = 1
  CASE "A"
    InterfaceMethod = 2
  CASE ELSE
    Ok = MsgBox%("UBSENDGL", "NOMETH4")
    GOTO EndTax
  END SELECT
  CLOSE TxFile
  
  TRGetAcctStruct FundLen%, AcctLen%, DetLen%
  GJRecLen = LEN(GJRec(1))
  
  'Check Existing Journal to Make Sure It Is Empty
  
  REDIM GJRec1(1 TO 2) AS TrEditRecType
  GJRecLen = LEN(GJRec1(1))
  GJFile = FREEFILE
  OPEN "GLTRXED.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS GJFile LEN = GJRecLen
  NumEdTrans = LOF(GJFile) \ GJRecLen
  CLOSE
  
  IF NumEdTrans > 0 THEN
    Ok = MsgBox%("UBSENDGL", "JEDNZERO")
    GOTO EndTax
  END IF
  
  GOSUB LoadTaxPayAccts
  IF NOT (OK2Load) THEN
    Ok = MsgBox%("UBSENDGL", "TXADACTS")
    GOTO EndTax
  END IF
  
  Ok = MsgBox%("UBSENDGL", "MUSTEXT4")
  SELECT CASE Ok
  CASE 2
    GOTO EndTax
  END SELECT
  
  'Now Process the Transactions
  ClearBox

  QPrintRC "Searching Cash Transactions.", 11, 26, 126
  QPrintRC "New Transactions:", 13, 29, Cnf.HiLite
  
  TranFile = FREEFILE
  OPEN "TAXTRANS.DAT" FOR RANDOM SHARED AS TranFile LEN = TaxTranRecLen
  NumOfTRecs& = LOF(TranFile) \ TaxTranRecLen
  FOR TCnt& = NumOfTRecs& TO 1 STEP -1
    GET #TranFile, TCnt&, TaxTrans(1)
    IF LEN(QPTrim$(TaxTrans(1).Posted2GL)) = 0 OR QPTrim$(TaxTrans(1).Posted2GL) = "N" THEN
      'Store trans rec numbers and dates in array
      IF TaxTrans(1).TransDate <= ThruDate% THEN
        FoundCnt = FoundCnt + 1
        REDIM PRESERVE TranInfo(FoundCnt) AS TranRecInfoType
        TranInfo(FoundCnt).TranDate = TaxTrans(1).TransDate
        TranInfo(FoundCnt).TranRecNo = TCnt&
      END IF
    ELSE
      NGCnt = NGCnt + 1
    END IF
    P2S$ = STR$(FoundCnt)
    QPrintRC P2S$, 13, 47, Cnf.HiLite
    'SmallPause
    IF NGCnt >= 2500 THEN EXIT FOR
  NEXT
  
  IF FoundCnt = 0 THEN
    CLOSE
    Ok = MsgBox%("UBSENDGL", "NONEFOND")
    GOTO EndTax
  END IF
  
  SortT TranInfo(1), FoundCnt, 0, 6, 0, -1      'sort'em by date. oldest first
  
  GJFile = FREEFILE
  OPEN "GLTRXED.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS GJFile LEN = GJRecLen
  
  FirstTran = 1
  ThisDate = TranInfo(1).TranDate
  FOR Cnt = 1 TO FoundCnt
    IF ThisDate <> TranInfo(Cnt).TranDate THEN
      ThisDate = TranInfo(Cnt).TranDate
      LastTran = Cnt - 1
      GOSUB ProcessThisBunchTX
      FirstTran = Cnt
    END IF
  NEXT Cnt
  LastTran = FoundCnt
  GOSUB ProcessThisBunchTX
  
  'transactions as interfaced
  
  FOR Cnt = 1 TO FoundCnt
    GET #TranFile, TranInfo(Cnt).TranRecNo, TaxTrans(1)
    TaxTrans(1).Posted2GL = "Y"
    PUT #TranFile, TranInfo(Cnt).TranRecNo, TaxTrans(1)
  NEXT Cnt
  CLOSE
  SLEEP 2
  GOTO EndTax

LoadTaxPayAccts:
  IF FileSize&("TAXGLACT.DAT") > 0 THEN
    OK2Load = True
    TXGLFile = FREEFILE
    OPEN "TAXGLACT.DAT" FOR RANDOM SHARED AS TXGLFile LEN = TaxAcctRecLen
    GET TXGLFile, 1, TaxGLAccts(1)
  END IF
  CLOSE
RETURN
  
LoadTaxBillAccts:
  IF FileSize&("TAXGLBAC.DAT") > 0 THEN
    OK2Load = True
    TXGLFile = FREEFILE
    OPEN "TAXGLACT.DAT" FOR RANDOM SHARED AS TXGLFile LEN = TaxAcctRecLen
    GET TXGLFile, 1, TaxGLAccts(1)
  END IF
  CLOSE
RETURN
  
EndTax:
  EXIT SUB
  
ProcessThisBunchTX:      'Initialize for This Set
  FOR RevCnt = 1 TO 31
    BillPrinAmt#(RevCnt) = 0
    PayPrinAmt#(RevCnt) = 0
    AdjPrinAmt#(RevCnt) = 0
    BillIntAmt#(RevCnt) = 0
    PayIntAmt#(RevCnt) = 0
    AdjIntAmt#(RevCnt) = 0
    BillAdvAmt#(RevCnt) = 0
    PayAdvAmt#(RevCnt) = 0
    AdjAdvAmt#(RevCnt) = 0
  NEXT
  
  FOR PCnt = FirstTran TO LastTran
    IF PCnt = FirstTran THEN
      WorkDate = TranInfo(PCnt).TranDate
    END IF
    GET #TranFile, TranInfo(PCnt).TranRecNo, TaxTrans(1)
    'Now Decipher by Type and Year
    IF TaxTrans(1).TranType = 1 THEN
      'Tax Bill
      IF AcctMeth$ = "A" THEN
        TaxYear = TaxTrans(1).TaxYear
        TaxYear = TaxYear - 1979                'Reduce Based on 1980 being = 1
        BillPrinAmt#(TaxYear) = BillPrinAmt#(TaxYear) + TaxTrans(1).Amount
      END IF
    END IF
    
    IF TaxTrans(1).TranType = 2 THEN
      'TaxPayment
      GET #TranFile, TaxTrans(1).BelongTo, TaxTrans(1)
      TaxYear = TaxTrans(1).TaxYear
      TaxYear = TaxYear - 1979  'Reduce Based on 1980 being = 1
      IF TaxYear < 0 THEN
        TaxYear = 0
      END IF
      GET #TranFile, TranInfo(PCnt).TranRecNo, TaxTrans(1)
      PayPrinAmt#(TaxYear) = PayPrinAmt#(TaxYear) + TaxTrans(1).Revenue.Principle1Pd
      PayIntAmt#(TaxYear) = PayIntAmt#(TaxYear) + TaxTrans(1).Revenue.InterestPd
      PayAdvAmt#(TaxYear) = PayAdvAmt#(TaxYear) + TaxTrans(1).Revenue.CollectionPd
    END IF
    
    IF TaxTrans(1).TranType = 3 THEN
      'Do Nothing : Not Used
    END IF
    
    IF TaxTrans(1).TranType = 4 THEN
      'Interest
      IF AcctMeth$ = "A" THEN
        GET #TranFile, TaxTrans(1).BelongTo, TaxTrans(1)
        TaxYear = TaxTrans(1).TaxYear
        TaxYear = TaxYear - 1979                'Reduce Based on 1980 being = 1
        IF TaxYear < 0 THEN
          TaxYear = 0
        END IF
        GET #TranFile, TranInfo(PCnt).TranRecNo, TaxTrans(1)
        BillIntAmt#(TaxYear) = BillIntAmt#(TaxYear) + TaxTrans(1).Amount
      END IF
    END IF
    
    IF TaxTrans(1).TranType = 5 THEN
      'Do Nothing : Not Used for NC Used for Penalty in Virginia
    END IF
    
    IF TaxTrans(1).TranType = 6 THEN
      'Collection Cost
      IF AcctMeth$ = "A" THEN
        GET #TranFile, TaxTrans(1).BelongTo, TaxTrans(1)
        TaxYear = TaxTrans(1).TaxYear
        TaxYear = TaxYear - 1979                'Reduce Based on 1980 being = 1
        GET #TranFile, TranInfo(PCnt).TranRecNo, TaxTrans(1)
        BillAdvAmt#(TaxYear) = BillAdvAmt#(TaxYear) + TaxTrans(1).Amount
      END IF
    END IF
    
    IF TaxTrans(1).TranType = 7 THEN
      'TaxAdj
      GET #TranFile, TaxTrans(1).BelongTo, TaxTrans(1)
      TaxYear = TaxTrans(1).TaxYear
      TaxYear = TaxYear - 1979  'Reduce Based on 1980 being = 1
      GET #TranFile, TranInfo(PCnt).TranRecNo, TaxTrans(1)
      AdjPrinAmt#(TaxYear) = AdjPrinAmt#(TaxYear) + TaxTrans(1).Revenue.Principle1Pd
      AdjIntAmt#(TaxYear) = AdjIntAmt#(TaxYear) + TaxTrans(1).Revenue.InterestPd
      AdjAdvAmt#(TaxYear) = AdjAdvAmt#(TaxYear) + TaxTrans(1).Revenue.CollectionPd
    END IF
    
    IF TaxTrans(1).TranType = 8 THEN
      'Misc Charge (Adv Coll)
      IF AcctMeth$ = "A" THEN
        GET #TranFile, TaxTrans(1).BelongTo, TaxTrans(1)
        TaxYear = TaxTrans(1).TaxYear
        TaxYear = TaxYear - 1979                'Reduce Based on 1980 being = 1
        GET #TranFile, TranInfo(PCnt).TranRecNo, TaxTrans(1)
        BillAdvAmt#(TaxYear) = BillAdvAmt#(TaxYear) + TaxTrans(1).Amount
      END IF
    END IF
  NEXT PCnt
  
  'Now Post and Return for Next Day
  'TranType1 Billing
  FOR T% = 1 TO 31
    IF BillPrinAmt#(T%) > 0 THEN
      TXGLFile = FREEFILE
      OPEN "TAXGLBAC.DAT" FOR RANDOM SHARED AS TXGLFile LEN = TaxAcctRecLen
      GET TXGLFile, 1, TaxGLAccts(1)
      IF LEN(QPTrim$(TaxGLAccts(1).TaxAcct(T%).TaxDBAcct)) <= 0 THEN
        GOTO NoGoodSkip1
      END IF
      ThisAcct = FindAcct(TaxGLAccts(1).TaxAcct(T%).TaxDBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).TaxDBAcct
      DebitAcctName$ = GetAcctTitle$(ThisAcct)
      DebitAmt# = BillPrinAmt#(T%)
      ThisAcct = FindAcct(TaxGLAccts(1).TaxAcct(T%).TaxCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).TaxCRAcct
      CreditAcctName$ = GetAcctTitle$(ThisAcct)
      CreditAmt# = BillPrinAmt#(T%)
      GOSUB PostToGeneralJournal
NoGoodSkip1:
      CLOSE TXGLFile
    END IF

  NEXT T%
  
  'TranType4 Interest
  FOR T% = 1 TO 31
    IF BillIntAmt#(T%) > 0 THEN
      TXGLFile = FREEFILE
      OPEN "TAXGLBAC.DAT" FOR RANDOM SHARED AS TXGLFile LEN = TaxAcctRecLen
      GET TXGLFile, 1, TaxGLAccts(1)
      IF LEN(QPTrim$(TaxGLAccts(1).TaxAcct(T%).IntDBAcct)) <= 0 THEN
        GOTO NoGoodSkip2
      END IF

      ThisAcct = FindAcct(TaxGLAccts(1).TaxAcct(T%).IntDBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).IntDBAcct
      DebitAcctName$ = GetAcctTitle$(ThisAcct)
      DebitAmt# = BillIntAmt#(T%)
      ThisAcct = FindAcct(TaxGLAccts(1).TaxAcct(T%).IntCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).IntCRAcct
      CreditAcctName$ = GetAcctTitle$(ThisAcct)
      CreditAmt# = BillIntAmt#(T%)
      GOSUB PostToGeneralJournal
NoGoodSkip2:
      CLOSE TXGLFile
    END IF

  NEXT T%
  'TranType6 or 8
  FOR T% = 1 TO 31
    IF BillAdvAmt#(T%) > 0 THEN
      TXGLFile = FREEFILE
      OPEN "TAXGLBAC.DAT" FOR RANDOM SHARED AS TXGLFile LEN = TaxAcctRecLen
      GET TXGLFile, 1, TaxGLAccts(1)
      ThisAcct = FindAcct(TaxGLAccts(1).TaxAcct(T%).TaxDBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).TaxDBAcct
      DebitAcctName$ = GetAcctTitle$(ThisAcct)
      DebitAmt# = BillAdvAmt#(T%)
      ThisAcct = FindAcct(TaxGLAccts(1).TaxAcct(T%).TaxCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).TaxCRAcct
      CreditAcctName$ = GetAcctTitle$(ThisAcct)
      CreditAmt# = BillAdvAmt#(T%)
      GOSUB PostToGeneralJournal
      CLOSE TXGLFile
    END IF
  NEXT T%
  
  'TranType2 Payments
  FOR T% = 1 TO 31
    IF PayPrinAmt#(T%) > 0 THEN
      TXGLFile = FREEFILE
      OPEN "TAXGLACT.DAT" FOR RANDOM SHARED AS TXGLFile LEN = TaxAcctRecLen
      GET TXGLFile, 1, TaxGLAccts(1)
      ThisAcct = FindAcct(TaxGLAccts(1).TaxAcct(T%).TaxDBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).TaxDBAcct
      DebitAcctName$ = GetAcctTitle$(ThisAcct)
      DebitAmt# = PayPrinAmt#(T%)
      ThisAcct = FindAcct(TaxGLAccts(1).TaxAcct(T%).TaxCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).TaxCRAcct
      CreditAcctName$ = GetAcctTitle$(ThisAcct)
      CreditAmt# = PayPrinAmt#(T%)
      GOSUB PostToGeneralJournal
      CLOSE TXGLFile
    END IF
  NEXT T%
  
  'TranType2 Payments (Interest)
  FOR T% = 1 TO 31
    IF PayIntAmt#(T%) > 0 THEN
      TXGLFile = FREEFILE
      OPEN "TAXGLACT.DAT" FOR RANDOM SHARED AS TXGLFile LEN = TaxAcctRecLen
      GET TXGLFile, 1, TaxGLAccts(1)
      ThisAcct = FindAcct(TaxGLAccts(1).TaxAcct(T%).IntDBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).IntDBAcct
      DebitAcctName$ = GetAcctTitle$(ThisAcct)
      DebitAmt# = PayIntAmt#(T%)
      ThisAcct = FindAcct(TaxGLAccts(1).TaxAcct(T%).IntCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).IntCRAcct
      CreditAcctName$ = GetAcctTitle$(ThisAcct)
      CreditAmt# = PayIntAmt#(T%)
      GOSUB PostToGeneralJournal
      CLOSE TXGLFile
    END IF
  NEXT T%
  
  'TranType2 Payments (Adv)
  FOR T% = 1 TO 31
    IF PayAdvAmt#(T%) > 0 THEN
      TXGLFile = FREEFILE
      OPEN "TAXGLACT.DAT" FOR RANDOM SHARED AS TXGLFile LEN = TaxAcctRecLen
      GET TXGLFile, 1, TaxGLAccts(1)
      ThisAcct = FindAcct(TaxGLAccts(1).TaxAcct(T%).AdvDBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).AdvDBAcct
      DebitAcctName$ = GetAcctTitle$(ThisAcct)
      DebitAmt# = PayAdvAmt#(T%)
      ThisAcct = FindAcct(TaxGLAccts(1).TaxAcct(T%).AdvCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).AdvCRAcct
      CreditAcctName$ = GetAcctTitle$(ThisAcct)
      CreditAmt# = PayAdvAmt#(T%)
      GOSUB PostToGeneralJournal
      CLOSE TXGLFile
    END IF
  NEXT T%
  
  
  'TranType7 Adjust
  FOR T% = 1 TO 31
    IF AdjPrinAmt#(T%) > 0 THEN
      TXGLFile = FREEFILE
      OPEN "TAXGLACT.DAT" FOR RANDOM SHARED AS TXGLFile LEN = TaxAcctRecLen
      GET TXGLFile, 1, TaxGLAccts(1)
      ThisAcct = FindAcct(TaxGLAccts(1).TaxAcct(T%).TaxDBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).TaxDBAcct
      DebitAcctName$ = GetAcctTitle$(ThisAcct)
      DebitAmt# = AdjPrinAmt#(T%)
      ThisAcct = FindAcct(TaxGLAccts(1).TaxAcct(T%).TaxCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).TaxCRAcct
      CreditAcctName$ = GetAcctTitle$(ThisAcct)
      CreditAmt# = AdjPrinAmt#(T%)
      GOSUB PostToGeneralJournal
      CLOSE TXGLFile
    END IF
  NEXT T%
  
  'TranType7 Adjust (Interest)
  FOR T% = 1 TO 31
    IF AdjIntAmt#(T%) > 0 THEN
      TXGLFile = FREEFILE
      OPEN "TAXGLACT.DAT" FOR RANDOM SHARED AS TXGLFile LEN = TaxAcctRecLen
      GET TXGLFile, 1, TaxGLAccts(1)
      ThisAcct = FindAcct(TaxGLAccts(1).TaxAcct(T%).IntDBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).IntDBAcct
      DebitAcctName$ = GetAcctTitle$(ThisAcct)
      DebitAmt# = AdjIntAmt#(T%)
      ThisAcct = FindAcct(TaxGLAccts(1).TaxAcct(T%).IntCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).IntCRAcct
      CreditAcctName$ = GetAcctTitle$(ThisAcct)
      CreditAmt# = AdjIntAmt#(T%)
      GOSUB PostToGeneralJournal
      CLOSE TXGLFile
    END IF
  NEXT T%
  'TranType7 Adjust (Adv)
  FOR T% = 1 TO 31
    IF AdjAdvAmt#(T%) > 0 THEN
      TXGLFile = FREEFILE
      OPEN "TAXGLACT.DAT" FOR RANDOM SHARED AS TXGLFile LEN = TaxAcctRecLen
      GET TXGLFile, 1, TaxGLAccts(1)
      ThisAcct = FindAcct(TaxGLAccts(1).TaxAcct(T%).AdvDBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).AdvDBAcct
      DebitAcctName$ = GetAcctTitle$(ThisAcct)
      DebitAmt# = AdjAdvAmt#(T%)
      ThisAcct = FindAcct(TaxGLAccts(1).TaxAcct(T%).AdvCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).AdvCRAcct
      CreditAcctName$ = GetAcctTitle$(ThisAcct)
      CreditAmt# = AdjAdvAmt#(T%)
      GOSUB PostToGeneralJournal
      CLOSE TXGLFile
    END IF
  NEXT T%
  
  
TXBunchReturn:
  RETURN
  
PostToGeneralJournal:
  'NOTE: Journal Rec 1 is the credit, Rec 2 is the debit
  REDIM GJRec1(1 TO 2) AS TrEditRecType
  GJRec1(1).AcctRec = CreditAcctRecord
  GJRec1(1).AcctNum = CreditAcctNumber$
  GJRec1(1).AcctName = CreditAcctName$
  GJRec1(1).TrDate = WorkDate
  GJRec1(1).Ref = Ref$
  GJRec1(1).CrAmt = CreditAmt#
  GJRec1(1).EType = "C"
  GJRec1(1).Desc = "FROM TAXES"
  GJRec1(1).Src = "IF"
  PUT #GJFile, , GJRec1(1)
  
  GJRec1(2).AcctRec = DebitAcctRecord
  GJRec1(2).AcctNum = DebitAcctNumber$
  GJRec1(2).AcctName = DebitAcctName$
  GJRec1(2).TrDate = WorkDate
  GJRec1(2).Ref = Ref$
  GJRec1(2).DrAmt = DebitAmt#
  GJRec1(2).EType = "D"
  GJRec1(2).Desc = "FROM TAXES"
  GJRec1(2).Src = "IF"
  PUT #GJFile, , GJRec1(2)
RETURN
  
  
TaxEnd:
  EXIT SUB
  RETURN

  
END SUB

SUB ExtractUB (ThruDate%)
  
  Today$ = DATE$
  Ref$ = "UB" + LEFT$(Today$, 2) + MID$(Today$, 4, 2) + RIGHT$(Today$, 2)
  
  REDIM TranInfo(1) AS TranRecInfoType
  
  Dash80$ = STRING$(80, "-")
  P2S$ = SPACE$(4)
  
  
  REDIM GJRec1(1 TO 2) AS TrEditRecType
  GJRecLen = LEN(GJRec1(1))
  GJFile = FREEFILE
  OPEN "GLTRXED.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS GJFile LEN = GJRecLen
  NumEdTrans = LOF(GJFile) \ GJRecLen
  CLOSE
  
  IF NumEdTrans > 0 THEN
    Ok = MsgBox%("UBSENDGL", "JEDNZERO")
    GOTO SendExitUB
  END IF
  
  Ok = MsgBox%("UBSENDGL", "MUSTEXT2")
  SELECT CASE Ok
  CASE 2
    GOTO SendExitUB
  END SELECT
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpFileNum, UBSetUpLen
  GET UBSetUpFileNum, 1, UBSetUpRec(1)
  AcctMeth$ = QPTrim$(UBSetUpRec(1).MethAcct)
  IF (LEN(AcctMeth$) = 0) THEN
    Ok = MsgBox%("UBSENDGL", "NOMETHOD")
    GOTO SendExitUB
  END IF
  
  SELECT CASE AcctMeth$
  CASE "C"
    InterfaceMethod = 1
  CASE "A"
    InterfaceMethod = 2
  CASE ELSE
    GOTO SendExitUB
  END SELECT
  
  RPTFile = FREEFILE
  OPEN "UBNOTFND.RPT" FOR OUTPUT AS RPTFile
  GOSUB NotFoundHeader
  
  'ShowProcessingScrn "Verifying GL Transfer Accounts"
  
  FOR RevCnt = 1 TO MaxRevsCnt
    TempRev$ = QPTrim$(UBSetUpRec(1).Revenues(RevCnt).RevName)
    IF LEN(TempRev$) = 0 THEN
      NumOfRevs = RevCnt - 1
      EXIT FOR
    ELSE
      REDIM PRESERVE GJInfo(1 TO RevCnt) AS GJXferRecType
      GJInfo(RevCnt).RevText = TempRev$
      GJInfo(RevCnt).BAcctInfo.DAcctNo = UBSetUpRec(1).BillAcct(RevCnt).DebitAcct
      GJInfo(RevCnt).BAcctInfo.CAcctNo = UBSetUpRec(1).BillAcct(RevCnt).CreditAcct
      GJInfo(RevCnt).PAcctInfo.DAcctNo = UBSetUpRec(1).PayAcct(RevCnt).DebitAcct
      GJInfo(RevCnt).PAcctInfo.CAcctNo = UBSetUpRec(1).PayAcct(RevCnt).CreditAcct
      IF UBSetUpRec(1).Revenues(RevCnt).UseDep = "Y" THEN
        GJInfo(RevCnt).DAcctInfo.DAcctNo = UBSetUpRec(1).DepAcct(RevCnt).DebitAcct
        GJInfo(RevCnt).DAcctInfo.CAcctNo = UBSetUpRec(1).DepAcct(RevCnt).CreditAcct
      END IF
    END IF
  NEXT
  
  'check to see if they are valid GL accounts
  GOSUB ValidateGLAccounts
  
  IF BadAcct THEN
    GOTO SendExitUB
  END IF
  
  ClearBox

  QPrintRC "Searching Cash Transactions.", 11, 26, 126
  QPrintRC "New Transactions:", 13, 29, Cnf.HiLite
  
  REDIM UBTransRec(1)  AS UBTransRecType
  UBTransRecLen = LEN(UBTransRec(1))
  
  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTransRecLen
  NumOfTRecs& = LOF(UBTran) \ UBTransRecLen
  
  FOR TCnt& = NumOfTRecs& TO 1 STEP -1
    GET #UBTran, TCnt&, UBTransRec(1)
    IF LEN(QPTrim$(UBTransRec(1).Posted2GL)) = 0 OR QPTrim$(UBTransRec(1).Posted2GL) = "N" THEN
      'Store trans rec numbers and dates in array
      IF UBTransRec(1).TransDate <= ThruDate% THEN
        FoundCnt = FoundCnt + 1
        REDIM PRESERVE TranInfo(FoundCnt) AS TranRecInfoType
        TranInfo(FoundCnt).TranDate = UBTransRec(1).TransDate
        TranInfo(FoundCnt).TranRecNo = TCnt&
      END IF
    ELSE
      'NGCnt = NGCnt + 1
    END IF
    RSET P2S$ = QPStrI$(FoundCnt)
    QPrintRC P2S$, 13, 47, Cnf.HiLite
    'SmallPause
    'IF NGCnt >= 2500 THEN EXIT FOR
  NEXT
  
  IF FoundCnt = 0 THEN
    CLOSE
    Ok = MsgBox%("UBSENDGL", "NONEFOND")
    GOTO SendExitUB
  END IF
  
  SortT TranInfo(1), FoundCnt, 0, 6, 0, -1
  'sort'em by date. oldest first
  'Array(1), NumElem, Dir, StructSize, MemOff, MemSize
  
  GJFile = FREEFILE
  OPEN "GLTRXED.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS GJFile LEN = GJRecLen
  
  FirstTran = 1
  ThisDate = TranInfo(1).TranDate
  FOR Cnt = 1 TO FoundCnt
    IF ThisDate <> TranInfo(Cnt).TranDate THEN
      ThisDate = TranInfo(Cnt).TranDate
      LastTran = Cnt - 1
      GOSUB ProcessThisBunchUB
      FirstTran = Cnt
    END IF
  NEXT
  LastTran = FoundCnt
  GOSUB ProcessThisBunchUB
  
  'transactions as interfaced
  
  FOR Cnt = 1 TO FoundCnt
    GET #UBTran, TranInfo(Cnt).TranRecNo, UBTransRec(1)
    UBTransRec(1).Posted2GL = "Y"
    PUT #UBTran, TranInfo(Cnt).TranRecNo, UBTransRec(1)
  NEXT
  CLOSE
  
  
  SLEEP 2
SendExitUB:
  
  EXIT SUB
  
NotFoundHeader:
  PageNo = PageNo + 1
  PRINT #RPTFile, "Utility Billing GL Transfer Invalid Account Listing."; TAB(70); "Page:"; PageNo
  PRINT #RPTFile, TOWNNAME$
  PRINT #RPTFile, "Report Date: "; DATE$
  PRINT #RPTFile, "Revenue           Acct. Type              Debit Acct.          Credit Acct."
  PRINT #RPTFile, Dash80$
  NumPrinted = 0
  RETURN
  
PrintBadAcct:
  IF LEN(QPTrim$(BadCAcct$)) = 0 THEN
    BadCAcct$ = "Undefined"
  END IF
  
  PRINT #RPTFile, GJInfo(RevCnt).RevText;
  
  SELECT CASE ActT
  CASE 1
    ActPg$ = "Billing"
  CASE 2
    ActPg$ = "Payment"
  CASE 3
    ActPg$ = "Deposit"
  END SELECT
  PRINT #RPTFile, TAB(22); ActPg$;
  PRINT #RPTFile, TAB(43); BadDAcct$; TAB(64); BadCAcct$
  RETURN
  
ProcessThisBunchUB:
  FOR RevCnt = 1 TO NumOfRevs
    GJInfo(RevCnt).BAcctInfo.CreditAmt = 0
    GJInfo(RevCnt).BAcctInfo.DebitAmt = 0
    GJInfo(RevCnt).PAcctInfo.CreditAmt = 0
    GJInfo(RevCnt).PAcctInfo.DebitAmt = 0
    GJInfo(RevCnt).DAcctInfo.CreditAmt = 0
    GJInfo(RevCnt).DAcctInfo.DebitAmt = 0
  NEXT
  
  FOR PCnt = FirstTran TO LastTran
    IF PCnt = FirstTran THEN
      WorkDate = TranInfo(PCnt).TranDate
    END IF
    GET #UBTran, TranInfo(PCnt).TranRecNo, UBTransRec(1)
    
    SELECT CASE InterfaceMethod
    CASE 1      'Cash Central
      SELECT CASE UBTransRec(1).TransType
      CASE TranUtilityBill      ' 1=Utility bill
        'no action
      CASE TranLateCharge       ' 2=late charge
        'no action
      CASE TranReconnectFee     ' 3=reconnect fee
        'no action
      CASE TranBillPayment      ' 4=Bill Payment
        FOR RevCnt = 1 TO NumOfRevs
          GJInfo(RevCnt).PAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).PAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
          GJInfo(RevCnt).PAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).PAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
        NEXT
      CASE TranAppliedDeposit   ' 5=Applied Deposit
        'no action
      CASE TranPenaltyCharge    ' 6=Penalty Charge
        'no action
      CASE TranDepositPayment   ' 7=Deposit Payment
        FOR RevCnt = 1 TO NumOfRevs
          GJInfo(RevCnt).DAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).DAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
          GJInfo(RevCnt).DAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).DAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
        NEXT
      CASE TranDraftPayment     ' 8=Draft Payment
        FOR RevCnt = 1 TO NumOfRevs
          GJInfo(RevCnt).PAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).PAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
          GJInfo(RevCnt).PAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).DAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
        NEXT
      CASE TranRefundDeposit    ' 9=Refund Deposit
        'no action
      CASE TranBeginBalance     '10=Beginning Balance
        'no action
      CASE TranUpwardAdjustment '11=Upward Adjustments
        'no action
      CASE TranDownwardAdjustment               '12=Downward Adjustments
        'no action
      END SELECT
      
    CASE 2      'Accural
      SELECT CASE UBTransRec(1).TransType
      CASE TranUtilityBill      ' 1=Utility bill
        FOR RevCnt = 1 TO NumOfRevs
          GJInfo(RevCnt).BAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).BAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
          GJInfo(RevCnt).BAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).BAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
        NEXT
      CASE TranLateCharge       ' 2=late charge
        FOR RevCnt = 1 TO NumOfRevs
          GJInfo(RevCnt).BAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).BAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
          GJInfo(RevCnt).BAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).BAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
        NEXT
      CASE TranReconnectFee     ' 3=reconnect fee
        FOR RevCnt = 1 TO NumOfRevs
          GJInfo(RevCnt).BAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).BAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
          GJInfo(RevCnt).BAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).BAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
        NEXT
      CASE TranBillPayment      ' 4=Bill Payment
        FOR RevCnt = 1 TO NumOfRevs
          GJInfo(RevCnt).PAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).PAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
          GJInfo(RevCnt).PAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).PAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
        NEXT
      CASE TranAppliedDeposit   ' 5=Applied Deposit
        'no action
        'FOR RevCnt = 1 TO NumOfRevs
        '  GJInfo(RevCnt).dacctInfo.CreditAmt = Round#(GJInfo(RevCnt).dacctInfo.CreditAmt - UBTransRec(1).RevAmt(RevCnt))
        '  GJInfo(RevCnt).pAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).pAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
        'NEXT
        
      CASE TranPenaltyCharge    ' 6=Penalty Charge
        FOR RevCnt = 1 TO NumOfRevs
          GJInfo(RevCnt).BAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).BAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
          GJInfo(RevCnt).BAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).BAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
        NEXT
      CASE TranDepositPayment   ' 7=Deposit Payment
        FOR RevCnt = 1 TO NumOfRevs
          GJInfo(RevCnt).DAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).DAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
          GJInfo(RevCnt).DAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).DAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
        NEXT
      CASE TranDraftPayment     ' 8=Draft Payment
        FOR RevCnt = 1 TO NumOfRevs
          GJInfo(RevCnt).PAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).PAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
          GJInfo(RevCnt).PAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).PAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
        NEXT
      CASE TranRefundDeposit    ' 9=Refund Deposit
        'no action
        '  FOR RevCnt = 1 TO NumOfRevs
        '    GJInfo(RevCnt).dacctInfo.CreditAmt = Round#(GJInfo(RevCnt).dacctInfo.CreditAmt - UBTransRec(1).RevAmt(RevCnt))
        '    GJInfo(RevCnt).dacctInfo.DebitAmt = Round#(GJInfo(RevCnt).dacctInfo.DebitAmt - UBTransRec(1).RevAmt(RevCnt))
        '  NEXT
      CASE TranBeginBalance     '10=Beginning Balance
        'no action
      CASE TranUpwardAdjustment '11=Upward Adjustments
        FOR RevCnt = 1 TO NumOfRevs
          GJInfo(RevCnt).BAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).BAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
          GJInfo(RevCnt).BAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).BAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
        NEXT
      CASE TranDownwardAdjustment               '12=Downward Adjustments
        FOR RevCnt = 1 TO NumOfRevs
          GJInfo(RevCnt).BAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).BAcctInfo.CreditAmt - UBTransRec(1).RevAmt(RevCnt))
          GJInfo(RevCnt).BAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).BAcctInfo.DebitAmt - UBTransRec(1).RevAmt(RevCnt))
        NEXT
      END SELECT
    END SELECT
    'SmallPause
    
  NEXT
  
  
  'NOTE: Journal Rec 1 is the credit, Rec 2 is the debit
  FOR RevCnt = 1 TO NumOfRevs
    REDIM GJRec1(1 TO 2) AS TrEditRecType
    IF GJInfo(RevCnt).BAcctInfo.CreditAmt <> 0 THEN
      GJRec1(1).AcctRec = GJInfo(RevCnt).BAcctInfo.CRecNo
      GJRec1(1).AcctNum = GJInfo(RevCnt).BAcctInfo.CAcctNo
      GJRec1(1).AcctName = GJInfo(RevCnt).BAcctInfo.CTitle
      GJRec1(1).TrDate = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).CrAmt = GJInfo(RevCnt).BAcctInfo.CreditAmt
      GJRec1(1).EType = "C"
      GJRec1(1).Desc = "FROM UTILITIES"
      GJRec1(1).Src = "IF"
      PUT #GJFile, , GJRec1(1)
    END IF
    IF GJInfo(RevCnt).BAcctInfo.DebitAmt <> 0 THEN
      GJRec1(2).AcctRec = GJInfo(RevCnt).BAcctInfo.DRecNo
      GJRec1(2).AcctNum = GJInfo(RevCnt).BAcctInfo.DAcctNo
      GJRec1(2).AcctName = GJInfo(RevCnt).BAcctInfo.DTitle
      GJRec1(2).TrDate = WorkDate
      GJRec1(2).Ref = Ref$
      GJRec1(2).DrAmt = GJInfo(RevCnt).BAcctInfo.DebitAmt
      GJRec1(2).EType = "D"
      GJRec1(2).Desc = "FROM UTILITIES"
      GJRec1(2).Src = "IF"
      PUT #GJFile, , GJRec1(2)
    END IF
  NEXT
  
  FOR RevCnt = 1 TO NumOfRevs
    REDIM GJRec1(1 TO 2) AS TrEditRecType
    IF GJInfo(RevCnt).PAcctInfo.CreditAmt <> 0 THEN
      GJRec1(1).AcctRec = GJInfo(RevCnt).PAcctInfo.CRecNo
      GJRec1(1).AcctNum = GJInfo(RevCnt).PAcctInfo.CAcctNo
      GJRec1(1).AcctName = GJInfo(RevCnt).PAcctInfo.CTitle
      GJRec1(1).TrDate = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).CrAmt = GJInfo(RevCnt).PAcctInfo.CreditAmt
      GJRec1(1).EType = "C"
      GJRec1(1).Desc = "FROM UTILITIES"
      GJRec1(1).Src = "IF"
      PUT #GJFile, , GJRec1(1)
    END IF
    IF GJInfo(RevCnt).PAcctInfo.DebitAmt <> 0 THEN
      GJRec1(2).AcctRec = GJInfo(RevCnt).PAcctInfo.DRecNo
      GJRec1(2).AcctNum = GJInfo(RevCnt).PAcctInfo.DAcctNo
      GJRec1(2).AcctName = GJInfo(RevCnt).PAcctInfo.DTitle
      GJRec1(2).TrDate = WorkDate
      GJRec1(2).Ref = Ref$
      GJRec1(2).DrAmt = GJInfo(RevCnt).PAcctInfo.DebitAmt
      GJRec1(2).EType = "D"
      GJRec1(2).Desc = "FROM UTILITIES"
      GJRec1(2).Src = "IF"
      PUT #GJFile, , GJRec1(2)
    END IF
  NEXT
  
  FOR RevCnt = 1 TO NumOfRevs
    REDIM GJRec1(1 TO 2) AS TrEditRecType
    IF GJInfo(RevCnt).DAcctInfo.CreditAmt <> 0 THEN
      GJRec1(1).AcctRec = GJInfo(RevCnt).DAcctInfo.CRecNo
      GJRec1(1).AcctNum = GJInfo(RevCnt).DAcctInfo.CAcctNo
      GJRec1(1).AcctName = GJInfo(RevCnt).DAcctInfo.CTitle
      GJRec1(1).TrDate = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).CrAmt = GJInfo(RevCnt).DAcctInfo.CreditAmt
      GJRec1(1).EType = "C"
      GJRec1(1).Desc = "FROM UTILITIES"
      GJRec1(1).Src = "IF"
      PUT #GJFile, , GJRec1(1)
    END IF
    IF GJInfo(RevCnt).DAcctInfo.DebitAmt <> 0 THEN
      GJRec1(2).AcctRec = GJInfo(RevCnt).DAcctInfo.DRecNo
      GJRec1(2).AcctNum = GJInfo(RevCnt).DAcctInfo.DAcctNo
      GJRec1(2).AcctName = GJInfo(RevCnt).DAcctInfo.DTitle
      GJRec1(2).TrDate = WorkDate
      GJRec1(2).Ref = Ref$
      GJRec1(2).DrAmt = GJInfo(RevCnt).DAcctInfo.DebitAmt
      GJRec1(2).EType = "D"
      GJRec1(2).Desc = "FROM UTILITIES"
      GJRec1(2).Src = "IF"
      PUT #GJFile, , GJRec1(2)
    END IF
  NEXT
  
UBBunchReturn:
  RETURN
  
ValidateGLAccounts:
  BadAcct = False
  FOR RevCnt = 1 TO NumOfRevs
    'Billing Accounts
    IF InterfaceMethod = 2 THEN
      'NOTE: We Only check billing accounts if Accural method
      ActT = 1
      ThisAcct = FindAcct(GJInfo(RevCnt).BAcctInfo.DAcctNo)
      IF ThisAcct <= 0 THEN
        BadDAcct$ = GJInfo(RevCnt).BAcctInfo.DAcctNo
        BadAcct = True
      ELSE
        GJInfo(RevCnt).BAcctInfo.DRecNo = ThisAcct
        GJInfo(RevCnt).BAcctInfo.DTitle = GetAcctTitle$(ThisAcct)
        BadDAcct$ = "     OK"
      END IF
      ThisAcct = FindAcct(GJInfo(RevCnt).BAcctInfo.CAcctNo)
      IF ThisAcct <= 0 THEN
        BadCAcct$ = GJInfo(RevCnt).BAcctInfo.CAcctNo
        BadAcct = True
      ELSE
        GJInfo(RevCnt).BAcctInfo.CRecNo = ThisAcct
        GJInfo(RevCnt).BAcctInfo.CTitle = GetAcctTitle$(ThisAcct)
        BadCAcct$ = "     OK"
      END IF
      GOSUB PrintBadAcct
    END IF
    
    'Payment Accounts
    ActT = 2
    ThisAcct = FindAcct(GJInfo(RevCnt).PAcctInfo.DAcctNo)
    IF ThisAcct <= 0 THEN
      BadDAcct$ = GJInfo(RevCnt).PAcctInfo.DAcctNo
      BadAcct = True
    ELSE
      GJInfo(RevCnt).PAcctInfo.DRecNo = ThisAcct
      GJInfo(RevCnt).PAcctInfo.DTitle = GetAcctTitle$(ThisAcct)
      BadDAcct$ = "     OK"
    END IF
    ThisAcct = FindAcct(GJInfo(RevCnt).PAcctInfo.CAcctNo)
    IF ThisAcct <= 0 THEN
      BadCAcct$ = GJInfo(RevCnt).PAcctInfo.CAcctNo
      BadAcct = True
    ELSE
      GJInfo(RevCnt).PAcctInfo.CRecNo = ThisAcct
      GJInfo(RevCnt).PAcctInfo.CTitle = GetAcctTitle$(ThisAcct)
      BadCAcct$ = "     OK"
    END IF
    GOSUB PrintBadAcct
    
    'Deposit Accounts
    ActT = 3
    IF UBSetUpRec(1).Revenues(RevCnt).UseDep = "Y" THEN
      ThisAcct = FindAcct(GJInfo(RevCnt).DAcctInfo.DAcctNo)
      IF ThisAcct <= 0 THEN
        BadDAcct$ = GJInfo(RevCnt).DAcctInfo.DAcctNo
        BadAcct = True
      ELSE
        GJInfo(RevCnt).DAcctInfo.DRecNo = ThisAcct
        GJInfo(RevCnt).DAcctInfo.DTitle = GetAcctTitle$(ThisAcct)
        BadDAcct$ = "     OK"
      END IF
    ELSE
      BadDAcct$ = "    N/A"
    END IF
    IF UBSetUpRec(1).Revenues(RevCnt).UseDep = "Y" THEN
      ThisAcct = FindAcct(GJInfo(RevCnt).DAcctInfo.CAcctNo)
      IF ThisAcct <= 0 THEN
        BadCAcct$ = GJInfo(RevCnt).DAcctInfo.CAcctNo
        BadAcct = True
      ELSE
        GJInfo(RevCnt).DAcctInfo.CRecNo = ThisAcct
        GJInfo(RevCnt).DAcctInfo.CTitle = GetAcctTitle$(ThisAcct)
        BadCAcct$ = "     OK"
      END IF
    ELSE
      BadCAcct$ = "    N/A"
    END IF
    GOSUB PrintBadAcct
  NEXT
  CLOSE RPTFile
  
  IF BadAcct THEN
    Ok = MsgBox%("UBSENDGL", "UBADACTS")
    EntryPoint = 0
    PrintRptFile "GL Transfer Invalid Account List.", "UBNOTFND.RPT", 1, RetCode, EntryPoint
  END IF
  KILL "UBNOTFND.RPT"
RETURN
  
END SUB

SUB IFInit
  
  REDIM Form$(0, 0)             'Form array holds data while editing
  REDIM Fld(0) AS FieldInfo     'Field info Array
  REDIM Frm(1) AS FormInfo      'Form editing info
  
  LibLoadDisplayForm "GL.QSL", "SETUPIF", Form$(), Fld()
  
  Form$(1, 0) = "N"
  Form$(2, 0) = "N"
  Form$(3, 0) = "N"
  Form$(4, 0) = "N"
  Form$(5, 0) = "N"
  Action = 1
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key '--SaveButton
      '--Fld 2 = OutputDevice
      F1$ = Form$(1, 0)
      F2$ = Form$(2, 0)
      F3$ = Form$(3, 0)
      F4$ = Form$(4, 0)
      F5$ = Form$(5, 0)
      GOSUB InitIF
      DONE = True
      
    CASE ESC
      EXIT SUB
      
    END SELECT
    
  LOOP UNTIL DONE
  
  EXIT SUB
  
InitIF:
  LibLoadDisplayForm "GL.QSL", "PROCESIF", Form$(), Fld()
  IF F1$ = "Y" THEN
    QPrintRC "INITIALIZING CASH TRANSACTIONS.        ", 12, 23, 15
    REDIM CMTransRec(1)  AS CMTransRecType
    CMTransRecLen = LEN(CMTransRec(1))
    CMTran = FREEFILE
    OPEN "CMTRANS.DAT" FOR RANDOM SHARED AS CMTran LEN = CMTransRecLen
    NumOfTRecs& = LOF(CMTran) \ CMTransRecLen
    FOR TCnt& = NumOfTRecs& TO 1 STEP -1
      GET #CMTran, TCnt&, CMTransRec(1)
      CMTransRec(1).Trans2GL = "Y"
      PUT #CMTran, TCnt&, CMTransRec(1)
    NEXT TCnt&
    CLOSE CMTran
  END IF
  
  IF F2$ = "Y" THEN
    QPrintRC "INITIALIZING UTILITY TRANSACTIONS.     ", 12, 23, 15
    REDIM UBTransRec(1)  AS UBTransRecType
    UBTransRecLen = LEN(UBTransRec(1))
    UBTran = FREEFILE
    OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTransRecLen
    NumOfTRecs& = LOF(UBTran) \ UBTransRecLen
    FOR TCnt& = NumOfTRecs& TO 1 STEP -1
      GET #UBTran, TCnt&, UBTransRec(1)
      UBTransRec(1).Posted2GL = "Y"
      PUT #UBTran, TCnt&, UBTransRec(1)
    NEXT TCnt&
    CLOSE UBTran
  END IF
  
  IF F3$ = "Y" THEN
    QPrintRC "INITIALIZING TAX TRANSACTIONS.         ", 12, 23, 15
    REDIM TaxTrans(1) AS TaxTransactionType
    TaxTranRecLen = LEN(TaxTrans(1))
    TranFile = FREEFILE
    OPEN "TAXTRANS.DAT" FOR RANDOM SHARED AS TranFile LEN = TaxTranRecLen
    NumOfRec& = LOF(TranFile) \ TaxTranRecLen
    FOR Cnt& = 1 TO NumOfRec&
      GET #TranFile, Cnt&, TaxTrans(1)
      TaxTrans(1).Posted2GL = "Y"
      PUT #TranFile, Cnt&, TaxTrans(1)
    NEXT Cnt&
    CLOSE TranFile
    SLEEP 2
  END IF
  
  IF F4$ = "Y" THEN
    QPrintRC "INITIALIZING BUS. LICENSE TRANSACTIONS.", 12, 23, 15
    REDIM ARTransRec(1) AS ARTransRecType
    ARTransRecLen = LEN(ARTransRec(1))
    ARTransFile = FREEFILE
    OPEN "ARTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARTransFile LEN = ARTransRecLen
    NumOfTransRecs& = LOF(ARTransFile) \ ARTransRecLen
    FOR TCnt& = 1 TO NumOfTransRecs&
      GET ARTransFile, TCnt&, ARTransRec(1)
      ARTransRec(1).Posted2GL = "Y"
      PUT ARTransFile, TCnt&, ARTransRec(1)
    NEXT TCnt&
    CLOSE ARTransFile
  END IF
  IF F5$ = "Y" THEN
    QPrintRC "INITIALIZING DECAL TRANSACTIONS.       ", 12, 23, 15
    SLEEP 2
  END IF
  PRINT CHR$(7);
  QPrintRC "DONE!!!                                ", 12, 23, 15
  SLEEP 3
  
  RETURN
  
END SUB

SUB LoadUBSetUpFile (UBSetUpFileNum, UBSetUpLen)
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  UBSetUpLen = LEN(UBSetUpRec(1))
  UBSetUpFileNum = FREEFILE
  OPEN "UBSETUP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBSetUpFileNum LEN = UBSetUpLen
  
END SUB

