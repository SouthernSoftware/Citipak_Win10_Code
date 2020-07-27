DEFINT A-Z
'Remington VA Tax Conver for Real Estate
DECLARE SUB BalanceListing ()
DECLARE SUB OpenTaxCustFile (NumOfTaxRecs%, TaxFile%)
DECLARE SUB OpenTaxPropFile (NumOfPropRecs%, PropTaxFile%)
DECLARE SUB OpenTaxPersFile (NumOfPersRecs%, PersTaxFile%)
DECLARE SUB DisplayTaxScrn (ScrnName$)
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE SUB CustomerListing ()
DECLARE SUB TAXCustomerMenu ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB ClearBack ()
DECLARE SUB SendDist2GL ()
DECLARE SUB ClearScrn ()
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION WEnvTest% ()
DECLARE FUNCTION Round# (b#)
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
  
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED PersRec(1) AS PersonalRecType
  DIM SHARED TBCustRec(1) AS TBCustType
  DIM SHARED TBCustRec1(1) AS TBCust1Type
  DIM SHARED TaxTrans(1) AS TaxTransactionType
  

  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
  DIM TransRec&(250)
  TaxFile = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
  TransFile = FREEFILE
  OPEN "TaxTrans.dat" FOR RANDOM AS TransFile LEN = LEN(TaxTrans(1))
  

  FOR Cnt& = 1 TO NumOfTaxRecs
  GET TaxFile, Cnt&, TaxCustRec(1)
  IF Cnt& = 70 THEN STOP
  TransRecord& = TaxCustRec(1).LastTrans
  TCnt& = 0
   WHILE TransRecord& <> 0
     GET TransFile, TransRecord&, TaxTrans(1)
     TCnt& = TCnt& + 1
     TransRec&(TCnt&) = TransRecord&
     TransRecord& = TaxTrans(1).LastTrans
   WEND
   GOSUB FixThisOne
   NEXT Cnt&
   CLOSE
   END


FixThisOne:
    IF TCnt& = 0 THEN RETURN
    FOR LCnt& = 1 TO TCnt&
      GET TransFile, TransRec&(LCnt&), TaxTrans(1)
      IF TaxTrans(1).TranType = 1 THEN
       BillRec& = TransRec&(LCnt&)
       TaxTrans(1).Revenue.Principle1Pd = 0
       TaxTrans(1).Revenue.Principle2Pd = 0
       TaxTrans(1).Revenue.Principle3Pd = 0
       TaxTrans(1).Revenue.Principle4Pd = 0
       TaxTrans(1).Revenue.Principle5Pd = 0
       TaxTrans(1).Revenue.InterestPd = 0
       TaxTrans(1).Revenue.PenaltyPd = 0
       TaxTrans(1).Revenue.CollectionPd = 0
       PUT TransFile, BillRec&, TaxTrans(1)
      'Now Go Thru Each Record and See if Any Paymnets or Abatements Belong to This Bill Rec
       FOR Snt& = 1 TO TCnt&
       GET TransFile, TransRec&(Snt&), TaxTrans(1)
       IF TaxTrans(1).BelongTo = BillRec& AND (TaxTrans(1).TranType = 3 OR TaxTrans(1).TranType = 2) THEN
        tp1# = TaxTrans(1).Revenue.Principle1Pd
        tp2# = TaxTrans(1).Revenue.Principle2Pd
        tp3# = TaxTrans(1).Revenue.Principle3Pd
        tp4# = TaxTrans(1).Revenue.Principle4Pd
        tp5# = TaxTrans(1).Revenue.Principle5Pd
        tpi# = TaxTrans(1).Revenue.InterestPd
        tpp# = TaxTrans(1).Revenue.PenaltyPd
        tpc# = TaxTrans(1).Revenue.CollectionPd
          GET TransFile, BillRec&, TaxTrans(1)
       TaxTrans(1).Revenue.Principle1Pd = TaxTrans(1).Revenue.Principle1Pd + tp1#
       TaxTrans(1).Revenue.Principle2Pd = TaxTrans(1).Revenue.Principle2Pd + tp2#
       TaxTrans(1).Revenue.Principle3Pd = TaxTrans(1).Revenue.Principle3Pd + tp3#
       TaxTrans(1).Revenue.Principle4Pd = TaxTrans(1).Revenue.Principle4Pd + tp4#
       TaxTrans(1).Revenue.Principle5Pd = TaxTrans(1).Revenue.Principle5Pd + tp5#
       TaxTrans(1).Revenue.InterestPd = TaxTrans(1).Revenue.InterestPd + tpi#
       TaxTrans(1).Revenue.PenaltyPd = TaxTrans(1).Revenue.PenaltyPd + tpp#
       TaxTrans(1).Revenue.CollectionPd = TaxTrans(1).Revenue.CollectionPd + tpc#
       PUT TransFile, BillRec&, TaxTrans(1)
      END IF
       NEXT Snt&
       END IF
      NEXT LCnt&
      RETURN


END SUB

SUB OpenTaxCustFile (NumOfTaxRecs, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
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

