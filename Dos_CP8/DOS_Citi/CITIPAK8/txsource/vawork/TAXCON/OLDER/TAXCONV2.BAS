DEFINT A-Z
DECLARE SUB balancelisting ()
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
  
  
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED TaxCust(1) AS TaxCustType
  DIM SHARED PersRec(1) AS PersonalRecType
  DIM SHARED TaxTrans(1) AS TaxTransactionType
  DIM SHARED RealRec(1) AS PropertyRecType
  DIM SHARED RABILL(1) AS RABILL
  
  STACK 5000
  balancelisting
  
  END

SUB balancelisting

REM start here *******************
CLS
  OpenTaxCustFile NumOfTaxRecs, TaxFile

  PersFile = FREEFILE
  OPEN "TAXPERS.DAT" FOR RANDOM SHARED AS PersFile LEN = LEN(PersRec(1))
  
  TransFile = FREEFILE
  OPEN "TaxTrans.dat" FOR RANDOM AS TransFile LEN = LEN(TaxTrans(1))

  FOR Cnt& = 1 TO LOF(TransFile) / LEN(TaxTrans(1))
  GET TransFile, Cnt&, TaxTrans(1)
  TBill# = VAL(MID$(TaxTrans(1).Description, 12, 5))
  CustRec& = TaxTrans(1).CustomerRec
  IF TBill# >= 98000 AND TBill# <= 98999 THEN
  IF TaxTrans(1).BillType = "P" THEN

      TaxTrans(1).TaxYear = 0
      TaxTrans(1).TranType = 1  '1=Bill 2=Payment 3=Release 4=Interest 5=Penalty 6=Collection/Ad Cost Billing
      TaxTrans(1).BillType = "P"                'R=Real P=Personal Property C=Combined (NC/GA)
      TaxTrans(1).Amount = 0
      TaxTrans(1).Revenue.Principle1 = 0
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
      TaxTrans(1).Description = "Voided Tax Bill #"
      TaxTrans(1).Posted2GL = "N"
      TaxTrans(1).LastTrans = 0
      TaxTrans(1).BelongTo = 0
      TaxTrans(1).Padding = ""

      PUT #TransFile, Cnt&, TaxTrans(1)

     GET #TaxFile, CustRec&, TaxCustRec(1)
        TaxCustRec(1).LastTrans = 0
     PUT #TaxFile, CustRec&, TaxCustRec(1)




         PRecord& = TaxCustRec(1).FirstPersRec
        WHILE PRecord& <> 0
         GET #PersFile, PRecord&, PersRec(1)
          PersRec(1).LastYrPrinted = 0
         PUT #PersFile, PRecord&, PersRec(1)
         PRecord& = PersRec(1).NextRec
        WEND

       END IF
      END IF

      NEXT Cnt&
      CLOSE
      EXIT SUB


END SUB

SUB OpenTaxCustFile (NumOfTaxRecs, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
END SUB

SUB OpenTaxPersFile (NumOfPersRecs, REALTaxFile)
  REALTaxFile = FREEFILE
  OPEN "TAXREAL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #REALTaxFile LEN = LEN(RealRec(1))
  NumOfPersRecs = LOF(REALTaxFile) / LEN(RealRec(1))
  
END SUB

