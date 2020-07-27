DEFINT A-Z
'TOWN OF VICTORIA CONVERSION FROM LUNENBURG COUNTY REAL ESTATE AFTER BILLS RUN
DECLARE SUB BalanceListing ()
DECLARE SUB OpenTaxCustFile (NumOfTaxRecs%, TAXFILE%)
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
  '$INCLUDE: 'lctax.bi'
  '$INCLUDE: 'Taxfiles.BI'
  '$INCLUDE: 'PROPAbst.BI'
  
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType

  
  DIM SHARED TBcustRec(1) AS TBRECust

  DIM SHARED TaxTran(1) AS TaxTransactionType

  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS

   TaxTranRecLen = LEN(TaxTran(1))
   OpenTaxPropFile NumOfPropRecs, PropTaxFile

   'Init Prop File to Zero Values
   FOR Cnt& = 1 TO LOF(PropTaxFile) / LEN(PropertyRec(1))
   GET PropTaxFile, Cnt&, PropertyRec(1)
     PropertyRec(1).PROPVALU = 0
     PropertyRec(1).EXMPSENI = 0
   PUT PropTaxFile, Cnt&, PropertyRec(1)
   NEXT Cnt&


   OPEN "F:\TBRECUST.DAT" FOR RANDOM AS #11 LEN = LEN(TBcustRec(1))
   TAXFILE = FREEFILE

   OPEN "RETXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TAXFILE LEN = LEN(TaxCustRec(1))
   NumOfTaxRecs = LOF(TAXFILE) / LEN(TaxCustRec(1))

   FOR x = 1 TO LOF(11) / LEN(TBcustRec(1))
   GET 11, x, TBcustRec(1)
   Updated = 0

   IF VAL(TBcustRec(1).District) = 10 AND LEFT$(LTRIM$(TBcustRec(1).MPCurYrBillNumber), 2) = "99" THEN

    FOR Snt& = 1 TO LOF(TAXFILE) / LEN(TaxCustRec(1))
       GET TAXFILE, Snt&, TaxCustRec(1)
    IF TaxCustRec(1).CountyAcct = x THEN
      Record! = Snt&
      GOSUB UpdateCustomer
      Updated = 1
      EXIT FOR
    END IF

    NEXT Snt&

    IF Updated = 0 THEN STOP

   END IF 'end search for victoria accounts
   NEXT x

   CLOSE
   STOP

   

       
UpdateCustomer:
   PropRec& = TaxCustRec(1).FirstPropRec
   GOSUB UpdateReal

   'Set Trans & Bills to Zero
     TaxTran = FREEFILE
     OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
     TaxTran(1).TransDate = Date2Num%("12-01-1999")
     TaxTran(1).TaxYear = 1999
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = "R"
     TaxTran(1).Amount = TBcustRec(1).MPCurYrTaxAmount
     TaxTran(1).Revenue.Principle1 = TBcustRec(1).MPCurYrTaxAmount
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Description = "Tax Bill" + TBcustRec(1).MPCurYrBillNumber
     TaxTran(1).Posted2Gl = "Y"
     TaxTran(1).CustomerRec = Record!
     TaxTran(1).LastTrans = TaxCustRec(1).LastTrans
     TaxTran(1).BelongTo = 0
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TAXFILE, Record!, TaxCustRec(1)
     TaxCustRec(1).LastTrans = NextTran&
     PUT TAXFILE, Record!, TaxCustRec(1)


    CLOSE TaxTran
     RETURN

UpdateReal:
     GET PropTaxFile, PropRec&, PropertyRec(1)
     PropertyRec(1).PROPVALU = TBcustRec(1).LandValue
     PropertyRec(1).EXMPSENI = TBcustRec(1).BldgValue
     PropertyRec(1).PROPNOT1 = TBcustRec(1).PropDesc1
     PropertyRec(1).PROPNOT2 = TBcustRec(1).PropDesc2
     PropertyRec(1).PROPNOT3 = TBcustRec(1).PropDesc3
     PropertyRec(1).LastYrPrinted = 1999
     PUT PropTaxFile, PropRec&, PropertyRec(1)
     RETURN

END SUB

SUB OpenTaxCustFile (NumOfTaxRecs, TAXFILE)
  
  TAXFILE = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TAXFILE LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TAXFILE) / LEN(TaxCustRec(1))
  
END SUB

SUB OpenTaxPropFile (NumOfPropRecs, PropTaxFile)
  PropTaxFile = FREEFILE
  OPEN "TAXPROP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PropTaxFile LEN = LEN(PropertyRec(1))
  NumOfPropRecs = LOF(PropTaxFile) / LEN(PropertyRec(1))
END SUB

