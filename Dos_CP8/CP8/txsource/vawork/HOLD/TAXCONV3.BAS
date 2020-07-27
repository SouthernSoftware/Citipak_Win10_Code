DEFINT A-Z
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
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile

   OPEN "TBCUST.DAT" FOR RANDOM AS #11 LEN = 512
   FOR x = 1 TO LOF(11) / 512
   GET 11, x, TBCustRec(1)
   IF VAL(TBCustRec(1).Acct) = 0 THEN GOTO NextOne

   PersonalValue# = TBCustRec(1).mh + TBCustRec(1).fm + TBCustRec(1).ls + TBCustRec(1).op + TBCustRec(1).ME
   PersonalValue# = INT((PersonalValue# * 100) + .5) / 100

   RealValue# = TBCustRec(1).pland + TBCustRec(1).pbldg
   RealValue# = INT((RealValue# * 100) + .5) / 100
   
  IF RealValue# > 0 THEN
   TaxFile = FREEFILE
   OPEN "RETXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
   NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
   Record! = NumOfTaxRecs + 1
   PRINT LTRIM$(TBCustRec(1).firstname + " " + TBCustRec(1).LName)
   TName$ = LTRIM$(TBCustRec(1).firstname + " " + TBCustRec(1).LName)
   TaxCustRec(1).FNAME = UCASE$(LTRIM$(TBCustRec(1).firstname))
   TaxCustRec(1).LName = UCASE$(LTRIM$(TBCustRec(1).LName))
   TaxCustRec(1).SName = UCASE$(LTRIM$(TBCustRec(1).LName))
   TaxCustRec(1).Acct = Record!
   TaxCustRec(1).CountyAcct = 0
   TaxCustRec(1).Addr1 = UCASE$(TBCustRec(1).Addr1)
   TaxCustRec(1).Addr2 = UCASE$(TBCustRec(1).Addr2)
   TaxCustRec(1).City = UCASE$(TBCustRec(1).City)
   TaxCustRec(1).STATE = UCASE$(TBCustRec(1).STATE)
   TaxCustRec(1).Zip = UCASE$(TBCustRec(1).Zip)
   TaxCustRec(1).HPHONE = TBCustRec(1).PHONE
   TaxCustRec(1).WPHONE = ""
   TaxCustRec(1).CSSN = LEFT$(TBCustRec(1).SSN, 9)
   TaxCustRec(1).SSSN = ""
   TaxCustRec(1).Active = "Y"
   TaxCustRec(1).Interest = "Y"
   TaxCustRec(1).TaxExempt = "N"
   TaxCustRec(1).Penalty = "Y"
   TaxCustRec(1).LastTrans = 0
   TaxCustRec(1).FirstPropRec = 0
   TaxCustRec(1).FirstPersRec = 0
   TaxCustRec(1).PIN = x
   TaxCustRec(1).Deleted = 0
   TaxCustRec(1).FileVer = 8
   TaxCustRec(1).OPENDATE = Date2Num%("01-01-1998")
   PUT TaxFile, Record!, TaxCustRec(1)
   PropNumb = 1
   GOSUB UpdateReal
   GOSUB UpdateTrans
   CLOSE TaxFile
   END IF

   IF PersonalValue# > 0 THEN
    TaxFile = FREEFILE
    OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
    NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
    Record! = NumOfTaxRecs + 1
    PRINT LTRIM$(TBCustRec(1).firstname + " " + TBCustRec(1).LName)
    TName$ = LTRIM$(TBCustRec(1).firstname + " " + TBCustRec(1).LName)
    TaxCustRec(1).FNAME = UCASE$(LTRIM$(TBCustRec(1).firstname))
    TaxCustRec(1).LName = UCASE$(LTRIM$(TBCustRec(1).LName))
    TaxCustRec(1).SName = UCASE$(LTRIM$(TBCustRec(1).LName))
    TaxCustRec(1).Acct = Record!
    TaxCustRec(1).CountyAcct = 0
    TaxCustRec(1).Addr1 = UCASE$(TBCustRec(1).Addr1)
    TaxCustRec(1).Addr2 = UCASE$(TBCustRec(1).Addr2)
    TaxCustRec(1).City = UCASE$(TBCustRec(1).City)
    TaxCustRec(1).STATE = UCASE$(TBCustRec(1).STATE)
    TaxCustRec(1).Zip = TBCustRec(1).Zip
    TaxCustRec(1).HPHONE = TBCustRec(1).PHONE
    TaxCustRec(1).WPHONE = ""
    TaxCustRec(1).CSSN = LEFT$(TBCustRec(1).SSN, 9)
    TaxCustRec(1).SSSN = ""
    TaxCustRec(1).Active = "Y"
    TaxCustRec(1).Interest = "Y"
    TaxCustRec(1).TaxExempt = "N"
    TaxCustRec(1).Penalty = "Y"
    TaxCustRec(1).LastTrans = 0
    TaxCustRec(1).FirstPropRec = 0
    TaxCustRec(1).FirstPersRec = 0
    TaxCustRec(1).PIN = x
    TaxCustRec(1).Deleted = 0
    TaxCustRec(1).FileVer = 8
    TaxCustRec(1).OPENDATE = Date2Num%("01-01-1998")
    PUT TaxFile, Record!, TaxCustRec(1)
    PropNumb = 1
    GOSUB Updatepersonal
    CLOSE TaxFile
   END IF
NextOne:
   NEXT x
   CLOSE
   EXIT SUB


Updatepersonal:
   PersRec(1).PERSVAL = TBCustRec(1).op
   PersRec(1).MCVALUE = TBCustRec(1).fm
   PersRec(1).CVALUE = TBCustRec(1).ls
   PersRec(1).MTVALUE = TBCustRec(1).ME
   PersRec(1).MHVALUE = TBCustRec(1).mh
upd1:
   PersRec(1).PROPPIN = LTRIM$(STR$(PropNumb))
   PersRec(1).PROPDATE = Date2Num%("03-01-1999")
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LATELIST = "N"
   PersRec(1).DESC1 = TBCustRec(1).pdesc
   PersRec(1).DESC2 = ""
   PersRec(1).DESC3 = ""
   PersRec(1).DESC4 = ""
   PersRec(1).DESC5 = ""
   PersRec(1).CustPin = Record!
   PersRec(1).NextRec = 0
   PersRec(1).LastYrPrinted = 1998
   PersRec(1).Deleted = 0
   PersRec(1).Blank = ""
   PRecord& = LOF(PersTaxFile) / LEN(PersRec(1)) + 1
   PUT PersTaxFile, PRecord&, PersRec(1)
   RETURN

UpdateReal:
    RERecord& = LOF(PropTaxFile) / LEN(PropertyRec(1)) + 1
    PropertyRec(1).REALPIN = TBCustRec(1).PIN
    PropertyRec(1).PROPDATE = Date2Num%("03-01-99")
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).MAP = LEFT$(TBCustRec(1).MAP, 6)
    PropertyRec(1).BLOCK = MID$(TBCustRec(1).MAP, 7, 6)
    PropertyRec(1).LOTNUMB = MID$(TBCustRec(1).MAP, 13, 6)
    PropertyRec(1).LOTACRE = "A"
    PropertyRec(1).PROPSIZE = TBCustRec(1).psize
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LATELIST = "N"
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).PROPVALU = TBCustRec(1).pland
    PropertyRec(1).EXMPSENI = TBCustRec(1).pbldg
    PropertyRec(1).EXMPOTHR = 0
    PropertyRec(1).PROPNOT1 = TBCustRec(1).pdesc
    PropertyRec(1).PROPNOT2 = RIGHT$(TBCustRec(1).pdesc, 8)
    PropertyRec(1).PROPNOT3 = TBCustRec(1).SubDiv
    PropertyRec(1).Fill1 = ""
    PropertyRec(1).CustPin = Record!
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 1999
    PropertyRec(1).Deleted = 0
    PropertyRec(1).Blank = ""
    PUT PropTaxFile, RERecord&, PropertyRec(1)
    RETURN

UpdateTrans:
  IF TBCustRec(1).CTax < .01 THEN RETURN
      TransFile = FREEFILE
      OPEN "TaxTrans.dat" FOR RANDOM AS TransFile LEN = LEN(TaxTrans(1))
      NextTrans& = LOF(TransFile) / LEN(TaxTrans(1)) + 1
 
'  GET TransFile, Cnt&, TaxTrans(1)
      CustRec& = Record!
      TaxTrans(1).TaxYear = 1999
      TaxTrans(1).TranType = 1  '1=Bill 2=Payment 3=Release 4=Interest 5=Penalty 6=Collection/Ad Cost Billing
      TaxTrans(1).BillType = "R"                'R=Real P=Personal Property C=Combined (NC/GA)
      TaxTrans(1).Amount = TBCustRec(1).CTax
      TaxTrans(1).Revenue.Principle1 = TBCustRec(1).CTax
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
      TaxTrans(1).Description = "TAX BILL #" + LTRIM$(STR$(TBCustRec(1).CNotice))
      TaxTrans(1).Posted2GL = "Y"
      TaxTrans(1).LastTrans = 0
      TaxTrans(1).BelongTo = Record!
      TaxTrans(1).Padding = ""

      PUT #TransFile, NextTrans&, TaxTrans(1)

     GET #TaxFile, CustRec&, TaxCustRec(1)
        TaxCustRec(1).LastTrans = NextTrans&
     PUT #TaxFile, CustRec&, TaxCustRec(1)
    CLOSE TransFile

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

