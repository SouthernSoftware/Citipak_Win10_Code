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
  DIM SHARED tbcustrecs(1) AS TBCustTypeShift
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
   GET 11, x, tbcustrecs(1)
   PRINT tbcustrecs(1).CTax


   PersonalValue# = TBCustRec(1).MH + TBCustRec(1).fm + TBCustRec(1).ls + TBCustRec(1).op + TBCustRec(1).me
   PersonalValue# = INT((PersonalValue# * 100) + .5) / 100

   RealValue# = TBCustRec(1).pland + TBCustRec(1).pbldg
   RealValue# = INT((RealValue# * 100) + .5) / 100
   IF PersonalValue# = 0 AND RealValue# = 0 AND tbcustrecs(1).CTax = 0 THEN GOTO Nextone
   CLOSE TaxFile
   RealFlag = 0: PersFlag = 0
   IF RealValue# > 0 THEN
    OPEN "RETXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
    NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
    RealFlag = 1
    ELSE
    OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
    NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
    PersFlag = 1
   END IF
   

    Record! = NumOfTaxRecs + 1
    PRINT LTRIM$(TBCustRec(1).firstname + " " + TBCustRec(1).LName)
    TName$ = LTRIM$(TBCustRec(1).firstname + " " + TBCustRec(1).LName)
    IF LEN(TName$) > 0 THEN
    
    TaxCustRec(1).FNAME = RTRIM$(LTRIM$(TBCustRec(1).firstname))
    FOR kk = 1 TO LEN(TaxCustRec(1).FNAME)
     IF MID$(TaxCustRec(1).FNAME, kk, 1) = "." THEN MID$(TaxCustRec(1).FNAME, kk, 1) = " "
    NEXT kk
    TaxCustRec(1).LName = RTRIM$(LTRIM$(TBCustRec(1).LName))
    FOR kk = 1 TO LEN(TaxCustRec(1).LName)
     IF MID$(TaxCustRec(1).LName, kk, 1) = "." THEN MID$(TaxCustRec(1).LName, kk, 1) = " "
    NEXT kk
    TaxCustRec(1).SName = LTRIM$(TBCustRec(1).LName)
    FOR kk = 1 TO LEN(TaxCustRec(1).SName)
     IF MID$(TaxCustRec(1).SName, kk, 1) = "." THEN MID$(TaxCustRec(1).SName, kk, 1) = " "
    NEXT kk
    
   TaxCustRec(1).acct = Record!
   TaxCustRec(1).CountyAcct = 0
   TaxCustRec(1).Addr1 = RTRIM$(TBCustRec(1).Addr1)
    FOR kk = 1 TO LEN(TaxCustRec(1).Addr1)
     IF MID$(TaxCustRec(1).Addr1, kk, 1) = "." THEN MID$(TaxCustRec(1).Addr1, kk, 1) = " "
    NEXT kk
   TaxCustRec(1).Addr2 = RTRIM$(TBCustRec(1).Addr2)
    FOR kk = 1 TO LEN(TaxCustRec(1).Addr2)
     IF MID$(TaxCustRec(1).Addr2, kk, 1) = "." THEN MID$(TaxCustRec(1).Addr2, kk, 1) = " "
    NEXT kk
   TaxCustRec(1).City = RTRIM$(TBCustRec(1).City)
    FOR kk = 1 TO LEN(TaxCustRec(1).City)
     IF MID$(TaxCustRec(1).City, kk, 1) = "." THEN MID$(TaxCustRec(1).City, kk, 1) = " "
    NEXT kk
   TaxCustRec(1).STATE = RTRIM$(TBCustRec(1).STATE)
    FOR kk = 1 TO LEN(TaxCustRec(1).STATE)
     IF MID$(TaxCustRec(1).STATE, kk, 1) = "." THEN MID$(TaxCustRec(1).STATE, kk, 1) = " "
    NEXT kk
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
   TaxCustRec(1).PIN = Record!
   TaxCustRec(1).Employer = ""
   TaxCustRec(1).Bankrupt = "N"
   TaxCustRec(1).TownShip = ""
   TaxCustRec(1).CountyAcctString = LEFT$(TBCustRec(1).SSN, 9)
   TaxCustRec(1).CountyAcct = 0

   TaxCustRec(1).Deleted = 0
   TaxCustRec(1).FileVer = 8
   TaxCustRec(1).OPENDATE = Date2Num%("10-01-1999")
   PUT TaxFile, Record!, TaxCustRec(1)
   PropNumb = 1

   IF PersFlag = 1 THEN
    GOSUB updatepersonal
    GOSUB UpdateTrans
   END IF

   IF RealFlag = 1 THEN
    GOSUB UpdateReal
    GOSUB UpdateTrans
   END IF

Nextone:
   CLOSE TaxFile
   END IF
   NEXT x
   CLOSE
   EXIT SUB


updatepersonal:
   PersRec(1).PERSVAL = TBCustRec(1).op
   PersRec(1).MCVALUE = TBCustRec(1).fm
   PersRec(1).CVALUE = TBCustRec(1).ls
   PersRec(1).MTVALUE = TBCustRec(1).me
   PersRec(1).MHVALUE = TBCustRec(1).MH
upd1:
   PersRec(1).PROPPIN = LTRIM$(STR$(PropNumb))
   PersRec(1).PROPDATE = Date2Num%("03-01-1999")
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LATELIST = "N"
   PersRec(1).DESC1 = TBCustRec(1).pdesc
    FOR kk = 1 TO LEN(PersRec(1).DESC1)
     IF MID$(PersRec(1).DESC1, kk, 1) = "." THEN MID$(PersRec(1).DESC1, kk, 1) = " "
    NEXT kk
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
    FOR kk = 1 TO LEN(PropertyRec(1).PROPNOT1)
     IF MID$(PropertyRec(1).PROPNOT1, kk, 1) = "." THEN MID$(PropertyRec(1).PROPNOT1, kk, 1) = " "
    NEXT kk
    PropertyRec(1).PROPNOT2 = RIGHT$(TBCustRec(1).pdesc, 8)
    FOR kk = 1 TO LEN(PropertyRec(1).PROPNOT2)
     IF MID$(PropertyRec(1).PROPNOT2, kk, 1) = "." THEN MID$(PropertyRec(1).PROPNOT2, kk, 1) = " "
    NEXT kk
    PropertyRec(1).PROPNOT3 = TBCustRec(1).SubDiv
    FOR kk = 1 TO LEN(PropertyRec(1).PROPNOT3)
     IF MID$(PropertyRec(1).PROPNOT3, kk, 1) = "." THEN MID$(PropertyRec(1).PROPNOT3, kk, 1) = " "
    NEXT kk
    PropertyRec(1).Fill1 = ""
    PropertyRec(1).CustPin = Record!
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 1999
    PropertyRec(1).Deleted = 0
    PropertyRec(1).Blank = ""
    PUT PropTaxFile, RERecord&, PropertyRec(1)
    RETURN

UpdateTrans:
  IF tbcustrecs(1).CTax < .01 THEN RETURN
      TransFile = FREEFILE
      OPEN "TaxTrans.dat" FOR RANDOM AS TransFile LEN = LEN(TaxTrans(1))
      NextTrans& = LOF(TransFile) / LEN(TaxTrans(1)) + 1
      CustRec& = Record!
      TaxTrans(1).TaxYear = 1999
      TaxTrans(1).TransDate = Date2Num("10-01-1999")
      TaxTrans(1).TranType = 1                  '1=Bill 2=Payment 3=Release 4=Interest 5=Penalty 6=Collection/Ad Cost Billing
      TaxTrans(1).BillType = "R"                'R=Real P=Personal Property C=Combined (NC/GA)
      TaxTrans(1).Amount = tbcustrecs(1).CTax
      TaxTrans(1).Revenue.Principle1 = tbcustrecs(1).CTax
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
      TaxTrans(1).Description = "TAX BILL #" + LTRIM$(STR$(tbcustrecs(1).CNotice))
      TaxTrans(1).Posted2GL = "Y"
      TaxTrans(1).LastTrans = 0
      TaxTrans(1).BelongTo = Record!
      TaxTrans(1).Padding = ""

      PUT #TransFile, NextTrans&, TaxTrans(1)

     GET #TaxFile, Record!, TaxCustRec(1)
        TaxCustRec(1).LastTrans = NextTrans&
     PUT #TaxFile, Record!, TaxCustRec(1)
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

