DEFINT A-Z
'Real Estate Conversion from Text File Information
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
  DIM SHARED TaxTrans(1) AS TaxTransactionType
  DIM SHARED PROPERTYRec(1) AS PropertyRecType
  DIM SHARED REALRec(1) AS PropertyRecType
  DIM SHARED RAMAST(1) AS RAMAST
  DIM SHARED Elk(1) AS ElkType
  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
DIM a$(36)
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  Record! = NumOfTaxRecs

   CLS
   OPEN "ELKRE.TXT" FOR RANDOM ACCESS READ AS #11 LEN = 341

   FOR Cnt& = 1 TO LOF(11) / 341
   GET 11, Cnt&, Elk(1)
   PRINT

  
   'Decode Name Here
   kk = INSTR(Elk(1).nam, " ")
   IF kk = 0 THEN
    LN$ = Elk(1).nam
    fm$ = ""
    ELSE
    LN$ = LEFT$(Elk(1).nam, kk - 1)
    fm$ = LTRIM$(MID$(Elk(1).nam, kk + 1, LEN(Elk(1).nam) - kk))
   END IF
    SN$ = LN$
    LN$ = UCASE$(LN$)
    fm$ = UCASE$(fm$)
    SN$ = UCASE$(SN$)

    'For CityState Find 2nd Blk and assume its it
    kk1 = INSTR(Elk(1).City, " ")
    kk2 = INSTR(kk1 + 1, Elk(1).City, " ")
    City$ = LEFT$(Elk(1).City, kk2 - 1)


    Record! = Record! + 1
    TaxCustRec(1).FName = fm$
    TaxCustRec(1).LName = LN$
    TaxCustRec(1).SName = SN$
    
   TaxCustRec(1).Acct = Record!
   TaxCustRec(1).Addr1 = UCASE$(Elk(1).Add1)
   TaxCustRec(1).Addr2 = UCASE$(Elk(1).Add2)
   TaxCustRec(1).City = UCASE$(City$)
   TaxCustRec(1).State = ""
   TaxCustRec(1).Zip = Elk(1).Zip
   TaxCustRec(1).HPHONE = ""
   TaxCustRec(1).WPHONE = ""
   TaxCustRec(1).CSSN = ""
   TaxCustRec(1).SSSN = ""
   TaxCustRec(1).CountyAcctString = Elk(1).MAP
   TaxCustRec(1).CountyAcct = 0
   TaxCustRec(1).Active = "Y"
   TaxCustRec(1).Interest = "Y"
   TaxCustRec(1).TaxExempt = "N"
   TaxCustRec(1).Penalty = "Y"
   TaxCustRec(1).LastTrans = 0
   TaxCustRec(1).FirstPropRec = 0
   TaxCustRec(1).FirstPersRec = 0
   TaxCustRec(1).PIN = Record!
   TaxCustRec(1).Deleted = 0
   TaxCustRec(1).FileVer = 8
   TaxCustRec(1).OPENDATE = Date2Num%("01-01-1999")
   PUT TaxFile, Record!, TaxCustRec(1)
   PropNumb = 1
   GOSUB updatereal
   GOSUB UpdateTransFile

NEXTONE:
   NEXT Cnt&
   CLOSE
   EXIT SUB


updatereal:
    RERecord& = LOF(PropTaxFile) / LEN(REALRec(1)) + 1
    PROPERTYRec(1).REALPIN = Elk(1).MAP + LTRIM$(Elk(1).Deed)
    PROPERTYRec(1).PROPDATE = Date2Num%("01-01-98")
    PROPERTYRec(1).GISPOS = ""
    PROPERTYRec(1).MAP = Elk(1).MAP
    PROPERTYRec(1).BLOCK = ""
    PROPERTYRec(1).LOTNUMB = ""
    PROPERTYRec(1).LOTACRE = "A"
    PROPERTYRec(1).PROPSIZE = VAL(Elk(1).Siz) / 100
    PROPERTYRec(1).PROPDISC = "N"
    PROPERTYRec(1).LATELIST = "N"
    PROPERTYRec(1).MORTCODE = ""
    PROPERTYRec(1).PROPVALU = VAL(Elk(1).LandVal)
    PROPERTYRec(1).EXMPSENI = VAL(Elk(1).BldgVal)
    PROPERTYRec(1).EXMPOTHR = VAL(Elk(1).DisVal)
    PROPERTYRec(1).PROPNOT1 = Elk(1).DESC1
    PROPERTYRec(1).PROPNOT2 = Elk(1).DESC2
    PROPERTYRec(1).PROPNOT3 = Elk(1).DESC3
    PROPERTYRec(1).Fill1 = ""
    PROPERTYRec(1).CustPin = Record!
    PROPERTYRec(1).NextRec = 0
    PROPERTYRec(1).LastYrPrinted = 1999
    PROPERTYRec(1).Deleted = 0
    PROPERTYRec(1).Blank = ""
    PUT PropTaxFile, RERecord&, PROPERTYRec(1)
    RETURN

UpdateTransFile:
  TransFile = FREEFILE
  OPEN "TaxTrans.dat" FOR RANDOM AS TransFile LEN = LEN(TaxTrans(1))
  NextRec& = LOF(TransFile) / LEN(TaxTrans(1)) + 1
      BillNo = BillNo + 1
      TaxTrans(1).TaxYear = 1999
      TaxTrans(1).TranType = 1  '1=Bill 2=Payment 3=Release 4=Interest 5=Penalty 6=Collection/Ad Cost Billing
      TaxTrans(1).BillType = "R"                'R=Real P=Personal Property C=Combined (NC/GA)
      TaxTrans(1).Amount = VAL(Elk(1).TaxAmt) / 100
      TaxTrans(1).Revenue.Principle1 = VAL(Elk(1).TaxAmt) / 100
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
      TaxTrans(1).Description = "Tax Bill #" + STR$(BillNo)
      TaxTrans(1).Posted2GL = "N"
      TaxTrans(1).CustomerRec = Record!
      TaxTrans(1).LastTrans = TaxCustRec(1).LastTrans
      TaxTrans(1).BelongTo = 0
      TaxTrans(1).Padding = ""
      PUT #TransFile, NextRec&, TaxTrans(1)

      GET #TaxFile, Record!, TaxCustRec(1)
       TaxCustRec(1).LastTrans = NextRec&
      PUT #TaxFile, Record!, TaxCustRec(1)

      CLOSE TransFile

      RETURN



 
 END SUB

SUB OpenTaxCustFile (NumOfTaxRecs, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "RETXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
END SUB

SUB OpenTaxPersFile (NumOfPersRecs, REALTaxFile)
  REALTaxFile = FREEFILE
  OPEN "TAXREAL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #REALTaxFile LEN = LEN(REALRec(1))
  NumOfPersRecs = LOF(REALTaxFile) / LEN(REALRec(1))
  
END SUB

SUB OpenTaxPropFile (NumOfPropRecs, PropTaxFile)
  PropTaxFile = FREEFILE
  OPEN "TAXPROP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PropTaxFile LEN = LEN(PROPERTYRec(1))
  NumOfPropRecs = LOF(PropTaxFile) / LEN(PROPERTYRec(1))
END SUB

