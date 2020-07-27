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
  '$INCLUDE: 'TaxConv.BI'
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

   RealValue# = TBCustRec(1).pland + TBCustRec(1).pbldg
   RealValue# = INT((RealValue# * 100) + .5) / 100
   
  IF RealValue# > 0 THEN
   TaxFile = FREEFILE
   OPEN "RETXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
   NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
   Record! = NumOfTaxRecs + 1
   PRINT LTRIM$(TBCustRec(1).FIRSTNAME + " " + TBCustRec(1).LNAME)
   TName$ = LTRIM$(TBCustRec(1).FIRSTNAME + " " + TBCustRec(1).LNAME)
   TaxCustRec(1).FName = UCASE$(LTRIM$(TBCustRec(1).FIRSTNAME))
   TaxCustRec(1).LNAME = UCASE$(LTRIM$(TBCustRec(1).LNAME))
   TaxCustRec(1).SName = UCASE$(LTRIM$(TBCustRec(1).LNAME))
   TaxCustRec(1).Acct = Record!
   TaxCustRec(1).CountyAcct = 0
   TaxCustRec(1).Addr1 = UCASE$(TBCustRec(1).Addr1)
   TaxCustRec(1).Addr2 = UCASE$(TBCustRec(1).Addr2)
   TaxCustRec(1).City = UCASE$(TBCustRec(1).City)
   TaxCustRec(1).State = UCASE$(TBCustRec(1).State)
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
   TaxCustRec(1).OPENDATE = Date2Num%("06-01-2003")
   PUT TaxFile, Record!, TaxCustRec(1)
   PropNumb = 1
   GOSUB UpdateReal
   CLOSE TaxFile
   END IF
NextOne:
   NEXT x
   CLOSE
   EXIT SUB


UpdateReal:
    RERecord& = LOF(PropTaxFile) / LEN(PropertyRec(1)) + 1
    PropertyRec(1).REALPIN = TBCustRec(1).PMAP
    PropertyRec(1).PROPDATE = Date2Num%("06-21-2003")
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).MAP = LEFT$(TBCustRec(1).PMAP, 6)
    PropertyRec(1).BLOCK = MID$(TBCustRec(1).PMAP, 7, 6)
    PropertyRec(1).LOTNUMB = MID$(TBCustRec(1).PMAP, 13, 6)
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
    PropertyRec(1).PROPNOT3 = TBCustRec(1).Pdesc3
    PropertyRec(1).CustPin = Record!
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 1999
    PropertyRec(1).Deleted = 0
    PropertyRec(1).Blank = ""
    PUT PropTaxFile, RERecord&, PropertyRec(1)
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

