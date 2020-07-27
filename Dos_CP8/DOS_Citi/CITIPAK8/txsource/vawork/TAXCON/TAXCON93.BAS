DEFINT A-Z
'Real Estate Conversion Dumfries VA
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
  '$INCLUDE: 'tax.bi'
  '$INCLUDE: 'PROPAbst.BI'
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED REALRec(1) AS PropertyRecType
  DIM SHARED TBCust(1) AS TaxCustomerRecType
  
  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS

  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile

   CLS

 OPEN "tbcust.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #11 LEN = LEN(TBCust(1))
   
 FOR Cnt! = 1 TO LOF(11) / LEN(TBCust(1))
 
 GET 11, Cnt!, TBCust(1)
   
  IF TBCust(1).LandValue + TBCust(1).BldgValue > 0 AND ASC(LEFT$(TBCust(1).Account, 1)) > 32 THEN
   t# = t# + TBCust(1).LandValue + TBCust(1).BldgValue
   TaxCustRec(1).Acct = Cnt!
   TaxCustRec(1).FNAME = ""
   TaxCustRec(1).LName = TBCust(1).NameLine1
   TaxCustRec(1).SName = TBCust(1).NameLine1
   TaxCustRec(1).Addr1 = TBCust(1).Address1
   TaxCustRec(1).Addr2 = ""
   TaxCustRec(1).CITY = TBCust(1).Address2
   TaxCustRec(1).State = TBCust(1).State
   TaxCustRec(1).Zip = TBCust(1).ZipCode
   TaxCustRec(1).HPHONE = ""
   TaxCustRec(1).WPHONE = ""
   TaxCustRec(1).CSSN = TBCust(1).SocSec
   TaxCustRec(1).SSSN = ""
   TaxCustRec(1).CountyAcctString = ""
   TaxCustRec(1).CountyAcct = 0
   TaxCustRec(1).Active = "Y"
   TaxCustRec(1).Interest = "Y"
   TaxCustRec(1).TaxExempt = "N"
   TaxCustRec(1).Penalty = "Y"
   TaxCustRec(1).LastTrans = 0
   TaxCustRec(1).FirstPropRec = 0
   TaxCustRec(1).FirstPersRec = 0
   TaxCustRec(1).PIN = Cnt!
   TaxCustRec(1).Deleted = 0
   TaxCustRec(1).FileVer = 8
   TaxCustRec(1).OPENDATE = Date2Num%("04-01-2003")
   PUT TaxFile, Cnt!, TaxCustRec(1)
    PropNumb = 1
    GOSUB updatereal
  ELSE
    TaxCustRec(1).FNAME = ""
    TaxCustRec(1).LName = ""
    TaxCustRec(1).SName = ""
    TaxCustRec(1).Acct = 0
    TaxCustRec(1).Addr1 = TBCust(1).Address1
    TaxCustRec(1).Addr2 = ""
    TaxCustRec(1).CITY = ""
    TaxCustRec(1).State = ""
    TaxCustRec(1).Zip = ""
    TaxCustRec(1).HPHONE = ""
    TaxCustRec(1).WPHONE = ""
    TaxCustRec(1).CSSN = ""
    TaxCustRec(1).SSSN = ""
    TaxCustRec(1).CountyAcctString = ""
    TaxCustRec(1).CountyAcct = 0
    TaxCustRec(1).Active = "Y"
    TaxCustRec(1).Interest = "Y"
    TaxCustRec(1).TaxExempt = "N"
    TaxCustRec(1).Penalty = "Y"
    TaxCustRec(1).LastTrans = 0
    TaxCustRec(1).FirstPropRec = 0
    TaxCustRec(1).FirstPersRec = 0
    TaxCustRec(1).PIN = 0
    TaxCustRec(1).Deleted = -1
    TaxCustRec(1).OPENDATE = Date2Num%("04-01-2003")
    PUT TaxFile, Cnt!, TaxCustRec(1)

  END IF

NEXTONE:
   NEXT Cnt!

   CLOSE
   STOP
   EXIT SUB


updatereal:
    RERecord& = LOF(PropTaxFile) / LEN(REALRec(1)) + 1
    PropertyRec(1).RealPin = TBCust(1).MapNumber
    PropertyRec(1).PropDate = Date2Num%("04/01/2003")
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).Map = TBCust(1).UseType
    PropertyRec(1).Block = ""
    PropertyRec(1).LotNumb = ""
    PropertyRec(1).LOTACRE = "A"
    PropertyRec(1).PROPSIZE = TBCust(1).Acres
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LATELIST = "N"
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).PROPVALU = TBCust(1).LandValue
    PropertyRec(1).EXMPSENI = TBCust(1).BldgValue
    PropertyRec(1).EXMPOTHR = 0
    PropertyRec(1).PROPNOT1 = TBCust(1).PropDesc1
    PropertyRec(1).PROPNOT2 = TBCust(1).PropDesc2
    PropertyRec(1).PROPNOT3 = TBCust(1).PropDesc3
    PropertyRec(1).StormFee = TBCust(1).StormFee
    PropertyRec(1).CustPin = Cnt!
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 2001
    PropertyRec(1).Deleted = 0
    PropertyRec(1).Blank = ""
    PUT PropTaxFile, RERecord&, PropertyRec(1)
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
  OPEN "TAXPROP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PropTaxFile LEN = LEN(PropertyRec(1))
  NumOfPropRecs = LOF(PropTaxFile) / LEN(PropertyRec(1))
END SUB

