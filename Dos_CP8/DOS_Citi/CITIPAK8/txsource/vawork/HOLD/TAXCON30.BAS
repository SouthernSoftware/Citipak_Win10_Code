DEFINT A-Z
'Real Estate Conversion from Text File Information ST PAUL VA
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
  '$INCLUDE: 'TAXCONV.BI'
  
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PROPERTYRec(1) AS PropertyRecType
  DIM SHARED REALRec(1) AS PropertyRecType
  DIM SHARED RAMAST(1) AS REALTYPE
  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
DIM A$(36)
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  Record! = NumOfTaxRecs

   CLS
   OPEN "REDATA.TXT" FOR RANDOM ACCESS READ WRITE SHARED AS #11 LEN = LEN(RAMAST(1))
   FOR CNT& = 1 TO LOF(11) / LEN(RAMAST(1))
   GET #11, CNT&, RAMAST(1)
 
 
    Record! = Record! + 1
    TaxCustRec(1).FName = RTRIM$(RAMAST(1).FIRSTNAME) + " " + RAMAST(1).INITIAL
    TaxCustRec(1).LName = RAMAST(1).LASTNAME
    TaxCustRec(1).SName = RAMAST(1).LASTNAME
    
   TaxCustRec(1).Acct = Record!
   TaxCustRec(1).Addr1 = RAMAST(1).ADDRESS1
   TaxCustRec(1).Addr2 = RAMAST(1).ADDRESS2
   TaxCustRec(1).CITY = RAMAST(1).CITY
   TaxCustRec(1).STATE = RAMAST(1).STATE
   TaxCustRec(1).ZIP = RAMAST(1).ZIPCODE
   TaxCustRec(1).HPHONE = ""
   TaxCustRec(1).WPHONE = ""
   TaxCustRec(1).CSSN = RAMAST(1).TAXID
   TaxCustRec(1).SSSN = ""
   TaxCustRec(1).CountyAcctString = RAMAST(1).TAXID
   TaxCustRec(1).CountyAcct = VAL(RAMAST(1).TAXID)
   TaxCustRec(1).Active = "Y"
   TaxCustRec(1).interest = "Y"
   TaxCustRec(1).TaxExempt = "N"
   TaxCustRec(1).Penalty = "Y"
   TaxCustRec(1).LastTrans = 0
   TaxCustRec(1).FirstPropRec = 0
   TaxCustRec(1).FirstPersRec = 0
   TaxCustRec(1).PIN = Record!
   TaxCustRec(1).Deleted = 0
   TaxCustRec(1).FileVer = 8
   TaxCustRec(1).OPENDATE = Date2Num%("07-01-1999")
   PUT TaxFile, Record!, TaxCustRec(1)
   PropNumb = 1
   GOSUB updatereal

NEXTONE:
   NEXT CNT&
   CLOSE
   EXIT SUB


updatereal:
    RERecord& = LOF(PropTaxFile) / LEN(REALRec(1)) + 1
    PROPERTYRec(1).REALPIN = RAMAST(1).MAP
    PROPERTYRec(1).PROPDATE = Date2Num%("07-01-99")
    PROPERTYRec(1).GISPOS = ""
    PROPERTYRec(1).MAP = ""
    PROPERTYRec(1).BLOCK = ""
    PROPERTYRec(1).LOTNUMB = ""
    PROPERTYRec(1).LOTACRE = "A"
    PROPERTYRec(1).PROPSIZE = VAL(RAMAST(1).ACRE) / 1000
    PROPERTYRec(1).PROPDISC = "N"
    PROPERTYRec(1).LATELIST = "N"
    PROPERTYRec(1).MORTCODE = ""
    PROPERTYRec(1).PROPVALU = VAL(RAMAST(1).LANDVALUE)
    PROPERTYRec(1).EXMPSENI = VAL(RAMAST(1).BLDGVALUE)
    PROPERTYRec(1).EXMPOTHR = VAL(RAMAST(1).EXEMPT)
    PROPERTYRec(1).PROPNOT1 = RAMAST(1).DESCRIPTION
    PROPERTYRec(1).PROPNOT2 = "CLASS " + RAMAST(1).CLASS
    PROPERTYRec(1).PROPNOT3 = RAMAST(1).REMARKS

    PROPERTYRec(1).Fill1 = ""
    PROPERTYRec(1).CustPin = Record!
    PROPERTYRec(1).NextRec = 0
    PROPERTYRec(1).LastYrPrinted = 1998
    PROPERTYRec(1).Deleted = 0
    PROPERTYRec(1).Blank = ""


   
   PUT PropTaxFile, RERecord&, PROPERTYRec(1)
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

