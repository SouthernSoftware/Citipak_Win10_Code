DEFINT A-Z
'Real Estate Conversion from Text File Information ELKTON VA 2001
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
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED REALRec(1) AS PropertyRecType
  DIM SHARED RAMAST(1) AS REALTYPE
  
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
   OPEN "ELKRE.TXT" FOR INPUT AS #10
1  CLS
   FOR c = 1 TO 11
   INPUT #10, a$(c)
   PRINT a$(c)
   NEXT c
   SLEEP 1

   'Search for Property
   FOR Snt& = 1 TO NumOfPropRecs
   GET PropTaxFile, Snt&, PropertyRec(1)
    IF RTRIM$(PropertyRec(1).PROPNOT1) = RTRIM$(a$(6)) THEN
      GET TaxFile, PropertyRec(1).CustPin, TaxCustRec(1)
      
      EXIT FOR
    END IF

    
   NEXT Snt&
   
      GOTO 1
   STOP


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
   TaxCustRec(1).Pin = Record!
   TaxCustRec(1).Deleted = 0
   TaxCustRec(1).FileVer = 8
   TaxCustRec(1).OPENDATE = Date2Num%("07-01-1999")
   PUT TaxFile, Record!, TaxCustRec(1)
   PropNumb = 1
   GOSUB updatereal

NEXTONE:
   
   CLOSE
   EXIT SUB


updatereal:
    RERecord& = LOF(PropTaxFile) / LEN(REALRec(1)) + 1
    PropertyRec(1).REALPIN = RAMAST(1).MAP
    PropertyRec(1).PROPDATE = Date2Num%("07-01-99")
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).MAP = ""
    PropertyRec(1).BLOCK = ""
    PropertyRec(1).LOTNUMB = ""
    PropertyRec(1).LOTACRE = "A"
    PropertyRec(1).PROPSIZE = VAL(RAMAST(1).ACRE) / 1000
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LATELIST = "N"
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).PROPVALU = VAL(RAMAST(1).LANDVALUE)
    PropertyRec(1).EXMPSENI = VAL(RAMAST(1).BLDGVALUE)
    PropertyRec(1).EXMPOTHR = VAL(RAMAST(1).EXEMPT)
    PropertyRec(1).PROPNOT1 = RAMAST(1).DESCRIPTION
    PropertyRec(1).PROPNOT2 = "CLASS " + RAMAST(1).CLASS
    PropertyRec(1).PROPNOT3 = RAMAST(1).REMARKS

    PropertyRec(1).Fill1 = ""
    PropertyRec(1).CustPin = Record!
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 1998
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

