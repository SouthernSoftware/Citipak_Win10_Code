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
  DIM SHARED PROPERTYRec(1) AS PropertyRecType
  DIM SHARED REALRec(1) AS PropertyRecType
  DIM SHARED RAMAST(1) AS RAMAST
  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  'OpenTaxPersFile NumOfPersRecs, PersTaxFile

   OPEN "RAMAST" FOR RANDOM AS #11 LEN = LEN(RAMAST(1))
   FOR x = 1 TO LOF(11) / LEN(RAMAST(1))
   GET 11, x, RAMAST(1)


 

   kk = INSTR(RAMAST(1).NAME, ",")

   IF kk = 0 THEN
    TaxCustRec(1).FNAME = ""
    TaxCustRec(1).LName = RAMAST(1).NAME
    TaxCustRec(1).SName = RAMAST(1).NAME
   ELSE
    Nme2$ = LTRIM$(RIGHT$(RAMAST(1).NAME, LEN(RAMAST(1).NAME) - kk))
    Nme1$ = LEFT$(RAMAST(1).NAME, kk - 1)
    IF LEFT$(Nme2$, 3) = "INC" THEN
     TaxCustRec(1).FNAME = ""
     TaxCustRec(1).LName = RAMAST(1).NAME
     TaxCustRec(1).SName = RAMAST(1).NAME
    ELSE
    TaxCustRec(1).FNAME = Nme2$
    TaxCustRec(1).LName = Nme1$
    TaxCustRec(1).SName = Nme1$

    END IF
   END IF
 
 
   Record! = Record! + 1
   ZIPCODE$ = LTRIM$(RAMAST(1).ZIP)
   IF LEN(ZIPCODE$) = 4 THEN ZIPCODE$ = "0" + ZIPCODE$
   TaxCustRec(1).Acct = Record!
   TaxCustRec(1).Addr1 = RAMAST(1).ADD1
   TaxCustRec(1).Addr2 = RAMAST(1).ADD2
   TaxCustRec(1).CITY = RAMAST(1).CITY
   TaxCustRec(1).STATE = RAMAST(1).STATE
   TaxCustRec(1).ZIP = RAMAST(1).ZIP
   TaxCustRec(1).HPHONE = ""
   TaxCustRec(1).WPHONE = ""

   

   TaxCustRec(1).CSSN = RAMAST(1).SSN
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
   GOSUB updatepersonal

NEXTONE:
   NEXT x
   CLOSE
   EXIT SUB


updatepersonal:
    RERecord& = LOF(PropTaxFile) / LEN(REALRec(1)) + 1
    PROPERTYRec(1).REALPIN = RAMAST(1).MAP
    PROPERTYRec(1).PROPDATE = Date2Num%("01-01-98")
    PROPERTYRec(1).GISPOS = ""
    PROPERTYRec(1).MAP = ""
    PROPERTYRec(1).BLOCK = ""
    PROPERTYRec(1).LOTNUMB = ""
    PROPERTYRec(1).LOTACRE = "A"
    PROPERTYRec(1).PROPSIZE = VAL(RAMAST(1).ACRES)
    PROPERTYRec(1).PROPDISC = "N"
    PROPERTYRec(1).LATELIST = "N"
    PROPERTYRec(1).MORTCODE = ""
    PROPERTYRec(1).PROPVALU = VAL(RAMAST(1).LANDVAL)
    PROPERTYRec(1).EXMPSENI = VAL(RAMAST(1).BLDGVAL)
    PROPERTYRec(1).EXMPOTHR = 0
    PROPERTYRec(1).PROPNOT1 = RAMAST(1).DESC1
    PROPERTYRec(1).PROPNOT2 = RAMAST(1).DESC2
    PROPERTYRec(1).PROPNOT3 = ""

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

