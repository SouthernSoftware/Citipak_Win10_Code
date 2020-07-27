DEFINT A-Z
'Real Estate Conversion from Text File Information (ALBERTA VA)
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
  '$INCLUDE: 'TAXCNVRT.BI'
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PROPERTYRec(1) AS PropertyRecType
  DIM SHARED REALRec(1) AS PropertyRecType
  DIM SHARED ALBERTA(1) AS AlbertaREType
  
  STACK 5000
'  BalanceListing
  
  END

SUB BalanceListing
CLS
DIM a$(36)
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile

   CLS
   OPEN "ALBTAXRE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #11 LEN = 1027

   FOR Cnt& = 1 TO LOF(11) / 1027
   GET 11, Cnt&, ALBERTA(1)
    NType = VAL(ALBERTA(1).CLASS)
   IF NType = 7 THEN GOTO NEXTONE
   
   RECORD! = RECORD! + 1
   'Decode Name Here

    NType = VAL(ALBERTA(1).CLASS)
   IF NType = 7 THEN
    LN$ = ALBERTA(1).NM
    fm$ = ""
    ELSE
    kk = INSTR(ALBERTA(1).NM, " ")
    LN$ = LEFT$(ALBERTA(1).NM, kk - 1)
    fm$ = MID$(LTRIM$(ALBERTA(1).NM), kk + 1, LEN(ALBERTA(1).NM) - kk)
   END IF
    SN$ = LN$
    LN$ = UCASE$(LN$)
    fm$ = UCASE$(fm$)
    SN$ = UCASE$(SN$)
    
    TaxCustRec(1).FName = fm$
    TaxCustRec(1).LName = LN$
    TaxCustRec(1).SName = SN$
    
   TaxCustRec(1).acct = RECORD!
   TaxCustRec(1).Addr1 = ALBERTA(1).Add1
   TaxCustRec(1).Addr2 = ALBERTA(1).ADD2
   TaxCustRec(1).CITY = ALBERTA(1).ADD3
   TaxCustRec(1).State = ""
   TaxCustRec(1).ZIP = ALBERTA(1).ZIP
   TaxCustRec(1).HPHONE = ""
   TaxCustRec(1).WPHONE = ""
   TaxCustRec(1).CSSN = ""
   TaxCustRec(1).SSSN = ""
   TaxCustRec(1).CountyAcctString = ""
   TaxCustRec(1).CountyAcct = VAL(ALBERTA(1).acct)
   TaxCustRec(1).Active = "Y"
   TaxCustRec(1).Interest = "Y"
   IF ALBERTA(1).CLASS = "7" THEN
    TaxCustRec(1).TaxExempt = "Y"
   ELSE
    TaxCustRec(1).TaxExempt = "N"
   END IF
   TaxCustRec(1).Penalty = "Y"
   TaxCustRec(1).LastTrans = 0
   TaxCustRec(1).FirstPropRec = 0
   TaxCustRec(1).FirstPersRec = 0
   TaxCustRec(1).pin = RECORD!
   TaxCustRec(1).Deleted = 0
   TaxCustRec(1).FileVer = 8
   TaxCustRec(1).OPENDATE = Date2Num%("10-01-1999")
   PUT TaxFile, RECORD!, TaxCustRec(1)
   PropNumb = 1
   GOSUB updatereal

NEXTONE:
   NEXT Cnt&
   CLOSE
   EXIT SUB


updatereal:
    RERecord& = LOF(PropTaxFile) / LEN(REALRec(1)) + 1
    PROPERTYRec(1).REALPIN = ALBERTA(1).Map
    PROPERTYRec(1).PROPDATE = Date2Num%("01-01-98")
    PROPERTYRec(1).GISPOS = ""
    PROPERTYRec(1).Map = ""
    PROPERTYRec(1).BLOCK = ALBERTA(1).BOOK
    PROPERTYRec(1).LOTNUMB = ALBERTA(1).PAGE
    PROPERTYRec(1).LOTACRE = "A"
    PROPERTYRec(1).PROPSIZE = 0
    PROPERTYRec(1).PROPDISC = "N"
    PROPERTYRec(1).LateList = "N"
    PROPERTYRec(1).MortCode = ""
    PROPERTYRec(1).PROPVALU = VAL(ALBERTA(1).LAND)
    PROPERTYRec(1).EXMPSENI = VAL(ALBERTA(1).BLDG)
    PROPERTYRec(1).EXMPOTHR = 0
    PROPERTYRec(1).PROPNOT1 = ALBERTA(1).DESC1
    PROPERTYRec(1).PROPNOT2 = ALBERTA(1).DESC2
    PROPERTYRec(1).PROPNOT3 = ""

    PROPERTYRec(1).Fill1 = ""
    PROPERTYRec(1).CustPin = RECORD!
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

