DEFINT A-Z
'Real Estate Conversion from Text File Information (Warsaw VA)
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
  DIM SHARED Warsaw(1) AS WarsawREType
  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
DIM a$(36)
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile

   CLS
   OPEN "WARSAW.ASC" FOR RANDOM ACCESS READ WRITE SHARED AS #11 LEN = 332

   FOR Cnt& = 1 TO LOF(11) / 332
   GET 11, Cnt&, Warsaw(1)

   'Decode Name Here
   NType = VAL(Warsaw(1).Class)
   IF NType = 4 THEN
    LN$ = Warsaw(1).Name
    fm$ = ""
    ELSE
    kk = INSTR(Warsaw(1).Name, " ")
    LN$ = LEFT$(Warsaw(1).Name, kk - 1)
    fm$ = MID$(LTRIM$(Warsaw(1).Name), kk + 1, LEN(Warsaw(1).Name) - kk)
   END IF
    SN$ = LN$
    LN$ = UCASE$(LN$)
    fm$ = UCASE$(fm$)
    SN$ = UCASE$(SN$)
    Record! = Record! + 1
    TaxCustRec(1).FName = fm$
    TaxCustRec(1).LName = LN$
    TaxCustRec(1).SName = SN$
    
   TaxCustRec(1).Acct = Record!
   TaxCustRec(1).Addr1 = Warsaw(1).Add1
   TaxCustRec(1).Addr2 = Warsaw(1).Add2
   TaxCustRec(1).City = Warsaw(1).City
   TaxCustRec(1).State = Warsaw(1).State
   TaxCustRec(1).Zip = Warsaw(1).Zip
   TaxCustRec(1).HPHONE = ""
   TaxCustRec(1).WPHONE = ""
   TaxCustRec(1).CSSN = ""
   TaxCustRec(1).SSSN = ""
   TaxCustRec(1).CountyAcctString = Warsaw(1).Parcel
   TaxCustRec(1).CountyAcct = VAL(a$(1))
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
   TaxCustRec(1).OPENDATE = Date2Num%("10-01-1999")
   PUT TaxFile, Record!, TaxCustRec(1)
   PropNumb = 1
   GOSUB updatereal

NEXTONE:
   NEXT Cnt&
   CLOSE
   EXIT SUB


updatereal:
    RERecord& = LOF(PropTaxFile) / LEN(REALRec(1)) + 1
    PROPERTYRec(1).REALPIN = Warsaw(1).Parcel
    PROPERTYRec(1).PROPDATE = Date2Num%("01-01-98")
    PROPERTYRec(1).GISPOS = ""
    PROPERTYRec(1).MAP = ""
    PROPERTYRec(1).BLOCK = ""
    PROPERTYRec(1).LOTNUMB = ""
    PROPERTYRec(1).LOTACRE = "A"
    PROPERTYRec(1).PROPSIZE = 0
    PROPERTYRec(1).PROPDISC = "N"
    PROPERTYRec(1).LATELIST = "N"
    PROPERTYRec(1).MORTCODE = ""
    PROPERTYRec(1).PROPVALU = VAL(Warsaw(1).LandVal)
    PROPERTYRec(1).EXMPSENI = VAL(Warsaw(1).BldgVal)
    PROPERTYRec(1).EXMPOTHR = VAL(Warsaw(1).LandEx)
    PROPERTYRec(1).PROPNOT1 = Warsaw(1).Desc1
    PROPERTYRec(1).PROPNOT2 = Warsaw(1).Desc2
    PROPERTYRec(1).PROPNOT3 = "Size " + Warsaw(1).Acres

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

