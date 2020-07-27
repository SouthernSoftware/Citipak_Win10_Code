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
  DIM SHARED ChilRec(1) AS chilrealtype
  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  
CLS
   OPEN "CHILRE.TXT" FOR RANDOM ACCESS READ AS #11 LEN = 324
FOR Cnt& = 1 TO LOF(11) / 332
GET 11, Cnt&, ChilRec(1)


PRINT ChilRec(1).CtyAcct
PRINT ChilRec(1).NME
PRINT ChilRec(1).Add1
PRINT ChilRec(1).Add2
PRINT ChilRec(1).Add3
PRINT ChilRec(1).Zip
PRINT ChilRec(1).MAP
PRINT ChilRec(1).Siz
PRINT ChilRec(1).DESC1
PRINT ChilRec(1).DESC2
PRINT ChilRec(1).DESC3
PRINT ChilRec(1).LandVal
PRINT ChilRec(1).BldgVal
PRINT ChilRec(1).Bk
PRINT ChilRec(1).Page
PRINT STRING$(79, "-")
   'SLEEP 1
   
   kk = INSTR(ChilRec(1).NME, ",")

   IF kk = 0 THEN
    TaxCustRec(1).FNAME = ""
    TaxCustRec(1).LName = ChilRec(1).NME
    TaxCustRec(1).SName = ChilRec(1).NME
   ELSE
    Nme2$ = LTRIM$(RIGHT$(ChilRec(1).NME, LEN(ChilRec(1).NME) - kk))
    Nme1$ = LEFT$(ChilRec(1).NME, kk - 1)
    TaxCustRec(1).FNAME = Nme2$
    TaxCustRec(1).LName = Nme1$
    TaxCustRec(1).SName = Nme1$
   END IF
 
 
   Record! = Record! + 1
   ZipCode$ = LTRIM$(ChilRec(1).Zip)
   IF LEN(ZipCode$) = 4 THEN ZipCode$ = "0" + ZipCode$
   IF LEN(QPTrim$(ChilRec(1).Add3)) = 0 THEN
   NoAdd2 = 1
   City$ = ChilRec(1).Add2
   ELSE
   NoAdd2 = 0
   City$ = ChilRec(1).Add3
   END IF

   TaxCustRec(1).acct = Record!
   TaxCustRec(1).Addr1 = ChilRec(1).Add1
   IF NoAdd2 = 0 THEN
   TaxCustRec(1).Addr2 = ChilRec(1).Add2
   ELSE
   TaxCustRec(1).Addr2 = ""
   END IF
   TaxCustRec(1).City = City$
   TaxCustRec(1).STATE = ""
   TaxCustRec(1).Zip = ZipCode$
   TaxCustRec(1).HPHONE = ""
   TaxCustRec(1).WPHONE = ""
   TaxCustRec(1).CSSN = ""
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
   TaxCustRec(1).OPENDATE = Date2Num%("10-01-1999")
   PUT TaxFile, Record!, TaxCustRec(1)
   PropNumb = 1
   GOSUB updatepersonal

NEXTONE:
   NEXT Cnt&
   CLOSE
   EXIT SUB


updatepersonal:
    RERecord& = LOF(PropTaxFile) / LEN(REALRec(1)) + 1
    PROPERTYRec(1).REALPIN = ChilRec(1).MAP
    PROPERTYRec(1).PROPDATE = Date2Num%("10-01-99")
    PROPERTYRec(1).GISPOS = ""
    PROPERTYRec(1).MAP = RIGHT$(ChilRec(1).Bk, 4)
    PROPERTYRec(1).BLOCK = RIGHT$(ChilRec(1).Page, 4)
    PROPERTYRec(1).LOTNUMB = ""
    PROPERTYRec(1).LOTACRE = "A"
    PROPERTYRec(1).PROPSIZE = VAL(ChilRec(1).Siz)
    PROPERTYRec(1).PROPDISC = "N"
    PROPERTYRec(1).LATELIST = "N"
    PROPERTYRec(1).MORTCODE = ""
    PROPERTYRec(1).PROPVALU = VAL(ChilRec(1).LandVal)
    PROPERTYRec(1).EXMPSENI = VAL(ChilRec(1).BldgVal)
    PROPERTYRec(1).EXMPOTHR = 0
    PROPERTYRec(1).PROPNOT1 = ChilRec(1).DESC1
    PROPERTYRec(1).PROPNOT2 = ChilRec(1).DESC2
    PROPERTYRec(1).PROPNOT3 = ChilRec(1).DESC3

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

