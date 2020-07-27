DEFINT A-Z
' data conversion for shenandoah real estate 2001
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
  '$INCLUDE: 'TaxConv.BI'

  
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PROPERTYRec(1) AS PropertyRecType
  DIM SHARED REALRec(1) AS PropertyRecType
  
  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  

   OPEN "ReMAST05.dat" FOR INPUT AS #11
1 LINE INPUT #11, a$
  CLS
  PRINT a$
  PRINT STRING$(79, "-")
   nme$ = LTRIM$(MID$(a$, 23, 35))
   A1$ = LTRIM$(MID$(a$, 58, 35))
   A2$ = LTRIM$(MID$(a$, 93, 35))
   city$ = RTRIM$(MID$(a$, 127, 35))
   city$ = LTRIM$(LEFT$(city$, LEN(city$) - 6))
   Zp$ = LTRIM$(MID$(a$, 164, 5))
   map$ = MID$(a$, 4, 19)
   PRINT nme$, city$

   
   
 

   kk = INSTR(nme$, " ")

   IF kk = 0 THEN
    TaxCustRec(1).FNAME = ""
    TaxCustRec(1).LNAME = nme$
    TaxCustRec(1).SName = nme$
   ELSE
    Nme2$ = LTRIM$(RIGHT$(nme$, LEN(nme$) - kk))
    Nme1$ = LEFT$(nme$, kk - 1)
    TaxCustRec(1).FNAME = Nme2$
    TaxCustRec(1).LNAME = Nme1$
    TaxCustRec(1).SName = Nme1$
   END IF
 
 
   Record! = Record! + 1
   ZipCode$ = LTRIM$(Zp$)
   IF LEN(ZipCode$) = 4 THEN ZipCode$ = "0" + ZipCode$
   TaxCustRec(1).ACCT = Record!
   TaxCustRec(1).Addr1 = A1$
   TaxCustRec(1).Addr2 = A2$
   TaxCustRec(1).city = LEFT$(city$, LEN(city$) - 2)
   TaxCustRec(1).State = RIGHT$(city$, 2)
   TaxCustRec(1).Zip = Zp$
   TaxCustRec(1).HPHONE = ""
   TaxCustRec(1).WPHONE = ""

   

   TaxCustRec(1).CSSN = ""
   TaxCustRec(1).SSSN = ""
 
   TaxCustRec(1).Active = "Y"
   TaxCustRec(1).Interest = "Y"

   IF VAL(MID$(a$, 288, 3)) >= 70 AND VAL(MID$(a$, 288, 3)) <= 79 THEN
    TaxCustRec(1).TaxExempt = "Y"
     ELSE
    TaxCustRec(1).TaxExempt = "N"
   END IF

   TaxCustRec(1).Penalty = "Y"
   TaxCustRec(1).LastTrans = 0
   TaxCustRec(1).FirstPropRec = 0
   TaxCustRec(1).FirstPersRec = 0
   TaxCustRec(1).pin = Record!
   TaxCustRec(1).Deleted = 0
   TaxCustRec(1).FileVer = 8
   TaxCustRec(1).OPENDATE = Date2Num%("12-31-2001")
   PUT TaxFile, Record!, TaxCustRec(1)
   PropNumb = 1
   GOSUB updatepersonal

NEXTONE:
   GOTO 1
   CLOSE
   EXIT SUB


updatepersonal:
    RERecord& = LOF(PropTaxFile) / LEN(REALRec(1)) + 1
    PROPERTYRec(1).REALPIN = map$
    PROPERTYRec(1).PROPDATE = Date2Num%("01-01-2001")
    PROPERTYRec(1).GISPOS = ""
    PROPERTYRec(1).map = ""
    PROPERTYRec(1).BLOCK = ""
    PROPERTYRec(1).LOTNUMB = ""
    PROPERTYRec(1).LOTACRE = "A"
    PROPERTYRec(1).PROPSIZE = VAL(MID$(a$, 278, 8))
    PROPERTYRec(1).PROPDISC = "N"
    PROPERTYRec(1).LateList = "N"
    PROPERTYRec(1).MORTCODE = ""
    PROPERTYRec(1).PROPVALU = VAL(MID$(a$, 293, 10))
    PROPERTYRec(1).EXMPSENI = VAL(MID$(a$, 303, 10))
    PROPERTYRec(1).EXMPOTHR = 0
    PROPERTYRec(1).PROPNOT1 = MID$(a$, 173, 31)
    PROPERTYRec(1).PROPNOT2 = MID$(a$, 208, 31)
    PROPERTYRec(1).PROPNOT3 = MID$(a$, 243, 35)

    PROPERTYRec(1).Fill1 = ""
    PROPERTYRec(1).CustPin = Record!
    PROPERTYRec(1).NextRec = 0
    PROPERTYRec(1).LastYrPrinted = 2000
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

