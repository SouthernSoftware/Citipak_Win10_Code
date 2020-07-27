'Cedar Bluff TAX DATA CONV REAL ESTATE 2003
DEFINT A-Z
DECLARE SUB BalanceListing ()
DECLARE SUB OpenTaxCustFile (numoftaxrecs%, TaxFile%)
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
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED REALRec(1) AS PropertyRecType
  
  
  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
  OpenTaxCustFile numoftaxrecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile



    OPEN "CBREAL" FOR INPUT AS #11
1   CLS
    LINE INPUT #11, a$
    PRINT a$
    acct$ = MID$(a$, 11, 8)
    N$ = MID$(a$, 18, 35)
    Add1$ = MID$(a$, 53, 30)
    add2$ = MID$(a$, 83, 30)
    add3$ = MID$(a$, 113, 30)
    Zip$ = MID$(a$, 144, 5)
    map$ = MID$(a$, 239, 16)
    IF VAL(MID$(a$, 348, 2)) = 7 THEN
      GOTO 1
     ELSE
      rv# = rv# + VAL(MID$(a$, 917, 9))
      RealVal# = VAL(MID$(a$, 917, 9))
      bv# = bv# + VAL(MID$(a$, 927, 9))
      BldgVal# = VAL(MID$(a$, 927, 9))
    END IF

FOR x = 1 TO numoftaxrecs
GET TaxFile, x, TaxCustRec(1)
 IF VAL(acct$) = TaxCustRec(1).CountyAcct THEN
  kk = INSTR(N$, ",")

   IF kk = 0 THEN
    TaxCustRec(1).FNAME = ""
    TaxCustRec(1).LName = N$
    TaxCustRec(1).SName = N$
   ELSE
    Nme2$ = LTRIM$(RIGHT$(N$, LEN(N$) - kk))
    Nme1$ = LEFT$(N$, kk - 1)
    TaxCustRec(1).FNAME = Nme2$
    TaxCustRec(1).LName = Nme1$
    TaxCustRec(1).SName = Nme1$
   END IF

   TaxCustRec(1).Addr1 = Add1$
   IF LEN(QPTrim$(add3$)) = 0 THEN
    TaxCustRec(1).Addr2 = ""
    TaxCustRec(1).City = add2$
     ELSE
    TaxCustRec(1).Addr2 = add2$
    TaxCustRec(1).City = add3$
   END IF
   TaxCustRec(1).STATE = ""
   TaxCustRec(1).Zip = Zip$
   PUT TaxFile, x, TaxCustRec(1)
   GOTO 1
 END IF
NEXT x
GOTO 1

   CLOSE
   END
   'Must Add Acct Here
   
   kk = INSTR(N$, ",")

   IF kk = 0 THEN
    TaxCustRec(1).FNAME = ""
    TaxCustRec(1).LName = N$
    TaxCustRec(1).SName = N$
   ELSE
    Nme2$ = LTRIM$(RIGHT$(N$, LEN(N$) - kk))
    Nme1$ = LEFT$(N$, kk - 1)
    TaxCustRec(1).FNAME = Nme2$
    TaxCustRec(1).LName = Nme1$
    TaxCustRec(1).SName = Nme1$
   END IF
 
 
   Record! = LOF(TaxFile) / LEN(TaxCustRec(1)) + 1

   TaxCustRec(1).acct = Record!
   TaxCustRec(1).Addr1 = Add1$
   IF LEN(QPTrim$(add3$)) = 0 THEN
    TaxCustRec(1).Addr2 = ""
    TaxCustRec(1).City = add2$
     ELSE
    TaxCustRec(1).Addr2 = add2$
    TaxCustRec(1).City = add3$
   END IF
   TaxCustRec(1).STATE = ""
   TaxCustRec(1).Zip = Zip$
   TaxCustRec(1).HPHONE = ""
   TaxCustRec(1).WPHONE = ""
   TaxCustRec(1).CSSN = ""
   TaxCustRec(1).SSSN = ""
   TaxCustRec(1).CountyAcctString$ = a$
   TaxCustRec(1).CountyAcct = VAL(a$)
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
   TaxCustRec(1).OPENDATE = Date2Num%("10-01-2002")
   PUT TaxFile, Record!, TaxCustRec(1)
   PropNumb = 1
   GOSUB updatepersonal

NEXTONE:
   GOTO 1
   CLOSE
   EXIT SUB


updatepersonal:
    RERecord& = LOF(PropTaxFile) / LEN(REALRec(1)) + 1
    PropertyRec(1).REALPIN = map$
    PropertyRec(1).PROPDATE = Date2Num%("10-01-2002")
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).map = ""
    PropertyRec(1).BLOCK = ""
    PropertyRec(1).LOTNUMB = ""
    PropertyRec(1).LOTACRE = "A"
    PropertyRec(1).PROPSIZE = 0
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LATELIST = "N"
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).PROPVALU = RealVal#
    PropertyRec(1).EXMPSENI = BldgVal#
    PropertyRec(1).EXMPOTHR = 0
    PropertyRec(1).PROPNOT1 = DESC1$
    PropertyRec(1).PROPNOT2 = DESC2$
    PropertyRec(1).PROPNOT3 = DESC3$
    PropertyRec(1).Fill1 = ""
    PropertyRec(1).CustPin = Record!
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 2001
    PropertyRec(1).Deleted = 0
    PropertyRec(1).Blank = ""
    PUT PropTaxFile, RERecord&, PropertyRec(1)
    RETURN



END SUB

SUB OpenTaxCustFile (numoftaxrecs, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "RETXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  numoftaxrecs = LOF(TaxFile) / LEN(TaxCustRec(1))
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

