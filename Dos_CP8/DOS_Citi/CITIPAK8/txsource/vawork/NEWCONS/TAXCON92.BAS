
'Cedar Bluff 2003 Real Estate Conversion from CBTAXMST
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
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED REALRec(1) AS PropertyRecType
  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  Record! = NumOfTaxRecs

CLS
   OPEN "CBTAXMST" FOR INPUT AS #10
1 LINE INPUT #10, a$
    ac = ac + 1
    Acct$ = MID$(a$, 6, 12)
    pdesc$ = MID$(a$, 164, 25)
    PDESC2$ = MID$(a$, 189, 25)
    PDESC3$ = MID$(a$, 215, 25)
    PMap$ = MID$(a$, 239, 20)
    PLand$ = MID$(a$, 360, 10)
    PBldg$ = MID$(a$, 370, 10)


 class = VAL(MID$(a$, 348, 2))
 IF class = 7 THEN
  GOTO 1
 END IF

    land# = land# + VAL(PLand$)
    Bldg# = Bldg# + VAL(PBldg$)


  'Find Account if exists
  FOR Cnt& = 1 TO NumOfTaxRecs
  GET TaxFile, Cnt&, TaxCustRec(1)
  IF TaxCustRec(1).CountyAcct = VAL(Acct$) AND TaxCustRec(1).FirstPropRec <> 0 THEN
   GOSUB updateacct
   GOTO 1
  END IF
  NEXT Cnt&


  'Must Add Acct Here


 
   nam$ = MID$(a$, 18, 35)
   nam$ = RTRIM$(nam$)
   namlen = LEN(nam$)
   kk = INSTR(nam$, " ")
   IF kk = 0 THEN
    LNam$ = nam$
     FirstNam$ = ""
    ELSE
     LNam$ = LEFT$(nam$, kk - 2)
     FirstNam$ = RIGHT$(nam$, (namlen - kk))
   END IF

 A1$ = MID$(a$, 53, 30)
 Cy$ = MID$(a$, 83, 13)
 ST$ = MID$(a$, 96, 2)
 Zp$ = MID$(a$, 144, 5)

    TaxCustRec(1).FNAME = FirstNam$
    TaxCustRec(1).LName = LNam$
    TaxCustRec(1).SName = LNam$
  
   Record! = Record! + 1

   TaxCustRec(1).Acct = Record!
   TaxCustRec(1).Addr1 = A1$
   TaxCustRec(1).Addr2 = ""
   TaxCustRec(1).City = Cy$
   TaxCustRec(1).STATE = ST$
   TaxCustRec(1).Zip = Zp$
   TaxCustRec(1).CountyAcct = VAL(Acct$)
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
   TaxCustRec(1).PIN = Record!
   TaxCustRec(1).Deleted = 0
   TaxCustRec(1).FileVer = 8
   TaxCustRec(1).OPENDATE = Date2Num%("10-01-2001")
   PUT TaxFile, Record!, TaxCustRec(1)
   PropNumb = 1
    RERecord& = LOF(PropTaxFile) / LEN(REALRec(1)) + 1
    PropertyRec(1).REALPIN = PMap$
    PropertyRec(1).PROPDATE = Date2Num%("10-01-2001")
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).MAP = ""
    PropertyRec(1).BLOCK = ""
    PropertyRec(1).LOTNUMB = ""
    PropertyRec(1).LOTACRE = "A"
    PropertyRec(1).PROPSIZE = 0
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LATELIST = "N"
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).PROPVALU = VAL(PLand$)
    PropertyRec(1).EXMPSENI = VAL(PBldg$)
    PropertyRec(1).EXMPOTHR = 0
    PropertyRec(1).PROPNOT1 = PDESC1$
    PropertyRec(1).PROPNOT2 = PDESC2$
    PropertyRec(1).PROPNOT3 = PDESC3$
    PropertyRec(1).CustPin = Record!
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 2000
    PropertyRec(1).Deleted = 0
    PropertyRec(1).blank = ""
    PUT PropTaxFile, RERecord&, PropertyRec(1)
    GOTO 1
   CLOSE
   EXIT SUB


updateacct:
    GET PropTaxFile, TaxCustRec(1).FirstPropRec, PropertyRec(1)
    PropertyRec(1).PROPDATE = Date2Num%("10-01-2001")
    PropertyRec(1).PROPVALU = VAL(PLand$)
    PropertyRec(1).EXMPSENI = VAL(PBldg$)
    PropertyRec(1).EXMPOTHR = 0
    PropertyRec(1).PROPNOT1 = PDESC1$
    PropertyRec(1).PROPNOT2 = PDESC2$
    PropertyRec(1).PROPNOT3 = PDESC3$
    PUT PropTaxFile, TaxCustRec(1).FirstPropRec, PropertyRec(1)
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

