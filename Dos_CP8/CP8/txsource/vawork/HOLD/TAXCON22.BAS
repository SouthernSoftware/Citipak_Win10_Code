DEFINT A-Z
'Convert Master PP Records from TXT File Hillsville VA
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
DECLARE FUNCTION Round# (b#)
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


  'Stephens City Va Tax PP Conversion of Master (Customers)

  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'Taxfiles.BI'
  '$INCLUDE: 'PROPAbst.BI'
  
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED PersRec(1) AS PersonalRecType
  DIM SHARED TBCustRec(1) AS TBCustType
  DIM SHARED TBCustRec1(1) AS TBCust1Type
  DIM SHARED TaxTrans(1) AS TaxTransactionType
  

  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
  DIM A$(31)
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile

   OPEN "DIST2.TXT" FOR INPUT AS #11
   CLS
   b = 0
1  FOR cc = 1 TO 31
    INPUT #11, A$(cc)
    PRINT cc, A$(cc)
    NEXT cc
    Tax# = Tax# + VAL(A$(15))
    g = g + 1
   PRINT STRING$(79, "-")

  'First Search for Existing Acct
   FOR Snt& = 1 TO NumOfTaxRecs
   GET TaxFile, Snt&, TaxCustRec(1)
   IF RTRIM$(TaxCustRec(1).CountyAcctString) = RTRIM$(A$(1)) THEN
    'If Found then Update Property Only
     AcctRecord! = Snt&
     GOSUB UpdatePersProperty
     GOTO NextOne
   END IF
   NEXT Snt&
  
  'If Not Found then Add Customer & Property

    Record! = NumOfTaxRecs + 1
    NumOfTaxRecs = Record!
    kk = INSTR(A$(2), " ")
    IF kk > 0 THEN
     LNM$ = LEFT$(A$(2), kk - 1)
     F$ = MID$(A$(2), kk + 1, LEN(A$(2)) - kk)
     ELSE
     LNM$ = A$(2)
     F$ = ""
    END IF
    TaxCustRec(1).FNAME = LTRIM$(F$)
    TaxCustRec(1).LName = LTRIM$(LNM$)
    TaxCustRec(1).SName = LTRIM$(LNM$)
    TaxCustRec(1).Acct = Record!
    TaxCustRec(1).Addr1 = A$(8)
    TaxCustRec(1).Addr2 = ""
    TaxCustRec(1).City = A$(10)
    TaxCustRec(1).STATE = A$(11)
    TaxCustRec(1).Zip = A$(12)
    TaxCustRec(1).HPHONE = ""
    TaxCustRec(1).WPHONE = ""
    TaxCustRec(1).CSSN = LEFT$(A$(1), 9)
    TaxCustRec(1).SSSN = LEFT$(A$(4), 9)
    TaxCustRec(1).Active = "Y"
    TaxCustRec(1).Interest = "Y"
    TaxCustRec(1).TaxExempt = "N"
    TaxCustRec(1).Penalty = "Y"
    TaxCustRec(1).CountyAcct = 0
    TaxCustRec(1).CountyAcctString = A$(1)
    TaxCustRec(1).LastTrans = 0
    TaxCustRec(1).FirstPropRec = 0
    TaxCustRec(1).FirstPersRec = 0
    TaxCustRec(1).PIN = Record!
    TaxCustRec(1).Deleted = 0
    TaxCustRec(1).FileVer = 8
    TaxCustRec(1).OPENDATE = Date2Num%("08-01-1999")
    PUT TaxFile, Record!, TaxCustRec(1)
    PropNumb = 1
    AcctRecord! = Record!
    GOSUB UpdatePersProperty
 
NextOne:
   GOTO 1
   CLOSE
   EXIT SUB

UpdatePersProperty:
   VType = VAL(A$(14))
   PersRec(1).MCVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).MHVALUE = 0
   PersRec(1).PERSVAL = 0
   PersRec(1).CValue = 0

   IF VType = 3 THEN
    PersRec(1).MCVALUE = VAL(A$(15))
    GOTO upd1
   END IF
   IF VType = 4 THEN
    PersRec(1).MTVALUE = VAL(A$(15))
    GOTO upd1
   END IF
   IF VType = 5 THEN
    PersRec(1).MHVALUE = VAL(A$(15))
    GOTO upd1
   END IF
   
   PersRec(1).PERSVAL = VAL(A$(15))
upd1:
TVal# = TVal# + VAL(A$(15))
   PersRec(1).PROPPIN = LTRIM$(STR$(PropNumb))
   PersRec(1).PROPDATE = Date2Num%("11-01-1999")
   PersRec(1).CValue = 0
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LATELIST = "N"
    PersRec(1).DESC1 = A$(23)
    PersRec(1).DESC2 = A$(19) + "/" + A$(21) + "/" + A$(22)
    PersRec(1).DESC3 = A$(24)
    PersRec(1).DESC4 = LEFT$(A$(20), 4)
    IF VAL(A$(26)) = 1 THEN
     PersRec(1).DESC5 = "Y" + "/"
    ELSE
     PersRec(1).DESC5 = "N" + "/"
    END IF
   
   PersRec(1).CustPin = AcctRecord!
   PersRec(1).NextRec = 0
   PersRec(1).LastYrPrinted = 1998
   PersRec(1).Deleted = 0
   PersRec(1).Blank = ""
   PRecord& = LOF(PersTaxFile) / LEN(PersRec(1)) + 1
   PUT PersTaxFile, PRecord&, PersRec(1)
   RETURN



END SUB

SUB OpenTaxCustFile (NumOfTaxRecs, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
  
END SUB

SUB OpenTaxPersFile (NumOfPersRecs, PersTaxFile)
  PersTaxFile = FREEFILE
  OPEN "TAXPERS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PersTaxFile LEN = LEN(PersRec(1))
  NumOfPersRecs = LOF(PersTaxFile) / LEN(PersRec(1))
  
END SUB

SUB OpenTaxPropFile (NumOfPropRecs, PropTaxFile)
  PropTaxFile = FREEFILE
  OPEN "TAXPROP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PropTaxFile LEN = LEN(PropertyRec(1))
  NumOfPropRecs = LOF(PropTaxFile) / LEN(PropertyRec(1))
END SUB

