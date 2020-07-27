DEFINT A-Z
'Real Estate Conversion from Text File Information Hillsville Disk 1 (real1)
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
DIM A$(36)
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  Record! = NumOfTaxRecs

   CLS
   GOSUB InitProp
   OPEN "HILLREAL.TXT" FOR INPUT AS #11
1 FOR C = 1 TO 20
   INPUT #11, A$(C)
  NEXT C
  'PRINT STRING$(79, "-")

  FOR Snt& = 1 TO NumOfTaxRecs
  GET TaxFile, Snt&, TaxCustRec(1)
  'PRINT TaxCustRec(1).CountyAcctstring
  IF QPTrim$(TaxCustRec(1).CountyAcctstring) = RTRIM$(A$(13)) THEN
    PropRecord! = TaxCustRec(1).FirstPropRec
    GOSUB UpdateReal1
   GOTO 1
  END IF

  NEXT Snt&
'STOP


  TOTREAL# = TOTREAL# + VAL(A$(15))
  TOTALBLDG# = TOTALBLDG# + VAL(A$(16))
  CC = CC + 1


   'Decode Name Here
   kk = INSTR(A$(3), " ")
   IF kk = 0 THEN
    LN$ = A$(3)
    FM$ = ""
    ELSE
    LN$ = LEFT$(A$(3), kk - 1)
    FM$ = LTRIM$(MID$(A$(3), kk + 1, LEN(A$(3)) - kk))
   END IF
    SN$ = LN$
    LN$ = UCASE$(LN$)
    FM$ = UCASE$(FM$)
    SN$ = UCASE$(SN$)
    Record! = Record! + 1
    TaxCustRec(1).FNAME = FM$
    TaxCustRec(1).LName = LN$
    TaxCustRec(1).SName = SN$
    
   TaxCustRec(1).Acct = Record!
   TaxCustRec(1).Addr1 = UCASE$(A$(5))
   TaxCustRec(1).Addr2 = UCASE$(A$(6))
   TaxCustRec(1).CITY = UCASE$(A$(7))
   TaxCustRec(1).STATE = UCASE$(A$(8))
   TaxCustRec(1).ZIP = A$(9)
   TaxCustRec(1).HPHONE = ""
   TaxCustRec(1).WPHONE = ""
   TaxCustRec(1).CSSN = ""
   TaxCustRec(1).SSSN = ""
   TaxCustRec(1).CountyAcctstring = A$(13)
   TaxCustRec(1).CountyAcct = VAL(A$(1))
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
   TaxCustRec(1).OPENDATE = Date2Num%("10-01-2000")
   PUT TaxFile, Record!, TaxCustRec(1)
   PropNumb = 1
   GOSUB UpdateReal

NEXTONE:
   GOTO 1
   CLOSE
   EXIT SUB


UpdateReal:
    RERecord& = LOF(PropTaxFile) / LEN(REALRec(1)) + 1
    PropertyRec(1).REALPIN = A$(13)
    PropertyRec(1).PROPDATE = Date2Num%("11-01-2000")
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).MAP = ""
    PropertyRec(1).BLOCK = ""
    PropertyRec(1).LOTNUMB = ""
    PropertyRec(1).LOTACRE = "A"
    PropertyRec(1).PROPSIZE = 0
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LATELIST = "N"
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).PROPVALU = VAL(A$(15))
    PropertyRec(1).EXMPSENI = VAL(A$(16))
    PropertyRec(1).EXMPOTHR = VAL(A$(17))
    PropertyRec(1).PROPNOT1 = A$(12)
    PropertyRec(1).PROPNOT2 = A$(14)
    PropertyRec(1).PROPNOT3 = ""

    PropertyRec(1).Fill1 = ""
    PropertyRec(1).CustPin = Record!
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 1998
    PropertyRec(1).Deleted = 0
    PropertyRec(1).Blank = ""


   
   PUT PropTaxFile, RERecord&, PropertyRec(1)
   RETURN
UpdateReal1:
    GET PropTaxFile, PropRecord!, PropertyRec(1)
    PropertyRec(1).REALPIN = A$(13)
    PropertyRec(1).PROPDATE = Date2Num%("11-01-2000")
    PropertyRec(1).PROPVALU = VAL(A$(15))
    PropertyRec(1).EXMPSENI = VAL(A$(16))
    PropertyRec(1).EXMPOTHR = VAL(A$(17))
    PropertyRec(1).PROPNOT1 = A$(12)
    PropertyRec(1).PROPNOT2 = A$(14)
    PropertyRec(1).PROPNOT3 = ""
    PropertyRec(1).CustPin = Snt&
    PropertyRec(1).LastYrPrinted = 1999
    PUT PropTaxFile, PropRecord!, PropertyRec(1)
   RETURN

InitProp:
 FOR Cnt& = 1 TO NumOfPropRecs
  GET PropTaxFile, Cnt&, PropertyRec(1)
  PropertyRec(1).PROPVALU = 0
  PropertyRec(1).EXMPSENI = 0
  PropertyRec(1).EXMPOTHR = 0
  PUT PropTaxFile, Cnt&, PropertyRec(1)
 NEXT Cnt&
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

