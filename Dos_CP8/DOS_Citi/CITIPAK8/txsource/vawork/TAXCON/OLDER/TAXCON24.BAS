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
  DIM SHARED PROPERTYRec(1) AS PropertyRecType
  DIM SHARED REALRec(1) AS PropertyRecType
  DIM SHARED RAMAST(1) AS RAMAST
  
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
   OPEN "REALES.TXT" FOR INPUT AS #11
1 FOR C = 1 TO 20
   INPUT #11, A$(C)
  ' PRINT A$(C)
  
  NEXT C
'  PRINT STRING$(79, "-")


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
   TaxCustRec(1).CountyAcctString = A$(13)
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
   TaxCustRec(1).OPENDATE = Date2Num%("01-01-1999")
   PUT TaxFile, Record!, TaxCustRec(1)
   PropNumb = 1
   GOSUB updatereal

NEXTONE:
   GOTO 1
   CLOSE
   EXIT SUB


updatereal:
    RERecord& = LOF(PropTaxFile) / LEN(REALRec(1)) + 1
    PROPERTYRec(1).REALPIN = A$(13)
    PROPERTYRec(1).PROPDATE = Date2Num%("01-01-99")
    PROPERTYRec(1).GISPOS = ""
    PROPERTYRec(1).MAP = ""
    PROPERTYRec(1).BLOCK = ""
    PROPERTYRec(1).LOTNUMB = ""
    PROPERTYRec(1).LOTACRE = "A"
    PROPERTYRec(1).PROPSIZE = 0
    PROPERTYRec(1).PROPDISC = "N"
    PROPERTYRec(1).LATELIST = "N"
    PROPERTYRec(1).MORTCODE = ""
    PROPERTYRec(1).PROPVALU = VAL(A$(15))
    PROPERTYRec(1).EXMPSENI = VAL(A$(16))
    PROPERTYRec(1).EXMPOTHR = VAL(A$(17))
    PROPERTYRec(1).PROPNOT1 = A$(12)
    PROPERTYRec(1).PROPNOT2 = A$(14)
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

