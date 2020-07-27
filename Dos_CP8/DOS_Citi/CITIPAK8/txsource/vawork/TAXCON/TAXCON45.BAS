DEFINT A-Z
'Real Estate Conversion from Text File Information  Shenandoah VA 2000
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
  '$INCLUDE: 'TAXCONV.BI'
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType


  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
DIM a$(36)
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile

   CLS
   OPEN "REDATA.TXT" FOR INPUT AS #11
1 FOR C = 1 TO 29
   INPUT #11, a$(C)
   NEXT C
  PRINT a$(1), a$(8)
  PRINT STRING$(79, "-")
  

   'Decode Name Here
   kk = INSTR(a$(8), " ")
   IF kk = 0 THEN
    LN$ = a$(8)
    fm$ = ""
    ELSE
    LN$ = LEFT$(a$(8), kk - 1)
    fm$ = LTRIM$(MID$(a$(8), kk + 1, LEN(a$(8)) - kk))
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
   TaxCustRec(1).Addr1 = UCASE$(a$(9))
   TaxCustRec(1).Addr2 = UCASE$(a$(10))
   TaxCustRec(1).City = UCASE$(a$(11))
   TaxCustRec(1).STATE = ""
   TaxCustRec(1).ZIP = ""
   TaxCustRec(1).HPHONE = ""
   TaxCustRec(1).WPHONE = ""
   TaxCustRec(1).CSSN = ""
   TaxCustRec(1).SSSN = ""
   'County String is Map Number with Spaces Taken Out
   CSTRING$ = QPTrim$(a$(1))
   CSTRING$ = CSTRING$ + QPTrim$(a$(2))
   CSTRING$ = CSTRING$ + QPTrim$(a$(3))
   CSTRING$ = CSTRING$ + QPTrim$(a$(4))
   CSTRING$ = CSTRING$ + QPTrim$(a$(5))
   CSTRING$ = CSTRING$ + QPTrim$(a$(6))
   CSTRING$ = CSTRING$ + QPTrim$(a$(7))
   



   TaxCustRec(1).CountyAcctString = CSTRING$
   TaxCustRec(1).CountyAcct = 0
   TaxCustRec(1).Active = "Y"
   TaxCustRec(1).interest = "Y"
   TaxCustRec(1).TaxExempt = "N"
   TaxCustRec(1).Penalty = "Y"
   TaxCustRec(1).LastTrans = 0
   TaxCustRec(1).FirstPropRec = 0
   TaxCustRec(1).FirstPersRec = 0
   TaxCustRec(1).PIN = x
   TaxCustRec(1).Deleted = 0
   TaxCustRec(1).FileVer = 8
   TaxCustRec(1).OPENDATE = Date2Num%("06-01-2000")
   PUT TaxFile, Record!, TaxCustRec(1)
   PropNumb = 1
   GOSUB updatereal

NEXTONE:
   GOTO 1
   CLOSE
   EXIT SUB


updatereal:
    RERecord& = LOF(PropTaxFile) / LEN(PropertyRec(1)) + 1
    PropertyRec(1).REALPIN = RTRIM$(a$(1)) + RTRIM$(a$(3)) + RTRIM$(a$(6))
    PropertyRec(1).PROPDATE = Date2Num%("06-01-2000")
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).MAP = RTRIM$(a$(1)) + RTRIM$(a$(2))
    PropertyRec(1).BLOCK = RTRIM$(a$(3)) + RTRIM$(a$(4)) + RTRIM$(a$(5))
    PropertyRec(1).LOTNUMB = RTRIM$(a$(6)) + RTRIM$(a$(7))
    PropertyRec(1).LOTACRE = "A"
    PropertyRec(1).PROPSIZE = VAL(a$(16))
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LATELIST = "N"
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).PROPVALU = VAL(a$(19))
    PropertyRec(1).EXMPSENI = VAL(a$(20))
    PropertyRec(1).EXMPOTHR = 0
    PropertyRec(1).PROPNOT1 = a$(13)
    PropertyRec(1).PROPNOT2 = a$(14)
    PropertyRec(1).PROPNOT3 = a$(15)

    PropertyRec(1).Fill1 = ""
    PropertyRec(1).CustPin = Record!
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 1999
    PropertyRec(1).Deleted = 0
    PropertyRec(1).Blank = ""


   
   PUT PropTaxFile, RERecord&, PropertyRec(1)
   RETURN



END SUB

SUB OpenTaxCustFile (NumOfTaxRecs, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "RETXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
END SUB

SUB OpenTaxPropFile (NumOfPropRecs, PropTaxFile)
  PropTaxFile = FREEFILE
  OPEN "TAXPROP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PropTaxFile LEN = LEN(PropertyRec(1))
  NumOfPropRecs = LOF(PropTaxFile) / LEN(PropertyRec(1))
END SUB

