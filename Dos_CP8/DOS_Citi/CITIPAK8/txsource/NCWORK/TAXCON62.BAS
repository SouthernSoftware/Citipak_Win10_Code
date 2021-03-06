'MAGGIE VALLEY NC FROM EXCEL TO ACCESS CONVERS
DEFINT A-Z
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE SUB BalanceListing ()
DECLARE SUB OpenTaxCustFile (NumOfTaxRecs%, TaxFile%)
DECLARE SUB OpenTaxPropFile (NumOfPropRecs%, PropTaxFile%)
DECLARE SUB OpenTaxPersFile (NumOfPersRecs%, PersTaxFile%)
DECLARE SUB DisplayTaxScrn (ScrnName$)
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
  DIM SHARED PersRec(1) AS PersonalRecType
  

  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
  DIM a$(120)
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile
  
   OPEN "MVALLEY.TXT" FOR INPUT AS #11
   Record = NumOfTaxRecs
   CLS
1 FOR X = 1 TO 26
10 INPUT #11, a$(X)
     PRINT a$(X),
  NEXT X
  R# = R# + VAL(a$(3))
  P# = P# + VAL(a$(4))
  e# = e# + VAL(a$(5))
 d# = d# + VAL(a$(6))


  GOTO 1



  PRINT STRING$(79, "-")

  
    kk = INSTR(a$(17), ",")
      IF kk > 0 THEN
       LN$ = UCASE$(LEFT$(a$(17), kk - 1))
       SN$ = UCASE$(LEFT$(a$(17), kk - 1))
       FF1$ = UCASE$(MID$(a$(17), kk + 1, LEN(a$(17)) - (kk)))
       ELSE
       LN$ = UCASE$(RTRIM$(a$(17)))
       SN$ = UCASE$(RTRIM$(a$(17)))
       FF1$ = ""
      END IF
'STOP
   Record = Record + 1
   TaxCustRec(1).Acct = Record
   TaxCustRec(1).OpenDate = Date2Num%("07-01-2001")
   TaxCustRec(1).FName = LTRIM$(FF1$)
   TaxCustRec(1).LName = LTRIM$(LN$)
   TaxCustRec(1).SName = LTRIM$(LN$)
   TaxCustRec(1).Addr1 = UCASE$(a$(19))
   TaxCustRec(1).Addr2 = UCASE$(a$(20))
   TaxCustRec(1).City = UCASE$(a$(21))
   TaxCustRec(1).State = UCASE$(a$(22))
   TaxCustRec(1).Zip = LEFT$(a$(23), 5)
   TaxCustRec(1).HPhone = ""
   TaxCustRec(1).WPhone = ""
   TaxCustRec(1).CSSN = ""
   TaxCustRec(1).SSSN = ""
   TaxCustRec(1).Active = "Y"
   TaxCustRec(1).Interest = "Y"
   TaxCustRec(1).TaxExempt = "N"
   TaxCustRec(1).Penalty = "N"
   TaxCustRec(1).CountyAcctString = a$(11)
   TaxCustRec(1).LastTrans = 0
   TaxCustRec(1).FirstPropRec = 0
   TaxCustRec(1).FirstPersRec = 0
   TaxCustRec(1).PIN = Record
   TaxCustRec(1).Deleted = 0
   TaxCustRec(1).FileVer = 8
   PUT TaxFile, Record, TaxCustRec(1)
   IF VAL(a$(3)) > 0 THEN
   GOSUB UpdateReal
   END IF
   IF VAL(a$(4)) > 0 THEN
   GOSUB UpdatePers
   END IF
  
NextRecd:
   GOTO 1
   CLOSE
   EXIT SUB
  
  

UpdateReal:
    PropertyRec(1).REALPIN = a$(13)
    PropertyRec(1).PROPDATE = Date2Num%("07-01-2001")
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).MAP = ""
    PropertyRec(1).BLOCK = ""
    PropertyRec(1).LOTNUMB = ""
    PropertyRec(1).LOTACRE = ""
    PropertyRec(1).PROPSIZE = 0
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LATELIST = "N"
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).PROPVALU = VAL(a$(3))
    PropertyRec(1).EXMPSENI = VAL(a$(5))
    PropertyRec(1).EXMPOTHR = 0
    PropertyRec(1).PROPNOT1 = UCASE$(a$(16))
    PropertyRec(1).PROPNOT2 = ""
    PropertyRec(1).PROPNOT3 = ""
    PropertyRec(1).Fill1 = ""
    PropertyRec(1).CustPin = Record
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 0
    PropertyRec(1).Deleted = 0
    PropertyRec(1).Blank = ""
    NextRecord = (LOF(PropTaxFile) / LEN(PropertyRec(1))) + 1
    PUT PropTaxFile, NextRecord, PropertyRec(1)
    RETURN

UpdatePers:
   PersRec(1).PROPPIN = "1"
   PersRec(1).PROPDATE = Date2Num("07-01-2001")
   PersRec(1).PERSVAL = VAL(a$(4))
   PersRec(1).MHVALUE = 0
   PersRec(1).MCVALUE = 0
   PersRec(1).CVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   IF a$(9) = "1" THEN
    PersRec(1).LATELIST = "Y"
   ELSE
    PersRec(1).LATELIST = "N"
   END IF
   PersRec(1).DESC1 = ""
   PersRec(1).DESC2 = ""
   PersRec(1).DESC3 = ""
   PersRec(1).DESC4 = ""
   PersRec(1).DESC5 = ""
   PersRec(1).CustPin = Record
   PersRec(1).NextRec = 0
   PersRec(1).LastYrPrinted = 0
   PersRec(1).Deleted = 0
   PersRec(1).Blank = ""
   NextRec = (LOF(PersTaxFile) / LEN(PersRec(1))) + 1
   PUT PersTaxFile, NextRec, PersRec(1)
   RETURN





END SUB

SUB OpenTaxCustFile (NumOfTaxRecs, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "TAXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
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

