DEFINT A-Z
'Convert PP Records from CSV File For Windsor 2003
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

  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
  DIM a$(34)
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile

   OPEN "windpers.txt" FOR INPUT AS #11
   CLS
   b = 0
1  FOR cc = 1 TO 34
    INPUT #11, a$(cc)
    NEXT cc
    PRINT a$(2), a$(3), a$(15), a$(16)
   PRINT STRING$(79, "-")
   Vtype$ = a$(16)




  'strip out - in ss#
   s1$ = ""
   ll = LEN(a$(9))
   FOR a = 1 TO ll
    IF MID$(a$(9), a, 1) = "-" THEN
    ELSE
    s1$ = s1$ + MID$(a$(9), a, 1)
    END IF
   NEXT a
   
  'strip out - in ss#
   s2$ = ""
   ll = LEN(a$(10))
   FOR a = 1 TO ll
    IF MID$(a$(10), a, 1) = "-" THEN
    ELSE
    s2$ = s2$ + MID$(a$(10), a, 1)
    END IF
   NEXT a

   'MAKE SURE ADDRESS IS IN A$(4)
    IF LEN(QPTrim$(a$(4))) = 0 THEN
     a$(4) = a$(5): a$(5) = ""
    END IF

   'FIX ZIP TO 5 CHAR
    a$(8) = LEFT$(a$(8), 5)


   IF NOT (Vtype$ = "A" OR Vtype$ = "B" OR Vtype$ = "C" OR Vtype$ = "E" OR Vtype$ = "H" OR Vtype$ = "I" OR Vtype$ = "M" OR Vtype$ = "O" OR Vtype$ = "P" OR Vtype$ = "Q" OR Vtype$ = "R" OR Vtype$ = "T" OR Vtype$ = "W") THEN
    GOTO 1
   END IF

  'First Search for Existing Acct
   FOR Snt& = 1 TO NumOfTaxRecs
   GET TaxFile, Snt&, TaxCustRec(1)
   IF TaxCustRec(1).CountyAcct = VAL(a$(2)) THEN


    kk = INSTR(a$(3), " ")
    IF kk > 0 THEN
     lnm$ = LEFT$(a$(3), kk - 1)
     F$ = RIGHT$(a$(3), LEN(a$(3)) - kk)
    ELSE
     lnm$ = a$(3)
     F$ = ""
    END IF
    TaxCustRec(1).FNAME = LTRIM$(F$)
    TaxCustRec(1).LName = LTRIM$(lnm$)
    TaxCustRec(1).SName = LTRIM$(lnm$)
    TaxCustRec(1).Addr1 = a$(4)
    TaxCustRec(1).Addr2 = a$(5)
    TaxCustRec(1).City = a$(6)
    TaxCustRec(1).STATE = a$(7)
    TaxCustRec(1).Zip = a$(8)
    TaxCustRec(1).CSSN = s1$
    TaxCustRec(1).SSSN = s2$

    TaxCustRec(1).OPENDATE = Date2Num%("11-19-2003")
    PUT TaxFile, Snt&, TaxCustRec(1)

    'If Found then Update Property Only
     Record! = Snt&
     GOSUB UpdatePersProperty
     GOTO NextOne
   END IF
   NEXT Snt&
  
  'If Not Found then Add Customer & Property

    Record! = LOF(TaxFile) / LEN(TaxCustRec(1)) + 1
    
    kk = INSTR(a$(3), " ")
    IF kk > 0 THEN
     lnm$ = LEFT$(a$(3), kk - 1)
     F$ = RIGHT$(a$(3), LEN(a$(3)) - kk)
    ELSE
     lnm$ = a$(3)
     F$ = ""
    END IF
    TaxCustRec(1).FNAME = LTRIM$(F$)
    TaxCustRec(1).LName = LTRIM$(lnm$)
    TaxCustRec(1).SName = LTRIM$(lnm$)
    TaxCustRec(1).Acct = Record!
    TaxCustRec(1).Addr1 = a$(4)
    TaxCustRec(1).Addr2 = a$(5)
    TaxCustRec(1).City = a$(6)
    TaxCustRec(1).STATE = a$(7)
    TaxCustRec(1).Zip = a$(8)
    TaxCustRec(1).HPHONE = ""
    TaxCustRec(1).WPHONE = ""
    TaxCustRec(1).CSSN = s1$
    TaxCustRec(1).SSSN = s2$
    TaxCustRec(1).Active = "Y"
    TaxCustRec(1).Interest = "Y"
    TaxCustRec(1).TaxExempt = "N"
    TaxCustRec(1).Penalty = "Y"
    TaxCustRec(1).CountyAcct = VAL(a$(2))
    TaxCustRec(1).CountyAcctString = a$(2)
    TaxCustRec(1).LastTrans = 0
    TaxCustRec(1).FirstPropRec = 0
    TaxCustRec(1).FirstPersRec = 0
    TaxCustRec(1).PIN = Record!
    TaxCustRec(1).Deleted = 0
    TaxCustRec(1).FileVer = 8
    TaxCustRec(1).OPENDATE = Date2Num%("11-19-2003")
    PUT TaxFile, Record!, TaxCustRec(1)
    PropNumb = 1
   GOSUB UpdatePersProperty
 
NextOne:
   GOTO 1
   CLOSE
   EXIT SUB

UpdatePersProperty:

   PersRec(1).MCVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).MHVALUE = 0
   PersRec(1).PERSVAL = 0
   PersRec(1).CValue = 0

   IF Vtype$ = "M" THEN
    PersRec(1).MTVALUE = VAL(a$(28))
    GOTO upd1
   END IF
   IF Vtype$ = "H" OR Vtype$ = "I" THEN
    PersRec(1).MHVALUE = VAL(a$(28))
    GOTO upd1
   END IF
   
   PersRec(1).PERSVAL = VAL(a$(28))
upd1:
   TVal# = TVal# + VAL(a$(28))
   
   mth = VAL(a$(19))
   MTHS$ = LTRIM$(STR$(mth))

   IF LEN(MTHS$) = 1 THEN MTHS$ = " " + MTHS$
   PersRec(1).PROPPIN = LTRIM$(STR$(PropNumb))
   PersRec(1).PROPDATE = Date2Num%("11-19-2003")
   PersRec(1).CValue = 0
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LATELIST = "N"
    PersRec(1).DESC1 = a$(22)
    PersRec(1).DESC2 = a$(18)
    PersRec(1).DESC3 = a$(24)
    PersRec(1).DESC4 = LEFT$(a$(17), 4)
    IF a$(34) = "Y" THEN
     PersRec(1).DESC5 = "Y" + MTHS$ + STRING$(22, " ") + "2003"
    ELSE
     PersRec(1).DESC5 = "N" + MTHS$ + STRING$(22, " ") + "2003"
    END IF
    
   

   PersRec(1).CustPin = Record!
   PersRec(1).NextRec = 0
   PersRec(1).LastYrPrinted = 2002
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

