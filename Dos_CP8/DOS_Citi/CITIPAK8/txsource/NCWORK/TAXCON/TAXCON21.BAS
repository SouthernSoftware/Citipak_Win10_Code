DEFINT A-Z
'Indian Trail 2003 (1st save .xls as .csv file then convert)
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

  OPEN "Indian~1.csv" FOR INPUT AS #11
  CLS

1 CLS
  FOR cnt1 = 1 TO 26
   INPUT #11, a$(cnt1)
  ' PRINT a$(cnt1)
  NEXT cnt1
  

   tr# = VAL(a$(16))
   te# = VAL(a$(15))
   tp# = VAL(a$(19))
   IF tr# = 0 AND te# = 0 AND tp# = 0 THEN GOTO 1
         
     LOCATE 3, 1
    PRINT "new = "; new
    PRINT "Updated= "; notnew


   a$(1) = LTRIM$(RTRIM$(a$(1)))
   alt$ = a$(1)
     
   IF ASC(RIGHT$(a$(2), 1)) > 64 THEN
    b! = ASC(RIGHT$(a$(2), 1))
    alt$ = LEFT$(a$(2), (LEN(a$(2)) - 1)) + LTRIM$(STR$(b! - 64))
   END IF


FOR Cnt& = 1 TO NumOfTaxRecs
     GET TaxFile, Cnt&, TaxCustRec(1)
     cssn$ = RTRIM$(TaxCustRec(1).cssn)
  
   IF RTRIM$(TaxCustRec(1).CountyAcctString) = a$(1) OR RTRIM$(cssn$) = a$(1) THEN
  
   'Redo Name
   'Split Name Here
   nm$ = RTRIM$(a$(2))
   kk1 = INSTR(nm$, " INC")
   kk2 = INSTR(nm$, " CO")
   kk3 = INSTR(nm$, " CORP")
   IF kk1 > 0 OR kk2 > 0 OR kk3 > 0 THEN
    LN$ = RTRIM$(nm$)
    SN$ = RTRIM$(nm$)
    FF1$ = ""
   GOTO NameUpd
   END IF

   kk = INSTR(nm$, " ")
   IF kk > 0 THEN
    LN$ = LEFT$(nm$, kk - 1)
    SN$ = LEFT$(nm$, kk - 1)
    FF1$ = MID$(nm$, kk + 1, LEN(nm$) - (kk + 1))
    ELSE
    LN$ = RTRIM$(nm$)
    SN$ = RTRIM$(nm$)
    FF1$ = ""

   END IF

                                                    
NameUpd:

   city$ = RTRIM$(a$(5))
   State$ = RIGHT$(city$, 2)
   
    IF LEN(city$) > 0 THEN city$ = RTRIM$(LEFT$(city$, LEN(city$) - 2))
   

   TaxCustRec(1).OpenDate = Date2Num(DATE$)
   TaxCustRec(1).FName = FF1$
   TaxCustRec(1).LName = LN$
   TaxCustRec(1).SName = LN$
   add1$ = RTRIM$(a$(3))
   IF LEN(add1$) = 0 THEN
    TaxCustRec(1).Addr1 = a$(4)
    TaxCustRec(1).Addr2 = ""
   ELSE
    TaxCustRec(1).Addr1 = a$(3)
    TaxCustRec(1).Addr2 = a$(4)
   END IF
   TaxCustRec(1).city = city$
   TaxCustRec(1).State = State$
   TaxCustRec(1).Zip = LEFT$(a$(6), 5)
   TaxCustRec(1).cssn = LTRIM$(a$(1))
   
   
   PUT TaxFile, Cnt&, TaxCustRec(1)


     RealValue# = VAL(a$(16))
    IF RealValue# > 0 THEN
      GOSUB UpdateReal
    END IF
     PersValue# = VAL(a$(19))
    IF PersValue# > 0 THEN
      GOSUB UpdatePers
    END IF
   notnew = notnew + 1
   GOTO 1

END IF
NEXT Cnt&
new = new + 1

  




'Must Add Record

   Record = (LOF(TaxFile) / LEN(TaxCustRec(1))) + 1
   Cnt& = Record
   'Split Name Here
   nm$ = RTRIM$(a$(2))
   kk1 = INSTR(nm$, " INC ")
   kk2 = INSTR(nm$, " CO ")
   kk3 = INSTR(nm$, " CORP")
   IF kk1 > 0 OR kk2 > 0 OR kk3 > 0 THEN
    LN$ = RTRIM$(nm$)
    SN$ = RTRIM$(nm$)
    FF1$ = ""
   GOTO NameUpd1
   END IF
   
   kk = INSTR(nm$, " ")
   IF kk > 0 THEN
    LN$ = LEFT$(nm$, kk - 1)
    SN$ = LEFT$(nm$, kk - 1)
    FF1$ = MID$(nm$, kk + 1, LEN(nm$) - (kk + 1))
    ELSE
    LN$ = RTRIM$(nm$)
    SN$ = RTRIM$(nm$)
    FF1$ = ""

   END IF

NameUpd1:

   city$ = RTRIM$(a$(5))
   State$ = RIGHT$(city$, 2)
   
   IF LEN(city$) > 2 THEN city$ = RTRIM$(LEFT$(city$, LEN(city$) - 2))

   TaxCustRec(1).Acct = Record
   TaxCustRec(1).OpenDate = Date2Num(DATE$)
   TaxCustRec(1).FName = FF1$
   TaxCustRec(1).LName = LN$
   TaxCustRec(1).SName = LN$
   add1$ = RTRIM$(a$(3))
   IF LEN(add1$) = 0 THEN
    TaxCustRec(1).Addr1 = a$(4)
    TaxCustRec(1).Addr2 = ""
   ELSE
    TaxCustRec(1).Addr1 = a$(3)
    TaxCustRec(1).Addr2 = a$(4)
   END IF
   TaxCustRec(1).city = city$
   TaxCustRec(1).State = State$
   TaxCustRec(1).Zip = a$(6)
   TaxCustRec(1).HPhone = ""
   TaxCustRec(1).WPhone = ""
   TaxCustRec(1).cssn = LTRIM$(a$(1))
   TaxCustRec(1).SSSN = ""
   TaxCustRec(1).Active = "Y"
   TaxCustRec(1).Interest = "Y"
   TaxCustRec(1).TaxExempt = "N"
   TaxCustRec(1).Penalty = "N"
   TaxCustRec(1).CountyAcctString = LTRIM$(a$(1))
   TaxCustRec(1).LastTrans = 0
   TaxCustRec(1).FirstPropRec = 0
   TaxCustRec(1).FirstPersRec = 0
   TaxCustRec(1).PIN = Record
   TaxCustRec(1).Deleted = 0
   TaxCustRec(1).FileVer = 8
   PUT TaxFile, Record, TaxCustRec(1)
   
   RealValue# = VAL(a$(16))
   IF RealValue# > 0 THEN
    GOSUB UpdateReal
   END IF

   PersValue# = VAL(a$(19))
   IF PersValue# > 0 THEN
    GOSUB UpdatePers
   END IF
   
   GOTO 1
   CLOSE
   EXIT SUB
  
  

UpdateReal:
    PropertyRec(1).REALPIN = a$(1)
    PropertyRec(1).PROPDATE = Date2Num%(DATE$)
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).MAP = ""
    PropertyRec(1).BLOCK = ""
    PropertyRec(1).LOTNUMB = ""
    PropertyRec(1).LOTACRE = ""
    PropertyRec(1).PROPSIZE = 0
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LATELIST = "N"
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).PROPVALU = VAL(a$(16))
    PropertyRec(1).EXMPSENI = VAL(a$(15))
    PropertyRec(1).EXMPOTHR = 0
    PropertyRec(1).PROPNOT1 = a$(17)
    PropertyRec(1).PROPNOT2 = ""
    PropertyRec(1).PROPNOT3 = "County Acct " + LTRIM$(a$(1))
    
    PropertyRec(1).CustPin = Cnt&
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 2002
    PropertyRec(1).Deleted = 0
    PropertyRec(1).Blank = ""
    NextRecord = (LOF(PropTaxFile) / LEN(PropertyRec(1))) + 1
    PUT PropTaxFile, NextRecord, PropertyRec(1)
    RETURN

UpdatePers:
   PersRec(1).PROPPIN = LTRIM$(a$(1))
   PersRec(1).PROPDATE = Date2Num%(DATE$)
   PersRec(1).PERSVAL = VAL(a$(19))
   PersRec(1).MHVALUE = 0
   PersRec(1).MCVALUE = 0
   PersRec(1).CVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   IF VAL(a$(11)) > 0 THEN
   PersRec(1).LATELIST = "Y"
   ELSE
   PersRec(1).LATELIST = "N"
   END IF
   PersRec(1).DESC1 = "County Acct " + LTRIM$(a$(1))
   PersRec(1).DESC2 = ""
   PersRec(1).DESC3 = ""
   PersRec(1).DESC4 = ""
   PersRec(1).DESC5 = ""
   PersRec(1).CustPin = Cnt&
   PersRec(1).NextRec = 0
   PersRec(1).LastYrPrinted = 2002
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

