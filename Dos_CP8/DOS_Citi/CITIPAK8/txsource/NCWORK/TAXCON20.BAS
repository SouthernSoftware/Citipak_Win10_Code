DEFINT A-Z
'Indian Trail 2001
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

  OPEN "Indiantr.dat" FOR INPUT AS #11
  CLS

10 LINE INPUT #11, a$

p = p + 1
LOCATE 3, 1: PRINT p
   IF VAL(MID$(a$, 118, 4)) = 0 THEN 10
 
   FOR Cnt& = 1 TO NumOfTaxRecs
     GET TaxFile, Cnt&, TaxCustRec(1)

   IF RTRIM$(TaxCustRec(1).CountyAcctString) = RTRIM$(LEFT$(a$, 11)) THEN

   'Redo Name
   'Split Name Here
   nm$ = RTRIM$(MID$(a$, 12, 34))
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
   TaxCustRec(1).OpenDate = Date2Num(DATE$)
   TaxCustRec(1).FName = FF1$
   TaxCustRec(1).LName = LN$
   TaxCustRec(1).SName = LN$
   add1$ = RTRIM$(MID$(a$, 46, 12))
   IF LEN(add1$) = 0 THEN
    TaxCustRec(1).Addr1 = MID$(a$, 58, 25)
    TaxCustRec(1).Addr2 = ""
   ELSE
    TaxCustRec(1).Addr1 = MID$(a$, 46, 12)
    TaxCustRec(1).Addr2 = MID$(a$, 58, 25)
   END IF
   TaxCustRec(1).City = MID$(a$, 83, 22)
   TaxCustRec(1).Zip = MID$(a$, 105, 5)
   TaxCustRec(1).CSSN = LEFT$(a$, 11)
   
   PUT TaxFile, Cnt&, TaxCustRec(1)







     RealValue# = VAL(MID$(a$, 169, 9))
    IF RealValue# > 0 THEN
      GOSUB UpdateReal
    END IF
     PersValue# = VAL(MID$(a$, 213, 9))
    IF PersValue# > 0 THEN
      GOSUB UpdatePers
    END IF
   GOTO 10

END IF
NEXT Cnt&

'Must Add Record

   Record = (LOF(TaxFile) / LEN(TaxCustRec(1))) + 1
   Cnt& = Record
   'Split Name Here
   nm$ = RTRIM$(MID$(a$, 12, 34))
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
   TaxCustRec(1).OpenDate = Date2Num(DATE$)
   
   TaxCustRec(1).Acct = Record
   TaxCustRec(1).FName = FF1$
   TaxCustRec(1).LName = LN$
   TaxCustRec(1).SName = LN$
   add1$ = RTRIM$(MID$(a$, 46, 12))
   IF LEN(add1$) = 0 THEN
    TaxCustRec(1).Addr1 = MID$(a$, 58, 25)
    TaxCustRec(1).Addr2 = ""
   ELSE
    TaxCustRec(1).Addr1 = MID$(a$, 46, 12)
    TaxCustRec(1).Addr2 = MID$(a$, 58, 25)
   END IF
   TaxCustRec(1).City = MID$(a$, 83, 22)
   TaxCustRec(1).State = ""
   TaxCustRec(1).Zip = MID$(a$, 105, 5)
   TaxCustRec(1).HPhone = ""
   TaxCustRec(1).WPhone = ""
   TaxCustRec(1).CSSN = LEFT$(a$, 11)
   TaxCustRec(1).SSSN = ""
   TaxCustRec(1).Active = "Y"
   TaxCustRec(1).Interest = "Y"
   TaxCustRec(1).TaxExempt = "N"
   TaxCustRec(1).Penalty = "N"
   TaxCustRec(1).CountyAcctString = LEFT$(a$, 11)
   TaxCustRec(1).LastTrans = 0
   TaxCustRec(1).FirstPropRec = 0
   TaxCustRec(1).FirstPersRec = 0
   TaxCustRec(1).PIN = Record
   TaxCustRec(1).Deleted = 0
   TaxCustRec(1).FileVer = 8
   PUT TaxFile, Record, TaxCustRec(1)
   
   RealValue# = VAL(MID$(a$, 169, 9))
   IF RealValue# > 0 THEN
    GOSUB UpdateReal
   END IF

   PersValue# = VAL(MID$(a$, 213, 9))
   IF PersValue# > 0 THEN
    GOSUB UpdatePers
   END IF
   
   GOTO 10
   CLOSE
   EXIT SUB
  
  

UpdateReal:
    PropertyRec(1).REALPIN = MID$(a$, 246, 10)
    PropertyRec(1).PROPDATE = Date2Num%("09-01-2001")
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).MAP = ""
    PropertyRec(1).BLOCK = ""
    PropertyRec(1).LOTNUMB = ""
    PropertyRec(1).LOTACRE = ""
    PropertyRec(1).PROPSIZE = 0
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LATELIST = "N"
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).PROPVALU = VAL(MID$(a$, 169, 9))
    PropertyRec(1).EXMPSENI = VAL(MID$(a$, 164, 5))
    PropertyRec(1).EXMPOTHR = 0
    PropertyRec(1).PROPNOT1 = MID$(a$, 178, 28)
    PropertyRec(1).PROPNOT2 = MID$(a$, 256, 10)
    PropertyRec(1).PROPNOT3 = "County Acct " + LEFT$(a$, 11)
    PropertyRec(1).Fill1 = ""
    PropertyRec(1).CustPin = Cnt&
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 2000
    PropertyRec(1).Deleted = 0
    PropertyRec(1).Blank = ""
    NextRecord = (LOF(PropTaxFile) / LEN(PropertyRec(1))) + 1
    PUT PropTaxFile, NextRecord, PropertyRec(1)
    RETURN

UpdatePers:
   PersRec(1).PROPPIN = RTRIM$(LEFT$(a$, 11))
   PersRec(1).PROPDATE = Date2Num%("09-01-2001")
   PersRec(1).PERSVAL = VAL(MID$(a$, 213, 9))
   PersRec(1).MHVALUE = 0
   PersRec(1).MCVALUE = 0
   PersRec(1).CVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LATELIST = "N"
   PersRec(1).DESC1 = "County Acct " + LEFT$(a$, 11)
   PersRec(1).DESC2 = ""
   PersRec(1).DESC3 = ""
   PersRec(1).DESC4 = ""
   PersRec(1).DESC5 = ""
   PersRec(1).CustPin = Cnt&
   PersRec(1).NextRec = 0
   PersRec(1).LastYrPrinted = 2000
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

