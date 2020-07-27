DEFINT A-Z
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION Num2Date$ (DateNumber%)

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
  DIM SHARED HRec(1) AS HamletTaxRec


  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
  
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile

   OPEN "mastseq" FOR RANDOM ACCESS READ WRITE SHARED AS #11 LEN = 577
   Record = 0

   FOR x = 1 TO LOF(11) / 577
    GET 11, x, HRec(1)
    Nam$ = RTRIM$(HRec(1).Name)
   IF LEN(Nam$) > 0 AND (VAL(HRec(1).REAL) > 0 OR VAL(HRec(1).Pers) > 0) THEN
   
     kk = INSTR(HRec(1).Name, ",")

    IF kk > 1 THEN
     LNm$ = LEFT$(HRec(1).Name, kk - 1)
     FirstName$ = MID$(HRec(1).Name, kk + 1, 30 - kk)
    ELSE
     FirstName$ = ""
     LNm$ = HRec(1).Name
    END IF
    ThisName$ = QPTrim$(FirstName$)
    ThisNam1$ = QPTrim$(LNm$)
    FoundName$ = ThisName$ + ThisNam1$


    'Now Check For Previous Account and Add to It if Found
    IF Record > 0 THEN
    GET TaxFile, LOF(TaxFile) / LEN(TaxCustRec(1)), TaxCustRec(1)
    LastNam$ = QPTrim$(TaxCustRec(1).LName)
    FirstNam$ = QPTrim$(TaxCustRec(1).FName)
    SearchNam$ = FirstNam$ + LastNam$
    SearchNam$ = QPTrim$(SearchNam$)
    IF SearchNam$ = FoundName$ THEN
    IF VAL(HRec(1).REAL) > 0 THEN
     GOSUB UpdateReal
    END IF
    IF VAL(HRec(1).Pers) > 0 THEN
     GOSUB UpdatePers
    END IF
    GOTO NexRecord
    END IF
   END IF
    PRINT HRec(1).Add1
    PRINT HRec(1).Add2
    PRINT HRec(1).Add3
    PRINT HRec(1).Add4



    CITYState$ = QPTrim$(HRec(1).Add4)

    IF LEN(CITYState$) = 0 THEN
     CITYState$ = RTRIM$(HRec(1).Add3)
    END IF

   'Split Name Here
    TaxCustRec(1).Acct = Record
    TaxCustRec(1).OpenDate = Date2Num(DATE$)
    TaxCustRec(1).FName = LTRIM$(FirstName$)
    TaxCustRec(1).LName = LTRIM$(LNm$)
    TaxCustRec(1).SName = LTRIM$(LNm$)
    TaxCustRec(1).Addr1 = HRec(1).Add1
    TaxCustRec(1).Addr2 = ""
    TaxCustRec(1).CITY = CITYState$
    TaxCustRec(1).STATE = ""
    TaxCustRec(1).Zip = RIGHT$(HRec(1).Zip, 5)
    TaxCustRec(1).HPHONE = ""
    TaxCustRec(1).WPHONE = ""
    TaxCustRec(1).CSSN = HRec(1).SSN
    TaxCustRec(1).SSSN = HRec(1).SSN1
    TaxCustRec(1).Active = "Y"
    TaxCustRec(1).Interest = "Y"
    TaxCustRec(1).TaxExempt = "N"
    TaxCustRec(1).Penalty = "N"
    TaxCustRec(1).CountyAcct = VAL(HRec(1).Acct)
    TaxCustRec(1).LastTrans = 0
    TaxCustRec(1).FirstPropRec = 0
    TaxCustRec(1).FirstPersRec = 0
    TaxCustRec(1).PIN = Record
    TaxCustRec(1).Deleted = 0
    TaxCustRec(1).FileVer = 8
    Record = Record + 1
    PUT TaxFile, Record, TaxCustRec(1)

    IF VAL(HRec(1).REAL) > 0 THEN
     GOSUB UpdateReal
    END IF
    IF VAL(HRec(1).Pers) > 0 THEN
     GOSUB UpdatePers
    END IF
   END IF
   
NexRecord:
   NEXT x
   CLOSE
   EXIT SUB
  
  

UpdateReal:
    PropertyRec(1).REALPIN = ""
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
    PropertyRec(1).PROPVALU = VAL(HRec(1).REAL)
    PropertyRec(1).EXMPSENI = VAL(HRec(1).DISCNT)
    PropertyRec(1).EXMPOTHR = 0
    PropertyRec(1).PROPNOT1 = QPTrim$(HRec(1).Add2)
    PropertyRec(1).PROPNOT2 = HRec(1).PROPDESC
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
   PersRec(1).PROPDATE = Date2Num%(DATE$)
   PersRec(1).PERSVAL = VAL(HRec(1).Pers)
   PersRec(1).MHVALUE = 0
   PersRec(1).MCVALUE = 0
   PersRec(1).CVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LATELIST = "N"
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

