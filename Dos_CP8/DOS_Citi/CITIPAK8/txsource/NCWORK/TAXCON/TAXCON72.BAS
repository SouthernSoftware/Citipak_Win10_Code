DEFINT A-Z
'Convert Data From County (Maxton : (Robeson Cty Real/Pers 2003))
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
  
  
  DIM SHARED TAXCUSTREC(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED PersRec(1) AS PersonalRecType
  


  
  STACK 5000
  BalanceListing
  
  END

DEFDBL A
SUB BalanceListing
DIM a$(20)
  
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile

  OPEN "maxton.txt" FOR INPUT AS #10
    Record = 0
 CLS
10 LINE INPUT #10, a$
   'PRINT A$
   
                 
   SDPin$ = MID$(a$, 390, 14)

   C = C + 1
    RealVal# = VAL(MID$(a$, 199, 8))
    PERSVAL# = VAL(MID$(a$, 468, 8))
    EVal1# = VAL(MID$(a$, 6, 6))
    EVal2# = VAL(MID$(a$, 12, 6))
    CUSTACCT$ = MID$(a$, 374, 12)
  

  FOR x = 1 TO NumOfTaxRecs
  GET TaxFile, x, TAXCUSTREC(1)

 ' PRINT x, TAXCUSTREC(1).CountyAcctString

  IF RTRIM$(CUSTACCT$) = RTRIM$(TAXCUSTREC(1).CountyAcctString) THEN
           Record! = x
     TAXCUSTREC(1).OpenDate = Date2Num(DATE$)
     PRINT TAXCUSTREC(1).lnAME + " " + TAXCUSTREC(1).FNAME, MID$(a$, 30, 30)
    ' SLEEP 1
     IF LEN(RTRIM$(MID$(a$, 60, 30))) = 0 THEN
     TAXCUSTREC(1).Addr1 = RTRIM$(MID$(a$, 90, 30))
     TAXCUSTREC(1).Addr2 = ""
      ELSE
     TAXCUSTREC(1).Addr1 = RTRIM$(MID$(a$, 60, 30))
     TAXCUSTREC(1).Addr2 = RTRIM$(MID$(a$, 90, 30))
     END IF
     TAXCUSTREC(1).City = MID$(a$, 120, 21)
     TAXCUSTREC(1).STATE = MID$(a$, 141, 2)
     TAXCUSTREC(1).Zip = MID$(a$, 490, 5)
     PUT TaxFile, Record!, TAXCUSTREC(1)
     GOSUB UpdateReal
     GOSUB UpdatePers
     GOTO 10
  END IF
  NEXT x
     
     Record! = LOF(TaxFile) / LEN(TAXCUSTREC(1)) + 1

     TAXCUSTREC(1).Acct = Record!
     TAXCUSTREC(1).OpenDate = Date2Num(DATE$)

     CUSTNAME$ = RTRIM$(MID$(a$, 30, 30))
     KK = INSTR(CUSTNAME$, " ")
     IF KK > 0 THEN
       TAXCUSTREC(1).FNAME = LTRIM$(MID$(CUSTNAME$, KK, (LEN(CUSTNAME$) - KK)))
       TAXCUSTREC(1).lnAME = LEFT$(CUSTNAME$, KK - 1)
       TAXCUSTREC(1).SName = LEFT$(CUSTNAME$, KK - 1)
      ELSE
       TAXCUSTREC(1).FNAME = ""
       TAXCUSTREC(1).lnAME = RTRIM$(CUSTNAME$)
       TAXCUSTREC(1).SName = RTRIM$(CUSTNAME$)
     END IF
     TAXCUSTREC(1).STATE = MID$(a$, 141, 2)
     TAXCUSTREC(1).Zip = MID$(a$, 490, 5)
     IF LEN(RTRIM$(MID$(a$, 60, 30))) = 0 THEN
     TAXCUSTREC(1).Addr1 = RTRIM$(MID$(a$, 90, 30))
     TAXCUSTREC(1).Addr2 = ""
      ELSE
     TAXCUSTREC(1).Addr1 = RTRIM$(MID$(a$, 60, 30))
     TAXCUSTREC(1).Addr2 = RTRIM$(MID$(a$, 90, 30))
     END IF
     TAXCUSTREC(1).City = MID$(a$, 120, 21)
     TAXCUSTREC(1).STATE = MID$(a$, 141, 2)
     TAXCUSTREC(1).Zip = MID$(a$, 490, 5)
     TAXCUSTREC(1).CountyAcct = 0
     TAXCUSTREC(1).CountyAcctString = MID$(a$, 374, 12)
     TAXCUSTREC(1).HPHONE = ""
    TAXCUSTREC(1).WPHONE = ""
    TAXCUSTREC(1).CSSN = ""
    TAXCUSTREC(1).SSSN = ""
    TAXCUSTREC(1).Active = "Y"
    TAXCUSTREC(1).Interest = "Y"
    TAXCUSTREC(1).TaxExempt = "N"
    TAXCUSTREC(1).Penalty = "N"
    TAXCUSTREC(1).LastTrans = 0
    TAXCUSTREC(1).FirstPropRec = 0
    TAXCUSTREC(1).FirstPersRec = 0
    TAXCUSTREC(1).Pin = Record!
    TAXCUSTREC(1).Deleted = 0
    TAXCUSTREC(1).FileVer = 8
    PUT TaxFile, Record!, TAXCUSTREC(1)
    GOSUB UpdateReal
    GOSUB UpdatePers
    GOTO 10


UpdateReal:
    PropertyRec(1).REALPIN = SDPin$
    PropertyRec(1).PROPDATE = Date2Num%(DATE$)
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).MAP = MID$(a$, 390, 6)
    PropertyRec(1).BLOCK = ""
    PropertyRec(1).LOTNUMB = ""
    PropertyRec(1).LOTACRE = MID$(a$, 192, 1)
    PropertyRec(1).PROPSIZE = VAL(MID$(a$, 184, 8))
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LATELIST = MID$(a$, 144, 1)
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).PROPVALU = RealVal#
    PropertyRec(1).EXMPSENI = EVal2#
    PropertyRec(1).EXMPOTHR = 0
    PropertyRec(1).PROPNOT1 = MID$(a$, 160, 24)
    PropertyRec(1).PROPNOT2 = ""
    PropertyRec(1).PROPNOT3 = MID$(a$, 390, 24)
    PropertyRec(1).CustPin = Record!
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 0
    PropertyRec(1).Deleted = 0
    PropertyRec(1).Blank = ""
    NextRecord = (LOF(PropTaxFile) / LEN(PropertyRec(1))) + 1
    PUT PropTaxFile, NextRecord, PropertyRec(1)
    RETURN

UpdatePers:

   IF PERSVAL# <= 0 THEN RETURN

   PersRec(1).PROPPIN = "1"
   PersRec(1).PROPDATE = Date2Num%(DATE$)
   PersRec(1).PERSVAL = PERSVAL#
   PersRec(1).MHVALUE = 0
   PersRec(1).MCVALUE = 0
   PersRec(1).CVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LATELIST = "N"
   PersRec(1).Desc1 = ""
   PersRec(1).Desc2 = ""
   PersRec(1).Desc3 = ""
   PersRec(1).DESC4 = ""
   PersRec(1).DESC5 = ""

   PersRec(1).CustPin = Record!
   PersRec(1).NextRec = 0
   PersRec(1).LastYrPrinted = 0
   PersRec(1).Deleted = 0
   PersRec(1).Blank = ""
   NextRec = (LOF(PersTaxFile) / LEN(PersRec(1))) + 1
   PUT PersTaxFile, NextRec, PersRec(1)
   RETURN



END SUB

DEFINT A
SUB OpenTaxCustFile (NumOfTaxRecs, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "TAXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TAXCUSTREC(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TAXCUSTREC(1))
  
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

