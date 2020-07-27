DEFINT A-Z
'Convert Data From County (Pink Hill 2002)
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
  


  
  STACK 5000
  BalanceListing
  
  END

DEFDBL A
SUB BalanceListing
DIM a$(20)
  
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile




  OPEN "taxtap.dat" FOR INPUT AS #10
    Record = 0
    CLS
10 LINE INPUT #10, a$
   



    RealVal# = VAL(MID$(a$, 198, 7))
    PersVal# = VAL(MID$(a$, 205, 8)) + VAL(MID$(a$, 213, 8)) + VAL(MID$(a$, 221, 8)) + VAL(MID$(a$, 229, 8)) + VAL(MID$(a$, 237, 8))



  FOR x = 1 TO NumOfTaxRecs
  GET TaxFile, x, TaxCustRec(1)
  IF VAL(a$(8)) = VAL(TaxCustRec(1).CountyAcctString) THEN
           Record! = x
           IF RealVal# > 0 THEN GOSUB UpdateReal
           IF PersVal# > 0 THEN GOSUB updatepers
           GOTO 10
  END IF
  NEXT x
  
nm$ = MID$(a$, 30, 30)
IF PersVal# > 0 THEN
LN$ = nm$
FSN$ = ""
SN$ = nm$
ELSE
   kk = INSTR(nm$, " ")
   IF kk > 0 THEN
     LN$ = LEFT$(nm$, kk)
     FSN$ = MID$(nm$, kk + 1, (LEN(nm$)) - kk)
     SN$ = LN$
      ELSE
     LN$ = nm$
     FSN$ = ""
     SN$ = nm$
   END IF
END IF
 

     Record! = LOF(TaxFile) / LEN(TaxCustRec(1)) + 1
     TaxCustRec(1).Acct = Record!
     TaxCustRec(1).OpenDate = Date2Num(DATE$)
     TaxCustRec(1).FName = LTRIM$(FSN$)
     TaxCustRec(1).LName = LTRIM$(LN$)
     TaxCustRec(1).SName = LTRIM$(SN$)

     IF LEN(QPTrim$(MID$(a$, 60, 30))) = 0 THEN
       TaxCustRec(1).Addr1 = QPTrim$(MID$(a$, 90, 30))
       TaxCustRec(1).Addr2 = ""
      ELSE
       TaxCustRec(1).Addr1 = QPTrim$(MID$(a$, 60, 30))
       TaxCustRec(1).Addr2 = QPTrim$(MID$(a$, 90, 30))
     END IF
     TaxCustRec(1).City = QPTrim$(MID$(a$, 120, 22))
     TaxCustRec(1).STATE = RIGHT$(QPTrim$(TaxCustRec(1).City), 2)
     IF LEN(QPTrim$(TaxCustRec(1).City)) > 0 THEN TaxCustRec(1).City = LEFT$(QPTrim$(TaxCustRec(1).City), LEN(QPTrim$(TaxCustRec(1).City)) - 2)
     TaxCustRec(1).Zip = QPTrim$(MID$(a$, 24, 5))
     TaxCustRec(1).CountyAcct = 0
     TaxCustRec(1).CountyAcctString = QPTrim$(MID$(a$, 363, 6))
     TaxCustRec(1).HPHONE = ""
     TaxCustRec(1).WPHONE = ""
     TaxCustRec(1).CSSN = ""
     TaxCustRec(1).SSSN = ""
     TaxCustRec(1).Active = "Y"
     TaxCustRec(1).Interest = "Y"
     TaxCustRec(1).TaxExempt = "N"
     TaxCustRec(1).Penalty = "N"
     TaxCustRec(1).LastTrans = 0
     TaxCustRec(1).FirstPropRec = 0
     TaxCustRec(1).FirstPersRec = 0
     TaxCustRec(1).Pin = Record!
     TaxCustRec(1).Deleted = 0
     TaxCustRec(1).FileVer = 8
     PUT TaxFile, Record!, TaxCustRec(1)
     IF RealVal# > 0 THEN GOSUB UpdateReal
     IF PersVal# > 0 THEN GOSUB updatepers

    GOTO 10


UpdateReal:
    PropertyRec(1).REALPIN = QPTrim$(MID$(a$, 143, 12))
    PropertyRec(1).PROPDATE = Date2Num%(DATE$)
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).MAP = ""
    PropertyRec(1).BLOCK = ""
    PropertyRec(1).LOTNUMB = ""
    PropertyRec(1).LOTACRE = "A"
    PropertyRec(1).PROPSIZE = VAL(MID$(a$, 186, 4))
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LATELIST = "N"
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).PROPVALU = RealVal#
    PropertyRec(1).EXMPSENI = 0
    PropertyRec(1).EXMPOTHR = 0
    PropertyRec(1).PROPNOT1 = QPTrim$(MID$(a$, 158, 28))
    PropertyRec(1).PROPNOT2 = ""
    PropertyRec(1).PROPNOT3 = ""
    PropertyRec(1).Fill1 = ""
    PropertyRec(1).CustPin = Record!
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 0
    PropertyRec(1).Deleted = 0
    PropertyRec(1).Blank = ""
    NextRecord = (LOF(PropTaxFile) / LEN(PropertyRec(1))) + 1
    PUT PropTaxFile, NextRecord, PropertyRec(1)
    RETURN

updatepers:
   PersRec(1).PROPPIN = "1"
   PersRec(1).PROPDATE = Date2Num%(DATE$)
   PersRec(1).PersVal = PersVal#
   PersRec(1).MHVALUE = 0
   PersRec(1).MCVALUE = 0
   PersRec(1).CVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   IF MID$(a$, 142, 1) = "Y" THEN
   PersRec(1).LATELIST = "Y"
   ELSE
   PersRec(1).LATELIST = "N"
   END IF
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

