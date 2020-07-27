DEFINT A-Z
'Update Personal Prop Detail Records from Txt file
DECLARE SUB BalanceListing ()
DECLARE SUB OpenTaxCustFile (NUMOFTAXRECS%, TAXFILE%)
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
  DIM SHARED TBCustRec(1) AS TBCustType

  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
  DIM A$(25)
   OpenTaxCustFile NUMOFTAXRECS, TAXFILE
   OpenTaxPersFile NumOfPersRecs, PersTaxFile

   OPEN "PPProp1.txt" FOR INPUT AS #11
1 FOR CC = 1 TO 25
  INPUT #11, A$(CC)
  NEXT CC
  c = c + 1
  LOCATE 1, 1: PRINT c



   FOR Cnt& = 1 TO NUMOFTAXRECS
   GET TAXFILE, Cnt&, TaxCustRec(1)
   nf = 1
   IF VAL(TaxCustRec(1).CountyAcctString) = VAL(A$(2)) THEN
    GOSUB updatepersonal
    nf = 0
    EXIT FOR
   END IF
   NEXT Cnt&

Nextone:
   GOTO 1
   CLOSE
   EXIT SUB


updatepersonal:
    

   VType = VAL(A$(4))


   PersRec(1).MCVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).MHVALUE = 0
   PersRec(1).PERSVAL = 0

   IF VType = 12 THEN
    PersRec(1).MCVALUE = VAL(A$(5))
    GOTO upd1
   END IF
   IF VType = 13 THEN
    PersRec(1).MTVALUE = VAL(A$(5))
    GOTO upd1
   END IF
   IF VType = 11 THEN
    PersRec(1).MHVALUE = VAL(A$(5))
    GOTO upd1
   END IF
   IF VType = 6 THEN
    PersRec(1).CVALUE = VAL(A$(5))
    GOTO upd1
   END IF

   PersRec(1).PERSVAL = VAL(A$(5))
upd1:
   PersRec(1).PROPPIN = LTRIM$(STR$(PropNumb))
   PersRec(1).PROPDATE = Date2Num%("11-01-1998")
   PersRec(1).CVALUE = 0
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LATELIST = "N"
   IF VType = 1 OR VType = 2 THEN
    PersRec(1).DESC1 = A$(14)
    PersRec(1).DESC2 = A$(11) + "/" + A$(12) + "/" + A$(13)
    PersRec(1).DESC3 = A$(15)
    PersRec(1).DESC4 = A$(10)
    IF UCASE$(A$(20)) = "P" THEN
     PersRec(1).DESC5 = "Y" + "/"
    ELSE
     PersRec(1).DESC5 = "N" + "/"
    END IF
    ELSE
    PersRec(1).DESC1 = A$(14)
    PersRec(1).DESC2 = A$(11) + "/" + A$(12) + "/" + A$(13)
    PersRec(1).DESC3 = A$(15)
    PersRec(1).DESC4 = A$(10)
    PersRec(1).DESC5 = "N"
   END IF
   
   PersRec(1).CustPin = Cnt&
   PersRec(1).NextRec = 0
   PersRec(1).LastYrPrinted = 1998
   PersRec(1).Deleted = 0
   PersRec(1).Blank = ""
   PRecord& = LOF(PersTaxFile) / LEN(PersRec(1)) + 1
   PUT PersTaxFile, PRecord&, PersRec(1)
   RETURN



END SUB

SUB OpenTaxCustFile (NUMOFTAXRECS, TAXFILE)
  
  TAXFILE = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TAXFILE LEN = LEN(TaxCustRec(1))
  NUMOFTAXRECS = LOF(TAXFILE) / LEN(TaxCustRec(1))
  
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

