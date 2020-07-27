DEFINT A-Z
'Update Personal Prop Detail Records from Txt file Gordonsville 2000 OTXLNK.DAT
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
  DIM SHARED PersRec(1) AS PersonalRecType
  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
  DIM a$(25)
   OpenTaxCustFile NUMOFTAXRECS, TAXFILE
   OpenTaxPersFile NumOfPersRecs, PersTaxFile

   OPEN "OTXLNK.DAT" FOR INPUT AS #11
1 LINE INPUT #11, a$
   PRINT a$
   PRINT STRING$(79, "-")
     Acct! = VAL(MID$(a$, 6, 8))
     Typ$ = MID$(a$, 19, 2)
     Amt# = VAL(MID$(a$, 49, 10))
     Make$ = RTRIM$(MID$(a$, 70, 10))
     Model$ = RTRIM$(MID$(a$, 80, 15))
     Body$ = RTRIM$(MID$(a$, 95, 15))
     Yr$ = MID$(a$, 111, 4)
     Com$ = MID$(a$, 126, 25)
     Wght$ = LTRIM$(MID$(a$, 234, 10))
     PPTRA$ = MID$(a$, 564, 1)

   IF Amt# = 0 THEN GOTO 1
   t = t + 1


   FOR Cnt& = 1 TO NUMOFTAXRECS
   GET TAXFILE, Cnt&, TaxCustRec(1)
   NF = 1
   'IF Acct! = 633 THEN PRINT TaxCustRec(1).CountyAcct: SLEEP 1
   IF TaxCustRec(1).CountyAcct = Acct! THEN
    GOSUB updatepersonal
    NF = 0
    EXIT FOR
   END IF
   NEXT Cnt&
IF NF = 1 THEN STOP
Nextone:
   GOTO 1
   CLOSE
   EXIT SUB


updatepersonal:
    
  

   PersRec(1).MCVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).MHVALUE = 0
   PersRec(1).PERSVAL = 0

   IF Typ$ = "MT" THEN
    t# = t# + Amt#
    PersRec(1).MTVALUE = Amt#
    GOTO upd1
   END IF
   IF Typ$ = "MH" THEN
    m# = m# + Amt#
    PersRec(1).MHVALUE = Amt#
    GOTO upd1
   END IF

   PersRec(1).PERSVAL = Amt#
   p# = p# + Amt#
upd1:
   PersRec(1).PROPPIN = LTRIM$(STR$(PropNumb))
   PersRec(1).PROPDATE = Date2Num%("11-01-2000")
   PersRec(1).CVALUE = 0
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LATELIST = "N"
   PersRec(1).DESC1 = RTRIM$(MID$(a$, 21, 20))
   PersRec(1).DESC2 = Make$ + "/" + Model$ + "/" + Body$
   PersRec(1).DESC3 = Yr$
   PersRec(1).DESC4 = Wght$
   PersRec(1).DESC5 = PPTRA$ + STRING$(24, " ") + "2000"
   PersRec(1).CustPin = Cnt&
   PersRec(1).NextRec = 0
   PersRec(1).LastYrPrinted = 1999
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

