DEFINT A-Z
'Update Personal Prop Detail Records from Txt file Shenandoah 2001
DECLARE SUB BalanceListing ()
DECLARE SUB OpenTaxCustFile (NUMOFTAXRECS%, TaxFile%)
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
  '$INCLUDE: 'TaxConv.Bi'
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED PersRec(1) AS PersonalRecType
  DIM SHARED Rec(1) AS SType
  

  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
   DIM A$(20)
   OpenTaxCustFile NUMOFTAXRECS, TaxFile
   OpenTaxPersFile NumOfPersRecs, PersTaxFile
   
   OPEN "PPDETAIL.TXT" FOR RANDOM ACCESS READ WRITE SHARED AS #11 LEN = 389
1  FOR Cnt& = 1 TO LOF(11) / 389
   LOCATE 1, 1: PRINT Cnt&, LOF(11) / 389
   GET 11, Cnt&, Rec(1)
   IF MID$(Rec(1).RecDesc, 78, 1) = "I" THEN GOTO nextone
   IF VAL(MID$(Rec(1).RecDesc, 59, 7)) = 0 THEN GOTO nextone

   Acct! = VAL(MID$(Rec(1).RecDesc, 5, 7))
 
   FOR Snt& = 1 TO NUMOFTAXRECS
   GET TaxFile, Snt&, TaxCustRec(1)
    IF TaxCustRec(1).CountyAcct = Acct! THEN
      GOSUB UpdatePers
      EXIT FOR
    END IF
   NEXT Snt&
nextone:
   NEXT Cnt&
   CLOSE
   EXIT SUB

UpdatePers:

   VType = VAL(MID$(Rec(1).RecDesc, 12, 2))


   PersRec(1).MCVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).MHVALUE = 0
   PersRec(1).PERSVAL = 0

   IF VType = 30 THEN
    PersRec(1).MTVALUE = VAL(MID$(Rec(1).RecDesc, 137, 9))
    GOTO upd1
   END IF
   IF VType = 50 THEN
    PersRec(1).MHVALUE = VAL(MID$(Rec(1).RecDesc, 137, 9))
    GOTO upd1
   END IF
   IF VType = 60 THEN
    PersRec(1).CVALUE = VAL(MID$(Rec(1).RecDesc, 137, 9))
    GOTO upd1
   END IF

   PersRec(1).PERSVAL = VAL(MID$(Rec(1).RecDesc, 137, 9))
upd1:
   PersRec(1).PROPPIN = LTRIM$(STR$(PropNumb))
   PersRec(1).PROPDATE = Date2Num%("11-01-2001")
   PersRec(1).CVALUE = 0
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LATELIST = "N"
   IF VType = 10 THEN
    PersRec(1).DESC1 = MID$(Rec(1).RecDesc, 16, 22)
    PersRec(1).DESC2 = MID$(Rec(1).RecDesc, 79, 30)
    PersRec(1).DESC3 = ""
    PersRec(1).DESC4 = MID$(Rec(1).RecDesc, 79, 4)
    IF MID$(Rec(1).RecDesc, 196, 1) = "Y" THEN
     PersRec(1).DESC5 = "Y" + "/" + STRING$(24, " ") + "2001"
    ELSE
     PersRec(1).DESC5 = "N" + "/" + STRING$(24, " ") + "2001"
    END IF
    ELSE
    PersRec(1).DESC1 = MID$(Rec(1).RecDesc, 16, 22)
    PersRec(1).DESC2 = MID$(Rec(1).RecDesc, 79, 30)
    PersRec(1).DESC3 = ""
    PersRec(1).DESC4 = ""
    PersRec(1).DESC5 = "N" + STRING$(24, " ") + "2001"
   END IF

   PersRec(1).CustPin = Snt&
   PersRec(1).NextRec = 0
   PersRec(1).LastYrPrinted = 2000
   PersRec(1).Deleted = 0
   PersRec(1).Blank = ""
   PRecord& = LOF(PersTaxFile) / LEN(PersRec(1)) + 1
   PUT PersTaxFile, PRecord&, PersRec(1)
   RETURN


RETURN


END SUB

SUB OpenTaxCustFile (NUMOFTAXRECS, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NUMOFTAXRECS = LOF(TaxFile) / LEN(TaxCustRec(1))
  
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

