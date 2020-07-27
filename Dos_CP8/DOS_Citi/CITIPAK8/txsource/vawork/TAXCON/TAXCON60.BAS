DEFINT A-Z
'Update Personal Prop Detail Records from Txt file  New Market Va  2000
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
  DIM a$(87)
  OpenTaxCustFile NUMOFTAXRECS, TAXFILE
  OpenTaxPersFile NumOfPersRecs, PersTaxFile

  OPEN "DST05PP.TXT" FOR INPUT AS #11
1 FOR CC = 1 TO 87
  INPUT #11, a$(CC)
  NEXT CC
   FOR Cnt& = 1 TO NUMOFTAXRECS
    GET TAXFILE, Cnt&, TaxCustRec(1)
    nf = 1
   IF TaxCustRec(1).CountyAcct = 929 THEN
   STOP
   ' GOSUB updatepersonal
    nf = 0
    EXIT FOR
   END IF
   NEXT Cnt&
    IF nf = 1 THEN STOP
Nextone:
   GOTO 1
   CLOSE
   EXIT SUB


updatepersonal:

  kk = INSTR(a$(10), "STATUTORY")
  IF kk > 0 THEN RETURN
  kk = INSTR(a$(11), "STATUTORY")
  IF kk > 0 THEN RETURN
  kk = INSTR(a$(12), "STATUTORY")
  IF kk > 0 THEN RETURN


   PersRec(1).CVALUE = 0
   PersRec(1).MCVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).MHVALUE = 0
   PersRec(1).PERSVAL = 0

   IF a$(5) = "MH" THEN
    PersRec(1).MHVALUE = VAL(a$(8))
    GOTO upd1
   END IF
   IF a$(5) = "MC" THEN
    PersRec(1).MCVALUE = VAL(a$(8))
    GOTO upd1
   END IF
   IF a$(5) = "MT" THEN
    PersRec(1).MTVALUE = VAL(a$(8))
    GOTO upd1
   END IF
   PersRec(1).PERSVAL = VAL(a$(8))
upd1:
   a# = a# + VAL(a$(8))
   PersRec(1).PROPPIN = ""
   PersRec(1).PROPDATE = Date2Num%("01-01-2000")
   PersRec(1).CVALUE = 0
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LATELIST = "N"
   PersRec(1).DESC1 = a$(6)
   PersRec(1).DESC2 = a$(10) + "/" + a$(11) + "/" + a$(12)
   PersRec(1).DESC3 = a$(29)
   PersRec(1).DESC4 = a$(13)
   IF a$(79) = "Y" THEN
    PersRec(1).DESC5 = "Y" + "/" + STRING$(24, " ") + "2000"
   ELSE
    PersRec(1).DESC5 = "N" + "/" + STRING$(24, " ") + "2000"
   END IF
   PersRec(1).CustPin = Cnt&
   PersRec(1).NextRec = 0
   PersRec(1).LastYrPrinted = 0
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

