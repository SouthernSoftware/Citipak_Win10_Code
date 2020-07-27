DEFINT A-Z
'Update Personal Prop Detail Records from Txt file  New Market Va  2002
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
  DIM A$(87)
  OpenTaxCustFile NUMOFTAXRECS, TAXFILE
  OpenTaxPersFile NumOfPersRecs, PersTaxFile

  OPEN "NEWMRKPD.TXT" FOR INPUT AS #11
1 LINE INPUT #11, A$
   A$(3) = MID$(A$, 6, 8)
   A$(5) = MID$(A$, 19, 2)
   A$(8) = MID$(A$, 49, 10)
   A$(6) = RTRIM$(MID$(A$, 21, 18))
   A$(10) = RTRIM$(MID$(A$, 70, 10))
   A$(11) = RTRIM$(MID$(A$, 80, 15))
   A$(12) = RTRIM$(MID$(A$, 95, 15))
   A$(13) = MID$(A$, 111, 4)
   A$(29) = MID$(A$, 237, 7)
   A$(79) = MID$(A$, 564, 1)




   FOR Cnt& = 1 TO NUMOFTAXRECS
    GET TAXFILE, Cnt&, TaxCustRec(1)
    nf = 1
   IF TaxCustRec(1).CountyAcct = VAL(A$(3)) THEN
    GOSUB updatepersonal
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

'  kk = INSTR(A$(10), "STATUTORY")
'  IF kk > 0 THEN RETURN
'  kk = INSTR(A$(11), "STATUTORY")
'  IF kk > 0 THEN RETURN
'  kk = INSTR(A$(12), "STATUTORY")
'  IF kk > 0 THEN RETURN


   PersRec(1).CVALUE = 0
   PersRec(1).MCVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).MHVALUE = 0
   PersRec(1).PERSVAL = 0

   IF A$(5) = "MH" THEN
    PersRec(1).MHVALUE = VAL(A$(8))
    GOTO upd1
    mh# = mh# + VAL(A$(8))
   END IF
   IF A$(5) = "MC" THEN
    PersRec(1).MCVALUE = VAL(A$(8))
    RETURN
   END IF
   IF A$(5) = "MT" THEN
    PersRec(1).MTVALUE = VAL(A$(8))
    mt# = mt# + VAL(A$(8))
    GOTO upd1
   END IF
   PersRec(1).PERSVAL = VAL(A$(8))
   CAR# = CAR# + VAL(A$(8))
upd1:
   A# = A# + VAL(A$(8))
   PersRec(1).PROPPIN = ""
   PersRec(1).PROPDATE = Date2Num%("01-01-2002")
   PersRec(1).CVALUE = 0
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LATELIST = "N"
   PersRec(1).DESC1 = A$(6)
   PersRec(1).DESC2 = A$(10) + "/" + A$(11) + "/" + A$(12)
   PersRec(1).DESC3 = A$(29)
   PersRec(1).DESC4 = A$(13)
   IF A$(79) = "Y" THEN
    PersRec(1).DESC5 = "Y" + "/" + STRING$(24, " ") + "2002"
   ELSE
    PersRec(1).DESC5 = "N" + "/" + STRING$(24, " ") + "2002"
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

