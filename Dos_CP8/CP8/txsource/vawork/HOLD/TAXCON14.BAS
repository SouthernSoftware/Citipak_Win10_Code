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
  TAXFILE = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TAXFILE LEN = LEN(TaxCustRec(1))
  NUMOFTAXRECS = LOF(TAXFILE) / LEN(TaxCustRec(1))

  TaxFile1 = FREEFILE
  OPEN "PPTXCUS1.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile1 LEN = LEN(TaxCustRec(1))
  

  PersTaxFile = FREEFILE
  OPEN "TAXPERS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PersTaxFile LEN = LEN(PersRec(1))
  NumOfPersRecs = LOF(PersTaxFile) / LEN(PersRec(1))

  PersTaxFile1 = FREEFILE
  OPEN "TAXPERS1.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PersTaxFile1 LEN = LEN(PersRec(1))
  


   FOR Cnt& = 1 TO NUMOFTAXRECS
   GET TAXFILE, Cnt&, TaxCustRec(1)
   IF TaxCustRec(1).Deleted <> -1 THEN
    GOSUB updatepersonal
   END IF
   NEXT Cnt&

Nextone:
   CLOSE
   EXIT SUB


updatepersonal:
   NumOfTaxRecs1 = LOF(TaxFile1) / LEN(TaxCustRec(1)) + 1
   PUT TaxFile1, NumOfTaxRecs1, TaxCustRec(1)

   PRecord& = TaxCustRec(1).FirstPersRec
   WHILE PRecord& <> 0
   GET PersTaxFile, PRecord&, PersRec(1)
   PersRec(1).CustPin = NumOfTaxRecs1
   PRecord& = PersRec(1).NextRec

   NumOfPersRecs1 = LOF(PersTaxFile1) / LEN(PersRec(1)) + 1
   PUT PersTaxFile1, NumOfPersRecs1, PersRec(1)
   WEND
   
   RETURN



END SUB

SUB OpenTaxCustFile (NUMOFTAXRECS, TAXFILE)
  
  TAXFILE = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TAXFILE LEN = LEN(TaxCustRec(1))
  NUMOFTAXRECS = LOF(TAXFILE) / LEN(TaxCustRec(1))
  TaxFile1 = FREEFILE
  OPEN "PPTXCUS1.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile1 LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs1 = LOF(TaxFile1) / LEN(TaxCustRec(1))
  
END SUB

SUB OpenTaxPersFile (NumOfPersRecs, PersTaxFile)
  PersTaxFile = FREEFILE
  OPEN "TAXPERS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PersTaxFile LEN = LEN(PersRec(1))
  NumOfPersRecs = LOF(PersTaxFile) / LEN(PersRec(1))
  PersTaxFile1 = FREEFILE
  OPEN "TAXPERS1.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PersTaxFile1 LEN = LEN(PersRec(1))
  NumOfPersRecs1 = LOF(PersTaxFile1) / LEN(PersRec(1))
  
END SUB

SUB OpenTaxPropFile (NumOfPropRecs, PropTaxFile)
  PropTaxFile = FREEFILE
  OPEN "TAXPROP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PropTaxFile LEN = LEN(PropertyRec(1))
  NumOfPropRecs = LOF(PropTaxFile) / LEN(PropertyRec(1))
END SUB

