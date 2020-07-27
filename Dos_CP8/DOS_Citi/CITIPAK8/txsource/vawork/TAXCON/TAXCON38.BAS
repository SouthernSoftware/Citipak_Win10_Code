
DEFINT A-Z
'Update Shenandoah Detail Records
DECLARE SUB BalanceListing ()
DECLARE SUB OpenTaxCustFile (NumOfTaxRecs%, TaxFile%)
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
  '$INCLUDE: 'TAXCONV.BI'
  
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PersRec(1) AS PersonalRecType
  
  

  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
  DIM a$(31)
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile
  OPEN "PPDETAIL.TXT" FOR RANDOM ACCESS READ WRITE SHARED AS #11 LEN = 389
  FIELD 11, 389 AS a$

  FOR CNT& = 1 TO LOF(11) / 389
  GET 11, CNT&
  PRINT LEFT$(a$, 78)
  acct$ = MID$(a$, 5, 7)
  'IF VAL(acct$) = 11401 THEN STOP
  IF MID$(a$, 78, 1) = "A" AND VAL(MID$(a$, 60, 6)) > 0 THEN
  ELSE
  GOTO nextone
  END IF
  
   'SEARCH RECORDS FOR MATCH
   FOR Snt& = 1 TO NumOfTaxRecs
   GET TaxFile, Snt&, TaxCustRec(1)
   IF VAL(acct$) = TaxCustRec(1).CountyAcct THEN
    AcctRecord! = Snt&
    GOSUB UpdatePersProperty
    EXIT FOR
   END IF
   NEXT Snt&
 
nextone:
   NEXT CNT&
    CLOSE
   EXIT SUB

UpdatePersProperty:

   Cat$ = MID$(a$, 12, 2)
   Amt# = VAL(MID$(a$, 137, 9))

  
   PersRec(1).MCVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).MHVALUE = 0
   PersRec(1).PersVal = 0
   PersRec(1).CVALUE = 0

   IF Cat$ = "50" THEN
    PersRec(1).MHVALUE = Amt#
    GOTO upd1
   END IF
   IF Cat$ = "60" THEN
    PersRec(1).CVALUE = Amt#
    GOTO upd1
   END IF
   
   PersRec(1).PersVal = Amt#
upd1:

   PersRec(1).PROPPIN = "1"
   PersRec(1).PROPDATE = Date2Num%("11-01-2001")
   PersRec(1).CVALUE = 0
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LATELIST = "N"
   IF MID$(a$, 196, 1) = "Y" THEN
    PersRec(1).DESC1 = LTRIM$(MID$(a$, 16, 22))
    PersRec(1).DESC2 = LTRIM$(MID$(a$, 83, 51))
    PersRec(1).DESC3 = ""
    PersRec(1).Desc4 = LTRIM$(MID$(a$, 79, 4))
    PersRec(1).Desc5 = "Y" + "/" + STRING$(24, " ") + "2001"
     ELSE
    PersRec(1).DESC1 = LTRIM$(MID$(a$, 16, 22))
    PersRec(1).DESC2 = LTRIM$(MID$(a$, 79, 51))
    PersRec(1).DESC3 = ""
    PersRec(1).Desc4 = ""
    PersRec(1).Desc5 = "N" + "/" + STRING$(24, " ") + "2001"
   END IF
   PersRec(1).CustPin = AcctRecord!
   PersRec(1).NextRec = 0
   PersRec(1).LastYrPrinted = 2000
   PersRec(1).Deleted = 0
   PersRec(1).Blank = ""
   PRecord& = LOF(PersTaxFile) / LEN(PersRec(1)) + 1
   PUT PersTaxFile, PRecord&, PersRec(1)
   RETURN



END SUB

SUB OpenTaxCustFile (NumOfTaxRecs, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
  
END SUB

SUB OpenTaxPersFile (NumOfPersRecs, PersTaxFile)
  PersTaxFile = FREEFILE
  OPEN "TAXPERS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PersTaxFile LEN = LEN(PersRec(1))
  NumOfPersRecs = LOF(PersTaxFile) / LEN(PersRec(1))
  
END SUB

