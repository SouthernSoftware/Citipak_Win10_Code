DEFINT A-Z
'Real Estate Conversion from Text File Information  INDEPENDENCE VA 2001
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
  DIM SHARED REALRec(1) AS PropertyRecType
  
  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
  ON ERROR RESUME NEXT
  CLS
  DIM a$(41)
  OpenTaxCustFile NumOfTaxRecs, TaxFile

   CLS
   PRINT "Updating County ID's"
   
   OPEN "INREAL04.TXT" FOR INPUT AS #11

1  FOR C = 1 TO 41
    INPUT #11, a$(C)
    IF ERR = 62 THEN
      CLOSE
      LOCATE 7, 1
      PRINT "Processing Complete."
      
      CALL KillFile("ifixpins.exe")
      CALL KillFile("inreal04.txt")
      END
    ELSEIF ERR > 0 THEN
      CLOSE
      PRINT "Unexpected error, Code:"; ERR
      END
    END IF
   NEXT C

   IF NOT DoneHeader THEN
     DoneHeader = True
     GOTO 1
   END IF

   ProcCnt = ProcCnt + 1
   LOCATE 5, 1
   PRINT "Processed:"; ProcCnt

   spin$ = RTRIM$(a$(14))

   LOCATE 5, 1
  
  'Find Exiting Acct
   DidEm = False
   FOR Snt& = 1 TO NumOfTaxRecs
     GET TaxFile, Snt&, TaxCustRec(1)
     CPin& = VAL(a$(1))

     IF TaxCustRec(1).CountyAcct = CPin& THEN
       DidEm = True
       TaxCustRec(1).CountyAcctString = "OLD:" + STR$(TaxCustRec(1).CountyAcct)
       TaxCustRec(1).CountyAcct = VAL(a$(26))
       PUT TaxFile, Snt&, TaxCustRec(1)
       EXIT FOR
     END IF
     
   NEXT Snt&
   IF DidEm = False THEN
     BadCnt = BadCnt + 1
   END IF
GOTO 1


END SUB

SUB OpenTaxCustFile (NumOfTaxRecs, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "RETXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
END SUB

SUB OpenTaxPersFile (NumOfPersRecs, REALTaxFile)
  REALTaxFile = FREEFILE
  OPEN "TAXREAL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #REALTaxFile LEN = LEN(REALRec(1))
  NumOfPersRecs = LOF(REALTaxFile) / LEN(REALRec(1))
  
END SUB

SUB OpenTaxPropFile (NumOfPropRecs, PropTaxFile)
  PropTaxFile = FREEFILE
  OPEN "TAXPROP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PropTaxFile LEN = LEN(PropertyRec(1))
  NumOfPropRecs = LOF(PropTaxFile) / LEN(PropertyRec(1))
END SUB

