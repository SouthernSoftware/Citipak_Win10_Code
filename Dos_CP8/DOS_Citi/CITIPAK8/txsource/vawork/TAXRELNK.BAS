DEFINT A-Z
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
  DIM SHARED PROPERTYRec(1) AS PropertyRecType
  DIM SHARED REALRec(1) AS PropertyRecType
  
  DIM SHARED TransRec(1) AS TaxTransactionType
  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
 OpenTaxCustFile NumOfTaxRecs, TaxFile
 FOR s& = 1 TO NumOfTaxRecs
 GET TaxFile, s&, TaxCustRec(1)
 TaxCustRec(1).LastTrans = 0
 PUT TaxFile, s&, TaxCustRec(1)
 NEXT s&


  TxFile = FREEFILE
   OPEN "TAXTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TxFile LEN = LEN(TransRec(1))
   FOR Cnt& = 1 TO LOF(TxFile) / LEN(TransRec(1))
   GET TxFile, Cnt&, TransRec(1)
    IF TransRec(1).BillType = "R" THEN
     GET TaxFile, TransRec(1).CustomerRec, TaxCustRec(1)
      TaxCustRec(1).LastTrans = Cnt&
     PUT TaxFile, TransRec(1).CustomerRec, TaxCustRec(1)
     GOTO nex
    END IF
    IF NOT (TransRec(1).BillType = "P" OR TransRec(1).BillType = "R") THEN
     CustomerRecord& = TransRec(1).CustomerRec
     GET TxFile, TransRec(1).BelongTo, TransRec(1)
     IF TransRec(1).BillType = "R" THEN
       GET TaxFile, CustomerRecord&, TaxCustRec(1)
        TaxCustRec(1).LastTrans = Cnt&
       PUT TaxFile, CustomerRecord&, TaxCustRec(1)
       GOTO nex
     END IF
   END IF
nex:
   NEXT Cnt&
   CLOSE

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
  OPEN "TAXPROP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PropTaxFile LEN = LEN(PROPERTYRec(1))
  NumOfPropRecs = LOF(PropTaxFile) / LEN(PROPERTYRec(1))
END SUB

