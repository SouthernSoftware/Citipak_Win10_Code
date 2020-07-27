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
  DIM SHARED RAMAST(1) AS RAMAST
  DIM SHARED TransRec(1) AS TaxTransactionType
  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
   TXFILE = FREEFILE
   OPEN "TAXTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TXFILE LEN = LEN(TransRec(1))
   LASTREC = LOF(TXFILE) / LEN(TransRec(1))
   FOR Cnt! = 1 TO LASTREC
   GET TXFILE, Cnt!, TransRec(1)
   IF TransRec(1).CustomerRec = 277 THEN
   IF TransRec(1).TranTYpe = 2 OR TransRec(1).TranTYpe = 3 THEN
    LPRINT TransRec(1).Amount, TransRec(1).Revenue.Principle1Pd, Num2Date$(TransRec(1).TransDate)
    Paid# = TransRec(1).Amount
    BelongTo = TransRec(1).BelongTo
    GET TXFILE, TransRec(1).BelongTo, TransRec(1)

   IF TransRec(1).TaxYear = 1999 AND TransRec(1).BillType = "R" THEN
     totp# = totp# + TransRec(1).Amount
     
     TransRec(1).Revenue.Principle4Pd = TransRec(1).Revenue.Principle4Pd + Paid#
     PUT TXFILE, BelongTo, TransRec(1)

    GET TXFILE, Cnt!, TransRec(1)
     tot# = tot# + TransRec(1).Amount
     tot1# = tot1# + TransRec(1).Revenue.Principle1Pd
   END IF

   END IF
   
                                      END IF
   NEXT Cnt!
   
    












         STOP
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

