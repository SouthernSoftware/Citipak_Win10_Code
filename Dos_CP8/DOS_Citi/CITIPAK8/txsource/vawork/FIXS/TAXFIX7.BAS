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
'Changes Date on Interest Charge From 10/01/99 to 07/01/99

CLS
   TxFile = FREEFILE
   OPEN "TAXTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TxFile LEN = LEN(TransRec(1))
   LASTREC = LOF(TxFile) / LEN(TransRec(1))
   GET TxFile, 4020, TransRec(1)
    TransRec(1).Amount = 134.21
    TransRec(1).Revenue.Principle1Pd = 101
    TransRec(1).Revenue.InterestPd = 28.16
    TransRec(1).Revenue.PenaltyPd = 5.05
    MBill = TransRec(1).BelongTo
   PUT TxFile, 4020, TransRec(1)

   GET TxFile, MBill, TransRec(1)
    TransRec(1).Revenue.Principle1Pd = 101
    TransRec(1).Revenue.InterestPd = 29.04
    TransRec(1).Revenue.PenaltyPd = 5.05
   PUT TxFile, MBill, TransRec(1)

   GET TxFile, 4019, TransRec(1)
    TransRec(1).Amount = 123.65
    TransRec(1).Revenue.Principle1Pd = 101
    TransRec(1).Revenue.InterestPd = 17.6
    TransRec(1).Revenue.PenaltyPd = 5.05
    MBill = TransRec(1).BelongTo
   PUT TxFile, 4019, TransRec(1)
   GET TxFile, MBill, TransRec(1)
    TransRec(1).Revenue.Principle1Pd = 101
    TransRec(1).Revenue.InterestPd = 18.48
    TransRec(1).Revenue.PenaltyPd = 5.05
   PUT TxFile, MBill, TransRec(1)

   GET TxFile, 4018, TransRec(1)
    TransRec(1).Amount = 127.22
    TransRec(1).Revenue.Principle1Pd = 113.62
    TransRec(1).Revenue.InterestPd = 7.92
    TransRec(1).Revenue.PenaltyPd = 5.68
    MBill = TransRec(1).BelongTo
   PUT TxFile, 4018, TransRec(1)
   GET TxFile, MBill, TransRec(1)
    TransRec(1).Revenue.Principle1Pd = 113.62
    TransRec(1).Revenue.InterestPd = 8.91
    TransRec(1).Revenue.PenaltyPd = 5.68
   PUT TxFile, MBill, TransRec(1)



   GET TxFile, 4023, TransRec(1)
    TransRec(1).Amount = 4.9
    TransRec(1).Revenue.Principle1Pd = 3.75
    TransRec(1).Revenue.InterestPd = .96
    TransRec(1).Revenue.PenaltyPd = .19
    MBill = TransRec(1).BelongTo
   PUT TxFile, 4023, TransRec(1)

   GET TxFile, MBill, TransRec(1)
    TransRec(1).Revenue.Principle1Pd = 3.75
    TransRec(1).Revenue.InterestPd = .99
    TransRec(1).Revenue.PenaltyPd = .19
   PUT TxFile, MBill, TransRec(1)

   GET TxFile, 4022, TransRec(1)
    TransRec(1).Amount = 4.54
    TransRec(1).Revenue.Principle1Pd = 3.75
    TransRec(1).Revenue.InterestPd = .6
    TransRec(1).Revenue.PenaltyPd = .19
    MBill = TransRec(1).BelongTo
   PUT TxFile, 4022, TransRec(1)
   GET TxFile, MBill, TransRec(1)
    TransRec(1).Revenue.Principle1Pd = 3.75
    TransRec(1).Revenue.InterestPd = .63
    TransRec(1).Revenue.PenaltyPd = .19
   PUT TxFile, MBill, TransRec(1)

   GET TxFile, 4021, TransRec(1)
    TransRec(1).Amount = 3.86
    TransRec(1).Revenue.Principle1Pd = 3.45
    TransRec(1).Revenue.InterestPd = .24
    TransRec(1).Revenue.PenaltyPd = .17
    MBill = TransRec(1).BelongTo
   PUT TxFile, 4021, TransRec(1)
   GET TxFile, MBill, TransRec(1)
    TransRec(1).Revenue.Principle1Pd = 3.45
    TransRec(1).Revenue.InterestPd = .27
    TransRec(1).Revenue.PenaltyPd = .17
   PUT TxFile, MBill, TransRec(1)




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

