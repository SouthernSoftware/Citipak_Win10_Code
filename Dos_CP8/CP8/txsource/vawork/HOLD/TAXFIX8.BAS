  DECLARE FUNCTION Num2Date$ (DateNumber%)
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'Taxfiles.BI'
  '$INCLUDE: 'PROPAbst.BI'
  '$INCLUDE: 'TAXRPTTY.BI'
  
  CLS
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED PersRec(1) AS PersonalRecType
  DIM SHARED TaxTrans(1) AS TaxTransactionType

  TaxCustRecLen = LEN(TaxCustRec(1))
  TaxTranRecLen = LEN(TaxTrans(1))

  CLS

  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
  FOR Cnt& = 1 TO LOF(TaxTran) / TaxTranRecLen
   GET TaxTran, Cnt&, TaxTrans(1)

  
  IF TaxTrans(1).TranType = 1 AND TaxTrans(1).TaxYear = 1999 AND TaxTrans(1).BillType = "P" THEN
    TaxTrans(1).Revenue.Principle1 = 0
    TaxTrans(1).Revenue.Principle2 = 0
    TaxTrans(1).Revenue.Principle3 = 0
    TaxTrans(1).Revenue.Principle4 = 0
    TaxTrans(1).Revenue.Principle5 = 0
    TaxTrans(1).Revenue.Principle1Pd = 0
    TaxTrans(1).Revenue.Principle2Pd = 0
    TaxTrans(1).Revenue.Principle3Pd = 0
    TaxTrans(1).Revenue.Principle4Pd = 0
    TaxTrans(1).Revenue.Principle5Pd = 0
    TaxTrans(1).Amount = 0
    TaxTrans(1).Description = "Voided"
    TaxTrans(1).TaxYear = 0
    TaxTrans(1).Posted2GL = "Y"
    PUT TaxTran, Cnt&, TaxTrans(1)
    END IF
    'Found the Bill
   NEXT Cnt&
   CLOSE

END

