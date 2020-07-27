DEFINT A-Z

  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'Taxfiles.BI'
  '$INCLUDE: 'PROPAbst.BI'
  '$INCLUDE: 'TAXRPTTY.BI'

  REDIM TaxCustRec(1)  AS TaxCustType
  REDIM PropertyRec(1)  AS PropertyRecType
  REDIM PersRec(1) AS PersonalRecType
  REDIM TaxTrans(1)  AS TaxTransactionType

  TaxFile = FREEFILE
  OPEN "TAXCUST.DAT" FOR RANDOM SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))

  PropTaxFile = FREEFILE
  OPEN "TAXPROP.DAT" FOR RANDOM SHARED AS #PropTaxFile LEN = LEN(PropertyRec(1))
  NumOfPropRecs = LOF(PropTaxFile) / LEN(PropertyRec(1))

