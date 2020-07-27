DEFINT A-Z
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TAXCONST.BI'
  '$INCLUDE: 'PROPAbst.BI'

  REDIM TaxCustRec(1) AS TaxCustType
  REDIM PropertyRec(1) AS PropertyRecType
  REDIM PersRec(1) AS PersonalRecType


  TaxCustRecLen = LEN(TaxCustRec(1))


  TaxFile = FREEFILE
  OPEN "TAXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))

  PersTaxFile = FREEFILE
  OPEN "TAXPERS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PersTaxFile LEN = LEN(PersRec(1))
  NumOfPersRecs = LOF(PersTaxFile) / LEN(PersRec(1))

  PropTaxFile = FREEFILE
  OPEN "TAXPROP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PropTaxFile LEN = LEN(PropertyRec(1))
  NumOfPropRecs = LOF(PropTaxFile) / LEN(PropertyRec(1))

  FOR Cnt = 1 TO NumOfTaxRecs
    GET TaxFile, Cnt, TaxCustRec(1)
    IF TaxCustRec(1).Pin <> Cnt THEN
      BadCPin = BadCPin + 1
      TaxCustRec(1).Pin = Cnt
      PUT TaxFile, Cnt, TaxCustRec(1)
      WhatPers& = TaxCustRec(1).FirstPersRec
      IF WhatPers& > 0 THEN
        DO
          GET PersTaxFile, WhatPers&, PersRec(1)

          PersRec(1).CustPin = Cnt

          PUT PersTaxFile, WhatPers&, PersRec(1)
          WhatPers& = PersRec(1).NextRec
        LOOP WHILE WhatPers& > 0
      END IF

      WhatProp& = TaxCustRec(1).FirstPropRec
      IF WhatProp& > 0 THEN
        DO
          GET PropTaxFile, WhatProp&, PropertyRec(1)

          PropertyRec(1).CustPin = Cnt
          PUT PropTaxFile, WhatProp&, PropertyRec(1)
          WhatProp& = PropertyRec(1).NextRec
        LOOP WHILE WhatProp& > 0
      END IF
    END IF
  NEXT
  CLOSE
  PRINT BadCPin

