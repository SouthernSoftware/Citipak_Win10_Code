DEFINT A-Z
DECLARE FUNCTION Date2Num (TheDate$)

'$INCLUDE: 'TaxCust.BI'                   'Tax Trans Layout

  WhatDate = Date2Num("03-13-2000")
  REDIM TaxTrans(1) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTrans(1))
  TranFile = FREEFILE
  OPEN "TAXTRANS.DAT" FOR RANDOM SHARED AS TranFile LEN = TaxTranRecLen
  NumOfRec& = LOF(TranFile) \ TaxTranRecLen
  FOR Cnt& = 1 TO NumOfRec&
    GET #TranFile, Cnt&, TaxTrans(1)
    IF TaxTrans(1).TransDate <= WhatDate THEN
      TaxTrans(1).Posted2GL = "Y"
    ELSE
      TaxTrans(1).Posted2GL = ""
    END IF
    PUT #TranFile, Cnt&, TaxTrans(1)

  NEXT Cnt&
  CLOSE TranFile

