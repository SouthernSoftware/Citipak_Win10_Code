DEFINT A-Z
DECLARE SUB KillFile (FileName$)
DECLARE FUNCTION Date2Num (TheDate$)

  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TAXCONST.BI'

  CLS
  WhatDate = Date2Num("10-01-2005")

  REDIM TaxTran(1 TO 2) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))

  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS #1 LEN = TaxTranRecLen
  NumOfRec& = LOF(1) / TaxTranRecLen
  FOR Cnt& = 1 TO NumOfRec&
    LOCATE 5, 1: PRINT "Processing:"; Cnt&;
    GET #1, Cnt&, TaxTran(1)
    IF (TaxTran(1).TransDate >= WhatDate) AND (TaxTran(1).TranType = 1) THEN
      Fixed = Fixed + 1
      IF (TaxTran(1).Revenue.Principle1 <> TaxTran(1).Amount) THEN
        TaxTran(1).Amount = TaxTran(1).Revenue.Principle1
      END IF
      PUT #1, Cnt&, TaxTran(1)
    END IF
  NEXT
  CLOSE
CLS
PRINT "Fixed:"; Fixed
CALL KillFile("fixshore.exe")

