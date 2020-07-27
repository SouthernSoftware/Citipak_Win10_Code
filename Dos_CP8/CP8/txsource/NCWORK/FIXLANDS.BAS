DEFINT A-Z
DECLARE SUB KillFile (FileName$)
DECLARE FUNCTION Date2Num (TheDate$)

  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TAXCONST.BI'

  CLS
  WhatDate = Date2Num("02-28-2000")

  REDIM TaxTran(1 TO 2) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))

  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS #1 LEN = TaxTranRecLen
  NumOfRec& = LOF(1) / TaxTranRecLen
  FOR Cnt& = 1 TO NumOfRec&
    LOCATE 5, 1: PRINT "Processing:"; Cnt&;
    GET #1, Cnt&, TaxTran(1)
    IF TaxTran(1).TransDate = WhatDate THEN
      Fixed = Fixed + 1
      TaxTran(1).TransDate = -32767
      TaxTran(1).Amount = 0
      TaxTran(1).BelongTo = 0
      TaxTran(1).CustomerRec = 0
      TaxTran(1).LastTrans = 0
      PUT #1, Cnt&, TaxTran(1)
    END IF
  NEXT
  CLOSE
CLS
PRINT "Fixed:"; Fixed
CALL KillFile("fixlands.exe")

