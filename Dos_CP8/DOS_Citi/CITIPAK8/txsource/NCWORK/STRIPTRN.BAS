DEFINT A-Z
DECLARE SUB KillFile (FileName$)
DECLARE FUNCTION Date2Num (TheDate$)

  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TAXCONST.BI'

  KillFile "taxtrans.bak"
  NAME "Taxtrans.dat" AS "taxtrans.bak"

  CLS
  WhatDate = Date2Num("08-21-2007")

  REDIM TaxTran(1) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))

  OPEN "TaxTRANS.bak" FOR RANDOM SHARED AS #1 LEN = TaxTranRecLen
  OPEN "TaxTRANS.dat" FOR RANDOM SHARED AS #2 LEN = TaxTranRecLen


  NumOfRec& = LOF(1) / TaxTranRecLen
  FOR Cnt& = 1 TO NumOfRec&
    LOCATE 5, 1: PRINT "Processing:"; Cnt&;
    GET #1, Cnt&, TaxTran(1)
    IF TaxTran(1).TransDate = WhatDate AND TaxTran(1).TranType = 1 THEN
      Fixed& = Fixed& + 1
    ELSE
      PUT #2, , TaxTran(1)
    END IF
  NEXT

  CLOSE
  PRINT
  PRINT
  PRINT "Removed:"; Fixed&

  CALL KillFile("striptrn.exe")

