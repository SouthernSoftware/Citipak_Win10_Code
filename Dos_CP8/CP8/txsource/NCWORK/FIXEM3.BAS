DEFINT A-Z
DECLARE SUB KILLFile (FileName$)
DECLARE FUNCTION Date2Num (TheDate$)

  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TAXCONST.BI'

  CLS
  WhatDate = Date2Num("06-20-2000")

  REDIM TaxTran(1 TO 2) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))

  KILLFile "TaxTRANS.BAK"

  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS #1 LEN = TaxTranRecLen
  NumOfRec& = LOF(1) / TaxTranRecLen
  FOR Cnt& = 1 TO NumOfRec&
    LOCATE 5, 1: PRINT "Processing:"; Cnt&;
    GET #1, Cnt&, TaxTran(1)
    IF TaxTran(1).TransDate = WhatDate AND TaxTran(1).TranType = 4 THEN
      GET #1, TaxTran(1).BelongTo, TaxTran(2)
      TaxTran(2).Revenue.Interest = TaxTran(2).Revenue.Interest - TaxTran(1).Revenue.Interest
      PUT #1, TaxTran(1).BelongTo, TaxTran(2)
    END IF
  NEXT
  CLOSE

  NAME "TaxTRANS.DAT" AS "TaxTRANS.BAK"
  OPEN "TaxTRANS.BAK" FOR RANDOM SHARED AS #1 LEN = TaxTranRecLen
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS #2 LEN = TaxTranRecLen

  NumOfRec& = LOF(1) / TaxTranRecLen
  FOR Cnt& = 1 TO NumOfRec&
    LOCATE 5, 1: PRINT "Processing:"; Cnt&;
    GET #1, Cnt&, TaxTran(1)
    IF Cnt& = 46 THEN
      TaxTran(1).Revenue.Principle1PD = 0
      PUT #1, 46, TaxTran(1)
    END IF

    IF TaxTran(1).TransDate = WhatDate AND TaxTran(1).TranType = 4 THEN
      Found = Found + 1
    ELSE
      PUT #2, , TaxTran(1)
    END IF
  NEXT
  PRINT
  PRINT
  PRINT "REMOVED:"; Found



  CALL KILLFile("fixem.exe")

