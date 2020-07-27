DEFINT A-Z
DECLARE SUB KillFile (FileName$)
DECLARE FUNCTION Date2Num (TheDate$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TAXCONST.BI'

 ' KillFile "taxtrans.bak"
 ' NAME "Taxtrans.dat" AS "taxtrans.bak"

  CLS
  'WhatDate = Date2Num("12-31-1979")

  REDIM TaxTran(1) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))

'  OPEN "TaxTRANS.bak" FOR RANDOM SHARED AS #1 LEN = TaxTranRecLen
  OPEN "TaxTRANS.dat" FOR RANDOM SHARED AS #1 LEN = TaxTranRecLen

  NumOfRec& = LOF(1) / TaxTranRecLen
  FOR Cnt& = 1 TO NumOfRec&
    LOCATE 5, 1: PRINT "Processing:"; Cnt&;
    GET #1, Cnt&, TaxTran(1)
    IF TaxTran(1).TranType = 1 THEN
      TranYear = VAL(RIGHT$(Num2Date(TaxTran(1).TransDate), 4))
      BillYear = TaxTran(1).TaxYear
      IF TranYear = 0 THEN
        IF TranYear <> BillYear THEN
          TRYear = Date2Num("09-01-" + QPTrim$(STR$(BillYear)))
          TaxTran(1).TransDate = TRYear
          PUT #1, Cnt&, TaxTran(1)
          Fixed& = Fixed& + 1
        END IF
      END IF
    END IF
  NEXT

  CLOSE
  PRINT
  PRINT
  PRINT "Removed:"; Fixed&

  CALL KillFile("fixlland.exe")

