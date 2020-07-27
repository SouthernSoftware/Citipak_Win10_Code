DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
  
  '$INCLUDE: 'newcust.BI'
  '$INCLUDE: 'ubtrans.BI'

  TDate = Date2Num%("05-31-1997")

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  CLS
  PRINT
  PRINT
  PRINT "Processing. . ."
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen

  UBTran = FREEFILE
  OPEN "UBTrans.dat" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen

  FOR Cnt& = 1 TO NumOfRecs&
    LOCATE , 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    GET UBFile, Cnt&, UBCustRec(1)
    WhatTrans& = UBCustRec(1).LastTrans
    DO WHILE WhatTrans& > 0
      GET UBTran, WhatTrans&, UBTranRec(1)
      IF (UBTranRec(1).TransType = TranUtilityBill) AND (UBTranRec(1).TransDate = TDate) THEN
        FOR MCnt = 1 TO 7
          UBCustRec(1).LocMeters(MCnt).PrevRead = UBTranRec(1).PrevRead(MCnt)
        NEXT
        PUT UBFile, Cnt&, UBCustRec(1)
        EXIT DO
      END IF
      WhatTrans& = UBTranRec(1).PrevTrans
    LOOP
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Done."

