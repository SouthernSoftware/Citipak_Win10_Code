DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)
  
  '$INCLUDE: 'ubtrans.BI'
  CLS
  PRINT "Correcting Balance. . ."

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  UBFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBFile LEN = UBTranRecLen
  LOCATE 5, 1
  GET UBFile, 26678, UBTranRec(1)
  UBTranRec(1).RunBalance = 198.43
  PUT UBFile, 26678, UBTranRec(1)

  CLOSE

LOCATE 9, 1
PRINT "Done"

