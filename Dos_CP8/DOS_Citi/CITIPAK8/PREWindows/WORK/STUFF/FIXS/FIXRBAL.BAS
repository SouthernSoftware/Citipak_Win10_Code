DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)

  '$INCLUDE: 'ubtrans.BI'

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  CLS
  LOCATE 3, 1, 0
  PRINT "Correcting Runbalance. . ."

  PRINT
  PRINT

  
  UBTFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTFile LEN = UBTranRecLen
  GET UBTFile, 13129, UBTranRec(1)
  UBTranRec(1).RunBalance = -5
  PUT UBTFile, 13129, UBTranRec(1)

SkiptoNext:
  CLOSE

LOCATE 9, 1
PRINT "Done"

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

