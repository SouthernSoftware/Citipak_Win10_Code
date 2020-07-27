DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Date2Num (TheDate$)
DECLARE FUNCTION uRound# (DblNum#)

  '$INCLUDE: 'ubtrans.BI'

  CLS
  PRINT
  PRINT "Correcting Transaction Balances. . ."
  REDIM UBTrans(1) AS UBTransRecType
  UBTransRecLen = LEN(UBTrans(1))
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBTRANS.dat" FOR RANDOM SHARED AS UBFile LEN = UBTransRecLen


    GET UBFile, 9848, UBTrans(1)
    UBTrans(1).RunBalance = 0
    PUT UBFile, 9848, UBTrans(1)



  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Corrected:"; 1

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

