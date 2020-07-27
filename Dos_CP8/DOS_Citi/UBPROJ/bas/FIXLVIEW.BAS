DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)
  
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.BI'

  CONST False = 0, True = NOT False
  
  REDIM UBCustREc(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustREc(1))


  CLS
  LOCATE 3, 1, 0
  PRINT "Correcting Customer Information. . ."

  PRINT
  PRINT

  UBCFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBCFile LEN = UBCustRecLen
  CNumOfRecs& = LOF(UBCFile) \ UBCustRecLen
  
  GET UBCFile, 2296, UBCustREc(1)
  FOR Cnt = 4 TO 15
    UBCustREc(1).Serv(Cnt).RateCode = ""
    UBCustREc(1).Serv(Cnt).RMtrType = ""
  NEXT
  PUT UBCFile, 2296, UBCustREc(1)

SkiptoNext:
  CLOSE

LOCATE 9, 1
PRINT "Processing Complete."

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

