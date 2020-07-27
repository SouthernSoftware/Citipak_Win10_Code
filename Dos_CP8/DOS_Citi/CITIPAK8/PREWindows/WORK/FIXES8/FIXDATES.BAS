DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)
  
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.BI'

  CONST False = 0, True = NOT False
  
  REDIM UBTransRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTransRec(1))


  CLS
  LOCATE 3, 1, 0
  PRINT "Correcting Customer Information. . ."

  PRINT
  PRINT

  UBCFile = FREEFILE
  OPEN "UBtrans.dat" FOR RANDOM SHARED AS UBCFile LEN = UBTranRecLen
  TNumOfRecs& = LOF(UBCFile) \ UBTranRecLen
  
  FOR CCnt& = 1 TO TNumOfRecs&
    GET UBCFile, CCnt&, UBTransRec(1)
    BadCnt = 0
    LOCATE 5, 1
    PRINT "Processing:"; CCnt&; " of"; TNumOfRecs&;
    UBTransRec(1).TransDate = UBTransRec(1).TransDate + 365
    PUT UBCFile, CCnt&, UBTransRec(1)
  NEXT

SkiptoNext:
  CLOSE

LOCATE 9, 1
PRINT "Processing Complete."

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

