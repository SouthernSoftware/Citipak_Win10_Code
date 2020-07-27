DEFINT A-Z
DECLARE FUNCTION uRound# (N#)
DECLARE FUNCTION FileSize& (FileName$)
  
  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Processing. . ."
  REDIM UBCustRec(2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
    LOCATE , 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    UBCustRec(1).DelFlag = -1
    PUT UBFile, 2917, UBCustRec(1)
    PUT UBFile, 2918, UBCustRec(1)
    PUT UBFile, 2919, UBCustRec(1)
    PUT UBFile, 2920, UBCustRec(1)
    PUT UBFile, 2921, UBCustRec(1)
    PUT UBFile, 2922, UBCustRec(1)
    PUT UBFile, 2923, UBCustRec(1)

  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Processing complete:"; Fixed

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

