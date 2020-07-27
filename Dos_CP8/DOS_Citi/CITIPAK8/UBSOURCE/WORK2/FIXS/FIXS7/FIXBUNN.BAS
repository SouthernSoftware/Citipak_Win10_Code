DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)
  
  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Purging Damaged Customers. . ."
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  UBCustRec(1).DelFlag = -1

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs = LOF(UBFile) \ UBCustRecLen
  FOR Cnt = 400 TO 403
    PUT #UBFile, Cnt, UBCustRec(1)
  NEXT

  FOR Cnt = 421 TO 432
    PUT #UBFile, Cnt, UBCustRec(1)
  NEXT

  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Processing Complete"

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

