DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)
  
  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Changing Customers. . ."
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfCust& = LOF(UBFile) / UBCustRecLen
  FOR zz = 1 TO NumOfCust&
    LOCATE , 1
    PRINT "Processing:"; zz; " of"; NumOfCust&;
    GET #UBFile, zz, UBCustRec(1)
    UBCustRec(1).LATEFEE = "Y"
    PUT #UBFile, zz, UBCustRec(1)
  NEXT

  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Done."

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

