DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)
  
  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Correcting Zipcodes. . ."
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  UBCustRec(1).DelFlag = -1

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRec = LOF(UBFile) / UBCustRecLen

  FOR zz = 1 TO NumOfRec
    GET #UBFile, zz, UBCustRec(1)
    Zip$ = UBCustRec(1).ZipCode
    Zip$ = LEFT$(Zip$, 5)
    LSET UBCustRec(1).ZipCode = Zip$
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

