DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)
  
  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Changing Customers Status. . ."
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

    GET #UBFile, 383, UBCustRec(1)
    UBCustRec(1).Status = "B"
    PUT #UBFile, 383, UBCustRec(1)

    GET #UBFile, 134, UBCustRec(1)
    UBCustRec(1).Status = "B"
    PUT #UBFile, 134, UBCustRec(1)

    GET #UBFile, 589, UBCustRec(1)
    UBCustRec(1).Status = "B"
    PUT #UBFile, 589, UBCustRec(1)



  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Fixed: 3"

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

