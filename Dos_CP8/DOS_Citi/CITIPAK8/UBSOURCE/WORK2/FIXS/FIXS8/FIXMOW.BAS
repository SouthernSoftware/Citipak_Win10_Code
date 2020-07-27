DEFINT A-Z
DECLARE FUNCTION uRound# (N#)
DECLARE FUNCTION FileSize& (FileName$)
  
  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Processing. . ."
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
    GET UBFile, 12982, UBCustRec(1)
    UBCustRec(1).CurrBalance = 0
    UBCustRec(1).PrevBalance = 0
    FOR rcnt = 1 TO 15
      UBCustRec(1).CurrRevAmts(rcnt) = 0
    NEXT
    PUT UBFile, 12982, UBCustRec(1)
    GET UBFile, 13087, UBCustRec(1)
    UBCustRec(1).CurrBalance = 0
    UBCustRec(1).PrevBalance = 0
    FOR rcnt = 1 TO 15
      UBCustRec(1).CurrRevAmts(rcnt) = 0
    NEXT
    PUT UBFile, 13087, UBCustRec(1)
    GET UBFile, 12865, UBCustRec(1)
    UBCustRec(1).CurrBalance = 0
    UBCustRec(1).PrevBalance = 0
    FOR rcnt = 1 TO 15
      UBCustRec(1).CurrRevAmts(rcnt) = 0
    NEXT
    PUT UBFile, 12865, UBCustRec(1)


  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Processing complete."

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

