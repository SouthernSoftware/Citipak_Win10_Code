DEFINT A-Z
DECLARE FUNCTION uRound# (N#)
DECLARE FUNCTION FileSize& (FileName$)
  
  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Processing Customers. . ."
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  FOR Cnt& = 1 TO NumOfRecs&
    LOCATE , 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    GET UBFile, Cnt&, UBCustRec(1)
    IF UBCustRec(1).CurrRevAmts(5) > 0 THEN
      UBCustRec(1).CurrRevAmts(1) = uRound#(UBCustRec(1).CurrRevAmts(1) + UBCustRec(1).CurrRevAmts(5))
      UBCustRec(1).CurrRevAmts(5) = 0
      PUT UBFile, Cnt&, UBCustRec(1)
      Fixed = Fixed + 1
    END IF
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Corrected:"; Fixed

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

