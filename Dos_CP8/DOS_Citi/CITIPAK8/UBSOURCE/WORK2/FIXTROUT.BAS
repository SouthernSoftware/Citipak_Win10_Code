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
  FOR Cnt& = 1 TO NumOfRecs&
    LOCATE , 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    GET UBFile, Cnt&, UBCustRec(1)

    FOR zz = 1 TO 2
      IF UBCustRec(1).Monthly(zz).AMTOWED < 0 THEN
        UBCustRec(1).Monthly(zz).AMTOWED = 0
      END IF
      IF UBCustRec(1).Monthly(zz).TotAmtPD THEN
        UBCustRec(1).Monthly(zz).TotAmtPD = 0
      END IF
      IF UBCustRec(1).Monthly(zz).PayAmt THEN
        UBCustRec(1).Monthly(zz).PayAmt = 0
      END IF
      IF UBCustRec(1).Monthly(zz).RevSource THEN
        UBCustRec(1).Monthly(zz).RevSource = 0
      END IF
    NEXT
    PUT UBFile, Cnt&, UBCustRec(1)
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Processing complete"

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

