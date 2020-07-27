DECLARE FUNCTION uRound# (N#)
DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
  
  CONST False = 0, True = NOT False

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
    IF UBCustRec(1).Book = "08" THEN
      IF DoOne THEN
        DoOne = False
        UBCustRec(1).PumpCode = "W"
        PUT UBFile, Cnt&, UBCustRec(1)
      ELSE
        DoOne = True
      END IF
    END IF
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Done."

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

