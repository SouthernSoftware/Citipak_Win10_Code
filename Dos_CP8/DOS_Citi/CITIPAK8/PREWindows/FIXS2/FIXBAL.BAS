DECLARE FUNCTION uRound# (N#)
DEFINT A-Z
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
    BalTot# = uRound#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
    IF BalTot# = 0 THEN
      FOR rCnt = 1 TO 15
        UBCustRec(1).CurrRevAmts(rCnt) = 0
      NEXT
      PUT UBFile, Cnt&, UBCustRec(1)
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

