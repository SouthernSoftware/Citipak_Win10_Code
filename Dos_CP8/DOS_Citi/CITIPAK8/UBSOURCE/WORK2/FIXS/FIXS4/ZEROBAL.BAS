DEFINT A-Z
DECLARE FUNCTION uRound# (N#)
DECLARE FUNCTION FileSize& (FileName$)
  
  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Processing Customers Balances. . ."
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
    IF uRound#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance) = 0 THEN
      FOR RCnt = 1 TO 15
        UBCustRec(1).CurrRevAmts(RCnt) = 0
      NEXT
      Fixed = Fixed + 1
      PUT UBFile, Cnt&, UBCustRec(1)
    END IF
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Fixed:"; Fixed

