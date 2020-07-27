DEFINT A-Z
DECLARE FUNCTION uRound# (N#)
DECLARE FUNCTION FileSize& (FileName$)
  
  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Clearing Customers Balances. . ."
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
    UBCustRec(1).DepositAmt = 0
    UBCustRec(1).CurrBalance = 0
    UBCustRec(1).PrevBalance = 0
    FOR RCnt = 1 TO 15
      UBCustRec(1).CurrRevAmts(RCnt) = 0
    NEXT
    PUT UBFile, Cnt&, UBCustRec(1)
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Finished."

