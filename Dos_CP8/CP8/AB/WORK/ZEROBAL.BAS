DEFINT A-Z
DECLARE FUNCTION Round# (N#)
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
    'IF Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance) = 0 THEN
      FOR RCnt = 1 TO 15
        UBCustRec(1).CurrRevAmts(RCnt) = 0
        UBCustRec(1).PrevRevAmts(RCnt) = 0
      NEXT
      UBCustRec(1).CurrBalance = 0
      UBCustRec(1).PrevBalance = 0
      UBCustRec(1).DepositAmt = 0
      PUT UBFile, Cnt&, UBCustRec(1)
    'END IF
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Finished."

CALL killfile("zerobal.exe")

