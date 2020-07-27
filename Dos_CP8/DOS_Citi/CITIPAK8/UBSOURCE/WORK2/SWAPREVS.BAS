DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
  
  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Swapping Revenues. . ."

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
      SWAP UBCustRec(1).CurrRevAmts(1), UBCustRec(1).CurrRevAmts(2)
      SWAP UBCustRec(1).Serv(1), UBCustRec(1).Serv(2)
    PUT UBFile, Cnt&, UBCustRec(1)
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Done."

