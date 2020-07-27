DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'newcust.bi'

  REDIM UBCustRec(1)  AS NewUBCustRecType
  UBCustLen = LEN(UBCustRec(1))
  CLS
  LOCATE 2, 1, 0

  PRINT "Processing. . ."

  
  UBFile = FREEFILE
  OPEN "ubcust.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustLen
  NumOfRec& = LOF(UBFile) \ UBCustLen
  FOR Cnt& = 1 TO NumOfRec&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRec&;
    GET UBFile, Cnt&, UBCustRec(1)
    SWAP UBCustRec(1).USEDRAFT, UBCustRec(1).AcctType
    PUT UBFile, Cnt&, UBCustRec(1)
  NEXT
  CLOSE
  LOCATE 8, 1
  PRINT "Processing Complete."
  PRINT
  CALL KillFile("swapdrft.exe")

