DEFINT A-Z


  '$INCLUDE: 'newcust.bi'

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfCust& = LOF(UBFile) / UBCustRecLen
  FOR Cnt = 1 TO 10000
    LOCATE 5, 1
    PRINT "Adding:"; NumOfCust& + Cnt;
    GET UBFile, Cnt, UBCustRec(1)
    PUT UBFile, NumOfCust& + Cnt, UBCustRec(1)
  NEXT
  CLOSE

  'PRINT NumOfCust&

