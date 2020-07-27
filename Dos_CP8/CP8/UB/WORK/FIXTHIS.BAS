DEFINT A-Z

  '$INCLUDE: 'newcust.BI'

  CLS

  PRINT
  PRINT
  PRINT "Correcting Account 192. . ."
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  GET UBFile, 445, UBCustRec(1)
  UBCustRec(1).DelFlag = 0
  FOR Cnt = 399 TO 404
    UBCustRec(1).CustPIN = Cnt
    PUT UBFile, Cnt, UBCustRec(1)
  NEXT
  FOR Cnt = 448 TO 456
    UBCustRec(1).CustPIN = Cnt
    PUT UBFile, Cnt, UBCustRec(1)
  NEXT

  'UBCustRec(1).CustPIN

  'PUT UBFile, 192, UBCustRec(1)
  CLOSE UBFile
  PRINT
  PRINT
  PRINT
  PRINT "Done."

  CALL KILLFile("FIX192.EXE")

