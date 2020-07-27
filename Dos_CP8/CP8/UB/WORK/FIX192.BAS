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
  GET UBFile, 192, UBCustRec(1)
  UBCustRec(1).DelFlag = 0
  PUT UBFile, 192, UBCustRec(1)
  CLOSE UBFile
  PRINT
  PRINT
  PRINT
  PRINT "Done."

  CALL KILLFile("FIX192.EXE")

