DEFINT A-Z

  '$INCLUDE: 'newcust.BI'

  CLS

  PRINT
  PRINT
  PRINT "Correcting Account 1505. . ."
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  GET UBFile, 1505, UBCustRec(1)
  UBCustRec(1).Status = "I"
  PUT UBFile, 1505, UBCustRec(1)
  CLOSE UBFile
  PRINT
  PRINT
  PRINT
  PRINT "Done."

  CALL KILLFile("FIX1505.EXE")

