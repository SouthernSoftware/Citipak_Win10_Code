DEFINT A-Z

  '$INCLUDE: 'newcust.BI'

  CLS

  PRINT
  PRINT
  PRINT "Correcting Account 1428. . ."
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  GET UBFile, 1428, UBCustRec(1)
  UBCustRec(1).Status = "F"
  PUT UBFile, 1428, UBCustRec(1)
  CLOSE UBFile
  PRINT
  PRINT
  PRINT
  PRINT "Done."

  CALL KILLFile("FIX1428.EXE")

