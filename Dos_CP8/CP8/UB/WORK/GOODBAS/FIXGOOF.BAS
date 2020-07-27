DEFINT A-Z

  '$INCLUDE: 'newcust.BI'

  CLS

  PRINT
  PRINT
  PRINT "UNdeleting Account 12307. . ."
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  GET UBFile, 12307, UBCustRec(1)
  UBCustRec(1).DelFlag = 0
  PUT UBFile, 12307, UBCustRec(1)
  CLOSE
  PRINT
  PRINT
  PRINT
  PRINT "Done"
  CALL KILLFile("FIXGOOF.EXE")

