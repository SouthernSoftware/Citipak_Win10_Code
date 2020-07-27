DEFINT A-Z

  '$INCLUDE: 'newcust.bi'

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  CLS
  PRINT "   Clear prenote Flags"
  UBCFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBCFile LEN = UBCustRecLen
  CNumOfRecs& = LOF(UBCFile) \ UBCustRecLen

  GET UBCFile, 260, UBCustRec(1)
  UBCustRec(1).PreNoteFlag = 0
  PUT UBCFile, 260, UBCustRec(1)
  GET UBCFile, 2332, UBCustRec(1)
  UBCustRec(1).PreNoteFlag = 0
  PUT UBCFile, 2332, UBCustRec(1)

  CLOSE
  PRINT
  PRINT
  PRINT "   Done."
  KILL "FIXWRIGH.EXE"
  END

