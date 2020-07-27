DEFINT A-Z

  '$INCLUDE: 'newcust.BI'


  CLS
  PRINT
  PRINT
  PRINT "Processing Accounts. . ."
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
    LSET UBCustRec(1).HHMSG1 = ""
    LSET UBCustRec(1).HHMSG2 = ""
    LSET UBCustRec(1).HHMSG3 = ""
    PUT UBFile, Cnt&, UBCustRec(1)
  NEXT
  CLOSE UBFile
  PRINT
  PRINT
  PRINT
  PRINT "Done."
  CALL KillFile("Stripmsg.exe")

