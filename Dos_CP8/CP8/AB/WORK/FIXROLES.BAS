DEFINT A-Z

  '$INCLUDE: 'newcust.BI'

  CLS

  PRINT
  PRINT
  PRINT "Processing Customers. . ."
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
    IF UBCustRec(1).DelFlag <> 0 THEN
      GOTO SkipThisOne
    END IF
    FOR MTCnt = 1 TO 2
      UBCustRec(1).Serv(MTCnt).RateCode = ""
      UBCustRec(1).Serv(MTCnt).RMtrType = ""
    NEXT
    PUT UBFile, Cnt&, UBCustRec(1)
SkipThisOne:
  NEXT
  CLOSE UBFile
PRINT
PRINT
PRINT
PRINT "Done."
CALL KILLFile("FIXROLES.EXE")

