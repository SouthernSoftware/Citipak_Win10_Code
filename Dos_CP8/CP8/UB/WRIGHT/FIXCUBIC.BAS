DEFINT A-Z

  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Converting Cubic Meters. . ."
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
    FOR MCnt = 1 TO 7
      IF UBCustRec(1).LocMeters(MCnt).CurRead > 0 THEN
        UBCustRec(1).LocMeters(MCnt).CurRead = UBCustRec(1).LocMeters(MCnt).PrevRead + 10
      END IF
    NEXT
    PUT UBFile, Cnt&, UBCustRec(1)
  NEXT
  CLOSE UBFile
PRINT
PRINT
PRINT
PRINT "Processing Complete."

