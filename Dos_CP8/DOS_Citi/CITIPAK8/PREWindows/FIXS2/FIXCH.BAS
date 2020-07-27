DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Processing Customers. . ."
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
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
    IF QPTrim$(UBCustRec(1).Serv(1).RateCode) = "CH" THEN
      UBCustRec(1).Serv(1).RateCode = "CHW"
      UBCustRec(1).Serv(2).RateCode = "CHS"
      PUT UBFile, Cnt&, UBCustRec(1)
    END IF
  NEXT
  IF UBCustRec(1).Status = "I" THEN
    FOR MTCnt = 1 TO 7
      IF UBCustRec(1).LocMeters(MTCnt).CurRead > 0 THEN
        UBCustRec(1).LocMeters(MTCnt).PrevRead = UBCustRec(1).LocMeters(MTCnt).CurRead
      END IF
    NEXT
    PUT UBFile, Cnt&, UBCustRec(1)
  END IF



  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Done."

