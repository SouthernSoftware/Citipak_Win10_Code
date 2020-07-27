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
    IF UBCustRec(1).Status = "B" THEN
      LSET UBCustRec(1) = UBCustRec(2)
      UBCustRec(2).DelFlag = -1
      GOTO putem
    END IF
    IF LEN(QPTrim$(UBCustRec(1).Serv(1).RateCode)) = 0 THEN
      FOR zz = 1 TO 7
        UBCustRec(1).LocMeters(zz).CurRead = -2147483647
        UBCustRec(1).LocMeters(zz).PrevRead = -2147483647
        UBCustRec(1).LocMeters(zz).ReadFlag = ""
        UBCustRec(1).LocMeters(zz).MTRNUM = ""
        UBCustRec(1).LocMeters(zz).MTRMulti = -32767
        UBCustRec(1).LocMeters(zz).MTRType = ""
        UBCustRec(1).LocMeters(zz).MTRUnit = ""
        UBCustRec(1).LocMeters(zz).NumUser = -32767
        UBCustRec(1).LocMeters(zz).InsDate = -32767
        UBCustRec(1).LocMeters(zz).CurDate = -32767
        UBCustRec(1).LocMeters(zz).PastDate = -32767
        UBCustRec(1).LocMeters(zz).AvgUse = -2147483647
        UBCustRec(1).LocMeters(zz).UseCnt = -32767
      NEXT
    END IF
    FOR zz = 1 TO 7
      IF UBCustRec(1).LocMeters(zz).MTRType = "C" THEN
        UBCustRec(1).LocMeters(zz).MTRType = "W"
      END IF
    NEXT
putem:
    PUT UBFile, Cnt&, UBCustRec(1)
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Done."

