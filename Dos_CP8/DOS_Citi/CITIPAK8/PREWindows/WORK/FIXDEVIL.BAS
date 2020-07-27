DEFINT A-Z
  ON ERROR RESUME NEXT
  '$INCLUDE: 'newcust.BI'

  CLS

  PRINT
  PRINT
  PRINT "Processing Meter Readings. . ."
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
    UBCustRec(1).LocMeters(1).MtrMulti = 1
    IF UBCustRec(1).LocMeters(1).PrevRead > 0 THEN
      UBCustRec(1).LocMeters(1).PrevRead = (UBCustRec(1).LocMeters(1).PrevRead * 1000)
    END IF
    IF UBCustRec(1).LocMeters(1).CurRead > 0 THEN
      UBCustRec(1).LocMeters(1).CurRead = (UBCustRec(1).LocMeters(1).CurRead * 1000)
    END IF
    PUT UBFile, Cnt&, UBCustRec(1)
SkipThisOne:
  NEXT
  CLOSE UBFile
PRINT
PRINT
PRINT
PRINT "Done."
CALL KILLFile("FIXDEVIL.EXE")

