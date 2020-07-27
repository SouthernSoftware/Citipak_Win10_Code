DEFINT A-Z
  'ON ERROR RESUME NEXT
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
    DidEm = 0
    LOCATE , 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    GET UBFile, Cnt&, UBCustRec(1)
    IF UBCustRec(1).DelFlag <> 0 THEN
      GOTO SkipThisOne
    END IF
    FOR MCnt = 1 TO 7
      IF UBCustRec(1).LocMeters(MCnt).MtrType = "T" THEN
        UBCustRec(1).LocMeters(MCnt).MtrMulti = 1000
        UBCustRec(1).LocMeters(MCnt).PrevRead = UBCustRec(1).LocMeters(MCnt).PrevRead / 1000
        UBCustRec(1).LocMeters(MCnt).CurRead = UBCustRec(1).LocMeters(MCnt).CurRead / 1000
        DidEm = -1
      END IF
    NEXT
    IF DidEm THEN
      PUT UBFile, Cnt&, UBCustRec(1)
      Fixed& = Fixed& + 1
    END IF
SkipThisOne:
  NEXT
  CLOSE UBFile
PRINT
PRINT
PRINT
PRINT "Fixed:"; Fixed&
CALL KILLFile("FIXDEVIL.EXE")

