DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
  
  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Processing Readings. . ."
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
    IF UBCustRec(1).Status = "A" THEN
      IF (UBCustRec(1).LocMeters(1).CurRead > 0) THEN
        IF UBCustRec(1).LocMeters(1).AvgUse = 0 THEN
          LPRINT "Acct:"; Cnt&
        ELSE
          TUse& = UBCustRec(1).LocMeters(1).AvgUse * 2
          UBCustRec(1).LocMeters(1).CurRead = UBCustRec(1).LocMeters(1).CurRead + TUse&
          PUT UBFile, Cnt&, UBCustRec(1)
        END IF
      END IF
    END IF
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Done."

