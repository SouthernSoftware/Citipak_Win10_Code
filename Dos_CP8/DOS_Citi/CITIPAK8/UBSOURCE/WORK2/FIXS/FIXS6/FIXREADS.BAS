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
      FOR MCnt = 1 TO 7
        IF UBCustRec(1).LocMeters(MCnt).CurRead > 0 THEN
          UBCustRec(1).LocMeters(MCnt).CurRead = UBCustRec(1).LocMeters(MCnt).CurRead * 1000
        END IF
      NEXT
      PUT UBFile, Cnt&, UBCustRec(1)
    END IF
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Done."

