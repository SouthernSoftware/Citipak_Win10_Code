DEFINT A-Z
  
  '$INCLUDE: 'newcust.BI'

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))


  CLS
  PRINT
  PRINT
  PRINT "Processing Book '6'. . ."
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

  NumOfRecs& = LOF(UBFile) \ UBCustRecLen

  FOR Cnt& = 1 TO NumOfRecs&
    LOCATE , 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    GET UBFile, Cnt&, UBCustRec(1)
    IF UBCustRec(1).Book = "06" THEN
        FOR MCnt = 1 TO 7
          IF UBCustRec(1).LocMeters(MCnt).MtrType = "E" THEN
            ReadAmt# = UBCustRec(1).LocMeters(MCnt).CurRead - UBCustRec(1).LocMeters(MCnt).PrevRead
            UBCustRec(1).LocMeters(MCnt).CurRead = UBCustRec(1).LocMeters(MCnt).CurRead + ReadAmt#
            PUT UBFile, Cnt&, UBCustRec(1)
            DidCnt = DidCnt + 1
          END IF
        NEXT
      END IF

  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Adjusted:"; DidCnt

