DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
  
  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Processing Meters. . ."

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
      DidOne = 0
      FOR MCnt = 1 TO 7
        IF UBCustRec(1).LocMeters(MCnt).MtrType = "C" THEN
          UBCustRec(1).LocMeters(MCnt).MtrType = "W"
          DidOne = -1
        END IF
      NEXT
      IF UBCustRec(1).Serv(1).RMTRTYPE = "C" THEN
        UBCustRec(1).Serv(1).RMTRTYPE = "W"
        DidOne = -1
      END IF
      IF DidOne THEN
        DidCnt = DidCnt + 1
        PUT UBFile, Cnt&, UBCustRec(1)
      END IF
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Corrected:"; DidCnt

