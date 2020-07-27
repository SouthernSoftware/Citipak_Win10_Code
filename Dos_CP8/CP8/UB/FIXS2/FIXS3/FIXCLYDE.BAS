DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Date2Num (TheDate$)

  '$INCLUDE: 'newcust.BI'

  BDate = Date2Num("12-31-1979")
  GDate = Date2Num("02-19-1999")
  
  CLS
  PRINT
  PRINT
  PRINT "Processing Invalid Reading Dates. . ."
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
    IF UBCustRec(1).DelFlag = 0 THEN
      DidOne = 0
      FOR MCnt = 1 TO 7
        IF UBCustRec(1).LocMeters(MCnt).CurDate = BDate AND UBCustRec(1).LocMeters(MCnt).ReadFlag = "Y" THEN
          DidOne = -1
          UBCustRec(1).LocMeters(MCnt).CurDate = GDate
        END IF
      NEXT
      IF DidOne = -1 THEN
        DidCnt = DidCnt + 1
        PUT UBFile, Cnt&, UBCustRec(1)
      END IF
    END IF
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Corrected:"; DidCnt

