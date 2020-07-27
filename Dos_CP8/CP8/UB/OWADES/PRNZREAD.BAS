DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)
  '$INCLUDE: 'newcust.BI'
  CONST True = -1, False = 0

  CLS
  PRINT
  PRINT
  PRINT "Zero Reads. . ."

  REDIM UBCustRec(1) AS NewUBCustRecType

  UBCustRecLen = LEN(UBCustRec(1))

  PRINT
  PRINT
  LPRINT "Zero Reads"
  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  FOR Cnt& = 1 TO NumOfRecs&
    LOCATE , 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    GET UBFile, Cnt&, UBCustRec(1)
    DidOne = False
    FOR MCnt = 1 TO 7
      IF LEN(QPTrim$(UBCustRec(1).LocMeters(MCnt).MtrType)) > 0 THEN
      IF (UBCustRec(1).LocMeters(MCnt).CurRead <= 0) OR (UBCustRec(1).LocMeters(MCnt).PrevRead <= 0) THEN
        DidOne = True
      END IF
      END IF
    NEXT
    IF DidOne THEN
      DidCnt = DidCnt + 1
      LPRINT UBCustRec(1).CustName
    END IF
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Zero Reads:"; DidCnt

