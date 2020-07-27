DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Processing. . ."
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
    DidEmOne = 0
    FOR MCnt = 1 TO 7
      SELECT CASE UBCustRec(1).LocMeters(MCnt).MtrType
        CASE "W", "S", "C", "T"
          DidEmOne = -1
          UBCustRec(1).LocMeters(MCnt).MtrMulti = 100
        CASE ELSE
      END SELECT
      PUT UBFile, Cnt&, UBCustRec(1)
      Fixed = Fixed + 1
    NEXT
    IF DidEmOne THEN
      PUT UBFile, Cnt&, UBCustRec(1)
      Fixed = Fixed + 1
    END IF
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Corrected: "; Fixed

