DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

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
    FOR MCnt = 1 TO 7
      MtrType$ = QPTrim$(UBCustRec(1).LocMeters(MCnt).MtrType)
      IF LEN(MtrType$) > 0 THEN
        IF UBCustRec(1).LocMeters(MCnt).MtrMulti = 100 THEN
          UBCustRec(1).LocMeters(MCnt).MtrMulti = 1
          Found = Found + 1
          'EXIT FOR
        END IF
      END IF
    NEXT
    'PUT UBFile, Cnt&, UBCustRec(1)
  NEXT
  CLOSE UBFile
PRINT
PRINT
PRINT
PRINT "Corrected:"; Found

