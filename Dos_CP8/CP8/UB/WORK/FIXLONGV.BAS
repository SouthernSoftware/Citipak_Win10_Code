DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'newcust.BI'

  CLS

  PRINT
  PRINT
  PRINT "Processing Meter Types. . ."
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
    FOR MTCnt = 1 TO 7
      MType$ = QPTrim$(UBCustRec(1).LocMeters(MTCnt).MTRType)
      SELECT CASE MType$
      CASE "T"
        MtrNumb$ = QPTrim$(UBCustRec(1).LocMeters(MTCnt).MtrNum)
        IF LEN(MtrNumb$) > 0 THEN
          MtrNumb$ = MtrNumb$ + "A"
          LSET UBCustRec(1).LocMeters(MTCnt).MtrNum = MtrNumb$
          Fixed# = Fixed# + 1
        END IF
      CASE "W", "S", "C"
        MtrNumb$ = QPTrim$(UBCustRec(1).LocMeters(MTCnt).MtrNum)
        IF LEN(MtrNumb$) > 0 THEN
          MtrNumb$ = MtrNumb$ + "M"
          LSET UBCustRec(1).LocMeters(MTCnt).MtrNum = MtrNumb$
          Fixed# = Fixed# + 1
        END IF
      CASE ELSE
      END SELECT
    NEXT
    PUT UBFile, Cnt&, UBCustRec(1)
SkipThisOne:
  NEXT
  CLOSE UBFile
  PRINT
  PRINT
  PRINT
  PRINT "Changed:"; Fixed#
  CALL KILLFile("FIXLONGV.EXE")

