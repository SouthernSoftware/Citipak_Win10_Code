DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'newcust.BI'

  CLS

  PRINT
  PRINT
  PRINT "Correcting Bill Cycles. . ."
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
    WCode$ = QPTrim$(UBCustRec(1).Serv(1).RateCode)
    SCode$ = QPTrim$(UBCustRec(1).Serv(2).RateCode)
    IF LEN(WCode$) > 0 THEN
      IF LEFT$(WCode$, 1) <> "W" THEN
        STOP
      END IF
    END IF
    IF LEN(SCode$) > 0 THEN
      IF LEFT$(SCode$, 1) <> "S" THEN
        STOP
      END IF
    END IF


'    PUT UBFile, Cnt&, UBCustRec(1)
SkipThisOne:
  NEXT
  CLOSE UBFile
PRINT
PRINT
PRINT
PRINT "Done"
CALL KILLFile("FIXBROAD.EXE")

