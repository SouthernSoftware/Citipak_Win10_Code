DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'newcust.BI'

  CLS

  PRINT
  PRINT
  PRINT "Correcting Meter Types. . ."
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
    HasOne = 0
    HasTwo = 0
    IF UBCustRec(1).DelFlag <> 0 THEN
      GOTO SkipThisOne
    END IF
    Serv1$ = QPTrim$(UBCustRec(1).Serv(1).RateCode)
    Serv2$ = QPTrim$(UBCustRec(1).Serv(2).RateCode)
    IF LEN(Serv1$) > 0 THEN
      HasOne = -1
    END IF
    IF LEN(Serv2$) > 0 THEN
      HasTwo = -1
    END IF

    IF HasOne AND HasTwo THEN
      MtrType$ = "C"
      IF UBCustRec(1).Serv(1).RMtrType = "T" THEN
        UBCustRec(1).Serv(1).RMtrType = MtrType$
      END IF
      IF UBCustRec(1).Serv(2).RMtrType = "T" THEN
        UBCustRec(1).Serv(2).RMtrType = MtrType$
      END IF
    ELSEIF HasOne THEN
      MtrType$ = "W"
      IF UBCustRec(1).Serv(1).RMtrType = "T" THEN
        UBCustRec(1).Serv(1).RMtrType = MtrType$
      END IF
    ELSEIF HasTwo THEN
      MtrType$ = "S"
      IF UBCustRec(1).Serv(2).RMtrType = "T" THEN
        UBCustRec(1).Serv(2).RMtrType = MtrType$
      END IF
    ELSE
      MtrType$ = ""
      GOTO SkipThisOne
    END IF

    PUT UBFile, Cnt&, UBCustRec(1)
SkipThisOne:
  NEXT
  CLOSE UBFile
PRINT
PRINT
PRINT
PRINT "Done"
CALL KILLFile("FIXwarrt.EXE")

