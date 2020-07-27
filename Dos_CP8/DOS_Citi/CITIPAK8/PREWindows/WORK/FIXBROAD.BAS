DEFINT A-Z

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
    IF UBCustRec(1).book = "03" THEN
      IF UBCustRec(1).locmeters(1).prevread < 0 OR UBCustRec(1).locmeters(1).Curread < 0 THEN
        LPRINT Cnt&
      END IF
    END IF

    UBCustRec(1).BillCycl = 1
'    PUT UBFile, Cnt&, UBCustRec(1)
SkipThisOne:
  NEXT
  CLOSE UBFile
PRINT
PRINT
PRINT
PRINT "Done"
CALL KILLFile("FIXBROAD.EXE")

