DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Processing Customers. . ."
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
    IF UBCustRec(1).Status <> "A" THEN
      IF UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance < 0 THEN
         LPRINT Cnt&, UBCustRec(1).CustName, USING "####.##"; UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance
      END IF
    END IF
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Done."

