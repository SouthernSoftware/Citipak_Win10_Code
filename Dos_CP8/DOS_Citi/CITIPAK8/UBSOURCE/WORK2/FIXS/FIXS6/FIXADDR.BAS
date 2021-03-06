DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
  
  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Processing Line 2 Address. . ."
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
    Addr1$ = QPTrim$(UBCustRec(1).Addr1)
    IF LEN(Addr1$) = 0 THEN
      UBCustRec(1).Addr1 = UBCustRec(1).Addr2
      PUT UBFile, Cnt&, UBCustRec(1)
    END IF
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Done."

