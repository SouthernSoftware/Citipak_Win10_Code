DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
  
  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Changing Zip 28105 to 28104. . ."
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfCust = LOF(UBFile) / UBCustRecLen
  FOR Cnt = 1 TO NumOfCust
    GET UBFile, Cnt, UBCustRec(1)
    LOCATE 5, 1
    PRINT "Processing:"; Cnt; " of"; NumOfCust;
    IF LEFT$(UBCustRec(1).ZipCode, 5) = "28105" THEN
      MID$(UBCustRec(1).ZipCode, 1, 5) = "28104"
      DidCnt = DidCnt + 1
      PUT UBFile, Cnt, UBCustRec(1)
    END IF
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Changed:"; DidCnt

