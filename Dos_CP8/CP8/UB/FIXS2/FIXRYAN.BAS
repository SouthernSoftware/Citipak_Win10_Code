DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
  
  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Correcting Customer Status. . ."
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfCust = LOF(UBFile) \ UBCustRecLen
  FOR Cnt = 1 TO NumOfCust
    LOCATE 5, 1
    PRINT "Processing:"; Cnt; "of "; NumOfCust;
    GET UBFile, Cnt, UBCustRec(1)
    SELECT CASE UCASE$(UBCustRec(1).Status)
    CASE "Y"
      UBCustRec(1).Status = "A"
    CASE "N"
      UBCustRec(1).Status = "I"
    END SELECT
    PUT UBFile, Cnt, UBCustRec(1)
  NEXT

  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Done"

