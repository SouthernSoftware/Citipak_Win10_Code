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
  LPRINT "Active Customers with Invalid RateCode 'SW0'"
  FOR Cnt& = 1 TO NumOfRecs&
    LOCATE , 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    GET UBFile, Cnt&, UBCustRec(1)
    IF UBCustRec(1).DelFlag = 0 AND UBCustRec(1).Status = "A" THEN
      RT$ = QPTrim$(UBCustRec(1).Serv(2).RateCode)
      IF RT$ = "SW0" THEN
        'UBCustRec(1).Serv(2).RateCode = "SW1"
        'PUT UBFile, Cnt&, UBCustRec(1)
        LPRINT Cnt&, ;
        Fixed = Fixed + 1
      END IF
    END IF
  NEXT
  CLOSE UBFile
  PRINT
  PRINT
  PRINT
  PRINT "Removed:"; Fixed
  LPRINT CHR$(12)

