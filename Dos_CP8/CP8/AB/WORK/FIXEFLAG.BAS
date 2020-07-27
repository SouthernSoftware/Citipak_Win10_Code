DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Date2Num (TheDate$)

  '$INCLUDE: 'newcust.BI'
  '$INCLUDE: 'ubtrans.BI'

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  CLS
  PRINT
  PRINT
  PRINT "Processing Customers. . ."
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  FOR Cnt& = 1 TO NumOfRecs&
    LOCATE , 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    GET UBFile, Cnt&, UBCustRec(1)
    UBCustRec(1).EstFlag = ""
    PUT UBFile, Cnt&, UBCustRec(1)
  NEXT
  CLOSE UBFile
  PRINT
  PRINT
  PRINT
  PRINT "Processing Transactions. . ."
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBTRANS.dat" FOR RANDOM SHARED AS UBFile LEN = UBTranRecLen

  NumOfRecs& = LOF(UBFile) \ UBTranRecLen
  FOR Cnt& = 1 TO NumOfRecs&
    LOCATE , 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    GET UBFile, Cnt&, UBTranRec(1)
    FOR zz = 1 TO 7
      UBTranRec(1).EstRead(zz) = ""
    NEXT
    PUT UBFile, Cnt&, UBTranRec(1)
  NEXT
  CLOSE
  PRINT
  PRINT
  PRINT
  PRINT "Processing Complete."

  CALL KillFile("fixeflag.exe")

