DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
  
  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Processing Flat Rates. . ."
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
    FOR MCnt = 1 TO 4
      'IF UCASE$(UBCustRec(1).FlatRates(MCnt).FRFREQ) = "E" THEN
       IF UBCustRec(1).FlatRates(MCnt).FRAmt < 0 THEN
         LPRINT Cnt&
       END IF
        'UBCustRec(1).FlatRates(MCnt).FRFREQ = "R"
        'Fixed = Fixed + 1
      'END IF
    NEXT
    'PUT UBFile, Cnt&, UBCustRec(1)
  NEXT
  CLOSE UBFile

  PRINT
  PRINT
PRINT
PRINT "Corrected:"; Fixed

