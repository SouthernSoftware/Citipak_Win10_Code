DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)
  
  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Processing Accounts 4038 thru 4176. . ."
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  FOR Cnt& = 4038 TO 4176
    LOCATE , 1
    PRINT "Processing:"; Cnt&;
    GET UBFile, Cnt&, UBCustRec(1)
    UBCustRec(1).FlatRates(1).FRDESC = "LANDFILL"
    UBCustRec(1).FlatRates(1).FRAMT = 5
    UBCustRec(1).FlatRates(1).FRFREQ = "R"
    UBCustRec(1).FlatRates(1).REVSRC = 1
    UBCustRec(1).FlatRates(1).NumMin = 3

    UBCustRec(1).FlatRates(2).FRDESC = "LANDFILL"
    UBCustRec(1).FlatRates(2).FRAMT = 30
    UBCustRec(1).FlatRates(2).FRFREQ = "N"
    UBCustRec(1).FlatRates(2).REVSRC = 1
    UBCustRec(1).FlatRates(2).NumMin = 1
    PUT UBFile, Cnt&, UBCustRec(1)
  NEXT
  PRINT
  PRINT "Processing Accounts 4177 thru 4342. . ."
  FOR Cnt& = 4177 TO 4342
    LOCATE , 1
    PRINT "Processing:"; Cnt&;
    GET UBFile, Cnt&, UBCustRec(1)
    UBCustRec(1).FlatRates(1).FRDESC = "LANDFILL"
    UBCustRec(1).FlatRates(1).FRAMT = 5
    UBCustRec(1).FlatRates(1).FRFREQ = "R"
    UBCustRec(1).FlatRates(1).REVSRC = 1
    UBCustRec(1).FlatRates(1).NumMin = 3
    PUT UBFile, Cnt&, UBCustRec(1)
  NEXT

  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Done."

