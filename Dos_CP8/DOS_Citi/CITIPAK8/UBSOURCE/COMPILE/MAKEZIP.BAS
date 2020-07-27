DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Processing. . ."
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
    Temp$ = "0000"
    Zip$ = QPTrim$(UBCustRec(1).ZipCode)
    Temp$ = Temp$ + QPTrim$(STR$(Cnt&))
    Temp$ = RIGHT$(Temp$, 4)
    Zip$ = Zip$ + Temp$
    UBCustRec(1).ZipCode = Zip$
    PUT UBFile, Cnt&, UBCustRec(1)

  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Done."

