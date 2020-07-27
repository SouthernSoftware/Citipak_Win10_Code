DEFINT A-Z
DECLARE FUNCTION Round# (N#)
DECLARE FUNCTION FileSize& (FileName$)
  
  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Zeroing Flats. . ."
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
    UBCustRec(1).pumpcode = "0"

    UBCustRec(1).USERCODE1 = "0"
    UBCustRec(1).USERCODE2 = "0"

    UBCustRec(1).FlatRates(1).FrAmt = 0
    UBCustRec(1).HHMSG1 = ""
    UBCustRec(1).HHMSG2 = ""
    UBCustRec(1).HHMSG3 = ""
    PUT UBFile, Cnt&, UBCustRec(1)
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Finished."

CALL killfile("fixemup.exe")

