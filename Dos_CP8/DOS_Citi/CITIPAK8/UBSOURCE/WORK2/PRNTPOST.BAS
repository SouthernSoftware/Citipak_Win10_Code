DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
  
  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Looking for Postal Routes. . ."
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

  NumOfRec = LOF(UBFile) \ UBCustRecLen

  FOR Cnt = 1 TO NumOfRec
    GET UBFile, Cnt, UBCustRec(1)
    IF LEN(QPTrim$(UBCustRec(1).PostRte)) > 0 THEN
      LPRINT Cnt
    END IF
  NEXT

  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Done"

