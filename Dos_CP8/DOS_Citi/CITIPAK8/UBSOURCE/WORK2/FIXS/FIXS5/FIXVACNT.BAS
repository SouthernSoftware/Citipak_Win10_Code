DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Processing Address. . ."
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
    IF QPTrim$(UBCustRec(1).Search) = "VACANT" THEN
      'LPRINT Cnt&
      UBCustRec(1).Status = "I"
      PUT UBFile, Cnt&, UBCustRec(1)
      Fixed = Fixed + 1
    END IF
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Corrected: "; Fixed

