DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
  
  '$INCLUDE: 'newcust.BI'

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))


  CLS
  PRINT
  PRINT
  PRINT "Processing. . ."
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

  NumOfRecs& = LOF(UBFile) \ UBCustRecLen

  FOR Cnt& = 1 TO NumOfRecs&
    LOCATE , 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    GET UBFile, Cnt&, UBCustRec(1)

    FOR zz = 1 TO 7
      IF UBCustRec(1).LocMeters(zz).CurRead > 0 THEN
        IF UBCustRec(1).LocMeters(zz).MTRMulti <> 100 THEN
          UBCustRec(1).LocMeters(zz).MTRMulti = 100
          Found = Found + 1
          'EXIT FOR
        END IF
      END IF
    NEXT
    PUT UBFile, Cnt&, UBCustRec(1)
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Found:"; Found
KILL "fixmulti.exe"

