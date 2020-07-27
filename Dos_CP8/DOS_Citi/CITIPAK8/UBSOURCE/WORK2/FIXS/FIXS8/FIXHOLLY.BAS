DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Date2Num% (TheDate$)

  '$INCLUDE: 'newcust.BI'
  NewDate = Date2Num("08-20-1998")

  CLS
  PRINT
  PRINT
  PRINT "Processing Read Dates in book 05. . ."
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
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
    IF UBCustRec(1).Book = "05" THEN
      Found = Found + 1
      FOR zz = 1 TO 7
        IF UBCustRec(1).LocMeters(zz).CurDate > 0 THEN
          UBCustRec(1).LocMeters(zz).CurDate = NewDate
        END IF
      NEXT
      PUT UBFile, Cnt&, UBCustRec(1)
    END IF
  NEXT
  CLOSE UBFile
PRINT
PRINT
PRINT
PRINT "Changed:"; Found

