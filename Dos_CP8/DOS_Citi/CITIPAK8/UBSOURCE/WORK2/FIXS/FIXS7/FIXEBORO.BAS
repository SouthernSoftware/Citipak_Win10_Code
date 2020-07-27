DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Date2Num% (thedate$)


  '$INCLUDE: 'newcust.BI'

  NDate = Date2Num("05-22-1998")
  
  CLS
  PRINT
  PRINT
  PRINT "Processing Read Dates. . ."
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

    IF UBCustRec(1).Book = "01" THEN
      DidOne = 0
      FOR MCnt = 1 TO 7
        IF UBCustRec(1).LocMeters(MCnt).CurDate > 0 THEN
          DidOne = -1
          UBCustRec(1).LocMeters(MCnt).CurDate = NDate
        END IF
      NEXT
      IF DidOne = -1 THEN
        DidCnt = DidCnt + 1
      END IF
      PUT UBFile, Cnt&, UBCustRec(1)
    END IF
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Corrected:"; DidCnt

