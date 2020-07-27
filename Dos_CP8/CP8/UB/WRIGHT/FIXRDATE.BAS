DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Date2Num (TheDate$)

  '$INCLUDE: 'newcust.BI'

  RDate1 = Date2Num("03-17-1999")
  RDate2 = Date2Num("03-22-1999")

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
    DidEm = 0
    SELECT CASE UBCustRec(1).Book
    CASE "08"
      FOR MCnt = 1 TO 7
        UBCustRec(1).LocMeters(MCnt).CurDate = RDate1
      NEXT
      DidEm = -1
    CASE "09"
      FOR MCnt = 1 TO 7
        UBCustRec(1).LocMeters(MCnt).CurDate = RDate2
      NEXT
      DidEm = -1
    END SELECT
    IF DidEm THEN
      FixedCnt = FixedCnt + 1
      PUT UBFile, Cnt&, UBCustRec(1)
    END IF
  NEXT
  CLOSE UBFile
PRINT
PRINT
PRINT
PRINT "Fixed:"; FixedCnt

