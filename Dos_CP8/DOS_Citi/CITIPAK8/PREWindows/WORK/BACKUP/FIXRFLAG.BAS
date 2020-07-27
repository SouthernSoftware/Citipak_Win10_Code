DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Date2Num (TheDate$)

  '$INCLUDE: 'newcust.BI'

  RDate1 = Date2Num("12-31-1979")
  RDate2 = Date2Num("05-11-2001")

  CLS
  PRINT
  PRINT
  PRINT "Processing Read Flags. . ."
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
    CASE "01", "06", "07"
      FOR MCnt = 1 TO 7
         UBCustRec(1).LocMeters(MCnt).ReadFlag = ""
         DidEm = -1
      NEXT
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
  CALL KillFile("fixrflag.exe")

