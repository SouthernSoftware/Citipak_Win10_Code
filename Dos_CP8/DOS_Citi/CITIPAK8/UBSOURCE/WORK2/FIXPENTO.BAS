DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
  
  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Processing Customers. . ."
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
    IF UBCustRec(1).Status = "A" THEN
      DidCnt = DidCnt + 1
      SELECT CASE UBCustRec(1).Seq
        CASE 1 TO 13599
          UBCustRec(1).Book = "01"
        CASE 13600 TO 27899
          UBCustRec(1).Book = "02"
        CASE 27900 TO 42899
          UBCustRec(1).Book = "03"
        CASE 42900 TO 53299
          UBCustRec(1).Book = "04"
        CASE 53300 TO 63299
          UBCustRec(1).Book = "05"
        CASE 63300 TO 76399
          UBCustRec(1).Book = "06"
        CASE 76400 TO 86599
          UBCustRec(1).Book = "07"
        CASE 86600 TO 103099
          UBCustRec(1).Book = "08"
        CASE 103100 TO 118299
          UBCustRec(1).Book = "09"
        CASE 118300 TO 999999
          UBCustRec(1).Book = "10"
      END SELECT
      PUT UBFile, Cnt&, UBCustRec(1)
    END IF
  NEXT
  CLOSE UBFile
  PRINT
  PRINT
  PRINT
  PRINT "Changed:"; DidCnt
  KILL "fixpento.exe"

