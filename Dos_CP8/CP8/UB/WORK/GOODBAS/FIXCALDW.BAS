DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Date2Num (TheDate$)

  '$INCLUDE: 'newcust.BI'
  '$INCLUDE: 'ubtrans.BI'

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  CLS
  PRINT
  PRINT
  PRINT "Processing Customer book 6. . ."
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  FOR Cnt& = 1 TO NumOfRecs&
    LOCATE , 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    GET UBFile, Cnt&, UBCustRec(1)
    Book$ = UBCustRec(1).Book
    SELECT CASE Book$
    CASE "06"
      UBCustRec(1).EstFlag = ""
      FOR zz = 1 TO 7
        UBCustRec(1).LocMeters(zz).ReadFlag = ""
      NEXT
      PUT UBFile, Cnt&, UBCustRec(1)
    END SELECT
  NEXT
  CLOSE UBFile
  PRINT
  PRINT
  PRINT
  PRINT "Processing Complete."

  CALL KillFile("fixcaldw.exe")

