DEFINT A-Z

  '$INCLUDE: 'newcust.BI'


  CLS
  PRINT
  PRINT
  PRINT "Processing. . ."
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
    SELECT CASE UBCustRec(1).Book
    CASE "07", "08", "09", "10", "11"
      UBCustRec(1).BillCycl = 98
      PUT UBFile, Cnt&, UBCustRec(1)
      DidEm = DidEm + 1
    END SELECT
    
  NEXT
  CLOSE UBFile
PRINT
PRINT
PRINT
PRINT "Done."; DidEm

CALL KillFile("setcycle.exe")

