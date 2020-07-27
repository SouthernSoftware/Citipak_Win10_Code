DEFINT A-Z
  
  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Clearing Deleted Accounts. . ."
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  UBCustRec(2).DelFlag = -1

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

  NumOfRec = LOF(UBFile) \ UBCustRecLen
  FOR Cnt = 1 TO NumOfRec
    LOCATE , 1
    PRINT "  Processing:"; Cnt; " of"; NumOfRec;
    GET UBFile, Cnt, UBCustRec(1)
    IF UBCustRec(1).DelFlag <> 0 THEN
      PUT UBFile, Cnt, UBCustRec(2)
    END IF
  NEXT
  CLOSE UBFile

PRINT
SHELL "ubmisc reindex"
PRINT
PRINT "Done"

