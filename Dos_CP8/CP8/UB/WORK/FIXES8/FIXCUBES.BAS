DEFINT A-Z
  
  '$INCLUDE: 'newcust.BI'
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
   
  CLS
  LOCATE 3, 1, 0
  PRINT "Processing Books 11, 12, 13, 14. . ."

  PRINT
  PRINT

  UBCFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBCFile LEN = UBCustRecLen
  CNumOfRecs& = LOF(UBCFile) \ UBCustRecLen
  FOR Cnt& = 1 TO CNumOfRecs&
    GET UBCFile, Cnt&, UBCustRec(1)
    SELECT CASE UBCustRec(1).Book
    CASE "11", "12", "13", "14"
      FOR MtrCnt = 1 TO 7
        if UBCustRec(1).
      PUT UBCFile, Cnt&, UBCustRec(1)

    CASE ELSE
    END SELECT

  NEXT

LOCATE 9, 1
PRINT "Processing Complete."

