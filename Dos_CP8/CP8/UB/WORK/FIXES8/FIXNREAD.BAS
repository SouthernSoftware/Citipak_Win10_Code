DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'newcust.BI'

  CONST False = 0, True = NOT False
  
  CONST BlankLng& = -2147483647

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))


  CLS
  LOCATE 3, 1, 0
  PRINT "Correcting Customer Information. . ."

  PRINT
  PRINT

  UBCFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBCFile LEN = UBCustRecLen
  CNumOfRecs& = LOF(UBCFile) \ UBCustRecLen
  
  FOR CCnt& = 1 TO CNumOfRecs&
    GET UBCFile, CCnt&, UBCustRec(1)
    BadCnt = 0
    LOCATE 5, 1
    PRINT "Processing:"; CCnt&; " of"; CNumOfRecs&;
    FOR MCnt = 1 TO 7
      MTType$ = QPTrim$(UBCustRec(1).LocMeters(MCnt).MTRType)
      MTUnit$ = QPTrim$(UBCustRec(1).LocMeters(MCnt).MTRUnit)
      IF LEN(MTType$) = 0 AND LEN(MTUnit$) = 0 THEN
        UBCustRec(1).LocMeters(MCnt).CurRead = BlankLng&
        UBCustRec(1).LocMeters(MCnt).PrevRead = BlankLng&
      END IF
    NEXT
    PUT UBCFile, CCnt&, UBCustRec(1)
  NEXT

SkiptoNext:
  CLOSE

LOCATE 9, 1
PRINT "Customers Found:"; Cust

