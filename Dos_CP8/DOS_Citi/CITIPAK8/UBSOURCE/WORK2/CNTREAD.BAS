DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)
  
  '$INCLUDE: 'newcust.BI'

  CONST False = 0, True = NOT False
  
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
      IF UBCustRec(1).LocMeters(MCnt).CurRead = 0 AND UBCustRec(1).LocMeters(MCnt).PrevRead = 0 THEN
        Cust = Cust + 1
      END IF
    NEXT
  NEXT

SkiptoNext:
  CLOSE

LOCATE 9, 1
PRINT "Customers Found:"; Cust

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

