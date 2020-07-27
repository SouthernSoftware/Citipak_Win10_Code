DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)
  
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.BI'

  CONST False = 0, True = NOT False
  
  REDIM UBCustREc(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustREc(1))


  CLS
  LOCATE 3, 1, 0
  PRINT "Correcting Customer Information. . ."

  PRINT
  PRINT

  UBCFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBCFile LEN = UBCustRecLen
  CNumOfRecs& = LOF(UBCFile) \ UBCustRecLen
  
  FOR CCnt& = 1 TO CNumOfRecs&
    GET UBCFile, CCnt&, UBCustREc(1)
    BadCnt = 0
    LOCATE 5, 1
    PRINT "Processing:"; CCnt&; " of"; CNumOfRecs&;
    UBCustREc(1).LocMeters(1).MTRType = UBCustREc(1).Serv(1).RMtrType
    IF UBCustREc(1).LocMeters(1).CurRead < 0 THEN
      UBCustREc(1).LocMeters(1).CurRead = 0
    END IF
    IF UBCustREc(1).LocMeters(1).PrevRead < 0 THEN
      UBCustREc(1).LocMeters(1).PrevRead = 0
    END IF

    PUT UBCFile, CCnt&, UBCustREc(1)
  NEXT

SkiptoNext:
  CLOSE

LOCATE 9, 1
PRINT "Processing Complete."

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

