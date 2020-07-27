DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)
  
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.BI'

  CONST False = 0, True = NOT False
  
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))


  CLS
  LOCATE 3, 1, 0
  PRINT "Correcting Customer Information. . ."
  PRINT
  PRINT

  UBCFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBCFile LEN = UBCustRecLen
  NumCust& = LOF(UBCFile) \ UBCustRecLen
  '

  FOR CCnt& = 1 TO NumCust&
    LOCATE , 1
    PRINT "Processing:"; CCnt&; "of"; NumCust&;
    GET UBCFile, CCnt&, UBCustRec(1)
    IF UBCustRec(1).DelFlag = 0 THEN
      FOR zz = 1 TO 4


    PUT UBCFile, CCnt&, UBCustRec(1)
  '  PUT UBCFile, CCnt&, UBCustRec(1)
  NEXT

  CLOSE

LOCATE 9, 1
PRINT "Processing Complete."

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

