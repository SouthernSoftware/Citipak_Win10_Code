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
  PRINT "Processing User Code 2 Information. . ."
  PRINT
  PRINT

  UBCFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBCFile LEN = UBCustRecLen
  NumCust& = LOF(UBCFile) \ UBCustRecLen

  FOR CCnt& = 1 TO NumCust&
    LOCATE , 1
    PRINT "Processing:"; CCnt&; "of"; NumCust&;
    GET UBCFile, CCnt&, UBCustRec(1)
    LSET UBCustRec(1).UserCode1 = "3"
    PUT UBCFile, CCnt&, UBCustRec(1)
  NEXT

  CLOSE

LOCATE 9, 1
PRINT "Complete."
CALL KILLFILE("FIXCODES.EXE")

