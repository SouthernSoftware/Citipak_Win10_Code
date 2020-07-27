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
  PRINT "Change Billing Cycles. . ."
  PRINT
  PRINT

  UBCFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBCFile LEN = UBCustRecLen
  NumCust& = LOF(UBCFile) \ UBCustRecLen

  UBOFile = FREEFILE
  OPEN "OCUST.dat" FOR RANDOM SHARED AS UBOFile LEN = UBCustRecLen
  ONumCust& = LOF(UBOFile) \ UBCustRecLen
  
  FOR CCnt& = 1 TO NumCust&
    GET UBCFile, CCnt&, UBCustRec(1)
    GET UBOFile, CCnt&, UBCustRec(2)
    UBCustRec(1).Seq = UBCustRec(2).Seq
    PUT UBCFile, CCnt&, UBCustRec(1)
  NEXT

  CLOSE

LOCATE 9, 1
PRINT "Processing Complete."
CALL KillFile("fixcycle.exe")

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

