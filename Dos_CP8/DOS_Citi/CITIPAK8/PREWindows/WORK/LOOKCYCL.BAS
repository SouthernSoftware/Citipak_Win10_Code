DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)
  
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.BI'

  CONST False = 0, True = NOT False

  REDIM Cycles(-1 TO 99) AS INTEGER

  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))


  CLS
  LOCATE 3, 1, 0
  PRINT "Billing Cycles. . ."
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumCust& = LOF(UBFile) \ UBCustRecLen

  
  FOR CCnt& = 1 TO NumCust&
    GET UBFile, CCnt&, UBCustRec(1)
    IF UBCustRec(1).BillCycl >= 0 AND (NOT UBCustRec(1).DelFlag) THEN
      Cycles(UBCustRec(1).BillCycl) = Cycles(UBCustRec(1).BillCycl) + 1
      'PRINT CCnt&, UBCustRec(1).BillCycl, UBCustRec(1).Status
      'STOP
    ELSE
      IF NOT UBCustRec(1).DelFlag THEN
        Cycles(-1) = Cycles(-1) + 1
      END IF
    END IF
'    PUT UBCFile, CCnt&, UBCustRec(1)
  NEXT

  CLOSE
FOR zz = 0 TO 99
  IF Cycles(zz) > 0 THEN
    PRINT zz, Cycles(zz)
  END IF
NEXT
PRINT "??", Cycles(-1)
'LOCATE 9, 1
PRINT "Processing Complete."

CALL KillFile("fixcycle.exe")

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

