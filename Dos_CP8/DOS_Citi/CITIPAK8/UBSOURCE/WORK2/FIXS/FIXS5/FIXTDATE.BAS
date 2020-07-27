DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)
  
  '$INCLUDE: 'ubtrans.BI'
  ''INCLUDE: 'newcust.BI'

  TDate1 = Date2Num%("07-28-1997")
  TDate2 = Date2Num%("08-28-1997")

  CLS
  LOCATE 3, 1, 0
  PRINT "Correcting Billing Dates. . ."

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  UBFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBFile LEN = UBTranRecLen

  FOR Cnt& = 978 TO 1352
    Fixed = Fixed + 1
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&;
    GET UBFile, Cnt&, UBTranRec(1)
    UBTranRec(1).TransDate = TDate1
    PUT UBFile, Cnt&, UBTranRec(1)
  NEXT

  FOR Cnt& = 1901 TO 2477
    Fixed = Fixed + 1
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&;
    GET UBFile, Cnt&, UBTranRec(1)
    UBTranRec(1).TransDate = TDate2
    PUT UBFile, Cnt&, UBTranRec(1)
  NEXT


  CLOSE

LOCATE 9, 1
PRINT "  Fixed:"; Fixed;

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

