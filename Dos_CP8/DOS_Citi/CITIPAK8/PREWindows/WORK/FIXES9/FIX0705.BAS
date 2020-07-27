DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)

  GoodDate = Date2Num("07-05-2000")
  BadDate = Date2Num("12-31-1979")


  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.BI'

  CONST False = 0, True = NOT False

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBTranRec(1 TO 2) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  CLS
  LOCATE 3, 1, 0
  PRINT "Correcting Billing Information. . ."

  PRINT
  PRINT
  
  UBTFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTFile LEN = UBTranRecLen
  TNumOfRecs& = LOF(UBTFile) / UBTranRecLen

  FOR Cnt& = 869 TO TNumOfRecs&
    GET UBTFile, Cnt&, UBTranRec(1)
    IF UBTranRec(1).TransDate = BadDate THEN
      UBTranRec(1).TransDate = GoodDate
      PUT UBTFile, Cnt&, UBTranRec(1)
      Fixed = Fixed + 1
    END IF
  NEXT

SkiptoNext:
  CLOSE

call killfile("fix0705.exe")

LOCATE 9, 1
PRINT "Fixed:"; Fixed

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

