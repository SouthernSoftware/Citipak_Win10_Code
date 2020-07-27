DEFINT A-Z
DECLARE FUNCTION Date2Num% (Today$)

  GoodDate1 = Date2Num("02-22-2001")
  GoodDate2 = Date2Num("06-01-2000")
  BadDate = Date2Num("12-31-1979")
  
  '$INCLUDE: 'ubtrans.BI'

  ''$INCLUDE: 'newcust.BI'

  CONST False = 0, True = NOT False

  'REDIM UBCustRec(1) AS NewUBCustRecType
  'UBCustRecLen = LEN(UBCustRec(1))

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

  FOR Cnt& = 1 TO TNumOfRecs&
    GET UBTFile, Cnt&, UBTranRec(1)
    IF UBTranRec(1).TransType = 1 THEN
      IF UBTranRec(1).TransDate = BadDate THEN
        IF Cnt& > 6000 THEN
          IF Cnt& < 14000 THEN
            UBTranRec(1).TransDate = GoodDate1
            Fixed = Fixed + 1
          ELSE
            Fixed = Fixed + 1
            UBTranRec(1).TransDate = GoodDate2
          END IF
          PUT UBTFile, Cnt&, UBTranRec(1)
        END IF
      END IF
    END IF
  NEXT

SkiptoNext:
  CLOSE

LOCATE 9, 1
PRINT "Fixed:"; Fixed
CALL KillFile("fixmorvn.exe")

