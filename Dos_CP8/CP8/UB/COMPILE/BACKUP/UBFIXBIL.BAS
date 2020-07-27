DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)

  '$INCLUDE: 'ubtrans.BI'

  TranDate = Date2Num("03-01-1997")
  PastDate = Date2Num("03-20-1997")

  CLS
  LOCATE 3, 1, 0
  PRINT "Adjusting Billing Date. . ."
  REDIM UBTranRec(1) AS UBTransRecType

  UBTranRecLen = LEN(UBTranRec(1))

  UBFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBFile LEN = UBTranRecLen
  NumOfRecs& = LOF(UBFile) / UBTranRecLen
  FOR Cnt& = 1 TO NumOfRecs&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    GET UBFile, Cnt&, UBTranRec(1)
    IF (UBTranRec(1).TransDate = 0) AND (UBTranRec(1).TransType = TranUtilityBill) THEN
      Fixed = Fixed + 1
      UBTranRec(1).PastDueDate = PastDate
      UBTranRec(1).TransDate = TranDate
      UBTranRec(1).ReadDate = TranDate
      UBTranRec(1).BillDate = TranDate
      PUT UBFile, Cnt&, UBTranRec(1)
    END IF
  NEXT
  CLOSE UBFile

LOCATE 9, 1
PRINT "Fixed:"; Fixed

