DEFINT A-Z

DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)


  BadDate1 = Date2Num("12-31-1979")
  GoodDate1 = Date2Num("10-18-2006")
  TRDate = Date2Num("10-30-2006")
  '$INCLUDE: 'ubtrans.BI'

  CONST False = 0, True = NOT False

  REDIM UBTranRec(1 TO 2) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  CLS
  LOCATE 3, 1, 0
  PRINT "Correcting Transaction Read Dates..."

  PRINT
  PRINT
  UBTFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTFile LEN = UBTranRecLen

'  OPEN "TRList.txt" FOR OUTPUT AS 10

  TNumOfRecs& = LOF(UBTFile) / UBTranRecLen
  
  FOR Cnt& = 1 TO TNumOfRecs&  'TNumOfRecs& TO 1 STEP -1
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; "of"; TNumOfRecs&;
    GET UBTFile, Cnt&, UBTranRec(1)
    IF UBTranRec(1).TransType = TranUtilityBill THEN
      IF UBTranRec(1).TransDate = TRDate THEN
        SELECT CASE UBTranRec(1).ReadDate
        CASE BadDate1
          UBTranRec(1).ReadDate = GoodDate1
          Fixed& = Fixed& + 1
          PUT UBTFile, Cnt&, UBTranRec(1)
        END SELECT
      END IF
    END IF
  NEXT

SkiptoNext:
  CLOSE
  LOCATE 9, 1
  PRINT "Fixed:"; Fixed&
  CALL killfile("fixbowmn.exe")

