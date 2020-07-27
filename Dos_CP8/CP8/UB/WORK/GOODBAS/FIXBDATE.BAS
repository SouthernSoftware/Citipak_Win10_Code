DEFINT A-Z

DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)


  BadDate1 = Date2Num("12-31-1979")
  GoodDate1 = Date2Num("07-15-2005")

  '$INCLUDE: 'ubtrans.BI'

  CONST False = 0, True = NOT False

  REDIM UBTranRec(1 TO 2) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  CLS
  LOCATE 3, 1, 0
  PRINT "Correcting Transaction. . ."

  PRINT
  PRINT
  UBTFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTFile LEN = UBTranRecLen

'  OPEN "TRList.txt" FOR OUTPUT AS 10

  TNumOfRecs& = LOF(UBTFile) / UBTranRecLen
  
  FOR Cnt& = TNumOfRecs& TO 1 STEP -1
    LOCATE 5, 1: PRINT "Processing:"; Cnt&; "of"; TNumOfRecs&;
    GET UBTFile, Cnt&, UBTranRec(1)
    IF UBTranRec(1).TransType = TranUtilityBill THEN
      SELECT CASE UBTranRec(1).TransDate
      CASE BadDate1
        UBTranRec(1).TransDate = GoodDate1
        DoIt = True
      CASE ELSE
        'STOP
      END SELECT
      IF DoIt THEN
        Fixed = Fixed + 1
        PUT UBTFile, Cnt&, UBTranRec(1)
        EXIT FOR
        'DoIt =
      END IF
    END IF
  NEXT

SkiptoNext:
  CLOSE

  LOCATE 9, 1
  PRINT "Fixed:"; Fixed
  CALL killfile("fixbdate.exe")

