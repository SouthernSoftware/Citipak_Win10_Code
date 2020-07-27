DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)

  BadDate1 = Date2Num("12-31-1979")
  GoodDate1 = Date2Num("07-30-2003")
  GoodDate2 = Date2Num("12-30-1999")

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
  
  FOR Cnt& = TNumOfRecs& TO TNumOfRecs& - 10000 STEP -1
    LOCATE 5, 1: PRINT "Processing:"; Cnt&; "of"; TNumOfRecs&;
    GET UBTFile, Cnt&, UBTranRec(1)
    IF UBTranRec(1).TransType = TranUtilityBill THEN
      SELECT CASE UBTranRec(1).TransDate
      CASE BadDate1
        UBTranRec(1).TransDate = GoodDate1
        DoIt = True
      CASE ELSE
      END SELECT
      IF DoIt THEN
        Fixed = Fixed + 1
        PUT UBTFile, Cnt&, UBTranRec(1)
        DoIt = False
      END IF
    END IF
  NEXT

  LOCATE 5, 1: PRINT "Processing:                                     "
  FOR Cnt& = 1 TO TNumOfRecs& - 10000
    LOCATE 5, 1: PRINT "Processing:"; Cnt&; "of"; TNumOfRecs&;
    GET UBTFile, Cnt&, UBTranRec(1)
    IF UBTranRec(1).TransType = TranUtilityBill THEN
      SELECT CASE UBTranRec(1).TransDate
      CASE BadDate1
        UBTranRec(1).TransDate = GoodDate2
        DoIt = True
      CASE ELSE
      END SELECT
      IF DoIt THEN
        Fixed = Fixed + 1
        PUT UBTFile, Cnt&, UBTranRec(1)
        DoIt = False
      END IF
    END IF
  NEXT
  LOCATE 5, 1
  PRINT "Processing:"; TNumOfRecs&; "of"; TNumOfRecs&;

SkiptoNext:
  CLOSE

  LOCATE 9, 1
  PRINT "Fixed:"; Fixed
  CALL killfile("fixface.exe")

