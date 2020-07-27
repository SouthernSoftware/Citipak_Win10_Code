DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)
  
  '$INCLUDE: 'ubtrans.BI'
  CLS

  Date1a = Date2Num%("12-31-1979")
  Date2a = Date2Num%("04-30-1999")
  Date1b = Date2Num%("02-01-1999")
  Date2b = Date2Num%("05-01-1999")

  LOCATE 3, 1, 0
  PRINT "Correcting Transaction Dates. . ."

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  UBFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBFile LEN = UBTranRecLen
  TNumOfRecs& = LOF(UBFile) / UBTranRecLen

  FOR Cnt& = 1 TO TNumOfRecs&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; " of"; TNumOfRecs&;
    GET UBFile, Cnt&, UBTranRec(1)
    IF UBTranRec(1).TransType = 1 THEN
      SELECT CASE UBTranRec(1).TransDate
      CASE Date1a
        UBTranRec(1).TransDate = Date1b
        PUT UBFile, Cnt&, UBTranRec(1)
        Fixed = Fixed + 1
      CASE Date2a
        UBTranRec(1).TransDate = Date2b
        PUT UBFile, Cnt&, UBTranRec(1)
        Fixed = Fixed + 1
      END SELECT
    END IF
  NEXT

  CLOSE

LOCATE 9, 1
PRINT "Fixed:"; Fixed

