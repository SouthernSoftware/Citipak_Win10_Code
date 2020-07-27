DEFINT A-Z
DECLARE FUNCTION Date2Num% (TheDate$)

  '$INCLUDE: 'ubtrans.BI'
  CLS
  FromDate = Date2Num("03-09-2001")
  ThruDate = Date2Num("02-25-2001")

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  PRINT "   Change Transaction Dates"
  PRINT
  UBFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBFile LEN = UBTranRecLen
  TNumOfRecs& = LOF(UBFile) / UBTranRecLen
  FOR Cnt& = 1 TO TNumOfRecs&
    GET UBFile, Cnt&, UBTranRec(1)
    LOCATE 7, 1
    PRINT "Processing:"; Cnt&; "of"; TNumOfRecs&;
    IF UBTranRec(1).TransType = TranUtilityBill THEN
      IF UBTranRec(1).TransDate = FromDate THEN
        TTCnt = TTCnt + 1
        UBTranRec(1).TransDate = ToDate
        PUT UBFile, Cnt&, UBTranRec(1)
      END IF
    END IF
  NEXT
  CLOSE
  PRINT
  PRINT
  PRINT "   Changed: "; TTCnt
  CALL KillFile("fixjohns.exe")

