DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)
  
  '$INCLUDE: 'ubtrans.BI'
  CLS


  NDate1 = Date2Num%("11-01-2000")
  NDate2 = Date2Num%("11-30-2000")

  LOCATE 3, 1, 0
  PRINT "Processing Transactions . . ."

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  UBFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBFile LEN = UBTranRecLen
  TNumOfRecs& = LOF(UBFile) / UBTranRecLen
  FOR Cnt& = 1 TO TNumOfRecs&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; " of"; TNumOfRecs&;
    GET UBFile, Cnt&, UBTranRec(1)
    SELECT CASE UBTranRec(1).TransDate
    CASE IS < NDate1
      TTCnt& = TTCnt& + 1
      UBTranRec(1).Posted2GL = "Y"
      PUT UBFile, Cnt&, UBTranRec(1)
    CASE ELSE
    END SELECT
  NEXT

  CLOSE

LOCATE 9, 1
PRINT "  Found:"; TTCnt&;

