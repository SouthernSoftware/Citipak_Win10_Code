DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)
  
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.BI'

  CONST False = 0, True = NOT False

  BadDate = Date2Num("12-31-1979")
  TRDate1 = Date2Num("06-10-1999")
  TRDate2 = Date2Num("01-13-1999")

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  CLS
  LOCATE 3, 1, 0
  PRINT "Correcting Deposit Information. . ."

  PRINT
  PRINT

  UBCFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBCFile LEN = UBCustRecLen
  CNumOfRecs& = LOF(UBCFile) \ UBCustRecLen
  
  UBTFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTFile LEN = UBTranRecLen
  TNumOfRecs& = LOF(UBTFile) / UBTranRecLen

  FOR CCnt& = 1 TO CNumOfRecs&
    GET UBCFile, CCnt&, UBCustRec(1)
    BadCnt = 0
    LOCATE 5, 1
    PRINT "Processing:"; CCnt&; " of"; CNumOfRecs&;

    ThisTrans& = UBCustRec(1).LastTrans
    DO WHILE ThisTrans& > 0
      GET UBTFile, ThisTrans&, UBTranRec(1)
      IF UBTranRec(1).TransDate = BadDate THEN
        BadCnt = BadCnt + 1
        IF BadCnt = 1 THEN
          UBTranRec(1).TransDate = TRDate1
        ELSE
          UBTranRec(1).TransDate = TRDate2
        END IF
        PUT UBTFile, ThisTrans&, UBTranRec(1)
      END IF

      ThisTrans& = UBTranRec(1).PrevTrans
    LOOP
    IF BadCnt = 1 THEN
      LPRINT CCnt&
      'STOP
    END IF
  NEXT

SkiptoNext:
  CLOSE

LOCATE 9, 1
PRINT "Processing Complete."

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

