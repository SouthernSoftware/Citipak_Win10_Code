DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)
  
  '$INCLUDE: 'ubtrans.BI'
  ''INCLUDE: 'newcust.BI'

  TDate = Date2Num%("06-01-1997")
  NDate = Date2Num%("12-31-1979")

  CLS
  LOCATE 3, 1, 0
  PRINT "Correcting Billing Date. . ."

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  UBFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBFile LEN = UBTranRecLen
  TNumOfRecs& = LOF(UBFile) / UBTranRecLen

  UBRpt = FREEFILE
  OPEN "TRANS.RPT" FOR OUTPUT AS UBRpt

  FOR Cnt& = 1 TO TNumOfRecs&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; " of"; TNumOfRecs&;
    GET UBFile, Cnt&, UBTranRec(1)
    PRINT #UBRpt, Num2Date$(UBTranRec(1).TransDate); TAB(11); USING "#####"; Cnt&;
    PRINT #UBRpt, TAB(30); LEFT$(UBTranRec(1).TransDesc, 20)
  NEXT

  CLOSE

LOCATE 9, 1
'PRINT "  Fixed:"; Fixed;

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

