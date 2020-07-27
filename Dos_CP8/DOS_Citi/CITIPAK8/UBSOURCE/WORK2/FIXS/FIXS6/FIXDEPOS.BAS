DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)
  
  '$INCLUDE: 'ubtrans.BI'
  ''INCLUDE: 'newcust.BI'

  TDate = Date2Num%("08-20-1997")
  NDate = Date2Num%("08-20-1998")

  CLS
  LOCATE 3, 1, 0
  PRINT "Correcting Billing Date. . ."

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  UBFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBFile LEN = UBTranRecLen
  'OPEN "TRANDATE.LST" FOR OUTPUT AS #10
  TNumOfRecs& = LOF(UBFile) / UBTranRecLen

  FOR Cnt& = 1 TO TNumOfRecs&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; " of"; TNumOfRecs&;
    GET UBFile, Cnt&, UBTranRec(1)
    IF UBTranRec(1).TransType < 100 THEN STOP
    'IF UBTranRec(1).TransType = 106 OR UBTranRec(1).TransType = 7 THEN
      TTrans# = uRound#(TTrans# + UBTranRec(1).TransAmt)
      TTCnt = TTCnt + 1
      'UBTranRec(1).TransDate = TDate
      'PUT UBFile, Cnt&, UBTranRec(1)
    'END IF
  NEXT

  CLOSE

LOCATE 9, 1
PRINT "  Fixed:"; TTCnt; "  Total: "; USING "#####.##"; TTrans#

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

