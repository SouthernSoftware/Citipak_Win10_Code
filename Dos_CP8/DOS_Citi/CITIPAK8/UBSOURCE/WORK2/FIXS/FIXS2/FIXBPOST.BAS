DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION Round# (DblNum#)

  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'Newcust.BI'

  TranDate = Date2Num("06-30-1997")
  'PastDate = Date2Num("11-15-1996")

  CLS
  LOCATE 3, 1, 0
  'PRINT "Converting Customer Deposits. . ."
  REDIM UBTranRec(1) AS UBTransRecType
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBTranRecLen = LEN(UBTranRec(1))
  UBCustRecLen = LEN(UBCustRec(1))

  UBFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBFile LEN = UBTranRecLen
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen

  NumOfRecs& = LOF(UBCust) / UBCustRecLen

  FOR Cnt& = 1 TO NumOfRecs&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    GET UBCust, Cnt&, UBCustRec(1)
    'IF Cnt& = 170 THEN STOP
    ThisTrans& = UBCustRec(1).LastTrans
    IF ThisTrans& > 0 THEN
      GET UBFile, ThisTrans&, UBTranRec(1)
      IF UBTranRec(1).TransDate = TranDate THEN
        FOR RCnt = 1 TO 15
          UBCustRec(1).CurrRevAmts(RCnt) = Round#(UBCustRec(1).CurrRevAmts(RCnt) - UBTranRec(1).RevAmt(RCnt))
        NEXT
        TCur# = 0
        FOR RCnt = 1 TO 15
          TCur# = Round#(TCur# + UBCustRec(1).CurrRevAmts(RCnt))
        NEXT
        UBCustRec(1).CurrBalance = TCur#
        UBCustRec(1).PrevBalance = Round#(UBCustRec(1).PrevBalance - UBCustRec(1).CurrBalance)
        cnvrt = cnvrt + 1
        PUT UBCust, Cnt&, UBCustRec(1)
      END IF
    END IF
  NEXT
  CLOSE

LOCATE 9, 1
PRINT "Fixed:"; cnvrt

