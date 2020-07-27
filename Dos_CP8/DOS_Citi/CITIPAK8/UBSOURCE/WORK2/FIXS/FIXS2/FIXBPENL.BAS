DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION Round# (DblNum#)

  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'Newcust.BI'

  TranDate = Date2Num("07-02-1997")
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
    ThisTrans& = UBCustRec(1).LastTrans

    DO WHILE ThisTrans& > 0
      GET UBFile, ThisTrans&, UBTranRec(1)
      IF UBTranRec(1).TransType = TranUtilityBill THEN
        cnvrt = cnvrt + 1
        FOR MtrCnt = 1 TO 7
          UBCustRec(1).LocMeters(MtrCnt).CurRead = UBTranRec(1).CurRead(MtrCnt)
          UBCustRec(1).LocMeters(MtrCnt).PrevRead = UBTranRec(1).PrevRead(MtrCnt)
        NEXT
        PUT UBCust, Cnt&, UBCustRec(1)
        EXIT DO
      END IF
      ThisTrans& = UBTranRec(1).PrevTrans
    LOOP
  NEXT
  CLOSE

LOCATE 9, 1
PRINT "Fixed:"; cnvrt

