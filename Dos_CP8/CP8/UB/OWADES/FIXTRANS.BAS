DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)
  
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.BI'

  TDate = Date2Num%("05-15-1997")
  NDate = Date2Num%("05-15-1997")

  CLS
  LOCATE 3, 1, 0
  PRINT "Processing Transactions History. . ."

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  UBFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBFile LEN = UBTranRecLen

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  TNumOfRecs& = LOF(UBFile) / UBTranRecLen

  FOR Cnt& = 1 TO TNumOfRecs&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; " of"; TNumOfRecs&;
    GET UBFile, Cnt&, UBTranRec(1)
    IF (UBTranRec(1).TransType = TranPenaltyCharge) AND (UBTranRec(1).TransDate = TDate) THEN
      GET UBCust, UBTranRec(1).CustAcctNo, UBTranRec(1)
      IF uRound#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance) <= 1 THEN
        UBCustRec(1).CurrRevAmts(1) = uRound#(UBCustRec(1).CurrRevAmts(1) - 1)
        UBCustRec(1).CurrBalance = uRound#(UBCustRec(1).CurrBalance - 1)
        Fixed = Fixed + 1
        PUT UBCust, UBTranRec(1).CustAcctNo, UBCustRec(1)
                                'CustAcctNo
        UBTranRec(1).TransDate = -32767
        PUT UBFile, Cnt&, UBTranRec(1)
      END IF
    END IF
  NEXT

  CLOSE

LOCATE 9, 1
PRINT "  Fixed:"; Fixed

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

