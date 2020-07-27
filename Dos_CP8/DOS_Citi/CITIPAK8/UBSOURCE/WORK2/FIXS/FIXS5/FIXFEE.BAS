DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION fRound# (N#)
DECLARE FUNCTION Date2Num% (TheDate$)

  '$INCLUDE: 'newcust.BI'
  '$INCLUDE: 'ubtrans.BI'

  TDate = Date2Num(DATE$)

  REDIM UBCustRec(1) AS NewUBCustRecType
  REDIM UBTranRec(1) AS UBTransRecType

  UBCustRecLen = LEN(UBCustRec(1))
  UBTranRecLen = LEN(UBTranRec(1))

  CLS
  PRINT
  PRINT
  PRINT "Processing. . ."
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  FOR Cnt& = 1 TO NumOfRecs&
    LOCATE , 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    GET UBFile, Cnt&, UBCustRec(1)
    TFee# = fRound#(UBCustRec(1).MFEE1 + UBCustRec(1).MFEE2)
    IF TFee# > 0 THEN
      GOSUB MakeTransAction
      Fixed = Fixed + 1
    'ELSEIF TFee# < 0 THEN
    '  LPRINT Cnt&
    END IF
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Corrected: "; Fixed

END

MakeTransAction:

  REDIM UBTranRec(1)  AS UBTransRecType
  UBTranRec(1).TransDate = TDate
  UBTranRec(1).TransType = TranDepositPayment
  UBTranRec(1).TransDesc = "DEPOSIT PAYMENT"
  UBTranRec(1).CustLocation = Cnt&
  UBTranRec(1).OperatorNumber = 10
  UBTranRec(1).CustAcctNo = Cnt&
  UBTranRec(1).CustStatus = UBCustRec(1).Status
  UBTranRec(1).TransAmt = TFee#
  UBTranRec(1).CheckAmount = 0
  UBTranRec(1).CashAmount = TFee#
  UBTranRec(1).PayTypeCode = 1

  UBTranRec(1).RevAmt(1) = TFee#
  
  UBCustRec(1).DepositAmt = UBTranRec(1).TransAmt
  UBTranRec(1).PrevTrans = UBCustRec(1).LastTrans

  THandle = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS #THandle LEN = UBTranRecLen
  NextTransRec& = (LOF(THandle) \ UBTranRecLen) + 1

  PUT #THandle, NextTransRec&, UBTranRec(1)
  UBCustRec(1).LastTrans = NextTransRec&
  UBCustRec(1).MFEE1 = 0
  UBCustRec(1).MFEE2 = 0
  PUT UBFile, Cnt&, UBCustRec(1)
  CLOSE THandle

RETURN

  '****************************************************************************
  'Rounds a double precision value to nearest hundreth
  '****************************************************************************
FUNCTION fRound# (N#)
  fRound# = INT(N# * 100 + .5) / 100
END FUNCTION

