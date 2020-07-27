DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)
  
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.BI'

  CONST False = 0, True = NOT False

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
    LOCATE 5, 1
    PRINT "Processing:"; CCnt&; " of"; CNumOfRecs&;
    DTrans = False
    RTrans = False
    DAmt# = 0

    DidOne = False

    ThisTrans& = UBCustRec(1).LastTrans

    DO WHILE ThisTrans& > 0
      GET UBTFile, ThisTrans&, UBTranRec(1)
      SELECT CASE UBTranRec(1).TransType
        CASE TranDepositPayment, TranDepositPayment + 100
          IF UBTranRec(1).TransAmt > 0 THEN
            DTrans = True
            DAmt# = UBTranRec(1).TransAmt
            Half# = UBTranRec(1).TransAmt \ 2
            UBTranRec(1).RevAmt(1) = Half#
            UBTranRec(1).RevAmt(2) = Half#
            PUT UBTFile, ThisTrans&, UBTranRec(1)
          ELSE
            UBTranRec(1).TransType = 0
            PUT UBTFile, ThisTrans&, UBTranRec(1)
          END IF
        CASE TranAppliedDeposit, TranAppliedDeposit + 100, TranRefundDeposit, TranRefundDeposit + 100
          RTrans = True
      END SELECT
    ThisTrans& = UBTranRec(1).PrevTrans
  LOOP

  IF DTrans = True AND RTrans = True THEN
    UBCustRec(1).DepositAmt = 0
    PUT UBCFile, CCnt&, UBCustRec(1)
    'Fixed = Fixed + 1
  ELSEIF DTrans = True AND RTrans = False THEN
    UBCustRec(1).DepositAmt = DAmt#
    TTransAmt# = uRound#(TTransAmt# + DAmt#)
    PUT UBCFile, CCnt&, UBCustRec(1)
    Fixed = Fixed + 1
  ELSE
    IF UBCustRec(1).DepositAmt > 0 THEN
      UBCustRec(1).DepositAmt = 0
      'LPRINT "Acct:"; CCnt&
      PUT UBCFile, CCnt&, UBCustRec(1)
    ELSE

    END IF

  END IF

SkiptoNext:
  NEXT

  CLOSE

LOCATE 9, 1
PRINT "Found:"; Fixed; "  Total: "; USING "#####.##"; TTransAmt#

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

