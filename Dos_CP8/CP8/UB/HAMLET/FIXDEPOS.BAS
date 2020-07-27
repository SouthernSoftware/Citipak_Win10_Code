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
  WIDTH LPRINT 132
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
        CASE TranAppliedDeposit, TranAppliedDeposit + 100
        TRev# = 0
        FOR RevCnt = 1 TO 15
          TRev# = uRound(TRev# + UBTranRec(1).RevAmt(RevCnt))
        NEXT
        IF ABS(UBTranRec(1).TransAmt) <> TRev# THEN
          LPRINT CCnt&,
          Found = Found + 1
          FOR RevCnt = 1 TO 15
            UBTranRec(1).RevAmt(RevCnt) = 0
          NEXT
          UBTranRec(1).RevAmt(1) = ABS(UBTranRec(1).TransAmt)
          PUT UBTFile, ThisTrans&, UBTranRec(1)
        END IF
      END SELECT
    ThisTrans& = UBTranRec(1).PrevTrans
  LOOP
  NEXT

SkiptoNext:
  CLOSE

LOCATE 9, 1
PRINT "Found:"; Found

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

