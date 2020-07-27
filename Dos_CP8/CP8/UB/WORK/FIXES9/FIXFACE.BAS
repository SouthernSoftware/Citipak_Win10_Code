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

  REDIM UBTranRec(1 TO 2) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  Today = Date2Num(DATE$)

  CLS
  LOCATE 3, 1, 0
  PRINT "Correcting Deposit Information. . ."
  PRINT
  PRINT

  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen
  TNumOfRecs& = LOF(UBTran) / UBTranRecLen

  UBCFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBCFile LEN = UBCustRecLen
  CNumOfRecs& = LOF(UBCFile) \ UBCustRecLen
  FOR Cnt& = 1 TO CNumOfRecs&
    LOCATE 3, 1
    PRINT "Processing:"; Cnt&; "of"; CNumOfRecs&;
    GET #UBCFile, Cnt&, UBCustRec(1)
    IF UBCustRec(1).Status = "A" THEN
      IF UBCustRec(1).DepositAmt > 0 THEN
        NoDepTran = True
        ThisTrans& = UBCustRec(1).LastTrans
        FirstTrans = True
        DO WHILE ThisTrans& > 0
          GET #UBTran, ThisTrans&, UBTranRec(1)
          IF UBTranRec(1).TransType = TranDepositPayment THEN
            IF UBCustRec(1).DepositAmt <> UBTranRec(1).TransAmt THEN
              STOP
            END IF
            NoDepTran = False
            EXIT DO
          END IF
          ThisTrans& = UBTranRec(1).PrevTrans
        LOOP
        IF NoDepTran THEN
          GOSUB MakeDepTrans
          NoTRCnt = NoTRCnt + 1
        END IF
      END IF
    END IF
  NEXT

SkiptoNext:
  CLOSE

'CALL KILLFILE("FIXFACE.EXE")

LOCATE 9, 1
PRINT "Done."
PRINT "Found:"; NoTRCnt
END

MakeDepTrans:
  GET #UBTran, UBCustRec(1).LastTrans, UBTranRec(2)

  REDIM UBTrans(1) AS UBTransRecType
  UBTrans(1).TransDate = Today
  UBTrans(1).TransType = TranDepositPayment
  UBTrans(1).TransDesc = "DEP PAYMENT (CNVRT ERR)"
  UBTrans(1).TransAmt = UBCustRec(1).DepositAmt
  UBTrans(1).RevAmt(1) = UBTrans(1).TransAmt
  UBTrans(1).CustStatus = "A"
  UBTrans(1).PayTypeCode = 0
  UBTrans(1).OperatorNumber = 99
  UBTrans(1).CustAcctNo = Cnt&
  UBTrans(1).PrevTrans = 0
  UBTrans(1).RunBalance = UBTranRec(2).RunBalance
  UBTrans(1).CashAmount = UBTrans(1).TransAmt
  UBTrans(1).BillMsg = "CONVERSION CORRECTION"

  NextRec& = ((LOF(UBTran) / UBTranRecLen) + 1)

  PUT #UBTran, NextRec&, UBTrans(1)


RETURN

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

