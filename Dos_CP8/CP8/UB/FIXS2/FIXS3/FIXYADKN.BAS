DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (N#)

  '$INCLUDE: 'newcust.BI'
  '$INCLUDE: 'ubtrans.BI'

  CLS
  PRINT
  PRINT
  PRINT "Removing Transactions. . ."
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  REDIM UBTrans(1)   AS UBTransRecType
  UBTranRecLen = LEN(UBTrans(1))

  UBTran = FREEFILE
  OPEN "ubtrans.dat" FOR RANDOM AS #UBTran LEN = UBTranRecLen
  NumOfTran = LOF(UBTran) \ UBTranRecLen

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

  FOR Cnt = 1 TO NumOfTran
    GET #UBTran, Cnt, UBTrans(1)
    IF UBTrans(1).TransType = 12 THEN
      GET UBFile, UBTrans(1).CustAcctNo, UBCustRec(1)
      Tst# = 0
      IF UBTrans(1).TransAmt < 0 THEN
        FOR zz = 1 TO 15
          UBCustRec(1).CurrRevAmts(zz) = uRound#(UBCustRec(1).CurrRevAmts(zz) + UBTrans(1).RevAmt(zz))
          Tst# = uRound#(Tst# + UBCustRec(1).CurrRevAmts(zz))
        NEXT
      ELSE
        FOR zz = 1 TO 15
          UBCustRec(1).CurrRevAmts(zz) = uRound#(UBCustRec(1).CurrRevAmts(zz) - UBTrans(1).RevAmt(zz))
          Tst# = uRound#(Tst# - UBCustRec(1).CurrRevAmts(zz))
        NEXT
      END IF

      UBCustRec(1).CurrBalance = Tst#  'uRound#(UBCustRec(1).CurrBalance + UBTrans(1).TransAmt)
      PUT UBFile, UBTrans(1).CustAcctNo, UBCustRec(1)
      UBTrans(1).CustAcctNo = -1
      PUT #UBTran, Cnt, UBTrans(1)
    END IF
  NEXT

  CLOSE


PRINT
PRINT
PRINT
PRINT "Done"
'KILL "FIXHOLLY.EXE"

'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

