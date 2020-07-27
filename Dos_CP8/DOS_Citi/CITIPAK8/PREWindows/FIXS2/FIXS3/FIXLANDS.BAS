DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (N#)
DECLARE FUNCTION Date2Num (TheDate$)


  '$INCLUDE: 'newcust.BI'
  '$INCLUDE: 'ubtrans.BI'

  'BDate = Date2Num("02-24-1999")

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
  'NumOfTran = LOF(UBTran) \ UBTranRecLen

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfCust = LOF(UBFile) \ UBCustRecLen

  FOR Cnt = 1 TO NumOfCust
    GET #UBFile, Cnt, UBCustRec(1)
    LOCATE 7, 1
    PRINT "Processing:"; Cnt; " of"; NumOfCust;
    IF UBCustRec(1).Book = "36" OR UBCustRec(1).Book = "37" THEN
      IF UBCustRec(1).LastTrans > 0 THEN
        GET #UBTran, UBCustRec(1).LastTrans, UBTrans(1)
        IF UBTrans(1).TransType = 1 THEN 'AND UBTrans(1).TransDate = BDate THEN
          FOR zz = 1 TO 15
            UBCustRec(1).CurrRevAmts(zz) = uRound#(UBCustRec(1).CurrRevAmts(zz) - UBTrans(1).RevAmt(zz))
            UBTrans(1).RevAmt(zz) = 0
          NEXT
          UBTrans(1).TransAmt = 0
          UBTrans(1).CustAcctNo = 0
          PUT #UBTran, UBCustRec(1).LastTrans, UBTrans(1)
          UBCustRec(1).LastTrans = UBTrans(1).PrevTrans
          UBCustRec(1).CurrBalance = UBCustRec(1).PrevBalance
          UBCustRec(1).PrevBalance = 0
          TRCnt& = TRCnt& + 1
          UBCustRec(1).BillCycl = 98
          PUT #UBFile, Cnt, UBCustRec(1)
        END IF
      END IF
    END IF
  NEXT

  CLOSE


PRINT
PRINT
PRINT
PRINT "Found:"; TRCnt&
'KILL "FIXCELNA.EXE"

'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

