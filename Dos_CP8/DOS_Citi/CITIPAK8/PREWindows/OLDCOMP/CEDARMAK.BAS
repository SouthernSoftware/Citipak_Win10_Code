DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (N#)
DECLARE FUNCTION Date2Num (TheDate$)
DECLARE SUB KillFile (File$)

  
  '$INCLUDE: 'newcust.BI'
  '$INCLUDE: 'ubtrans.BI'

  Today = Date2Num(DATE$)

  CLS
  PRINT
  PRINT
  PRINT "Creating Balance Transactions. . ."
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  REDIM UBTranRec(1) AS UBTransRecType

  UBCustRecLen = LEN(UBCustRec(1))
  UBTranRecLen = LEN(UBTranRec(1))
  KillFile ("UBTRANS.DAT")

  PRINT
  PRINT

  UBTran = FREEFILE
  OPEN "UBTRANS.dat" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfCust = LOF(UBFile) \ UBCustRecLen
  FOR Cnt = 1 TO NumOfCust
    LOCATE 5, 1
    PRINT "Processing:"; Cnt; " of"; NumOfCust;
    GET UBFile, Cnt, UBCustRec(1)
    IF UBCustRec(1).DelFlag = 0 THEN
      TestBal# = uRound(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
      IF TestBal# <> 0 THEN
        GOSUB MakeTrans
      ELSE
        FOR zz = 1 TO 15
          UBCustRec(1).CurrRevAmts(zz) = 0
          UBCustRec(1).PrevRevAmts(zz) = 0
        NEXT
        UBCustRec(1).CurrBalance = 0
        UBCustRec(1).PrevBalance = 0
        UBCustRec(1).DepositAmt = 0
        UBCustRec(1).LastTrans = 0
        PUT UBFile, Cnt, UBCustRec(1)
      END IF
    END IF
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Done"

KILL "CEDARMAK.EXE"

END

MakeTrans:
 TranCnt = TranCnt + 1
 REDIM UBTranRec(1) AS UBTransRecType

 UBTranRec(1).TransDate = Today
 UBTranRec(1).TransType = TranBeginBalance
 UBTranRec(1).TransDesc = "BALANCE REBUILD"
 UBTranRec(1).TransAmt = TestBal#
 UBTranRec(1).RunBalance = TestBal#

 FOR zz = 1 TO 15
   UBTranRec(1).RevAmt(zz) = UBCustRec(1).CurrRevAmts(zz)
 NEXT
 UBTranRec(1).CustStatus = UBCustRec(1).Status
 UBTranRec(1).CustAcctNo = Cnt

 PUT UBTran, TranCnt, UBTranRec(1)

 UBCustRec(1).DepositAmt = 0
 UBCustRec(1).LastTrans = TranCnt
 PUT UBFile, Cnt, UBCustRec(1)



RETURN

'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

