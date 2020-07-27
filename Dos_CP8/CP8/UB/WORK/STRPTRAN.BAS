DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION uRound# (DblNum#)
CONST False = 0, True = NOT False

 ' '$INCLUDE: 'ubtrans.BI'
 ' '$INCLUDE: 'newcust.BI'

  '$INCLUDE: 'citiub.BI'
  TranDate = Date2Num("12-26-2007")
  REDIM RecArray(1 TO 1) AS LONG
  CLS
  LOCATE 3, 1, 0
  PRINT "Removing Transactions. . ."
  DIM testit AS testsizetype
  PRINT LEN(testit)
  END
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  UBFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBFile LEN = UBTranRecLen

  NUBFile = FREEFILE
  OPEN "NUBTRANS.DAT" FOR RANDOM SHARED AS NUBFile LEN = UBTranRecLen
  
  CUBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CUBFile LEN = UBCustRecLen

  TNumOfRecs& = LOF(UBFile) / UBTranRecLen
  FirstTime = True
  FOR cnt& = 1 TO TNumOfRecs&
    LOCATE 5, 1
    PRINT "Processing:"; cnt&; " of"; TNumOfRecs&;
    GET UBFile, cnt&, UBTranRec(1)
    IF UBTranRec(1).TransType = 4 THEN
      SELECT CASE UBTranRec(1).TransDate
      CASE TranDate
        GET CUBFile, UBTranRec(1).CustAcctNo, UBCustRec(1)

        'IF UBTranRec(1).CustAcctNo = 2318 THEN STOP

        'IF FirstTime THEN
        '  FirstTime = False
        '  ArrayCnt = 1
        '  REDIM PRESERVE RecArray(1 TO ArrayCnt) AS LONG
        '  RecArray(ArrayCnt) = UBTranRec(1).CustAcctNo
          FOR RCnt = 1 TO 15
            UBCustRec(1).CurrRevAmts(RCnt) = uRound#(UBCustRec(1).CurrRevAmts(RCnt) + UBTranRec(1).RevAmt(RCnt))
            UBCustRec(1).CurrBalance = uRound#(UBCustRec(1).CurrBalance + UBTranRec(1).RevAmt(RCnt))
          NEXT
          PUT CUBFile, UBTranRec(1).CustAcctNo, UBCustRec(1)

        'ELSE
        '  FoundIt = False
        '  FOR cnt = 1 TO ArrayCnt
        '    IF RecArray(cnt) = UBTranRec(1).CustAcctNo THEN
        '      FoundIt = True  'it allready been done
        '      EXIT FOR
        '    END IF
        '  NEXT
        '  IF FoundIt = False THEN
        '    ArrayCnt = ArrayCnt + 1
        '    REDIM PRESERVE RecArray(1 TO ArrayCnt) AS LONG
        '    RecArray(ArrayCnt) = UBTranRec(1).CustAcctNo
        '  END IF
        'END IF
        'IF FoundIt = False THEN
        '  FOR RCnt = 1 TO 15
        '    UBCustRec(1).CurrRevAmts(RCnt) = uRound#(UBCustRec(1).CurrRevAmts(RCnt) - UBTranRec(1).RevAmt(RCnt))
        '    UBCustRec(1).CurrBalance = uRound#(UBCustRec(1).CurrBalance + UBTranRec(1).RevAmt(RCnt))
        '  NEXT
        '  DidCnt = DidCnt + 1
        '  PUT CUBFile, UBTranRec(1).CustAcctNo, UBCustRec(1)
        'END IF
        Removed = Removed + 1
      CASE ELSE
        PUT NUBFile, , UBTranRec(1)
      END SELECT
    ELSE
      PUT NUBFile, , UBTranRec(1)
    END IF
  NEXT
  CLOSE

LOCATE 9, 1
PRINT "Removed:"; Removed
CALL KillFile("UBTRANS.DAT")
NAME "NUBTRANS.DAT" AS "UBTRANS.DAT"

CALL KillFile("STRPTRAN.EXE")

