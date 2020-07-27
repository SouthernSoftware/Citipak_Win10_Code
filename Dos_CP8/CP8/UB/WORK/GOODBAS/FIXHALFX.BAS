DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)

'  BadDate1 = Date2Num("12-31-1979")
  GoodDate1 = Date2Num("01-03-2006")

  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.bi'

  CONST False = 0, True = NOT False

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))     'Length of Cust Record Structure

  REDIM UBTranRec(1 TO 2) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  CLS
  LOCATE 3, 1, 0
  PRINT "Correcting Transaction. . ."

  PRINT
  PRINT
  UBTFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTFile LEN = UBTranRecLen

  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

'  OPEN "TRList.txt" FOR OUTPUT AS 10

  TNumOfRecs& = LOF(UBTFile) / UBTranRecLen
  
  FOR Cnt& = 1 TO TNumOfRecs&
    LOCATE 5, 1: PRINT "Processing:"; Cnt&; "of"; TNumOfRecs&;
    GET UBTFile, Cnt&, UBTranRec(1)
    IF UBTranRec(1).TransDate = GoodDate1 THEN
      IF UBTranRec(1).TransType = TranBillPayment THEN
        Fixed = Fixed + 1
        GET UBFile, UBTranRec(1).CustAcctNo, UBCustRec(1)
        FOR RevCnt = 1 TO 15
          UBCustRec(1).CurrRevAmts(RevCnt) = 0
        NEXT

        SELECT CASE UBTranRec(1).CustAcctNo
        CASE 1198
          UBCustRec(1).CurrRevAmts(1) = 0
          UBCustRec(1).CurrRevAmts(2) = -391.37
          UBTranRec(1).RunBalance = UBCustRec(1).CurrRevAmts(2)
          UBCustRec(1).CurrBalance = UBTranRec(1).RunBalance
          UBCustRec(1).PrevBalance = 0
        CASE ELSE
          UBCustRec(1).CurrBalance = 0
          UBCustRec(1).PrevBalance = 0
          UBTranRec(1).RunBalance = 0
        END SELECT
  
        PUT UBTFile, Cnt&, UBTranRec(1)          'update trans
        PUT UBFile, UBTranRec(1).CustAcctNo, UBCustRec(1)  'ubdate cust
      END IF
    END IF
  NEXT

SkiptoNext:
  CLOSE

  LOCATE 9, 1
  PRINT "Fixed:"; Fixed
  CALL KillFile("fixhalfx.exe")

        'FOR RevCnt = 1 TO 15
        '  UBCustRec(1).CurrRevAmts(RevCnt) = 0
        'NEXT
        'UBCustRec(1).CurrRevAmts(1) = uRound(UBCustRec(1).CurrRevAmts(1) - 5)
        'UBCustRec(1).CurrRevAmts(3) = uRound(UBCustRec(1).CurrRevAmts(3) + 5)
        'UBTranRec(1).RevAmt(1) = 0
        'UBTranRec(1).RevAmt(3) = 5

'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

