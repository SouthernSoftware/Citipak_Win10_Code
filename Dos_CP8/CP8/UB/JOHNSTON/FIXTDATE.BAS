DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)

  BadDate = Date2Num("12-31-1979")
  GoodDate = Date2Num("09-15-2003")
  
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.BI'

  CONST False = 0, True = NOT False

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBTranRec(1 TO 2) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  CLS
  PRINT
  PRINT "  Searching Transactions. . ."

  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  
  UBTFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTFile LEN = UBTranRecLen
  TNumOfRecs& = LOF(UBTFile) / UBTranRecLen

  'FOR Cnt& = 1 TO TNumOfRecs& '(TNumOfRecs& - 600) TO TNumOfRecs&
  FOR Cnt& = (TNumOfRecs& - 7500) TO TNumOfRecs&
    LOCATE 9, 3: PRINT "Processing:"; Cnt&; "of"; TNumOfRecs&;
    GET UBTFile, Cnt&, UBTranRec(1)
'    IF UBTranRec(1).TransType = TranBillPayment THEN
      IF UBTranRec(1).TransDate <= 0 THEN '= BadDate THEN
        IF UBTranRec(1).CustAcctNo > 0 THEN
        STOP

'          IF UBTranRec(1).OperatorNumber = 99 THEN
            'GET UBFile, UBTranRec(1).CustAcctNo, UBCustRec(1)
            'IF UBCustRec(1).BillCycl = 92 THEN
'              UBTranRec(1).TransDate = GoodDate
              Fixed = Fixed + 1
'              PUT UBTFile, Cnt&, UBTranRec(1)
            'END IF
          END IF
        END IF
'      END IF
'    END IF
SkipCheck:
  NEXT

SkiptoNext:
  CLOSE

LOCATE 13, 1
PRINT "Fixed:"; Fixed
CALL killfile("fixtdate.exe")

'    st1! = TIMER
'    st2! = st1! + .08
'    DO
'    LOOP WHILE st2! > TIMER
'    IF Cnt& < 114500 THEN
'      GOTO SkipCheck
'    END IF

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

