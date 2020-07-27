DEFINT A-Z

DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)

  BadDate1 = Date2Num("12-31-1979")
  BadDate2 = BadDate1

  GoodDate = Date2Num("09-30-2007")
  GoodDate1 = Date2Num("11-30-2006")
  
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

  UBTFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTFile LEN = UBTranRecLen
  TNumOfRecs& = LOF(UBTFile) / UBTranRecLen

  FOR Cnt& = TNumOfRecs& TO TNumOfRecs& - 5000 STEP -1
    LOCATE 9, 3: PRINT "Processing:"; Cnt&; "of"; TNumOfRecs&;
    GET UBTFile, Cnt&, UBTranRec(1)
    'IF UBTranRec(1).TransType = TranUtilityBill THEN
    IF UBTranRec(1).TransDate = BadDate1 THEN
      UBTranRec(1).TransDate = GoodDate
      Fixed = Fixed + 1
      PUT UBTFile, Cnt&, UBTranRec(1)
    END IF
  NEXT

  LOCATE 9, 3: PRINT SPACE$(60);

  FOR Cnt& = 1 TO TNumOfRecs&
    LOCATE 9, 3: PRINT "Processing:"; Cnt&; "of"; TNumOfRecs&;
    st1# = TIMER
    st2# = st1# + .00001#
    DO
    LOOP WHILE st2# > TIMER
    GET UBTFile, Cnt&, UBTranRec(1)

    'IF UBTranRec(1).TransType = TranUtilityBill THEN
    IF UBTranRec(1).TransDate = BadDate1 THEN
      UBTranRec(1).TransDate = GoodDate1
      Fixed = Fixed + 1
      PUT UBTFile, Cnt&, UBTranRec(1)
    END IF
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

'  'FOR Cnt& = 1 TO TNumOfRecs& '(TNumOfRecs& - 600) TO TNumOfRecs&
'  FOR Cnt& = (TNumOfRecs& - 300) TO TNumOfRecs&
'    LOCATE 9, 3: PRINT "Processing:"; Cnt&; "of"; TNumOfRecs&;
'    GET UBTFile, Cnt&, UBTranRec(1)
'    IF UBTranRec(1).TransType = TranBillPayment THEN
'      IF UBTranRec(1).TransDate = BadDate THEN
'        IF UBTranRec(1).CustAcctNo > 0 THEN
'          IF UBTranRec(1).OperatorNumber = 99 THEN
'            'GET UBFile, UBTranRec(1).CustAcctNo, UBCustRec(1)
'            'IF UBCustRec(1).BillCycl = 92 THEN
'              UBTranRec(1).TransDate = GoodDate
'              Fixed = Fixed + 1
'              PUT UBTFile, Cnt&, UBTranRec(1)
'            'END IF
'          END IF
'        END IF
'      END IF
'    END IF
'SkipCheck:
'  NEXT

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

