DEFINT A-Z
DECLARE FUNCTION ruRound# (DblNum#)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)

  '$INCLUDE: 'newcust.BI'
  '$INCLUDE: 'ubtrans.BI'

  TDate = Date2Num%("10-24-1997")
  ZDate = Date2Num%("01-01-1980")

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))
  
  CLS
  PRINT
  PRINT
  PRINT "Processing. . ."
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen

  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen



'  FOR Cnt& = 1 TO NumOfRecs&
    'LOCATE , 1
    'PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    GET UBFile, 261, UBCustRec(1)

    GotCnt = 0
    First& = 0

      LastTran& = UBCustRec(1).LastTrans


      DO WHILE LastTran& > 0
        GET UBTran, LastTran&, UBTranRec(1)
        LastTran& = UBTranRec(1).PrevTrans
        GET UBTran, LastTran&, UBTranRec(1)
        UBTranRec(1).TransDate = ZDate
        'UBTranRec(1).CustAcctNo = 0
        PUT UBTran, LastTran&, UBTranRec(1)
        EXIT DO
      LOOP
    'IF GotCnt = 2 THEN
SkipEm:
  'NEXT
  CLOSE


    IF UBCustRec(1).PrevBalance < 0 THEN
      'STOP
    END IF
    IF UBCustRec(1).CurrBalance < 0 THEN
      'LPRINT LEFT$(UBCustRec(1).CustName, 25), UBCustRec(1).CurrBalance, UBCustRec(1).PrevBalance, UBCustRec(1).Status
      'STOP
    END IF
    'IF UBCustRec(1).Status = "B" THEN
      'IF UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance > 0 THEN
      '  Tot# = uRound(Tot# + UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
        'LPRINT Cnt&, UBCustRec(1).CustName, USING "####.##"; UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance
      'END IF
    'END IF
  'LPRINT
  'LPRINT "Total:"; USING "#####.##"; Tot#
'PRINT
'PRINT
'PRINT
'PRINT "Done."

  '****************************************************************************
  'Rounds a double precision value to nearest hundreth
  '****************************************************************************
FUNCTION ruRound# (N#)
  ruRound# = INT(N# * 100 + .5) / 100
END FUNCTION

