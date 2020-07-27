DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION uRound# (DblNum#)
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'Newcust.BI'

  WIDTH LPRINT 78
  TranDate = Date2Num(DATE$)
  'PastDate = Date2Num("11-15-1996")

  CLS
  LOCATE 3, 1, 0
  PRINT "Checking Customer Balances. . ."
  REDIM UBTranRec(1) AS UBTransRecType
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBTranRecLen = LEN(UBTranRec(1))
  UBCustRecLen = LEN(UBCustRec(1))


  'NAME "UBTRANS.DAT" AS "OUBTRANS.DAT"

  UBFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBFile LEN = UBTranRecLen

'  NUBFile = FREEFILE
'  OPEN "nUBTRANS.DAT" FOR RANDOM SHARED AS NUBFile LEN = UBTranRecLen

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen

  NumOfRecs& = LOF(UBCust) / UBCustRecLen

  FOR Cnt& = 1 TO NumOfRecs&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    GET UBCust, Cnt&, UBCustRec(1)
    'REDIM UBTranRec(1) AS UBTransRecType
    IF UBCustRec(1).DelFlag = 0 THEN
      TBal# = uRound#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
      RevBal# = 0

      FOR RevCnt = 1 TO 14
        RevBal# = uRound#(RevBal# + UBCustRec(1).CurrRevAmts(RevCnt))
      NEXT

      IF RevBal# <> TBal# THEN
        TCnt = TCnt + 1
        LOCATE 12, 1
        PRINT Cnt&
        SLEEP
        'UBCustRec(1).CurrBalance = RevBal#
        'UBCustRec(1).PrevBalance = 0
        'PUT UBCust, Cnt&, UBCustRec(1)
        'STOP
'        LPRINT Cnt&, QPTrim$(UBCustRec(1).CustName); TAB(40); RevBal#; TAB(50); TBal#; TAB(60); UBCustRec(1).Status
        'LPRINT Cnt&, ;
      END IF
    END IF
  NEXT
  CLOSE


LOCATE 9, 1
PRINT "Finished:"; TCnt
'LPRINT CHR$(12)
END

    IF RevBal# <> TBal# THEN
      'STOP
      LPRINT Cnt&, QPTrim$(UBCustRec(1).CustName); TAB(40); RevBal#; TAB(50); TBal#; TAB(60); UBCustRec(1).Status
    END IF

'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

