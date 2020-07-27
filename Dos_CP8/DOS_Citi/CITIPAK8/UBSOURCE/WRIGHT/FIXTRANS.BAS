DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)
  
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.BI'

  TDate = Date2Num%("10-31-1998")
  NDate = Date2Num%("10-31-1998")

  CLS
  LOCATE 3, 1, 0
  PRINT "Correcting Billing Date. . ."

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen

  UBFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBFile LEN = UBTranRecLen
  TNumOfRecs& = LOF(UBFile) / UBTranRecLen

  FOR Cnt& = 1 TO TNumOfRecs&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; " of"; TNumOfRecs&;
    GET UBFile, Cnt&, UBTranRec(1)
    IF (UBTranRec(1).TransType = TranUtilityBill) AND (UBTranRec(1).TransDate = NDate) THEN
      'FOR ZZ = 1 TO 7
      '  CurRead& = UBTranRec(1).CurRead(ZZ)
      '  PrevRead& = UBTranRec(1).PrevRead(ZZ)
      '  IF CurRead& - PrevRead& = 50 THEN
          GET #UBCust, UBTranRec(1).CustAcctNo, UBCustRec(1)
          'IF UBCustRec(1).PrevBalance <> 0 THEN
           'STOP
           'LPRINT UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb
            TTCnt = TTCnt + 1
            LOCATE 12, 1
            PRINT "                       ";
            LOCATE 12, 1
            PRINT "Last:"; UBTranRec(1).CustAcctNo;
          'END IF
      '  END IF
      'NEXT

      'PUT UBFile, Cnt&, UBTranRec(1)
    END IF
  NEXT

  CLOSE

LOCATE 9, 1
PRINT "  Found:"; TTCnt;

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

