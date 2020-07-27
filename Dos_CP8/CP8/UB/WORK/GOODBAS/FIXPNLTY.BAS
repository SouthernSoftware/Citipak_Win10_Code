DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)

'  BadDate1 = Date2Num("12-31-1979")
  GoodDate1 = Date2Num("05-11-2004")

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
      IF UBTranRec(1).TransType = TranPenaltyCharge THEN
        Fixed = Fixed + 1
        GET UBFile, UBTranRec(1).CustAcctNo, UBCustRec(1)
        UBCustRec(1).CurrRevAmts(1) = uRound(UBCustRec(1).CurrRevAmts(1) - 5)
        UBCustRec(1).CurrRevAmts(3) = uRound(UBCustRec(1).CurrRevAmts(3) + 5)
        UBTranRec(1).RevAmt(1) = 0
        UBTranRec(1).RevAmt(3) = 5
        PUT UBTFile, Cnt&, UBTranRec(1)          'update trans
        PUT UBFile, UBTranRec(1).CustAcctNo, UBCustRec(1)  'ubdate cust
      END IF
    END IF
  NEXT

SkiptoNext:
  CLOSE

  LOCATE 9, 1
  PRINT "Fixed:"; Fixed
  CALL KillFile("fixpnlty.exe")

'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

