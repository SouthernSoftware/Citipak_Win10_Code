DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)


  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.bi'

  CONST False = 0, True = NOT False

  REDIM UBTranRec(1 TO 2) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  RecNo& = 1541

  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  GET UBFile, RecNo&, UBCustRec(1)
  UBCustRec(1).CurrBalance = 0
  UBCustRec(1).PrevBalance = 0
  UBCustRec(1).LastTrans = 0
  FOR zz = 1 TO 15
    UBCustRec(1).CurrRevAmts(zz) = 0
    UBCustRec(1).PrevRevAmts(zz) = 0
  NEXT
  PUT UBFile, RecNo&, UBCustRec(1)

  CLOSE

  CLS
  LOCATE 3, 1, 0
  PRINT "Correcting Transaction. . ."

  PRINT
  PRINT

  Cnt& = 952167

  UBTFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTFile LEN = UBTranRecLen
  GET UBTFile, Cnt&, UBTranRec(1)
  UBTranRec(1).TransAmt = 0
  FOR zz = 1 TO 15
    UBTranRec(1).RevAmt(zz) = 0
  NEXT
  UBTranRec(1).CustAcctNo = 0

  PUT UBTFile, Cnt&, UBTranRec(1)

SkiptoNext:
  CLOSE

  LOCATE 9, 1
  PRINT "Done"
  CALL killfile("fixbtran.exe")

