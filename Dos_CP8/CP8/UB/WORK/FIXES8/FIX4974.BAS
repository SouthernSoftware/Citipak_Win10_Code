DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)
  
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.BI'
  ''$INCLUDE: 'fieldinf.BI'

CONST BlankInt% = -32767                        'Specifies blank integer
CONST BlankLng& = -2147483647                   '       "        long int.

  CONST False = 0, True = NOT False
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))


  CLS
  LOCATE 3, 1, 0
  PRINT "Correcting Customer Information. . ."

  PRINT
  PRINT

  UBCFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBCFile LEN = UBCustRecLen

  UBTFile = FREEFILE
  OPEN "UBTRANS.dat" FOR RANDOM SHARED AS UBTFile LEN = UBTranRecLen

  CNumOfRecs& = LOF(UBCFile) \ UBCustRecLen
  
  GET UBCFile, 74, UBCustRec(1)
  IF UBCustRec(1).LastTrans > 0 THEN
    GET UBTFile, UBCustRec(1).LastTrans, UBTranRec(1)
  END IF

  FOR cnt = 1 TO 15
    UBCustRec(1).CurrRevAmts(cnt) = 0
  NEXT
  UBCustRec(1).CurrBalance = 0
  UBCustRec(1).PrevBalance = 0

  IF UBCustRec(1).LastTrans > 0 THEN
    UBTranRec(1).RunBalance = 0
    PUT UBTFile, UBCustRec(1).LastTrans, UBTranRec(1)
  END IF
  PUT UBCFile, 74, UBCustRec(1)


  GET UBCFile, 49, UBCustRec(1)

  IF UBCustRec(1).LastTrans > 0 THEN
    GET UBTFile, UBCustRec(1).LastTrans, UBTranRec(1)
  END IF

  FOR cnt = 1 TO 15
    UBCustRec(1).CurrRevAmts(cnt) = 0
  NEXT
  UBCustRec(1).CurrBalance = 0
  UBCustRec(1).PrevBalance = 0

  IF UBCustRec(1).LastTrans > 0 THEN
    UBTranRec(1).RunBalance = 0
    PUT UBTFile, UBCustRec(1).LastTrans, UBTranRec(1)
  END IF
  
  PUT UBCFile, 49, UBCustRec(1)

SkiptoNext:
  CLOSE

LOCATE 9, 1
PRINT "Processing Complete."
KILL "FIX4974.EXE"

