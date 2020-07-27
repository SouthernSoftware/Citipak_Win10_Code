DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)

  '$INCLUDE: 'ubtrans.BI'
  
  '$INCLUDE: 'newcust.BI'

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
  CNumOfRecs& = LOF(UBCFile) \ UBCustRecLen
  
  GET UBCFile, 1134, UBCustRec(1)
  FOR Cnt = 2 TO 15
    UBCustRec(1).CurrRevAmts(Cnt) = 0
  NEXT
  UBCustRec(1).CurrRevAmts(1) = -45
  PUT UBCFile, 1134, UBCustRec(1)

  GET UBCFile, 1410, UBCustRec(1)
  FOR Cnt = 1 TO 15
    UBCustRec(1).CurrRevAmts(Cnt) = 0
  NEXT
  PUT UBCFile, 1410, UBCustRec(1)

  GET UBCFile, 2048, UBCustRec(1)
  FOR Cnt = 2 TO 15
    UBCustRec(1).CurrRevAmts(Cnt) = 0
  NEXT
  UBCustRec(1).CurrRevAmts(1) = -45
  PUT UBCFile, 2048, UBCustRec(1)

  GET UBCFile, 4708, UBCustRec(1)
  FOR Cnt = 2 TO 15
    UBCustRec(1).CurrRevAmts(Cnt) = 0
  NEXT
  UBCustRec(1).CurrRevAmts(1) = -80
  PUT UBCFile, 4708, UBCustRec(1)

  GET UBCFile, 3226, UBCustRec(1)
  UBCustRec(1).CurrBalance = 0
  UBCustRec(1).PrevBalance = 0
  FOR Cnt = 1 TO 15
    UBCustRec(1).CurrRevAmts(Cnt) = 0
  NEXT
  PUT UBCFile, 3226, UBCustRec(1)

  GET UBCFile, 5156, UBCustRec(1)
  FOR Cnt = 1 TO 15
    UBCustRec(1).CurrRevAmts(Cnt) = 0
  NEXT
  PUT UBCFile, 5156, UBCustRec(1)

  GET UBCFile, 5182, UBCustRec(1)
  FOR Cnt = 1 TO 15
    UBCustRec(1).CurrRevAmts(Cnt) = 0
  NEXT
  UBCustRec(1).CurrRevAmts(1) = 18.3
  UBCustRec(1).CurrRevAmts(2) = 19.3
  UBCustRec(1).CurrRevAmts(3) = 6.5
  UBCustRec(1).CurrRevAmts(4) = 1.99
  PUT UBCFile, 5182, UBCustRec(1)
  
  GET UBCFile, 5279, UBCustRec(1)
  FOR Cnt = 1 TO 15
    UBCustRec(1).CurrRevAmts(Cnt) = 0
  NEXT
  PUT UBCFile, 5279, UBCustRec(1)

  GET UBCFile, 6139, UBCustRec(1)
  FOR Cnt = 2 TO 15
    UBCustRec(1).CurrRevAmts(Cnt) = 0
  NEXT
  UBCustRec(1).CurrRevAmts(1) = -45
  PUT UBCFile, 6139, UBCustRec(1)

  GET UBCFile, 6173, UBCustRec(1)
  FOR Cnt = 2 TO 15
    UBCustRec(1).CurrRevAmts(Cnt) = 0
  NEXT
  UBCustRec(1).CurrRevAmts(1) = -45
  PUT UBCFile, 6173, UBCustRec(1)

  GET UBCFile, 6437, UBCustRec(1)
  FOR Cnt = 2 TO 15
    UBCustRec(1).CurrRevAmts(Cnt) = 0
  NEXT
  UBCustRec(1).CurrRevAmts(1) = -70
  UBCustRec(1).CurrBalance = -70
  PUT UBCFile, 6437, UBCustRec(1)

  GET UBCFile, 6457, UBCustRec(1)
  FOR Cnt = 2 TO 15
    UBCustRec(1).CurrRevAmts(Cnt) = 0
  NEXT
  UBCustRec(1).CurrRevAmts(1) = 0
  PUT UBCFile, 6457, UBCustRec(1)

  CLOSE

  'UBTFile = FREEFILE
  'OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTFile LEN = UBTranRecLen
  'GET UBTFile, 4328, UBTranRec(1)
  'SWAP UBTranRec(1).RevAmt(15), UBTranRec(1).RevAmt(1)
  'PUT UBTFile, 4328, UBTranRec(1)
  'CLOSE

  LOCATE 9, 1
  PRINT "Processing Complete."
  KILL "fixholly.exe"

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

