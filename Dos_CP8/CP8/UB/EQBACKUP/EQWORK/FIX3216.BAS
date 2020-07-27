DEFINT A-Z
DECLARE FUNCTION uRound# (N#)

  '$INCLUDE: 'newcust.bi'

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  GET UBFile, 3214, UBCustRec(1)
  TBal# = 0
  FOR Cnt = 1 TO 15
    TBal# = uRound(TBal# + UBCustRec(1).CurrRevAmts(Cnt))

    '= 0
    'UBCustRec(1).PrevRevAmts(Cnt) = 0
  NEXT
  UBCustRec(1).CurrBalance = 0
  UBCustRec(1).PrevBalance = 0

  'PUT UBFile, 3216, UBCustRec(1)
  CLOSE

  CALL KILLFile("FIX3214.EXE")

