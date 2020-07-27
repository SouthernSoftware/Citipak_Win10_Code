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
    UBCustRec(1).CurrRevAmts(Cnt) = 0
    UBCustRec(1).PrevRevAmts(Cnt) = 0
  NEXT

  UBCustRec(1).CurrRevAmts(1) = 10
  UBCustRec(1).CurrRevAmts(2) = 10
  UBCustRec(1).CurrRevAmts(3) = 42.06
  UBCustRec(1).CurrRevAmts(9) = 5
  UBCustRec(1).CurrRevAmts(6) = 4.21
  UBCustRec(1).CurrRevAmts(4) = 1
  UBCustRec(1).CurrRevAmts(5) = 1
  UBCustRec(1).CurrRevAmts(10) = .5
  UBCustRec(1).CurrRevAmts(7) = 15
  UBCustRec(1).CurrBalance = 88.77
  UBCustRec(1).PrevBalance = 0

  PUT UBFile, 3214, UBCustRec(1)
  CLOSE

  CALL KILLFile("FIX3214.EXE")

