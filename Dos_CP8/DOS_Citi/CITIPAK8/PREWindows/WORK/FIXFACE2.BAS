DEFINT A-Z
  '$INCLUDE: 'Newcust.BI'

  CLS
  LOCATE 3, 1, 0
  PRINT "Correcting. . ."
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  UBfile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBfile LEN = UBCustRecLen

  GET UBfile, 180, UBCustRec(1)
  FOR RevCnt = 1 TO 15
    UBCustRec(1).CurrRevAmts(RevCnt) = 0
  NEXT

  UBCustRec(1).CurrRevAmts(1) = -27.98
  UBCustRec(1).CurrBalance = -27.98
  UBCustRec(1).PrevBalance = 0
  PUT UBfile, 180, UBCustRec(1)

  GET UBfile, 588, UBCustRec(1)
  FOR RevCnt = 1 TO 15
    UBCustRec(1).CurrRevAmts(RevCnt) = 0
  NEXT

  UBCustRec(1).CurrRevAmts(1) = -175.74
  UBCustRec(1).CurrBalance = -175.74
  UBCustRec(1).PrevBalance = 0
  PUT UBfile, 588, UBCustRec(1)

  CLOSE
  PRINT
  PRINT
  PRINT "Done."
  CALL KillFile("fixface2.exe")

