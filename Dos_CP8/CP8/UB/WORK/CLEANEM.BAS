DEFINT A-Z

  '$INCLUDE: 'newcust.bi'

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  UBCustRec(1).DelFlag = -1

  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  PUT UBFile, 2510, UBCustRec(1)
  PUT UBFile, 2514, UBCustRec(1)
  PUT UBFile, 2515, UBCustRec(1)
  PUT UBFile, 2513, UBCustRec(1)
  PUT UBFile, 2517, UBCustRec(1)
  PUT UBFile, 2518, UBCustRec(1)
  CLOSE
  KILL "CLEANEM.EXE"


