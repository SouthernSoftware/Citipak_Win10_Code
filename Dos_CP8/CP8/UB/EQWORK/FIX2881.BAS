DEFINT A-Z
DECLARE FUNCTION uRound# (N#)

  '$INCLUDE: 'ubEPP.BI'
  '$INCLUDE: 'newcust.bi'

  REDIM UBEPPRec(1) AS UBCustEPPRecType
  UBEPPRecLen = LEN(UBEPPRec(1))

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  CLS
  PRINT
  PRINT "Correcting Acct: 2881. . ."
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  GET UBFile, 2881, UBCustRec(1)
  TBal# = 0
  FOR Cnt = 1 TO 15
    UBCustRec(1).CurrRevAmts(Cnt) = 0
    UBCustRec(1).PrevRevAmts(Cnt) = 0
  NEXT

  UBCustRec(1).CurrRevAmts(1) = 16.43
  UBCustRec(1).CurrRevAmts(4) = 1.64
  UBCustRec(1).CurrRevAmts(8) = 15
  UBCustRec(1).CurrBalance = 33.07
  UBCustRec(1).PrevBalance = 0

  PUT UBFile, 2881, UBCustRec(1)
  CLOSE

  UBFile = FREEFILE
  OPEN "UBCUSTEP.DAT" FOR RANDOM SHARED AS UBFile LEN = UBEPPRecLen
  GET UBFile, 2881, UBEPPRec(1)
  FOR Cnt = 1 TO 15
    UBEPPRec(1).EPInfo(Cnt).RevBal = 0
    UBEPPRec(1).EPInfo(Cnt).RevDis = 0
  NEXT

  UBEPPRec(1).EPInfo(3).RevBal = 0
  UBEPPRec(1).EPBalTot = 0
  UBEPPRec(1).EPDisTot = 0

  PUT UBFile, 2881, UBEPPRec(1)
  CLOSE

  PRINT
  PRINT
  PRINT
  PRINT "Done."
  CALL KILLFile("FIX2881.EXE")

