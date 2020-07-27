DEFINT A-Z


  ''$INCLUDE: 'ubEPP.BI'
  '$INCLUDE: 'newcust.bi'

  'REDIM UBEPPRec(1) AS UBCustEPPRecType
  'UBEPPRecLen = LEN(UBEPPRec(1))
  DIM Cnt AS INTEGER
  DIM CustCnt AS LONG
  DIM tCnt AS LONG
  DIM TotAmt AS DOUBLE
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  CLS
  PRINT
  PRINT "Correcting. . . "
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

  CustCnt = LOF(UBFile) / UBCustRecLen
  FOR tCnt = 1 TO CustCnt
    IF UBCustRec(1).DelFlag <> 0 THEN
      GOTO Skip2Here
    END IF

    GET UBFile, tCnt, UBCustRec(1)
    UBCustRec(1).CurrRevAmts(13) = 0
    UBCustRec(1).PrevRevAmts(13) = 0
    PUT UBFile, tCnt, UBCustRec(1)
    IF UBCustRec(1).PrevRevAmts(14) <> 0 THEN
      PRINT tCnt; USING "########.##"; UBCustRec(1).PrevRevAmts(14)
    END IF
Skip2Here:
  NEXT

  GET UBFile, 7, UBCustRec(1)
  FOR Cnt = 1 TO 15
    UBCustRec(1).CurrRevAmts(Cnt) = 0
    UBCustRec(1).PrevRevAmts(Cnt) = 0
  NEXT
  UBCustRec(1).DepositAmt = 0
  PUT UBFile, 7, UBCustRec(1)

  GET UBFile, 3564, UBCustRec(1)
  FOR Cnt = 1 TO 15
    UBCustRec(1).CurrRevAmts(Cnt) = 0
    UBCustRec(1).PrevRevAmts(Cnt) = 0
  NEXT
  UBCustRec(1).DepositAmt = 0
  PUT UBFile, 3564, UBCustRec(1)

  GET UBFile, 3483, UBCustRec(1)
  FOR Cnt = 1 TO 15
    UBCustRec(1).CurrRevAmts(Cnt) = 0
    UBCustRec(1).PrevRevAmts(Cnt) = 0
  NEXT
  UBCustRec(1).DepositAmt = 0
  PUT UBFile, 3483, UBCustRec(1)

  GET UBFile, 2997, UBCustRec(1)
  FOR Cnt = 1 TO 15
    UBCustRec(1).CurrRevAmts(Cnt) = 0
    UBCustRec(1).PrevRevAmts(Cnt) = 0
  NEXT
  UBCustRec(1).DepositAmt = 0
  PUT UBFile, 2997, UBCustRec(1)

  GET UBFile, 3163, UBCustRec(1)
  FOR Cnt = 1 TO 15
    UBCustRec(1).CurrRevAmts(Cnt) = 0
    UBCustRec(1).PrevRevAmts(Cnt) = 0
  NEXT
  UBCustRec(1).DepositAmt = 0
  PUT UBFile, 3163, UBCustRec(1)

  GET UBFile, 3102, UBCustRec(1)
  UBCustRec(1).CurrRevAmts(7) = 0
  UBCustRec(1).PrevRevAmts(7) = 0
  UBCustRec(1).CurrRevAmts(14) = 0

  UBCustRec(1).DepositAmt = 0
  FOR Cnt = 1 TO 15
    TotAmt = (TotAmt + UBCustRec(1).CurrRevAmts(Cnt))
  NEXT
  UBCustRec(1).CurrBalance = TotAmt
  PUT UBFile, 3102, UBCustRec(1)

  TotAmt = 0
  GET UBFile, 3161, UBCustRec(1)
  UBCustRec(1).CurrRevAmts(4) = 2.61
  FOR Cnt = 1 TO 15
    TotAmt = (TotAmt + UBCustRec(1).CurrRevAmts(Cnt))
  NEXT
  UBCustRec(1).CurrBalance = TotAmt
  UBCustRec(1).PrevBalance = 0
  PUT UBFile, 3161, UBCustRec(1)

  GET UBFile, 2901, UBCustRec(1)
  FOR Cnt = 1 TO 15
    UBCustRec(1).CurrRevAmts(Cnt) = 0
    UBCustRec(1).PrevRevAmts(Cnt) = 0
  NEXT
  UBCustRec(1).DepositAmt = 0
  UBCustRec(1).CurrBalance = 0
  UBCustRec(1).PrevBalance = 0
  PUT UBFile, 2901, UBCustRec(1)

  GET UBFile, 3158, UBCustRec(1)
  FOR Cnt = 1 TO 15
    UBCustRec(1).CurrRevAmts(Cnt) = 0
    UBCustRec(1).PrevRevAmts(Cnt) = 0
  NEXT
  UBCustRec(1).DepositAmt = 0
  UBCustRec(1).CurrBalance = 0
  UBCustRec(1).PrevBalance = 0
  PUT UBFile, 3158, UBCustRec(1)

  GET UBFile, 2997, UBCustRec(1)
  UBCustRec(1).PrevBalance = 0
  PUT UBFile, 2997, UBCustRec(1)

  GET UBFile, 3163, UBCustRec(1)
  UBCustRec(1).PrevBalance = 0
  UBCustRec(1).CurrBalance = 0
  PUT UBFile, 3163, UBCustRec(1)

  GET UBFile, 3564, UBCustRec(1)
  UBCustRec(1).PrevBalance = 0
  UBCustRec(1).CurrBalance = 0
  PUT UBFile, 3564, UBCustRec(1)

  GET UBFile, 5728, UBCustRec(1)
  UBCustRec(1).PrevBalance = 0
  UBCustRec(1).CurrBalance = 0
  PUT UBFile, 5728, UBCustRec(1)

  CLOSE

  PRINT
  PRINT
  PRINT "Done"
  CALL KillFile("RUN_1ST.exe")

SUB commented
'  TBal# = 0
'  FOR Cnt = 1 TO 15
'    UBCustRec(1).CurrRevAmts(Cnt) = 0
'    UBCustRec(1).PrevRevAmts(Cnt) = 0
'  NEXT
'
'  UBCustRec(1).CurrRevAmts(1) = 16.43
'  UBCustRec(1).CurrRevAmts(4) = 1.64
'  UBCustRec(1).CurrRevAmts(8) = 15
'  UBCustRec(1).CurrBalance = 33.07
'  UBCustRec(1).PrevBalance = 0
'
'  PUT UBFile, 2881, UBCustRec(1)
'  CLOSE
'
'  UBFile = FREEFILE
'  OPEN "UBCUSTEP.DAT" FOR RANDOM SHARED AS UBFile LEN = UBEPPRecLen
'  GET UBFile, 2881, UBEPPRec(1)
'  FOR Cnt = 1 TO 15
'    UBEPPRec(1).EPInfo(Cnt).RevBal = 0
'    UBEPPRec(1).EPInfo(Cnt).RevDis = 0
'  NEXT
'
'  UBEPPRec(1).EPInfo(3).RevBal = 0
'  UBEPPRec(1).EPBalTot = 0
'  UBEPPRec(1).EPDisTot = 0
'
'  PUT UBFile, 2881, UBEPPRec(1)
'  CLOSE
'
'  PRINT
'  PRINT
'  PRINT
'  PRINT "Done."
'  CALL KILLFile("FIX2881.EXE")

'    FOR zCnt = 1 TO 14
'      IF UBCustRec(1).CurrRevAmts(zCnt) > 0 THEN
'        bGotCur = 1
'      END IF
'      IF UBCustRec(1).PrevRevAmts(zCnt) > 0 THEN
'        bGotPre = 1
'      END IF
'    NEXT
'    IF bGotCur = 1 AND bGotPre = 1 THEN
'      rCnt% = rCnt% + 1
'      PRINT tCnt 'UBCustRec(1).CurrRevAmts(zCnt); UBCustRec(1).PrevRevAmts(zCnt)
'    END IF
'    PRINT UBCustRec(1).CurrRevAmts(zCnt); UBCustRec(1).PrevRevAmts(zCnt)
'
END SUB

