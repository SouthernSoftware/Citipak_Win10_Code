DEFINT A-Z

SUB Archive (ThruYear)
  
'  REDIM UBTranRec(1) AS UBTransRecType
'  UBTranRecLen = LEN(UBTranRec(1))
'  REDIM UBCustRec(1) AS NewUBCustRecType
'  UBCustRecLen = LEN(UBCustRec(1))
'
'  REDIM TranRecs(0 TO 1) AS ArcTranType
'
'  ShowWarning
'  PRINT "   Archive Transaction History thru Year:"; ThruYear
'  Ok$ = GetProceed$
'
'  SELECT CASE Ok$
'  CASE "Y"
'    PRINT Ok$
'    GOSUB GOArchiveHist
'    IF ThruYear <> -1 THEN
'      PRINT
'      PRINT
'      PRINT "   Transaction History Archived."
'    END IF
'  CASE ELSE
'    PRINT "N"
'    PRINT
'    PRINT "   ARCHIVE TRANSACTION ABORTED"
'  END SELECT
'EXIT SUB
'
'GOArchiveHist:
'  PRINT
'
'  FreeSpace# = DiskRoom&("")
'  IF FreeSpace# < (FileSize&("UBTRANS.DAT") * 2) THEN
'    PRINT "   INSUFFICIENT DISK FREE SPACE!!!"
'    PRINT
'    PRINT "   ARCHIVE TRANSACTION ABORTED"
'    ThruYear = -1
'    GOTO ArchiveExit
'  END IF
'
'  StartDate = Date2Num%("01-01-" + QPTrim$(STR$(ThruYear)))
'  EndDate = Date2Num%("12-31-" + QPTrim$(STR$(ThruYear)))
'
'  IF Exist("UBTrans.ARK") THEN
'    KillFile "UBTrans.ARK"
'  END IF
'  NAME "UBTrans.dat" AS "UBTrans.ARK"
'
'  UBTran = FREEFILE
'  OPEN "UBTrans.ARK" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen
'
'  NUBTran = FREEFILE
'  OPEN "UBTrans.DAT" FOR RANDOM SHARED AS NUBTran LEN = UBTranRecLen
'
'  UBFile = FREEFILE
'  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
'
'  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
'
'  FOR Cnt& = 1 TO NumOfRecs&
'    ShowProcessing Cnt&, NumOfRecs&
'    GET UBFile, Cnt&, UBCustRec(1)
'    TranCnt = 0
'    ThisTran& = UBCustRec(1).LastTrans
'    REDIM TranRecs(0 TO 1) AS ArcTranType
'    DO WHILE ThisTran& > 0
'      TranCnt = TranCnt + 1
'      REDIM PRESERVE TranRecs(0 TO TranCnt)  AS ArcTranType
'      GET UBTran, ThisTran&, UBTranRec(1)
'      TranRecs(TranCnt).RecNo = ThisTran&
'      IF UBTranRec(1).TransType > 100 THEN
'        TranRecs(TranCnt).TRType = UBTranRec(1).TransType - 100
'      ELSE
'        TranRecs(TranCnt).TRType = UBTranRec(1).TransType
'      END IF
'      TranRecs(TranCnt).TRDate = UBTranRec(1).TransDate
'      TranRecs(TranCnt).DidFlag = False
'      ThisTran& = UBTranRec(1).PrevTrans
'    LOOP
'    'first make sure they have transactions greater than this date range
'
'    CanDo = False
'    FOR TCnt = 1 TO TranCnt
'      IF TranRecs(TCnt).TRDate > EndDate THEN
'        CanDo = True
'        EXIT FOR
'      END IF
'    NEXT
'    IF CanDo THEN
'      FOR TCnt = 1 TO TranCnt
'        IF TranRecs(TCnt).TRDate <= EndDate THEN
'          SELECT CASE TranRecs(TCnt).TRType
'          'keep these transaction anyway
'          CASE TranAppliedDeposit, TranDepositPayment, TranRefundDeposit
'            GET UBTran, TranRecs(TCnt).RecNo, UBTranRec(1)
'            PUT NUBTran, , UBTranRec(1)
'          CASE ELSE
'            Arked& = Arked& + 1
'          END SELECT
'        ELSE
'          GET UBTran, TranRecs(TCnt).RecNo, UBTranRec(1)
'          PUT NUBTran, , UBTranRec(1)
'        END IF
'      NEXT
'    ELSE 'no trans greater then date; copy all of them
'      FOR TCnt = 1 TO TranCnt
'        GET UBTran, TranRecs(TCnt).RecNo, UBTranRec(1)
'        PUT NUBTran, , UBTranRec(1)
'      NEXT
'    END IF
'    'PUT UBFile, Cnt&, UBCustRec(1)
'  NEXT
'  CLOSE
'  PRINT
'  PRINT
'  PRINT "   Removed:"; Arked&; " Tranactions"
'ArchiveExit:
'
'RETURN

END SUB

