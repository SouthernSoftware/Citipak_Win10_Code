DECLARE SUB CheckTranFile ()

DECLARE SUB CheckCustFile ()
DECLARE SUB CheckRateTables ()
DECLARE SUB CheckSetUpFile ()
DECLARE FUNCTION FileSize& (FileName$)
DECLARE SUB KillFile (FileName$)

DEFINT A-Z

'$INCLUDE: 'ubsetup.bi'
'$INCLUDE: 'ubrate.bi'
'$INCLUDE: 'ubcust.bi'
'$INCLUDE: 'ubloc.bi'
'$INCLUDE: 'ubtrans.bi'

'$INCLUDE: 'oubtypes.bi'

  CLS
  PRINT "UBUTIL1 v8.1   Copyright 1996-97  Southern Software Inc."
  PRINT
  'CheckSetUpFile
  CheckRateTables
  CheckCustFile
  CheckTranFile

SUB CheckCustFile

  REDIM UBCustRec(1 TO 2) AS UBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  REDIM oUBCustRec(1) AS oUBCustRecType
  oUBCustRecLen = LEN(oUBCustRec(1))

  IF FileSize&("UBCUST.DAT") MOD UBCustRecLen THEN
    PRINT "Converting Customer File . . . ";
    Col = POS(1)
    Row = CSRLIN

    oUBFile = FREEFILE
    OPEN "UBCUST.DAT" FOR RANDOM SHARED AS oUBFile LEN = oUBCustRecLen
    nUBFile = FREEFILE
    OPEN "NUBCUST.DAT" FOR RANDOM SHARED AS nUBFile LEN = UBCustRecLen

    NumOfRecs = LOF(oUBFile) \ oUBCustRecLen

    FOR Cnt = 1 TO NumOfRecs
      LOCATE Row, Col: PRINT Cnt;
      LSET UBCustRec(1) = UBCustRec(2)
      GET oUBFile, Cnt, oUBCustRec(1)
      UBCustRec(1).STATUS = oUBCustRec(1).STATUS
      UBCustRec(1).OPENDATE = oUBCustRec(1).OPENDATE
      UBCustRec(1).SEARCH = oUBCustRec(1).SEARCH
      UBCustRec(1).CUSTNAME = oUBCustRec(1).CUSTNAME
      UBCustRec(1).ADDR1 = oUBCustRec(1).ADDR1
      UBCustRec(1).ADDR2 = oUBCustRec(1).ADDR2
      UBCustRec(1).CITY = oUBCustRec(1).CITY
      UBCustRec(1).STATE = oUBCustRec(1).STATE
      UBCustRec(1).ZIPCODE = oUBCustRec(1).ZIPCODE
      UBCustRec(1).CUSTTYPE = oUBCustRec(1).CUSTTYPE
      UBCustRec(1).PostRte = oUBCustRec(1).PostRte
      UBCustRec(1).BILLCOPY = oUBCustRec(1).BILLCOPY
      UBCustRec(1).HPHONE = oUBCustRec(1).HPHONE
      UBCustRec(1).WPHONE = oUBCustRec(1).WPHONE
      UBCustRec(1).SOSEC = oUBCustRec(1).SOSEC
      UBCustRec(1).DRVLIC = oUBCustRec(1).DRVLIC
      UBCustRec(1).TAXEXPT = oUBCustRec(1).TAXEXPT
      UBCustRec(1).SRCIT = oUBCustRec(1).SRCIT
      UBCustRec(1).EPP = oUBCustRec(1).EPP
      UBCustRec(1).EPPAMT = oUBCustRec(1).EPPAMT
      UBCustRec(1).CASHONLY = oUBCustRec(1).CASHONLY
      UBCustRec(1).LATEFEE = oUBCustRec(1).LATEFEE
      UBCustRec(1).CUTOFFYN = oUBCustRec(1).CUTOFFYN

      UBCustRec(1).Monthly(1).AMTOWED = oUBCustRec(1).AMTOWED
      UBCustRec(1).Monthly(1).TotAmtPD = oUBCustRec(1).AMTPD
      UBCustRec(1).Monthly(1).PayAmt = oUBCustRec(1).PAYMENT
      UBCustRec(1).Monthly(1).RevSource = oUBCustRec(1).REVSRCE

      UBCustRec(1).DRAFTYPE = oUBCustRec(1).DRAFTYPE
      UBCustRec(1).BANKNAME = oUBCustRec(1).BANKNAME
      UBCustRec(1).BANKLOC = oUBCustRec(1).BANKLOC
      UBCustRec(1).TRANSIT = oUBCustRec(1).TRANSIT
      UBCustRec(1).BANKACCT = oUBCustRec(1).BANKACCT
      UBCustRec(1).BILLCMNT = oUBCustRec(1).BILLCMNT
      UBCustRec(1).PAYCMNT = oUBCustRec(1).PAYCMNT
      UBCustRec(1).CustPIN = oUBCustRec(1).CustPIN
      UBCustRec(1).OwnFrstLocRec = oUBCustRec(1).OwnFrstLocRec
      UBCustRec(1).OwnLastLocRec = oUBCustRec(1).OwnLastLocRec
      UBCustRec(1).CustFrstLocRec = oUBCustRec(1).CustFrstLocRec
      UBCustRec(1).CustLastLocRec = oUBCustRec(1).CustLastLocRec
      UBCustRec(1).Pend2Rec = oUBCustRec(1).Pend2Rec
      UBCustRec(1).LastTrans = oUBCustRec(1).LastTrans
      UBCustRec(1).CurrBalance = oUBCustRec(1).CurrBalance
      UBCustRec(1).PrevBalance = oUBCustRec(1).PrevBalance
      FOR RevCnt = 1 TO 15
        UBCustRec(1).CurrRevAmts(RevCnt) = oUBCustRec(1).CurrRevAmts(RevCnt)
      NEXT
      UBCustRec(1).DepositAmt = oUBCustRec(1).DepositAmt
      UBCustRec(1).DelFlag = oUBCustRec(1).DelFlag
      UBCustRec(1).PreNoteFlag = oUBCustRec(1).PreNoteFlag
      UBCustRec(1).PrevLocRec = oUBCustRec(1).PrevLocRec
      UBCustRec(1).FillPad = oUBCustRec(1).FillPad

      PUT nUBFile, Cnt, UBCustRec(1)
    NEXT
    CLOSE
    KillFile "UBCUST.DAT"
    NAME "NUBCUST.DAT" AS "UBCUST.DAT"
    LOCATE Row, Col:
    PRINT " OK.   "
  ELSE
    PRINT "Current Customer file Structure is correct."
  END IF

END SUB

SUB CheckRateTables

  REDIM UBRateTblRec(1) AS UBRateTblRecType
  UBRateTblRecLen = LEN(UBRateTblRec(1))
  REDIM oUBRateTblRec(1) AS oUBRateTblRecType
  oUBRateTblRecLen = LEN(oUBRateTblRec(1))

  IF FileSize&("UBRATE.DAT") MOD UBRateTblRecLen THEN
    PRINT "Converting Rate Tables . . .";
    oUBFile = FREEFILE
    OPEN "UBRATE.DAT" FOR RANDOM SHARED AS oUBFile LEN = oUBRateTblRecLen
    nUBFile = FREEFILE
    OPEN "NUBRATE.DAT" FOR RANDOM SHARED AS nUBFile LEN = UBRateTblRecLen

    NumOfRates = LOF(oUBFile) \ oUBRateTblRecLen

    FOR Cnt = 1 TO NumOfRates
      GET oUBFile, Cnt, oUBRateTblRec(1)
      UBRateTblRec(1).RateCode = oUBRateTblRec(1).RateCode
      UBRateTblRec(1).RATEDESC = oUBRateTblRec(1).RATEDESC
      UBRateTblRec(1).MINAMT = oUBRateTblRec(1).MINAMT
      UBRateTblRec(1).MINUNITS = oUBRateTblRec(1).MINUNITS
      UBRateTblRec(1).DiscPct = 0
      FOR Brk = 1 TO 10
        LSET UBRateTblRec(1).TblBreaks(Brk) = oUBRateTblRec(1).TblBreaks(Brk)
      NEXT
      PUT nUBFile, Cnt, UBRateTblRec(1)
    NEXT
    CLOSE
    KillFile "UBRATE.DAT"
    NAME "NUBRATE.DAT" AS "UBRATE.DAT"
    PRINT " OK."
  ELSE
    PRINT "Current Rate Table Structure is correct."
  END IF
  

getoutofhere:

END SUB

SUB CheckSetUpFile

'REDIM oUBSetUp(1) AS oUBSetupRecType
'REDIM UBSetUp(1) AS UBSetupRecType
'
' OSetUpLen = LEN(oUBSetUp(1))
' NSetUpLen = LEN(UBSetUp(1))
'
' OPEN "UBSETUP.DAT" FOR RANDOM AS #1 LEN = NSetUpLen
' IF LOF(1) <> NSetUpLen THEN
'   GET #1, 1, oUBSetUp(1)
'   UBSetUp(1).UTILNAME = oUBSetUp(1).UTILNAME
'   UBSetUp(1).DEFCITY = oUBSetUp(1).DEFCITY
'   UBSetUp(1).DEFSTATE = oUBSetUp(1).DEFSTATE
'   UBSetUp(1).ZIPCODE = oUBSetUp(1).ZIPCODE
'   UBSetUp(1).FLYDeposit = oUBSetUp(1).FLYDeposit
'   UBSetUp(1).EQPAY = oUBSetUp(1).EQPAY
'   UBSetUp(1).WORKORD = oUBSetUp(1).WORKORD
'   UBSetUp(1).ESTREAD = oUBSetUp(1).ESTREAD
'   UBSetUp(1).BANKDFT = oUBSetUp(1).BANKDFT
'   UBSetUp(1).UseSeq = oUBSetUp(1).UseSeq
'   UBSetUp(1).BILLCYCL = oUBSetUp(1).BILLCYCL
'   UBSetUp(1).BARCODE = oUBSetUp(1).BARCODE
 '  UBSetUp(1).RemLookFlag = "Y"
'   UBSetUp(1).LowRead = oUBSetUp(1).LowRead
'   UBSetUp(1).HighRead = oUBSetUp(1).HighRead
'   UBSetUp(1).HHDEVICE = oUBSetUp(1).HHDEVICE
'   FOR Cnt = 1 TO 15
'     LSET UBSetUp(1).Revenues(Cnt) = oUBSetUp(1).Revenues(Cnt)
'   NEXT
'   PUT #1, 1, UBSetUp(1)
' END IF
' CLOSE
'
END SUB

SUB CheckTranFile

  REDIM NTrans(1 TO 2) AS UBTransRecType
  NTransLen = LEN(NTrans(1))

  REDIM OTrans(1) AS oUBTransRecType
  OTransLen = LEN(OTrans(1))

  IF FileSize&("UBTRANS.DAT") MOD NTransLen THEN
    PRINT "Converting Transacton File . . . ";
    Col = POS(1)
    Row = CSRLIN

    oUBFile = FREEFILE
    OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS oUBFile LEN = OTransLen
    nUBFile = FREEFILE
    OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS nUBFile LEN = NTransLen

    NumOfRecs& = LOF(oUBFile) \ OTransLen

    FOR TCnt& = NumOfRecs& TO 1 STEP -1
      GET oUBFile, TCnt&, OTrans(1)
      LSET NTrans(1) = NTrans(2)
      LOCATE Row, Col: PRINT TCnt&;
      NTrans(1).TransDate = OTrans(1).TransDate
      NTrans(1).TransType = OTrans(1).TransType
      NTrans(1).TransDesc = OTrans(1).TransDesc
      NTrans(1).TransAmt = OTrans(1).TransAmt
      FOR SCnt = 1 TO 15
        NTrans(1).RevAmt(SCnt) = OTrans(1).RevAmt(SCnt)
        NTrans(1).TaxAmt(SCnt) = OTrans(1).TaxAmt(SCnt)
      NEXT
      FOR MTCnt = 1 TO 7
        NTrans(1).CurRead(MTCnt) = OTrans(1).CurRead(MTCnt)
        NTrans(1).PrevRead(MTCnt) = OTrans(1).PrevRead(MTCnt)
        NTrans(1).EstRead(MTCnt) = OTrans(1).EstRead(MTCnt)
      NEXT

      NTrans(1).BillNumber = OTrans(1).BillNumber
      NTrans(1).ReadDate = OTrans(1).ReadDate
      NTrans(1).BillDate = OTrans(1).BillDate
      NTrans(1).PastDueDate = OTrans(1).PastDueDate
      NTrans(1).DraftDate = OTrans(1).DraftDate
      NTrans(1).CustLocation = OTrans(1).CustLocation
      NTrans(1).CustStatus = OTrans(1).CustStatus
      NTrans(1).CustCode = OTrans(1).CustCode
      NTrans(1).PostRoute = OTrans(1).PostRoute
      NTrans(1).BillCycle = OTrans(1).BillCycle
      NTrans(1).PayTypeCode = OTrans(1).PayTypeCode
      NTrans(1).OperatorNumber = OTrans(1).OperatorNumber
      NTrans(1).CustAcctNo = OTrans(1).CustAcctNo
      NTrans(1).PrevTrans = OTrans(1).PrevTrans
      NTrans(1).NextTrans = OTrans(1).NextTrans
      NTrans(1).ActiveFlag = OTrans(1).ActiveFlag
      NTrans(1).RunBalance = OTrans(1).RunBalance
      NTrans(1).CheckAmount = OTrans(1).CheckAmount
      NTrans(1).CashAmount = OTrans(1).CashAmount
      NTrans(1).BillMsg = OTrans(1).BillMsg
      NTrans(1).ApplyDepFlag = OTrans(1).ApplyDepFlag
      NTrans(1).Posted2GL = OTrans(1).Posted2GL
      NTrans(1).PrevDate = OTrans(1).PrevDate

      PUT nUBFile, TCnt&, NTrans(1)
    NEXT
    CLOSE
    'KillFile "UBTRANS.DAT"
    'NAME "NUBTRANS.DAT" AS "UBTRANS.DAT"
    LOCATE Row, Col:
    PRINT " OK.   "
  ELSE
    PRINT "Current Transaction file Structure is correct."
  END IF



'    NEXT
'    CLOSE


END SUB

