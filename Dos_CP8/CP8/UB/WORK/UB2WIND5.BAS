DEFINT A-Z
DECLARE SUB KILLFILE (filename$)
CONST true = -1, False = NOT true

'$INCLUDE: 'ub2wind5.bi'
'$INCLUDE: 'wadecust.bi'

  DIM UBCustVI AS NewUBCustRecTypeVI
  DIM UBCustVII AS NewUBCustRecTypeVII
  DIM oUBCust AS NewUBCustRecTypeVIII
  DIM NUBCust AS NewUBCustRecType

  VIRecLen = LEN(UBCustVI)
  VIIRecLen = LEN(UBCustVII)
  oreclen = LEN(oUBCust)
  nRecLen = LEN(NUBCust)

  KILLFILE "ubcust.old"
  GOSUB chkcheckbyte

  IF AbortFlag THEN
    GOTO Holy.Cow.Batman
  END IF

  CALL KILLFILE("ubcust.old")

  NAME "ubcust.dat" AS "ubcust.old"

  CLS
  IF Type2Flag THEN
    PRINT "Converting Customer File Type: 2 -> 4. . . "
  ELSEIF Type4Flag THEN
    PRINT "Converting Customer File Type: 3 -> 4. . . "
  ELSE
    PRINT "Converting Customer File Type: 1 -> 4. . . "
  END IF

  oUBCustFile = FREEFILE
  IF Type2Flag THEN
    OPEN "ubcust.old" FOR RANDOM AS oUBCustFile LEN = VIIRecLen
    NumORecs& = LOF(oUBCustFile) / VIIRecLen
  ELSEIF Type4Flag THEN
    OPEN "ubcust.old" FOR RANDOM AS oUBCustFile LEN = oreclen
    NumORecs& = LOF(oUBCustFile) / oreclen
  ELSE
    OPEN "ubcust.old" FOR RANDOM AS oUBCustFile LEN = VIRecLen
    NumORecs& = LOF(oUBCustFile) / VIRecLen
  END IF

  nUBCustFile = FREEFILE
  OPEN "ubcust.dat" FOR RANDOM SHARED AS nUBCustFile LEN = nRecLen

  FOR cnt& = 1 TO NumORecs&
    LOCATE 5, 1
    PRINT "Processing:"; cnt&; "Of"; NumORecs&;
    IF Type2Flag THEN
      GET #oUBCustFile, cnt&, UBCustVII
    ELSEIF Type4Flag THEN
      GET #oUBCustFile, cnt&, oUBCust
    ELSE
      GET #oUBCustFile, cnt&, UBCustVI
    END IF
    IF Type2Flag THEN
      GOSUB FromVerII
    ELSEIF Type4Flag THEN
      GOSUB FromVerIII
    ELSE
      GOSUB FromVerI
    END IF
    PUT #nUBCustFile, cnt&, NUBCust
  NEXT

  CLOSE
  LOCATE 9, 1
  PRINT "Processing Complete."

Holy.Cow.Batman:
  KILLFILE "ub2wind5.exe"

  END

chkcheckbyte:

  '*********************
  UBCustFile = FREEFILE
  OPEN "UBCust.dat" FOR RANDOM AS UBCustFile LEN = nRecLen
  GET #UBCustFile, 1, NUBCust
  IF NUBCust.ChkByte = CHR$(5) THEN
    GET #UBCustFile, 2, NUBCust
    IF NUBCust.ChkByte = CHR$(5) THEN
      'file is allready converted
      AbortFlag = true
    END IF
  END IF
  CLOSE UBCustFile
  IF AbortFlag THEN
    GOTO CheckByteReturn
  END IF
 '*************************

  UBCustFile = FREEFILE
  OPEN "UBCust.dat" FOR RANDOM AS UBCustFile LEN = oreclen
  GET #UBCustFile, 1, oUBCust
  IF oUBCust.ChkByte = CHR$(4) THEN
    GET #UBCustFile, 2, oUBCust
    IF oUBCust.ChkByte = CHR$(4) THEN
      'file is allready converted
      Type4Flag = true
     ' AbortFlag = true
    END IF
  END IF
  CLOSE UBCustFile
  IF Type4Flag THEN
    GOTO CheckByteReturn
  END IF

  'if we get to here we know the cust file isn't current
  'now see if cust file is type 2

  UBCustFile = FREEFILE        '
  OPEN "UBCust.dat" FOR RANDOM AS UBCustFile LEN = VIIRecLen
  GET #UBCustFile, 1, UBCustVII
  IF UBCustVII.ChkByte = CHR$(3) THEN
    GET #UBCustFile, 2, UBCustVII
    IF UBCustVII.ChkByte = CHR$(3) THEN
      Type2Flag = true   'cust file IS type 2
    END IF
  END IF
  CLOSE UBCustFile
  IF Type2Flag THEN
    GOTO CheckByteReturn
  END IF
  'if we get here then the cust file IS type 1
  'so AbortFlag is false and Type2Flag is false

CheckByteReturn:

RETURN

FromVerI: 'from dos
  NUBCust.BOOK = UBCustVI.BOOK
  NUBCust.SEQNUMB = UBCustVI.SEQNUMB
  NUBCust.Status = UBCustVI.Status
  NUBCust.OPENDATE = UBCustVI.OPENDATE
  NUBCust.SEARCH = UBCustVI.SEARCH
  NUBCust.CUSTNAME = UBCustVI.CUSTNAME
  NUBCust.ADDR1 = UBCustVI.ADDR1
  NUBCust.ADDR2 = UBCustVI.ADDR2
  NUBCust.SERVADDR = UBCustVI.SERVADDR
  NUBCust.CITY = UBCustVI.CITY
  NUBCust.STATE = UBCustVI.STATE
  NUBCust.ZIPCODE = UBCustVI.ZIPCODE
  NUBCust.HPHONE = UBCustVI.HPHONE
  NUBCust.WPHONE = UBCustVI.WPHONE
  NUBCust.SOSEC = UBCustVI.SOSEC
  NUBCust.DRVLIC = UBCustVI.DRVLIC
  NUBCust.CUSTTYPE = UBCustVI.CUSTTYPE
  NUBCust.Addr911 = UBCustVI.Addr911
  NUBCust.BillTo = UBCustVI.BillTo
  NUBCust.BILLCOPY = UBCustVI.BILLCOPY
  IF NUBCust.BILLCOPY <= 0 OR NUBCust.BILLCOPY > 99 THEN
    NUBCust.BILLCOPY = 1
  END IF
  NUBCust.POSTRTE = UBCustVI.POSTRTE
  NUBCust.BILLCYCL = UBCustVI.BILLCYCL
  NUBCust.ZONE = UBCustVI.ZONE
  NUBCust.SEQ = UBCustVI.SEQ
  NUBCust.CASHONLY = UBCustVI.CASHONLY
  NUBCust.LATEFEE = UBCustVI.LATEFEE
  NUBCust.CUTOFFYN = UBCustVI.CUTOFFYN
  NUBCust.TAXEXPT = UBCustVI.TAXEXPT
  NUBCust.SRCIT = UBCustVI.SRCIT
  NUBCust.EPPFlag = UBCustVI.EPPFlag
  NUBCust.GroupCodeRec = 0
  NUBCust.Filler1 = ""
  NUBCust.USEDRAFT = UBCustVI.USEDRAFT
  NUBCust.AcctType = UBCustVI.AcctType
  NUBCust.BANKNAME = UBCustVI.BANKNAME
  NUBCust.BANKLOC = UBCustVI.BANKLOC
  NUBCust.TRANSIT = UBCustVI.TRANSIT
  NUBCust.BANKACCT = UBCustVI.BANKACCT
  NUBCust.BILLCMNT = UBCustVI.BILLCMNT
  NUBCust.PAYCMNT = UBCustVI.PAYCMNT
  NUBCust.PUMPCODE = UBCustVI.PUMPCODE
  NUBCust.USERCODE1 = UBCustVI.USERCODE1
  NUBCust.USERCODE2 = UBCustVI.USERCODE2
  NUBCust.ProRatePCT = UBCustVI.ProRatePCT
  NUBCust.HHMSG1 = UBCustVI.HHMSG1
  NUBCust.HHMSG2 = UBCustVI.HHMSG2
  NUBCust.HHMSG3 = UBCustVI.HHMSG3
  FOR cnt = 1 TO 15
    NUBCust.Serv(cnt).RateCode = UBCustVI.Serv(cnt).RateCode
    NUBCust.Serv(cnt).RMtrType = UBCustVI.Serv(cnt).RMtrType
  NEXT
  FOR cnt = 1 TO 4
    NUBCust.FlatRates(cnt).FRDESC = UBCustVI.FlatRates(cnt).FRDESC
    NUBCust.FlatRates(cnt).FRAMT = UBCustVI.FlatRates(cnt).FRAMT
    NUBCust.FlatRates(cnt).FRFREQ = UBCustVI.FlatRates(cnt).FRFREQ
    NUBCust.FlatRates(cnt).REVSRC = UBCustVI.FlatRates(cnt).REVSRC
    NUBCust.FlatRates(cnt).NumMin = UBCustVI.FlatRates(cnt).NumMin
  NEXT
  FOR cnt = 1 TO 2
    NUBCust.Monthly(cnt).AMTOWED = UBCustVI.Monthly(cnt).AMTOWED
    NUBCust.Monthly(cnt).TotAmtPD = UBCustVI.Monthly(cnt).TotAmtPD
    NUBCust.Monthly(cnt).PayAmt = UBCustVI.Monthly(cnt).PayAmt
    NUBCust.Monthly(cnt).RevSource = UBCustVI.Monthly(cnt).RevSource
  NEXT
  NUBCust.MFEE1 = UBCustVI.MFEE1
  NUBCust.MFEE2 = UBCustVI.MFEE2
  FOR cnt = 1 TO 7
    NUBCust.LocMeters(cnt).MTRNUM = UBCustVI.LocMeters(cnt).MTRNUM
    NUBCust.LocMeters(cnt).MTRMulti = UBCustVI.LocMeters(cnt).MTRMulti
    NUBCust.LocMeters(cnt).MTRType = UBCustVI.LocMeters(cnt).MTRType
    NUBCust.LocMeters(cnt).MTRUnit = UBCustVI.LocMeters(cnt).MTRUnit
    NUBCust.LocMeters(cnt).NumUser = UBCustVI.LocMeters(cnt).NumUser
    NUBCust.LocMeters(cnt).InsDate = UBCustVI.LocMeters(cnt).InsDate
    NUBCust.LocMeters(cnt).CurRead = UBCustVI.LocMeters(cnt).CurRead
    NUBCust.LocMeters(cnt).PrevRead = UBCustVI.LocMeters(cnt).PrevRead
    NUBCust.LocMeters(cnt).CurDate = UBCustVI.LocMeters(cnt).CurDate
    NUBCust.LocMeters(cnt).PastDate = UBCustVI.LocMeters(cnt).PastDate
    NUBCust.LocMeters(cnt).ReadFlag = UBCustVI.LocMeters(cnt).ReadFlag
    NUBCust.LocMeters(cnt).AvgUse = UBCustVI.LocMeters(cnt).AvgUse
    NUBCust.LocMeters(cnt).UseCnt = UBCustVI.LocMeters(cnt).UseCnt
    NUBCust.LocMeters(cnt).MtrIDNO = ""
    NUBCust.LocMeters(cnt).MtrLat = 0
    NUBCust.LocMeters(cnt).MtrLng = 0
  NEXT
  NUBCust.CustPIN = UBCustVI.CustPIN
  NUBCust.LastTrans = UBCustVI.LastTrans
  NUBCust.CurrBalance = UBCustVI.CurrBalance
  NUBCust.PrevBalance = UBCustVI.PrevBalance
  FOR cnt = 1 TO 15
    NUBCust.CurrRevAmts(cnt) = UBCustVI.CurrRevAmts(cnt)
    NUBCust.PrevRevAmts(cnt) = UBCustVI.PrevRevAmts(cnt)
  NEXT
  NUBCust.DepositAmt = UBCustVI.DepositAmt
  NUBCust.DelFlag = UBCustVI.DelFlag
  NUBCust.PreNoteFlag = UBCustVI.PreNoteFlag
  NUBCust.WOLastTrans = UBCustVI.WOLastTrans
  NUBCust.EstFlag = ""    'UBCustVI.EstFlag
  NUBCust.MessageRec = UBCustVI.MessageRec
  NUBCust.OldRec = UBCustVI.OldRec
  NUBCust.EPPLastTran = UBCustVI.EPPLastTran
  NUBCust.NewNotes = UBCustVI.NewNotes
  NUBCust.DPCode = ""
  NUBCust.FillPad = "" 'UBCustVI.FillPad
  NUBCust.ChkByte = CHR$(5)

RETURN

FromVerII:
  NUBCust.BOOK = UBCustVII.BOOK
  NUBCust.SEQNUMB = UBCustVII.SEQNUMB
  NUBCust.Status = UBCustVII.Status
  NUBCust.OPENDATE = UBCustVII.OPENDATE
  NUBCust.SEARCH = UBCustVII.SEARCH
  NUBCust.CUSTNAME = UBCustVII.CUSTNAME
  NUBCust.ADDR1 = UBCustVII.ADDR1
  NUBCust.ADDR2 = UBCustVII.ADDR2
  NUBCust.SERVADDR = UBCustVII.SERVADDR
  NUBCust.CITY = UBCustVII.CITY
  NUBCust.STATE = UBCustVII.STATE
  NUBCust.ZIPCODE = UBCustVII.ZIPCODE
  NUBCust.HPHONE = UBCustVII.HPHONE
  NUBCust.WPHONE = UBCustVII.WPHONE
  NUBCust.SOSEC = UBCustVII.SOSEC
  NUBCust.DRVLIC = UBCustVII.DRVLIC
  NUBCust.CUSTTYPE = UBCustVII.CUSTTYPE
  NUBCust.Addr911 = UBCustVII.Addr911
  NUBCust.BillTo = UBCustVII.BillTo
  NUBCust.BILLCOPY = UBCustVII.BILLCOPY
  IF NUBCust.BILLCOPY <= 0 OR NUBCust.BILLCOPY > 99 THEN
    NUBCust.BILLCOPY = 1
  END IF
  NUBCust.POSTRTE = UBCustVII.POSTRTE
  NUBCust.BILLCYCL = UBCustVII.BILLCYCL
  NUBCust.ZONE = UBCustVII.ZONE
  NUBCust.SEQ = UBCustVII.SEQ
  NUBCust.CASHONLY = UBCustVII.CASHONLY
  NUBCust.LATEFEE = UBCustVII.LATEFEE
  NUBCust.CUTOFFYN = UBCustVII.CUTOFFYN
  NUBCust.TAXEXPT = UBCustVII.TAXEXPT
  NUBCust.SRCIT = UBCustVII.SRCIT
  NUBCust.EPPFlag = UBCustVII.EPPFlag
  NUBCust.GroupCodeRec = 0
  NUBCust.Filler1 = ""
  NUBCust.USEDRAFT = UBCustVII.USEDRAFT
  NUBCust.AcctType = UBCustVII.AcctType
  NUBCust.BANKNAME = UBCustVII.BANKNAME
  NUBCust.BANKLOC = UBCustVII.BANKLOC
  NUBCust.TRANSIT = UBCustVII.TRANSIT
  NUBCust.BANKACCT = UBCustVII.BANKACCT
  NUBCust.BILLCMNT = UBCustVII.BILLCMNT
  NUBCust.PAYCMNT = UBCustVII.PAYCMNT
  NUBCust.PUMPCODE = UBCustVII.PUMPCODE
  NUBCust.USERCODE1 = UBCustVII.USERCODE1
  NUBCust.USERCODE2 = UBCustVII.USERCODE2
  NUBCust.ProRatePCT = UBCustVII.ProRatePCT
  NUBCust.HHMSG1 = UBCustVII.HHMSG1
  NUBCust.HHMSG2 = UBCustVII.HHMSG2
  NUBCust.HHMSG3 = UBCustVII.HHMSG3
  FOR cnt = 1 TO 15
    NUBCust.Serv(cnt).RateCode = UBCustVII.Serv(cnt).RateCode
    NUBCust.Serv(cnt).RMtrType = UBCustVII.Serv(cnt).RMtrType
  NEXT
  FOR cnt = 1 TO 4
    NUBCust.FlatRates(cnt).FRDESC = UBCustVII.FlatRates(cnt).FRDESC
    NUBCust.FlatRates(cnt).FRAMT = UBCustVII.FlatRates(cnt).FRAMT
    NUBCust.FlatRates(cnt).FRFREQ = UBCustVII.FlatRates(cnt).FRFREQ
    NUBCust.FlatRates(cnt).REVSRC = UBCustVII.FlatRates(cnt).REVSRC
    NUBCust.FlatRates(cnt).NumMin = UBCustVII.FlatRates(cnt).NumMin
  NEXT
  FOR cnt = 1 TO 2
    NUBCust.Monthly(cnt).AMTOWED = UBCustVII.Monthly(cnt).AMTOWED
    NUBCust.Monthly(cnt).TotAmtPD = UBCustVII.Monthly(cnt).TotAmtPD
    NUBCust.Monthly(cnt).PayAmt = UBCustVII.Monthly(cnt).PayAmt
    NUBCust.Monthly(cnt).RevSource = UBCustVII.Monthly(cnt).RevSource
  NEXT
  NUBCust.MFEE1 = UBCustVII.MFEE1
  NUBCust.MFEE2 = UBCustVII.MFEE2
  FOR cnt = 1 TO 7
    NUBCust.LocMeters(cnt).MTRNUM = UBCustVII.LocMeters(cnt).MTRNUM
    NUBCust.LocMeters(cnt).MTRMulti = UBCustVII.LocMeters(cnt).MTRMulti
    NUBCust.LocMeters(cnt).MTRType = UBCustVII.LocMeters(cnt).MTRType
    NUBCust.LocMeters(cnt).MTRUnit = UBCustVII.LocMeters(cnt).MTRUnit
    NUBCust.LocMeters(cnt).NumUser = UBCustVII.LocMeters(cnt).NumUser
    NUBCust.LocMeters(cnt).InsDate = UBCustVII.LocMeters(cnt).InsDate
    NUBCust.LocMeters(cnt).CurRead = UBCustVII.LocMeters(cnt).CurRead
    NUBCust.LocMeters(cnt).PrevRead = UBCustVII.LocMeters(cnt).PrevRead
    NUBCust.LocMeters(cnt).CurDate = UBCustVII.LocMeters(cnt).CurDate
    NUBCust.LocMeters(cnt).PastDate = UBCustVII.LocMeters(cnt).PastDate
    NUBCust.LocMeters(cnt).ReadFlag = UBCustVII.LocMeters(cnt).ReadFlag
    NUBCust.LocMeters(cnt).AvgUse = UBCustVII.LocMeters(cnt).AvgUse
    NUBCust.LocMeters(cnt).UseCnt = UBCustVII.LocMeters(cnt).UseCnt
    NUBCust.LocMeters(cnt).MtrIDNO = ""
    NUBCust.LocMeters(cnt).MtrLat = 0
    NUBCust.LocMeters(cnt).MtrLng = 0
  NEXT
  NUBCust.CustPIN = UBCustVII.CustPIN
  NUBCust.LastTrans = UBCustVII.LastTrans
  NUBCust.CurrBalance = UBCustVII.CurrBalance
  NUBCust.PrevBalance = UBCustVII.PrevBalance
  FOR cnt = 1 TO 15
    NUBCust.CurrRevAmts(cnt) = UBCustVII.CurrRevAmts(cnt)
    NUBCust.PrevRevAmts(cnt) = UBCustVII.PrevRevAmts(cnt)
  NEXT
  NUBCust.DepositAmt = UBCustVII.DepositAmt
  NUBCust.DelFlag = UBCustVII.DelFlag
  NUBCust.PreNoteFlag = UBCustVII.PreNoteFlag
  NUBCust.WOLastTrans = UBCustVII.WOLastTrans
  NUBCust.EstFlag = ""    'UBCustVII.EstFlag
  NUBCust.MessageRec = UBCustVII.MessageRec
  NUBCust.OldRec = UBCustVII.OldRec
  NUBCust.EPPLastTran = UBCustVII.EPPLastTran
  NUBCust.NewNotes = UBCustVII.NewNotes
  NUBCust.DPCode = ""
  NUBCust.FillPad = "" 'UBCustVII.FillPad
  NUBCust.ChkByte = CHR$(5)


RETURN

'  oUBCustFile = FREEFILE
'  OPEN "ubcust.dat" FOR RANDOM SHARED AS oUBCustFile LEN = oRecLen
'  NumORecs& = LOF(oUBCustFile) / oRecLen

'  nUBCustFile = FREEFILE
 ' OPEN "ubcust.dat" FOR RANDOM SHARED AS nUBCustFile LEN = nRecLen


'  FOR cnt& = 1 TO NumORecs&
'    LOCATE 5, 1
'    PRINT "Processing:"; cnt&; "Of"; NumORecs&;
'      GET #oUBCustFile, cnt&, oUBCust

'      GOSUB FromVerIII
'    PUT #nUBCustFile, cnt&, NUBCust
'  NEXT


FromVerIII:
  NUBCust.BOOK = oUBCust.BOOK
  NUBCust.SEQNUMB = oUBCust.SEQNUMB
  NUBCust.Status = oUBCust.Status
  NUBCust.OPENDATE = oUBCust.OPENDATE
  NUBCust.SEARCH = oUBCust.SEARCH
  NUBCust.CUSTNAME = oUBCust.CUSTNAME
  NUBCust.ADDR1 = oUBCust.ADDR1
  NUBCust.ADDR2 = oUBCust.ADDR2
  NUBCust.SERVADDR = oUBCust.SERVADDR
  NUBCust.CITY = oUBCust.CITY
  NUBCust.STATE = oUBCust.STATE
  NUBCust.ZIPCODE = oUBCust.ZIPCODE
  NUBCust.HPHONE = oUBCust.HPHONE
  NUBCust.WPHONE = oUBCust.WPHONE
  NUBCust.SOSEC = oUBCust.SOSEC
  NUBCust.DRVLIC = oUBCust.DRVLIC
  NUBCust.CUSTTYPE = oUBCust.CUSTTYPE
  NUBCust.Addr911 = oUBCust.Addr911
  NUBCust.BillTo = oUBCust.BillTo
  NUBCust.BILLCOPY = oUBCust.BILLCOPY
  IF NUBCust.BILLCOPY <= 0 OR NUBCust.BILLCOPY > 99 THEN
    NUBCust.BILLCOPY = 1
  END IF
  NUBCust.POSTRTE = oUBCust.POSTRTE
  NUBCust.BILLCYCL = oUBCust.BILLCYCL
  NUBCust.ZONE = oUBCust.ZONE
  NUBCust.SEQ = oUBCust.SEQ
  NUBCust.CASHONLY = oUBCust.CASHONLY
  NUBCust.LATEFEE = oUBCust.LATEFEE
  NUBCust.CUTOFFYN = oUBCust.CUTOFFYN
  NUBCust.TAXEXPT = oUBCust.TAXEXPT
  NUBCust.SRCIT = oUBCust.SRCIT
  NUBCust.EPPFlag = oUBCust.EPPFlag
  NUBCust.GroupCodeRec = 0
  NUBCust.Filler1 = ""
  NUBCust.USEDRAFT = oUBCust.USEDRAFT
  NUBCust.AcctType = oUBCust.AcctType
  NUBCust.BANKNAME = oUBCust.BANKNAME
  NUBCust.BANKLOC = oUBCust.BANKLOC
  NUBCust.TRANSIT = oUBCust.TRANSIT
  NUBCust.BANKACCT = oUBCust.BANKACCT
  NUBCust.BILLCMNT = oUBCust.BILLCMNT
  NUBCust.PAYCMNT = oUBCust.PAYCMNT
  NUBCust.PUMPCODE = oUBCust.PUMPCODE
  NUBCust.USERCODE1 = oUBCust.USERCODE1
  NUBCust.USERCODE2 = oUBCust.USERCODE2
  NUBCust.ProRatePCT = oUBCust.ProRatePCT
  NUBCust.HHMSG1 = oUBCust.HHMSG1
  NUBCust.HHMSG2 = oUBCust.HHMSG2
  NUBCust.HHMSG3 = oUBCust.HHMSG3
  FOR cnt = 1 TO 15
    NUBCust.Serv(cnt).RateCode = oUBCust.Serv(cnt).RateCode
    NUBCust.Serv(cnt).RMtrType = oUBCust.Serv(cnt).RMtrType
  NEXT
  FOR cnt = 1 TO 4
    NUBCust.FlatRates(cnt).FRDESC = oUBCust.FlatRates(cnt).FRDESC
    NUBCust.FlatRates(cnt).FRAMT = oUBCust.FlatRates(cnt).FRAMT
    NUBCust.FlatRates(cnt).FRFREQ = oUBCust.FlatRates(cnt).FRFREQ
    NUBCust.FlatRates(cnt).REVSRC = oUBCust.FlatRates(cnt).REVSRC
    NUBCust.FlatRates(cnt).NumMin = oUBCust.FlatRates(cnt).NumMin
  NEXT
  FOR cnt = 1 TO 2
    NUBCust.Monthly(cnt).AMTOWED = oUBCust.Monthly(cnt).AMTOWED
    NUBCust.Monthly(cnt).TotAmtPD = oUBCust.Monthly(cnt).TotAmtPD
    NUBCust.Monthly(cnt).PayAmt = oUBCust.Monthly(cnt).PayAmt
    NUBCust.Monthly(cnt).RevSource = oUBCust.Monthly(cnt).RevSource
  NEXT
  NUBCust.MFEE1 = oUBCust.MFEE1
  NUBCust.MFEE2 = oUBCust.MFEE2
  FOR cnt = 1 TO 7
    NUBCust.LocMeters(cnt).MTRNUM = oUBCust.LocMeters(cnt).MTRNUM
    NUBCust.LocMeters(cnt).MTRMulti = oUBCust.LocMeters(cnt).MTRMulti
    NUBCust.LocMeters(cnt).MTRType = oUBCust.LocMeters(cnt).MTRType
    NUBCust.LocMeters(cnt).MTRUnit = oUBCust.LocMeters(cnt).MTRUnit
    NUBCust.LocMeters(cnt).NumUser = oUBCust.LocMeters(cnt).NumUser
    NUBCust.LocMeters(cnt).InsDate = oUBCust.LocMeters(cnt).InsDate
    NUBCust.LocMeters(cnt).CurRead = oUBCust.LocMeters(cnt).CurRead
    NUBCust.LocMeters(cnt).PrevRead = oUBCust.LocMeters(cnt).PrevRead
    NUBCust.LocMeters(cnt).CurDate = oUBCust.LocMeters(cnt).CurDate
    NUBCust.LocMeters(cnt).PastDate = oUBCust.LocMeters(cnt).PastDate
    NUBCust.LocMeters(cnt).ReadFlag = oUBCust.LocMeters(cnt).ReadFlag
    NUBCust.LocMeters(cnt).AvgUse = oUBCust.LocMeters(cnt).AvgUse
    NUBCust.LocMeters(cnt).UseCnt = oUBCust.LocMeters(cnt).UseCnt
    NUBCust.LocMeters(cnt).MtrIDNO = oUBCust.LocMeters(cnt).MtrIDNO
    NUBCust.LocMeters(cnt).MtrLat = oUBCust.LocMeters(cnt).MtrLat
    NUBCust.LocMeters(cnt).MtrLng = oUBCust.LocMeters(cnt).MtrLng
  NEXT
  NUBCust.CustPIN = oUBCust.CustPIN
  NUBCust.LastTrans = oUBCust.LastTrans
  NUBCust.CurrBalance = oUBCust.CurrBalance
  NUBCust.PrevBalance = oUBCust.PrevBalance
  FOR cnt = 1 TO 15
    NUBCust.CurrRevAmts(cnt) = oUBCust.CurrRevAmts(cnt)
    NUBCust.PrevRevAmts(cnt) = oUBCust.PrevRevAmts(cnt)
  NEXT
  NUBCust.DepositAmt = oUBCust.DepositAmt
  NUBCust.DelFlag = oUBCust.DelFlag
  NUBCust.PreNoteFlag = oUBCust.PreNoteFlag
  NUBCust.WOLastTrans = oUBCust.WOLastTrans
  NUBCust.EstFlag = ""    'oUBCust.EstFlag
  NUBCust.MessageRec = oUBCust.MessageRec
  NUBCust.OldRec = oUBCust.OldRec
  NUBCust.EPPLastTran = oUBCust.EPPLastTran
  NUBCust.NewNotes = oUBCust.NewNotes
  NUBCust.DPCode = oUBCust.DPCode
  NUBCust.FillPad = "" 'oUBCust.FillPad
  NUBCust.ChkByte = CHR$(5)


RETURN

