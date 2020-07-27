DEFINT A-Z
DECLARE SUB KILLFILE (filename$)
CONST true = -1, False = NOT true

'$INCLUDE: 'ub2wind5.bi'
'$INCLUDE: 'wadecust.bi'


  DIM oUBCust AS NewUBCustRecTypeWade

  DIM NUBCust AS NewUBCustRecType

  oreclen = LEN(oUBCust)
  nRecLen = LEN(NUBCust)

  KILLFILE "ubcust.old"
  'GOSUB chkcheckbyte

  'IF AbortFlag THEN
  '  GOTO Holy.Cow.Batman
  'END IF

  CALL KILLFILE("ubcust.old")

  NAME "ubcust.dat" AS "ubcust.old"

  CLS

  oUBCustFile = FREEFILE
    OPEN "ubcust.old" FOR RANDOM AS oUBCustFile LEN = oreclen
    NumORecs& = LOF(oUBCustFile) / oreclen
  nUBCustFile = FREEFILE
  OPEN "ubcust.dat" FOR RANDOM SHARED AS nUBCustFile LEN = nRecLen

  FOR cnt& = 1 TO NumORecs&
    LOCATE 5, 1
    PRINT "Processing:"; cnt&; "Of"; NumORecs&;
      GET #oUBCustFile, cnt&, oUBCust

      GOSUB FromVerIII
      PUT #nUBCustFile, cnt&, NUBCust
  NEXT

  CLOSE
  LOCATE 9, 1
  PRINT "Processing Complete."

Holy.Cow.Batman:
  KILLFILE "ub2wind5.exe"

  END



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
  'NUBCust.EPPFlag = oUBCust.EPPFlag
  NUBCust.GroupCodeRec = 0
  NUBCust.Filler1 = ""
  NUBCust.USEDRAFT = oUBCust.USEDRAFT
  'NUBCust.AcctType = oUBCust.AcctType
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
    'NUBCust.LocMeters(cnt).MtrIDNO = oUBCust.LocMeters(cnt).MtrIDNO
    'NUBCust.LocMeters(cnt).MtrLat = oUBCust.LocMeters(cnt).MtrLat
    'NUBCust.LocMeters(cnt).MtrLng = oUBCust.LocMeters(cnt).MtrLng
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
  'NUBCust.EPPLastTran = oUBCust.EPPLastTran
  NUBCust.NewNotes = oUBCust.NewNotes
  'NUBCust.DPCode = oUBCust.DPCode
  NUBCust.FillPad = "" 'oUBCust.FillPad
  NUBCust.ChkByte = CHR$(5)


RETURN

