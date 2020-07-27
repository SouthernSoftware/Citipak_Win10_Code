DEFINT A-Z
DECLARE SUB KILLfile (filename$)
CONST True = -1, False = NOT True

'$INCLUDE: 'ub2wind.bi'

  'DIM oubcust AS ONewUBCustRecType

  DIM UBCustVI AS NewUBCustRecTypeVI
  DIM UBCustVII AS NewUBCustRecTypeVII
  DIM NUBCust AS NewUBCustRecType

  VIRecLen = LEN(UBCustVI)
  VIIRecLen = LEN(UBCustVII)
  NRecLen = LEN(NUBCust)

'END
  KILLfile "ubcust.old"
'STOP
  GOSUB chkcheckbyte

  IF AbortFlag THEN
    GOTO Holy.Cow.Batman
  END IF

  NAME "ubcust.dat" AS "ubcust.old"

  CLS
  IF Type2Flag THEN
    PRINT "Converting Customer File Type: 2 -> 3. . . "
  ELSE
    PRINT "Converting Customer File Type: 1 -> 3. . . "
  END IF

  oUBCustFile = FREEFILE
  IF Type2Flag THEN
    OPEN "ubcust.old" FOR RANDOM AS oUBCustFile LEN = VIIRecLen
    NumORecs& = LOF(oUBCustFile) / VIIRecLen
  ELSE
    OPEN "ubcust.old" FOR RANDOM AS oUBCustFile LEN = VIRecLen
    NumORecs& = LOF(oUBCustFile) / VIRecLen
  END IF

  UBCustFile = FREEFILE
  OPEN "UBCust.dat" FOR RANDOM AS UBCustFile LEN = NRecLen

  FOR cnt& = 1 TO NumORecs&
    LOCATE 5, 1
    PRINT "Processing:"; cnt&; "Of"; NumORecs&;
    IF Type2Flag THEN
      GET #oUBCustFile, cnt&, UBCustVII
    ELSE
      GET #oUBCustFile, cnt&, UBCustVI
    END IF
    IF Type2Flag THEN
      GOSUB FromVerII
    ELSE
      GOSUB FromVerI
    END IF
    PUT #UBCustFile, cnt&, NUBCust
  NEXT

  CLOSE
  LOCATE 9, 1
  PRINT "Processing Complete."

Holy.Cow.Batman:
  KILLfile "ub2wind.exe"

  END

chkcheckbyte:

  UBCustFile = FREEFILE
  OPEN "UBCust.dat" FOR RANDOM AS UBCustFile LEN = NRecLen
  GET #UBCustFile, 1, NUBCust
  IF NUBCust.ChkByte = CHR$(4) THEN
    GET #UBCustFile, 2, NUBCust
    IF NUBCust.ChkByte = CHR$(4) THEN
      'file is allready converted
      AbortFlag = True
    END IF
  END IF
  CLOSE UBCustFile
  IF AbortFlag THEN
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
      Type2Flag = True   'cust file IS type 2
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

FromVerI:
  NUBCust.Book = UBCustVI.Book
  NUBCust.SEQNUMB = UBCustVI.SEQNUMB
  NUBCust.Status = UBCustVI.Status
  NUBCust.OPENDATE = UBCustVI.OPENDATE
  NUBCust.SEARCH = UBCustVI.SEARCH
  NUBCust.CustName = UBCustVI.CustName
  NUBCust.ADDR1 = UBCustVI.ADDR1
  NUBCust.ADDR2 = UBCustVI.ADDR2
  NUBCust.ServAddr = UBCustVI.ServAddr
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
  NUBCust.POSTRTE = UBCustVI.POSTRTE
  NUBCust.BILLCYCL = UBCustVI.BILLCYCL
  NUBCust.ZONE = UBCustVI.ZONE
  NUBCust.Seq = UBCustVI.Seq
  NUBCust.CASHONLY = UBCustVI.CASHONLY
  NUBCust.LATEFEE = UBCustVI.LATEFEE
  NUBCust.CUTOFFYN = UBCustVI.CUTOFFYN
  NUBCust.TAXEXPT = UBCustVI.TAXEXPT
  NUBCust.SRCIT = UBCustVI.SRCIT
  NUBCust.EPPFlag = UBCustVI.EPPFlag
  NUBCust.Filler1 = ""
  NUBCust.USEDRAFT = UBCustVI.USEDRAFT
  NUBCust.AcctType = UBCustVI.AcctType
  NUBCust.BankName = UBCustVI.BankName
  NUBCust.BANKLOC = UBCustVI.BANKLOC
  NUBCust.TRANSIT = UBCustVI.TRANSIT
  NUBCust.BankAcct = UBCustVI.BankAcct
  NUBCust.BILLCMNT = UBCustVI.BILLCMNT
  NUBCust.PAYCMNT = UBCustVI.PAYCMNT
  NUBCust.PumpCode = UBCustVI.PumpCode
  NUBCust.USERCODE1 = UBCustVI.USERCODE1
  NUBCust.USERCODE2 = UBCustVI.USERCODE2
  NUBCust.ProRatePCT = UBCustVI.ProRatePCT
  NUBCust.HHMSG1 = UBCustVI.HHMSG1
  NUBCust.HHMSG2 = UBCustVI.HHMSG2
  NUBCust.HHMSG3 = UBCustVI.HHMSG3
  FOR cnt = 1 TO 15
    NUBCust.Serv(cnt).RATECODE = UBCustVI.Serv(cnt).RATECODE
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
    NUBCust.LocMeters(cnt).MtrNum = UBCustVI.LocMeters(cnt).MtrNum
    NUBCust.LocMeters(cnt).MTRMulti = UBCustVI.LocMeters(cnt).MTRMulti
    NUBCust.LocMeters(cnt).MtrType = UBCustVI.LocMeters(cnt).MtrType
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
  NUBCust.FillPad = "" 'UBCustVI.FillPad
  NUBCust.ChkByte = CHR$(4)

RETURN

FromVerII:
  NUBCust.Book = UBCustVII.Book
  NUBCust.SEQNUMB = UBCustVII.SEQNUMB
  NUBCust.Status = UBCustVII.Status
  NUBCust.OPENDATE = UBCustVII.OPENDATE
  NUBCust.SEARCH = UBCustVII.SEARCH
  NUBCust.CustName = UBCustVII.CustName
  NUBCust.ADDR1 = UBCustVII.ADDR1
  NUBCust.ADDR2 = UBCustVII.ADDR2
  NUBCust.ServAddr = UBCustVII.ServAddr
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
  NUBCust.POSTRTE = UBCustVII.POSTRTE
  NUBCust.BILLCYCL = UBCustVII.BILLCYCL
  NUBCust.ZONE = UBCustVII.ZONE
  NUBCust.Seq = UBCustVII.Seq
  NUBCust.CASHONLY = UBCustVII.CASHONLY
  NUBCust.LATEFEE = UBCustVII.LATEFEE
  NUBCust.CUTOFFYN = UBCustVII.CUTOFFYN
  NUBCust.TAXEXPT = UBCustVII.TAXEXPT
  NUBCust.SRCIT = UBCustVII.SRCIT
  NUBCust.EPPFlag = UBCustVII.EPPFlag
  NUBCust.Filler1 = ""
  NUBCust.USEDRAFT = UBCustVII.USEDRAFT
  NUBCust.AcctType = UBCustVII.AcctType
  NUBCust.BankName = UBCustVII.BankName
  NUBCust.BANKLOC = UBCustVII.BANKLOC
  NUBCust.TRANSIT = UBCustVII.TRANSIT
  NUBCust.BankAcct = UBCustVII.BankAcct
  NUBCust.BILLCMNT = UBCustVII.BILLCMNT
  NUBCust.PAYCMNT = UBCustVII.PAYCMNT
  NUBCust.PumpCode = UBCustVII.PumpCode
  NUBCust.USERCODE1 = UBCustVII.USERCODE1
  NUBCust.USERCODE2 = UBCustVII.USERCODE2
  NUBCust.ProRatePCT = UBCustVII.ProRatePCT
  NUBCust.HHMSG1 = UBCustVII.HHMSG1
  NUBCust.HHMSG2 = UBCustVII.HHMSG2
  NUBCust.HHMSG3 = UBCustVII.HHMSG3
  FOR cnt = 1 TO 15
    NUBCust.Serv(cnt).RATECODE = UBCustVII.Serv(cnt).RATECODE
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
    NUBCust.LocMeters(cnt).MtrNum = UBCustVII.LocMeters(cnt).MtrNum
    NUBCust.LocMeters(cnt).MTRMulti = UBCustVII.LocMeters(cnt).MTRMulti
    NUBCust.LocMeters(cnt).MtrType = UBCustVII.LocMeters(cnt).MtrType
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
  NUBCust.FillPad = "" 'UBCustVII.FillPad
  NUBCust.ChkByte = CHR$(4)


RETURN

