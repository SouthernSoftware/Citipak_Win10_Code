DEFINT A-Z
  DECLARE FUNCTION MakePhone$ (Phone$)
  DECLARE FUNCTION FileSize& (FileName$)
  DECLARE FUNCTION QPTrim$ (Text$)
  DECLARE FUNCTION Date2Num% (TheDate$)

  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubrate.BI'
  '$INCLUDE: 'ubowner.BI'
  '$INCLUDE: 'UBTRANS.bi'

  CONST BlankInt% = -32767      'Specifies blank integer
  CONST BlankLng& = -2147483647                   '       "        long int.
  CONST BlankSng! = -3.402823E+38                 '       "        single
  CONST BlankDbl# = -1.79769313486231D+308      'Specifies blank double

  REDIM UBMessRec(1) AS UBMessRecType
  UBMessRecLen = LEN(UBMessRec(1))
  
  REDIM UBOwnerRec(1) AS UBOwnerRecType
  UBOwnerRecLen = LEN(UBOwnerRec(1))

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))     'Length of Cust Record Structure
  REDIM UBCustPIN(1) AS UBPINType      'Pin info array

  '"account","meter","desc","sname","mname","madd","mcity","mst","mzip"
  '"mphone","oname","oadd","ocity","ost","ozip","ophone","djoin","duespd"
  '"pdate","type","special","totdue","unbilled","t30","t60","t90","t120"
  '"gone","locked","inactive","notes","read3","date3","read2","date2"
  '"read1","date1","custno","race","seq","sewer","nmeters","scust"
  '"gallons","gal2","gal3"

    CLS

    OPEN "cust.txt" FOR INPUT AS #1

    DO
      'IF RCCnt& = 2373 THEN STOP
      RCCnt& = RCCnt& + 1

      'IF RCCnt& = 1000 THEN EXIT DO
      INPUT #1, Account$, meter$, Desc$, SName$, mName$, madd$, MCity$, Mst$, MZip$
      INPUT #1, MPhone$, OName$, OAdd$, OCity$, OSt$, OZip$, OPhone$, djoin$, duespd$
      INPUT #1, pdate$, CType$, special$, totdue$, unbilled$, t30$, t60$, t90$, t120$
      INPUT #1, Gone$, locked$, inactive$, notes$, Read3$, Date3$, Read2$, Date2$
      INPUT #1, Read1$, Date1$, custno$, race$, SEQ$, sewer$, nmeters$, scust$
      INPUT #1, gallons$, gal2$, gal3$

      'IF INSTR(mName$, "ROXA") > 0 THEN
      '  STOP
      'END IF
      GOSUB Parse2Cust

      LOCATE 5, 1
      PRINT "Record: "; RCCnt&
      LOCATE 6, 1
      PRINT SPACE$(78)
      LOCATE 6, 1
      PRINT Account$, mName$, notes$
      'FOR zzzz& = 1 TO 100000
      'NEXT
      'IF RCCnt& = 500 THEN EXIT DO
Skip2Next:
    LOOP UNTIL EOF(1)
Done:
  CLOSE

END

Parse2Cust:

  REDIM UBCustRec(1) AS NewUBCustRecType
  REDIM UBOwnerRec(1) AS UBOwnerRecType

  IF FileSize("UBCUSPIN.DAT") > 0 THEN 'If there was a previous pin then
    PHandle = FREEFILE
    OPEN "UBCUSPIN.DAT" FOR RANDOM SHARED AS #PHandle LEN = 4
    GET #PHandle, 1, UBCustPIN(1)      'get last pin used info
    CLOSE #PHandle
  END IF
  NextPIN& = UBCustPIN(1).PIN + 1      'Increment last pin used
  UBCustPIN(1).PIN = NextPIN&
  PHandle = FREEFILE                   'Save new last pin info
  OPEN "UBCUSPIN.DAT" FOR RANDOM SHARED AS #PHandle LEN = 4
  PUT #PHandle, 1, UBCustPIN(1)
  CLOSE #PHandle

  IF VAL(inactive$) <> 0 THEN
    UBCustRec(1).Status = "I"
  ELSE
    UBCustRec(1).Status = "A"
  END IF

  UBCustRec(1).CustPIN = NextPIN&
  
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  NextRec& = NumOfRecs& + 1

  UBOFile = FREEFILE
  OPEN "UBOWNER.DAT" FOR RANDOM SHARED AS UBOFile LEN = UBOwnerRecLen

  UBMess = FREEFILE
  OPEN "UBMESAGE.DAT" FOR RANDOM SHARED AS UBMess LEN = UBMessRecLen

  TBook$ = LEFT$(Account$, 2)
  IF TBook$ <> CBook$ THEN
    CBook$ = TBook$
    LocCount& = 0
  END IF

  CBook$ = "10"

  UBCustRec(1).Serv(1).RateCode = "WAT1"
  UBCustRec(1).Serv(1).RMtrType = "W"
  IF VAL(scust$) <> 0 THEN
    UBCustRec(1).Serv(1).RMtrType = "C"
    UBCustRec(1).Serv(2).RateCode = "SEW1"
    UBCustRec(1).Serv(2).RMtrType = "C"
  END IF

  FOR zz = 1 TO 7
    UBCustRec(1).LocMeters(zz).MtrMulti = BlankInt
    UBCustRec(1).LocMeters(zz).NumUser = BlankInt
    UBCustRec(1).LocMeters(zz).InsDate = BlankInt
    UBCustRec(1).LocMeters(zz).CurDate = BlankInt
    UBCustRec(1).LocMeters(zz).PastDate = BlankInt
    UBCustRec(1).LocMeters(zz).UseCnt = 0
    UBCustRec(1).LocMeters(zz).AvgUse = 0
    UBCustRec(1).LocMeters(zz).CurRead = BlankLng
    UBCustRec(1).LocMeters(zz).PrevRead = BlankLng
  NEXT

  UBCustRec(1).LocMeters(1).MTRType = UBCustRec(1).Serv(1).RMtrType
  UBCustRec(1).LocMeters(1).MtrMulti = 1
  NUsers = VAL(nmeters$)

  IF NUsers > 1 THEN
    UBCustRec(1).LocMeters(1).NumUser = NUsers
  ELSE
    UBCustRec(1).LocMeters(1).NumUser = 1
  END IF

  Read01& = VAL(Read1$)
  Read02& = VAL(Read2$)
  Read03& = VAL(Read3$)

  UBCustRec(1).LocMeters(1).MtrUnit = "G"
  IF Read01& > 0 THEN
    UBCustRec(1).LocMeters(1).CurRead = Read01&
    UBCustRec(1).LocMeters(1).PrevRead = Read02&
    UBCustRec(1).LocMeters(1).PastDate = Date2Num(LEFT$(Date2$, 10))
    UBCustRec(1).LocMeters(1).CurDate = Date2Num(LEFT$(Date1$, 10))
  ELSE
    UBCustRec(1).LocMeters(1).CurRead = Read02&
    UBCustRec(1).LocMeters(1).PrevRead = Read03&
    UBCustRec(1).LocMeters(1).PastDate = Date2Num(LEFT$(Date3$, 10))
    UBCustRec(1).LocMeters(1).CurDate = Date2Num(LEFT$(Date2$, 10))
  END IF

  UBCustRec(1).LocMeters(1).MtrNum = QPTrim$(meter$)

  UBCustRec(1).CashOnly = "N"
  UBCustRec(1).LATEFEE = "Y"
  UBCustRec(1).CUTOFFYN = "Y"
  UBCustRec(1).TAXEXPT = "N"
  UBCustRec(1).SRCIT = "N"
  UBCustRec(1).EPPFlag = "N"
  UBCustRec(1).Filler1 = ""
  UBCustRec(1).BillCopy = 1
  UBCustRec(1).USEDRAFT = "N"
  UBCustRec(1).BillTo = "C"

  SEQLng& = VAL(SEQ$)

  Location$ = "000000" + QPTrim$(STR$(SEQLng&))

  Location$ = RIGHT$(Location$, 6)

  UBCustRec(1).Book = CBook$
  UBCustRec(1).SeqNumb = Location$
  
  UBCustRec(1).Search = QPTrim$(SName$)
  UBCustRec(1).CustName = mName$
  UBCustRec(1).Addr1 = QPTrim$(madd$)
  UBCustRec(1).City = QPTrim$(MCity$)
  UBCustRec(1).ServAddr = UBCustRec(1).Addr1
  UBCustRec(1).State = QPTrim$(Mst$)
  UBCustRec(1).ZipCode = QPTrim$(MZip$)
  
  UBCustRec(1).HPhone = MakePhone$(MPhone$)

  UBCustRec(1).SEQ = VAL(SEQ$)
  OwnerName$ = QPTrim$(OName$)
  IF LEN(OwnerName$) > 20 THEN
    UBOwnerRec(1).OwnLName = RIGHT$(OwnerName$, 20)
    UBOwnerRec(1).OwnFName = LEFT$(OwnerName$, 15)
  ELSE
    UBOwnerRec(1).OwnLName = OwnerName$
  END IF
  UBOwnerRec(1).Addr1 = QPTrim$(OAdd$)
  UBOwnerRec(1).City = QPTrim$(OCity$)
  UBOwnerRec(1).State = QPTrim$(OSt$)
  UBOwnerRec(1).ZipCode = QPTrim$(OZip$)
  UBOwnerRec(1).HPhone = MakePhone$(OPhone$)
  UBOwnerRec(1).ChkByte = CHR$(1)

  UBCustRec(1).addr911 = custno$

  UBCustRec(1).MessageRec = 0          'Set Pointers To Zero
  UBCustRec(1).LastTrans = 0
  UBCustRec(1).CurrBalance = 0
  UBCustRec(1).PrevBalance = 0
  UBCustRec(1).ProratePCT = 100
  UBCustRec(1).DepositAmt = 0
  UBCustRec(1).WOLastTrans = 0

  IF LEN(CType$) > 0 THEN
    UBCustRec(1).CustType = CType$
  END IF

  CNote$ = QPTrim(notes$)
  IF LEN(CNote$) > 0 THEN
    UBMessRec(1).MessLine(1).Line = CNote$
    'UBMessRec(1).MessLine(1).LineDate =
    UBMessRec(1).CustRec = NextRec&
    MessageRecord = LOF(UBMess) / UBMessRecLen + 1
    UBCustRec(1).MessageRec = MessageRecord
    PUT UBMess, MessageRecord, UBMessRec(1)
  END IF

  PUT UBFile, NextRec&, UBCustRec(1)
  PUT UBOFile, NextRec&, UBOwnerRec(1)

  CLOSE UBFile, UBOFile, UBMess

RETURN

SUB CustMessageSystem (RecNo&)
  
  CustRec& = RecNo&
  
  
  
ExitMessageInquiry:
  EXIT SUB
  '***************
  
'LoadMessageInfo:
'  MessageRecord = UBCustRec(1).MessageRec
'  IF MessageRecord > 0 THEN
'    UBMess = FREEFILE
'    OPEN "UBMESAGE.DAT" FOR RANDOM SHARED AS UBMess LEN = UBMessRecLen
'    GET UBMess, MessageRecord, UBMessRec(1)
'    CLOSE
'    BCopy VARSEG(UBMessRec(1)), VARPTR(UBMessRec(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), UBMessRecLen, 0
'    CALL UnPackBuffer(0, 0, Form$(), Fld())
'  END IF
'RETURN
  
'SaveRecord:
'  UBMess = FREEFILE
'  OPEN "UBMESAGE.DAT" FOR RANDOM SHARED AS UBMess LEN = UBMessRecLen
'  IF MessageRecord = 0 THEN
'    MessageRecord = LOF(UBMess) / LEN(UBMessRec(1)) + 1
'  END IF
'
'  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(UBMessRec(1)), VARPTR(UBMessRec(1)), UBMessRecLen, 0
'  PUT UBMess, MessageRecord, UBMessRec(1)
'  CLOSE
'
'  UBCust = FREEFILE
'  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
'  GET UBCust, CustRec&, UBCustRec(1)
'  UBCustRec(1).MessageRec = MessageRecord
'  PUT UBCust, CustRec&, UBCustRec(1)
'  CLOSE
'RETURN
  
END SUB

FUNCTION MakePhone$ (Phone$)
  HPhone$ = ""
  TPhone$ = QPTrim$(Phone$)
  TLen = LEN(TPhone$)
  IF TLen = 0 THEN
    GOTO NOPhoneExit
  ELSE
    FOR Cnt = 1 TO TLen
      TChr$ = MID$(TPhone$, Cnt, 1)
      IF INSTR("1234567890", TChr$) > 0 THEN
        HPhone$ = HPhone$ + TChr$
      END IF
    NEXT
  END IF

NOPhoneExit:
  IF LEN(HPhone$) = 7 THEN
    MakePhone$ = "   " + HPhone$
  ELSE
    MakePhone$ = HPhone$
  END IF
END FUNCTION

