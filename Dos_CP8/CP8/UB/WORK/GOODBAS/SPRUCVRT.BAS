DEFINT A-Z
  DECLARE FUNCTION MakePhone$ (Phone$)
  DECLARE FUNCTION FileSize& (FileName$)
  DECLARE FUNCTION QPTrim$ (Text$)
  DECLARE FUNCTION Date2Num% (TheDate$)
  DECLARE FUNCTION Num2Date$ (DateNumber%)

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

'"ACCTID","ACCOUNT_CUSTID","ACCOUNT_SVCID","STATUS","CUSTNAME","ACCOUNT_SVCSTRNUM","ACCOUNT_SVCSTRNAME","ACCOUNT_METERID","LOADSEQ","DEPOSIT","WATERCHG","SEWERCHG","GARBCHG","RECONCHG","OTHERCHG","PENALTYCHG","BALFWD","CURRCHGS","PMTS"
'"ADJUSTS","LASTPMTEXT","CUTOFF","CUSTOMER_CUSTID","ACTIVITY","FIRSTNAME","LASTNAME","ADDR1","ADDR2","CITY","STATE","ZIP","POSTALSORT","PHONE","SSN","SERVICE_SVCID","LOCATIONID","SVCZONE","SERVICE_SVCSTRNUM","SERVICE_SVCSTRNAME","APTNO"
'"SERVICE_METERID","MTRINSTALL","METERSIZE","LOCNOTE","HHROUTE","READSEQ","READTYPE","READDATE","READTIME","CURREAD","PREVREAD","CONSP","AVGCSP","FLOWTHRU","FLINKCODE","LINKCONSP","TURNEDON","WRATE","SRATE","GRATE","CITYCODE","MULTIPLIER"


    CLS

    OPEN "account.txt" FOR INPUT AS #1

    INPUT #1, ACCTID$, ACCOUNTCUSTID$, ACCOUNTSVCID$, aSTATUS$, CustName$, AccountSVCStrNum$, AccountSVCStrName$, ACCOUNTMETERID$
    INPUT #1, LoadSeq$, DEPOSITAMT$, WATERCHG$, SEWERCHG$, GARBCHG$, RECONCHG$, OTHERCHG$, PENALTYCHG$, BALFWD$, CURRCHGS$, PMTS$
    INPUT #1, ADJUSTS$, LASTPMTEXT$, CUTOFF$, CUSTOMERCUSTID$, ACTIVITY$, cFirstName$, cLastName$, cADDR1$, cADDR2$, CCity$, CState$, cZIP$
    INPUT #1, POSTALSORT$, CPhone$, cSSN$, SERVICESVCID$, LOCATIONID$, SVCZone$, SERVICESVCSTRNUM$, SERVICESVCSTRNAME$, APTNO$
    INPUT #1, ServiceMeterID$, MTRInstall$, METERSIZE$, LOCNOTE$, HHROUTE$, ReadSEQ$, READTYPE$, ReadDate$, READTIME$, CurRead$, PrevRead$
    INPUT #1, CONSP$, AVGCSP$, FLOWTHRU$, FLINKCODE$, LINKCONSP$, TURNEDON$, WRate$, SRate$, GRate$, CityCode$, Multiplier$

    DO

      RCCnt& = RCCnt& + 1

      'FOR zzzz& = 1 TO 10000
      'NEXT

      INPUT #1, ACCTID$, ACCOUNTCUSTID$, ACCOUNTSVCID$, aSTATUS$, CustName$, AccountSVCStrNum$, AccountSVCStrName$, ACCOUNTMETERID$
      INPUT #1, LoadSeq$, DEPOSITAMT$, WATERCHG$, SEWERCHG$, GARBCHG$, RECONCHG$, OTHERCHG$, PENALTYCHG$, BALFWD$, CURRCHGS$, PMTS$
      INPUT #1, ADJUSTS$, LASTPMTEXT$, CUTOFF$, CUSTOMERCUSTID$, ACTIVITY$, cFirstName$, cLastName$, cADDR1$, cADDR2$, CCity$, CState$, cZIP$
      INPUT #1, POSTALSORT$, CPhone$, cSSN$, SERVICESVCID$, LOCATIONID$, SVCZone$, SERVICESVCSTRNUM$, SERVICESVCSTRNAME$, APTNO$
      INPUT #1, ServiceMeterID$, MTRInstall$, METERSIZE$, LOCNOTE$, HHROUTE$, ReadSEQ$, READTYPE$, ReadDate$, READTIME$, CurRead$, PrevRead$
      INPUT #1, CONSP$, AVGCSP$, FLOWTHRU$, FLINKCODE$, LINKCONSP$, TURNEDON$, WRate$, SRate$, GRate$, CityCode$, Multiplier$


      GOSUB Parse2Cust

      LOCATE 5, 1
      PRINT "Record: "; RCCnt&
      LOCATE 6, 1
      PRINT SPACE$(78)
      LOCATE 6, 1
      PRINT ACCTID$
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


  UBCustRec(1).CustPIN = NextPIN&
  
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  NextRec& = NumOfRecs& + 1

  UBOFile = FREEFILE
  OPEN "UBOWNER.DAT" FOR RANDOM SHARED AS UBOFile LEN = UBOwnerRecLen

'  UBMess = FREEFILE
'  OPEN "UBMESAGE.DAT" FOR RANDOM SHARED AS UBMess LEN = UBMessRecLen

  aSTATUS$ = QPTrim$(aSTATUS$)

  IF LEN(aSTATUS$) > 0 THEN
    UBCustRec(1).Status = aSTATUS$
  ELSE
    UBCustRec(1).Status = "I"
  END IF

  LoadSeq$ = QPTrim$(LoadSeq$)
  IF LEN(LoadSeq$) > 0 THEN
    Book$ = "0" + LEFT$(LoadSeq$, 1)
  ELSE
    Book$ = "00"
  END IF

  Location$ = "000000" + RIGHT$(LoadSeq$, 6)

  UBCustRec(1).Book = Book$
  UBCustRec(1).SeqNumb = RIGHT$(Location$, 6)

  WRate$ = QPTrim$(WRate$)
  SRate$ = QPTrim$(SRate$)
  GRate$ = QPTrim$(GRate$)
  CityCode$ = QPTrim$(CityCode$)
  READTYPE$ = QPTrim$(READTYPE$)

  IF LEN(WRate$) > 0 THEN
    UBCustRec(1).Serv(1).RateCode = "W" + WRate$ + CityCode$ + "1"
    IF READTYPE$ = "B" THEN
      UBCustRec(1).Serv(1).RMtrType = "T"
    ELSE
      UBCustRec(1).Serv(1).RMtrType = "W"
    END IF
  ELSE
    UBCustRec(1).Serv(1).RateCode = ""
    UBCustRec(1).Serv(1).RMtrType = ""
  END IF

  IF LEN(SRate$) > 0 THEN
    UBCustRec(1).Serv(2).RateCode = "S" + SRate$ + CityCode$ + "1"
    SELECT CASE UBCustRec(1).Serv(1).RMtrType
    CASE "T"                'Touch read
      UBCustRec(1).Serv(2).RMtrType = "T"
    CASE "W"                'has water change both to combined
      UBCustRec(1).Serv(1).RMtrType = "C"
      UBCustRec(1).Serv(2).RMtrType = "C"
    CASE ELSE               'no water only sewer
      UBCustRec(1).Serv(2).RMtrType = "S"
    END SELECT
  ELSE
    UBCustRec(1).Serv(2).RateCode = ""
    UBCustRec(1).Serv(2).RMtrType = ""
  END IF

  IF LEN(GRate$) THEN
    UBCustRec(1).Serv(3).RateCode = "GRB" + CityCode$
  ELSE
    UBCustRec(1).Serv(3).RateCode = ""
  END IF

  FOR zz = 1 TO 7  'clear meter info
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

  IF LEN(QPTrim$(UBCustRec(1).Serv(1).RMtrType)) > 0 THEN
    UBCustRec(1).LocMeters(1).MTRType = UBCustRec(1).Serv(1).RMtrType
    GOSUB SetupMeter
  ELSEIF LEN(QPTrim$(UBCustRec(1).Serv(2).RMtrType)) > 0 THEN
    UBCustRec(1).LocMeters(1).MTRType = UBCustRec(1).Serv(2).RMtrType
    GOSUB SetupMeter
  ELSE
    UBCustRec(1).LocMeters(1).MTRType = ""
  END IF
  UBCustRec(1).OpenDate = BlankInt%

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
  UBCustRec(1).PostRte = POSTALSORT$

  cLastName$ = QPTrim$(cLastName$)
  cFirstName$ = QPTrim$(cFirstName$)

  SName$ = cLastName$ + cFirstName$
  UBCustRec(1).Search = QPTrim$(SName$)
  SName$ = cFirstName$ + " " + cLastName$
  UBCustRec(1).CustName = QPTrim$(SName$)
  UBCustRec(1).Addr1 = QPTrim$(cADDR1$)
  UBCustRec(1).Addr2 = QPTrim$(cADDR2$)
  UBCustRec(1).City = QPTrim$(CCity$)

  SrvAdd& = VAL(AccountSVCStrNum$)

  IF SrvAdd& > 0 THEN
    ServAddr$ = STR$(SrvAdd&) + " " + QPTrim$(AccountSVCStrName$)
    ServAddr$ = QPTrim$(ServAddr$)
  ELSE
    ServAddr$ = QPTrim$(AccountSVCStrName$)
  END IF

  UBCustRec(1).ServAddr = ServAddr$
  UBCustRec(1).State = QPTrim$(CState$)
  UBCustRec(1).ZipCode = QPTrim$(cZIP$)
  
  UBCustRec(1).HPhone = MakePhone$(CPhone$)

  UBCustRec(1).SEQ = VAL(ReadSEQ$)
  UBCustRec(1).PumpCode = QPTrim$(SVCZone$)
  UBOwnerRec(1).OwnLName = QPTrim$(LastName$)
  UBOwnerRec(1).OwnFName = QPTrim$(FirstName$)
  UBOwnerRec(1).Addr1 = UBCustRec(1).Addr1
  UBOwnerRec(1).City = UBCustRec(1).City
  UBOwnerRec(1).State = UBCustRec(1).State
  UBOwnerRec(1).ZipCode = UBCustRec(1).ZipCode
  UBOwnerRec(1).HPhone = UBCustRec(1).HPhone
  UBOwnerRec(1).ChkByte = CHR$(1)

  UBCustRec(1).addr911 = custno$

  UBCustRec(1).MessageRec = 0          'Set Pointers To Zero
  UBCustRec(1).LastTrans = 0
  UBCustRec(1).CurrBalance = 0
  UBCustRec(1).PrevBalance = 0
  UBCustRec(1).ProratePCT = 100
  UBCustRec(1).DEPOSITAMT = 0
  UBCustRec(1).WOLastTrans = 0

'  IF LEN(CType$) > 0 THEN
'    UBCustRec(1).CustType = CType$
'  END IF

'  CNote$ = QPTrim(notes$)
'  IF LEN(CNote$) > 0 THEN
'    UBMessRec(1).MessLine(1).Line = CNote$
'    'UBMessRec(1).MessLine(1).LineDate =
'    UBMessRec(1).CustRec = NextRec&
'    MessageRecord = LOF(UBMess) / UBMessRecLen + 1
'    UBCustRec(1).MessageRec = MessageRecord
'    PUT UBMess, MessageRecord, UBMessRec(1)
'  END IF

  PUT UBFile, NextRec&, UBCustRec(1)
  PUT UBOFile, NextRec&, UBOwnerRec(1)

  CLOSE UBFile, UBOFile, UBMess

RETURN


SetupMeter:

  UBCustRec(1).LocMeters(1).MtrNum = QPTrim$(ServiceMeterID$)

  Multiplier$ = QPTrim$(Multiplier$)
  IF LEN(Multiplier$) > 0 THEN
    UBCustRec(1).LocMeters(1).NumUser = VAL(Multiplier$)
  ELSE
    UBCustRec(1).LocMeters(1).NumUser = 1
  END IF

  UBCustRec(1).LocMeters(1).MtrMulti = 1

  MTRInstall$ = QPTrim$(MTRInstall$)
  IF LEN(MTRInstall$) > 0 THEN
    UBCustRec(1).LocMeters(1).InsDate = Date2Num(MTRInstall$)
  END IF

  ReadDate$ = QPTrim$(ReadDate$)

  IF LEN(ReadDate$) = 4 THEN
    UBCustRec(1).LocMeters(1).CurDate = Date2Num(LEFT$(ReadDate$, 2) + "-" + RIGHT$(ReadDate$, 2) + "-2004")
  ELSE
    UBCustRec(1).LocMeters(1).CurDate = BlankInt
  END IF
  UBCustRec(1).LocMeters(1).PastDate = UBCustRec(1).LocMeters(1).CurDate
  UBCustRec(1).LocMeters(1).UseCnt = 1
  UBCustRec(1).LocMeters(1).AvgUse = VAL(AVGCSP$)
  UBCustRec(1).LocMeters(1).CurRead = VAL(QPTrim$(CurRead$))
  UBCustRec(1).LocMeters(1).PrevRead = VAL(QPTrim$(PrevRead$))
  UBCustRec(1).LocMeters(1).MtrUnit = "G"

RETURN

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

