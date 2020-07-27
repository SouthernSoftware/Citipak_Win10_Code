DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Date2Num (Text$)

  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubrate.BI'
  '$INCLUDE: 'ubowner.BI'

  CONST BlankInt% = -32767      'Specifies blank integer
  CONST BlankLng& = -2147483647                   '       "        long int.
  CONST BlankSng! = -3.402823E+38                 '       "        single
  CONST BlankDbl# = -1.79769313486231D+308      'Specifies blank double

  CONST False = 0, True = NOT False

'GOTO SkipToRightHere:
  CLS
  PRINT "Importing Customers. . ."
  REDIM UBOwnerRec(1) AS UBOwnerRecType
  UBOwnerRecLen = LEN(UBOwnerRec(1))
  UBOwnerRec(1).OwnLName = ""
  UBOwnerRec(1).OwnFName = ""
  UBOwnerRec(1).ADDR1 = ""
  UBOwnerRec(1).ADDR2 = ""
  UBOwnerRec(1).CITY = ""
  UBOwnerRec(1).STATE = ""
  UBOwnerRec(1).ZIPCODE = ""
  UBOwnerRec(1).HPHONE = ""
  UBOwnerRec(1).WPHONE = ""
  UBOwnerRec(1).ChkByte = CHR$(1)

  Cnt& = 487

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))     'Length of Cust Record Structure

  REDIM UBCustPIN(1) AS UBPINType      'Pin info array

  PHandle = FREEFILE
  OPEN "UBCUSPIN.DAT" FOR RANDOM SHARED AS #PHandle LEN = 4
  GET #PHandle, 1, UBCustPIN(1)      'get last pin used info
  NextPIN& = UBCustPIN(1).PIN       'Increment last pin used

  UBOwner = FREEFILE
  OPEN "UBOWNER.DAT" FOR RANDOM SHARED AS UBOwner LEN = UBOwnerRecLen

  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  NextRec& = NumOfRecs&
  InfoHandle = FREEFILE
  OPEN "signup2.csv" FOR INPUT AS InfoHandle

  INPUT #InfoHandle, ID$, GoDate$, PNAME$, PADDRESS$, OWNERFIRST$, OWNERLASTN$, CITY$, STATE$, ZIPCODE$, PHONE$, MAILINGADD$, MAILINGCIT$, STATE$, MAILZIP$
  DO UNTIL EOF(InfoHandle)
    INPUT #InfoHandle, ID$, GoDate$, PNAME$, PADDRESS$, OWNERFIRST$, OWNERLASTN$, CITY$, STATE$, ZIPCODE$, PHONE$, MAILINGADD$, MAILINGCIT$, STATE$, MAILZIP$
    LOCATE 5, 1
    NextRec& = NextRec& + 1
    NextPIN& = NextPIN& + 1
    Added& = Added& + 1
    Cnt& = Cnt& + 1
    PRINT "Adding:"; Added&;
    Loc$ = "000000"
    UBCustRec(1).Book = "80"
    UBLoc$ = Loc$ + QPTrim$(STR$(Cnt& * 20))
    UBCustRec(1).SeqNumb = RIGHT$(UBLoc$, 6)
    UBCustRec(1).Status = "A"
    UBCustRec(1).OpenDate = Date2Num(GoDate$)
    UBCustRec(1).Search = QPTrim$(OWNERLASTN$) + QPTrim$(OWNERFIRST$)
    UBCustRec(1).CustName = QPTrim$(OWNERFIRST$) + " " + QPTrim$(OWNERLASTN$)
    UBCustRec(1).ADDR1 = QPTrim$(MAILINGADD$)
    UBCustRec(1).ServAddr = QPTrim$(PADDRESS$)
    UBCustRec(1).CITY = QPTrim$(MAILINGCIT$)
    UBCustRec(1).STATE = QPTrim$(STATE$)
    UBCustRec(1).ZIPCODE = QPTrim$(MAILZIP$)
    UBCustRec(1).BillTo = "C"
    UBCustRec(1).BillCopy = 1
    UBCustRec(1).BillCycl = 80
    UBCustRec(1).CashOnly = "N"
    UBCustRec(1).LateFee = "Y"
    UBCustRec(1).CutOffYN = "Y"
    UBCustRec(1).TaxExpt = "N"
    UBCustRec(1).SrCIT = ""
    UBCustRec(1).EPPFlag = "N"
    UBCustRec(1).UseDraft = "N"
    UBCustRec(1).ProratePCT = 100
    UBCustRec(1).HHMsg1 = ""
    UBCustRec(1).HHMsg2 = ""
    UBCustRec(1).HHMsg3 = ""
    FOR Cnt = 1 TO 4
      UBCustRec(1).FlatRates(Cnt).FRDESC = ""
      UBCustRec(1).FlatRates(Cnt).FRAMT = 0
      UBCustRec(1).FlatRates(Cnt).FRFREQ = ""
      UBCustRec(1).FlatRates(Cnt).REVSRC = 0
      UBCustRec(1).FlatRates(Cnt).NumMin = 1
    NEXT
    UBCustRec(1).FlatRates(1).FRDESC = "TAP FEE"
    UBCustRec(1).FlatRates(1).FRAMT = 250
    UBCustRec(1).FlatRates(1).FRFREQ = "N"
    UBCustRec(1).FlatRates(1).REVSRC = 10
    UBCustRec(1).FlatRates(1).NumMin = 1
    FOR Cnt = 1 TO 2
      UBCustRec(1).Monthly(Cnt).AMTOWED = 0
      UBCustRec(1).Monthly(Cnt).TotAmtPD = 0
      UBCustRec(1).Monthly(Cnt).PayAmt = 0
      UBCustRec(1).Monthly(Cnt).RevSource = 0
    NEXT
    UBCustRec(1).MFee1 = 0
    UBCustRec(1).MFee2 = 0
    FOR Cnt = 1 TO 7
      UBCustRec(1).LocMeters(Cnt).MTRNUM = ""
      UBCustRec(1).LocMeters(Cnt).MTRMulti = 1
      UBCustRec(1).LocMeters(Cnt).MTRType = ""
      UBCustRec(1).LocMeters(Cnt).MTRUnit = ""
      UBCustRec(1).LocMeters(Cnt).NumUser = 1
      UBCustRec(1).LocMeters(Cnt).InsDate = BlankInt%
      UBCustRec(1).LocMeters(Cnt).CurRead = BlankLng&
      UBCustRec(1).LocMeters(Cnt).PrevRead = BlankLng&
      UBCustRec(1).LocMeters(Cnt).CurDate = BlankInt%
      UBCustRec(1).LocMeters(Cnt).PastDate = BlankInt%
      UBCustRec(1).LocMeters(Cnt).ReadFlag = "N"
      UBCustRec(1).LocMeters(Cnt).AvgUse = 0
      UBCustRec(1).LocMeters(Cnt).UseCnt = 0
    NEXT
    
    UBCustRec(1).CustPin = NextPIN&
    UBCustRec(1).LastTrans = 0
    UBCustRec(1).CurrBalance = 0
    UBCustRec(1).PrevBalance = 0
    FOR Cnt = 1 TO 15
      UBCustRec(1).Serv(Cnt).RATECODE = ""
      UBCustRec(1).Serv(Cnt).RMtrType = ""
      UBCustRec(1).CurrRevAmts(Cnt) = 0
      UBCustRec(1).PrevRevAmts(Cnt) = 0
    NEXT
    UBCustRec(1).DepositAmt = 0
    UBCustRec(1).DelFlag = 0
    UBCustRec(1).PreNoteFlag = 0
    UBCustRec(1).WOLastTrans = 0
    UBCustRec(1).EstFlag = ""
    UBCustRec(1).MessageRec = 0
    UBCustRec(1).OldRec = 0
    UBCustRec(1).EPPLastTran = 0
    UBCustRec(1).NewNotes = 0
    UBCustRec(1).Fillpad = ""
    PUT UBFile, NextRec&, UBCustRec(1)
    PUT UBOwner, NextRec&, UBOwnerRec(1)
  LOOP
  UBCustPIN(1).PIN = NextPIN&
  PUT #PHandle, 1, UBCustPIN(1)
  CLOSE
  PRINT
  PRINT
  PRINT
  PRINT
  PRINT "Import Complete."
  PRINT
  PRINT "Hi Peggy."
  CALL KillFile("ubaddem.exe")

'SkipToRightHere:
'  CLS
'  InfoHandle = FREEFILE
'  OPEN "signup.csv" FOR INPUT AS InfoHandle
'  INPUT #InfoHandle, ID$, GoDate$, PName$, PAddress$, OWNERFIRST$, OWNERLASTN$, CITY$, STATE$, ZIPCODE$, PHONE$, MAILINGADD$, MAILINGCIT$, STATE$, MailZip$
'  DO UNTIL EOF(InfoHandle)
'    INPUT #InfoHandle, ID$, GoDate$, PName$, PAddress$, OWNERFIRST$, OWNERLASTN$, CITY$, STATE$, ZIPCODE$, PHONE$, MAILINGADD$, MAILINGCIT$, STATE$, MailZip$
'    PRINT PAddress$, QPTrim$(OWNERFIRST$); " "; OWNERLASTN$, GoDate$
'    PRINT MAILINGADD$
'    DO
'      zz$ = INKEY$
'    LOOP UNTIL LEN(zz$) > 0
'    IF zz$ = CHR$(27) THEN EXIT DO
'  LOOP
'  CLOSE


  

