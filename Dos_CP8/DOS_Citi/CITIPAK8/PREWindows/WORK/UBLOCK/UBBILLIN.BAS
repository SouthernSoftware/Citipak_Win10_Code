DEFINT A-Z
DECLARE FUNCTION Date2Num (TheDate$)
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE FUNCTION ErrorScrn (WhatError%, Acct&)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION GetAdjFactor# ()
DECLARE FUNCTION GetBillBook% ()
DECLARE FUNCTION GetBillCycle% ()
DECLARE FUNCTION GetCustMeterType% (UBCustRec() AS ANY, ThisMeter%)
DECLARE FUNCTION GetNumOfRevs% ()
DECLARE FUNCTION GetNumRateRecs% ()
DECLARE FUNCTION GetRevCharge# (RateTbl AS ANY, TMeterConsp&, MeterMulti&)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValL& (LongNum$)
DECLARE FUNCTION Round# (N#)
DECLARE SUB BlockClear ()
DECLARE SUB ClearScrn ()
DECLARE SUB CursorOff ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB Get.Moose.OR.Key (Ky$, MooseButton%, MRow%, MCol%)
DECLARE SUB GetPreBillOrder (Choice%, ExitFlag%, SeqFlag$)
DECLARE SUB HideCursor ()
DECLARE SUB KillFile (FileName$)
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetUpLen%)
DECLARE SUB MakeBillFile (AbortFlag%, FuelAdjAmt#, ThisCycle%, ThisBook%)
DECLARE SUB MakeMowZipCodeIndex (IndexText$)
DECLARE SUB MakePostalIndex (IndexText$)
DECLARE SUB MakeSequenceIndex (IndexText$)
DECLARE SUB MakeZipCodeIndex (IndexText$)
DECLARE SUB PostBillTrans ()
DECLARE SUB PreBillReport ()
DECLARE SUB PrintHighLowReport ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB RateCodeErrScrn (RATECODE$)
DECLARE SUB RestScrn (ScrnArray())
DECLARE SUB SaveScrn (ScrnArray())
DECLARE SUB ShowCursor ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB UBLog (Text$)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB WaitForAction ()
DECLARE SUB QPSound (Freq, Dur)

  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'ubdraft.BI'
  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubrate.BI'
  
  STACK 8000
  
  CONST False = 0, True = NOT False
  
  CrLf$ = CHR$(13) + CHR$(10)
  FF$ = CHR$(12)
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen
  
  IF INSTR(UBSetUpRec(1).DEFCITY, "MCCOR") > 0 THEN
    MccorFlag = True
  END IF
  
'033198 Added Conway special handling
  
  DIM SHARED ConwayFlag AS INTEGER
  
  IF INSTR(UBSetUpRec(1).UTILNAME, "CONWAY") > 0 THEN
    ConwayFlag = True
  END IF
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 12)
  MChoice$(1) = " Meter Readings"
  MChoice$(2) = " Pre-Billing Report"
  MChoice$(3) = " Utility Bill Printing Menu"
  MChoice$(4) = " Post Billing Transactions"
  MChoice$(5) = " Bank Draft Processing"
  MChoice$(6) = " Penalty Processing Menu"
  MChoice$(7) = " Refund a Customers Deposit"
  MChoice$(8) = " Utility Billing Adjustments "
  MChoice$(9) = " Late Notice Processing "
  MChoice$(10) = " Exit to OS"
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 18   'limit the box length to go no lower than line
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  '--Find max menu width
  FOR Cnt = 1 TO UBOUND(MChoice$)
    TLen = LEN(MChoice$(Cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT
  
  '--Center Menu within Screen
  Row = ((25 - (UBOUND(MChoice$))) \ 2) + 1
  Col = ((80 - MaxLen) \ 2)
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    
    TitleBox 2, Col, MaxLen + 3, "Billing, Readings, Drafts ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN
      Choice = 0
      ExitFlag = True
      EXIT DO
    END IF
    
    SELECT CASE Choice
    CASE 1
      RUN "UBMTREAD"
    CASE 2
      PreBillReport
    CASE 3
      RUN "UBPRNBIL"
    CASE 4
      PostBillTrans
    CASE 5
      RUN "UBDRAFT"
    CASE 6
      IF MccorFlag THEN
        RUN "ubmpnlty"
      ELSE
        RUN "ubpenlty"
      END IF
    CASE 7
      RUN "UBREFUND"
    CASE 8
      RUN "UBADJUST"
    CASE 9
      RUN "UBLNPRNT"
    CASE 10
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP UNTIL ExitFlag
  
  RUN "UBMENU"

FUNCTION ErrorScrn (WhatError%, Acct&)
  
  ErrorScrn = True
  
  AcCol = 22
  REDIM TempArray(0) AS INTEGER
  SaveScrn TempArray()
  
  HideCursor
  BlockClear
  DisplayUBScrn "ERRSCRN1"
  
  SELECT CASE WhatError
  CASE 1
    QPrintRC "Has Invalid Reading!", 10, 37, -1
  CASE 2
    QPrintRC "Invalid Book Number!", 10, 37, -1
  CASE 3
    AcCol = 21
    QPrintRC "Has an INVALID RATE CODE!!", 10, 35, -1
  CASE 4
    QPrintRC "Has Mismatched Meters!", 10, 37, -1
  CASE 5
    QPrintRC "Has an INVALID Reading!", 10, 37, -1
  CASE 6
    QPrintRC "INVALID Flat Rate Info!", 10, 37, -1
  CASE 7
    QPrintRC "INVALID Monthly Billed Code!", 10, 35, -1
  CASE 8
    QPrintRC "Meters with NO RATE Code!", 10, 36, -1
  CASE 9
    QPrintRC "Invalid Customer Type!", 10, 36, -1
  END SELECT
  QPrintRC "ACCOUNT:" + STR$(Acct&), 10, AcCol, -1
  QPrintRC "Correct and Print Again.", 13, 28, -1
  
  ShowCursor
  Get.Moose.OR.Key Ky$, MooseButton%, MRow%, MCol%

  IF LEN(Ky$) = 2 THEN
    IF RIGHT$(Ky$, 1) = "g" THEN
      ErrorScrn = False
      'LPRINT Acct&
    END IF
  END IF
  RestScrn TempArray()
  ERASE TempArray
  
END FUNCTION

FUNCTION GetNumOfRevs%
  NumOfRevs = 15
  REDIM UBSetUpRec(1) AS UBSetupRecType
  UBSetUpLen = LEN(UBSetUpRec(1))
  FGetAH "UBSETUP.DAT", UBSetUpRec(1), UBSetUpLen, 1            'load it
  FOR RevCnt = 1 TO 15
    TempRev$ = QPTrim$(UBSetUpRec(1).Revenues(RevCnt).REVNAME)
    IF LEN(TempRev$) = 0 THEN
      NumOfRevs = RevCnt - 1
      EXIT FOR
    END IF
  NEXT
  GetNumOfRevs = NumOfRevs
  ERASE UBSetUpRec
END FUNCTION

SUB GetPreBillOrder (Choice, ExitFlag, SeqFlag$)
  
  REDIM MChoice$(1 TO 6)
  
  MChoice$(1) = "Customer Name Order"
  MChoice$(2) = "Account Number Order"
  MChoice$(3) = "Location Number Order"
  MChoice$(4) = "Postal Carrier Route Order"
  MChoice$(5) = "ZipCode Order"
  
  IF SeqFlag$ = "Y" THEN
    MChoice$(6) = "Sequence Number Order"
  END IF
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 17   'limit the box length to go no lower than line 20
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  '--Find max menu width
  FOR Cnt = 1 TO UBOUND(MChoice$)
    TLen = LEN(MChoice$(Cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT
  
  'Restart:
  '--Center Menu within Screen
  Row = ((25 - (UBOUND(MChoice$))) \ 2)
  Col = ((80 - MaxLen) \ 2) - 1
  
  
  DO            '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    TitleBox 2, Col, MaxLen + 3, "Pre-Billing Report ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN
      ExitFlag = True
      Choice = 0
    END IF
    
    EXIT DO
    
  LOOP
  
  
END SUB

FUNCTION GetRevCharge# (RateTbl AS UBRateTblRecType, TMeterConsp&, MeterMulti&) STATIC
  
  MinBillAmt# = RateTbl.MINAMT
  
  IF MinBillAmt# < -1000000 THEN
    MinBillAmt# = 0
    TAmt# = -1
    GOTO GotTAmt
  END IF
  
  LastTblCnt = 10
  FOR BCnt = 1 TO 10
    IF RateTbl.TblBreaks(BCnt).UNITAMT <= 0 THEN
      LastTblCnt = BCnt - 1
      EXIT FOR
    END IF
  NEXT
  
  MeterConsump& = TMeterConsp&
  
  TAmt# = 0
  
  IF LastTblCnt >= 2 THEN
    IF MeterConsump& >= RateTbl.TblBreaks(1).UNITS AND MeterConsump& <= RateTbl.TblBreaks(2).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(1).UNITS)
'special patch for cave junction
      IF UNITS& = 0 THEN
        UNITS& = 1
      END IF
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(1).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(2).UNITS - RateTbl.TblBreaks(1).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(1).UNITAMT))
    END IF
  ELSE          'no other rate breaks
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(1).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(1).UNITAMT))
    GOTO GotTAmt
  END IF
  
  'Break 2
  IF LastTblCnt >= 3 THEN
    IF MeterConsump& > RateTbl.TblBreaks(2).UNITS AND MeterConsump& <= RateTbl.TblBreaks(3).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(2).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(2).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(3).UNITS - RateTbl.TblBreaks(2).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(2).UNITAMT))
    END IF
  ELSE
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(2).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(2).UNITAMT))
    GOTO GotTAmt
  END IF
  
  'Break 3
  IF LastTblCnt >= 4 THEN
    IF MeterConsump& >= RateTbl.TblBreaks(3).UNITS AND MeterConsump& <= RateTbl.TblBreaks(4).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(3).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(3).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(4).UNITS - RateTbl.TblBreaks(3).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(3).UNITAMT))
    END IF
  ELSE
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(3).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(3).UNITAMT))
    GOTO GotTAmt
  END IF
  
  'Break 4
  IF LastTblCnt >= 5 THEN
    IF MeterConsump& >= RateTbl.TblBreaks(4).UNITS AND MeterConsump& <= RateTbl.TblBreaks(5).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(4).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(4).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(5).UNITS - RateTbl.TblBreaks(4).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(4).UNITAMT))
    END IF
  ELSE
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(4).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(4).UNITAMT))
    GOTO GotTAmt
  END IF
  
  'break 5
  IF LastTblCnt >= 6 THEN
    IF MeterConsump& >= RateTbl.TblBreaks(5).UNITS AND MeterConsump& <= RateTbl.TblBreaks(6).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(5).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(5).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(6).UNITS - RateTbl.TblBreaks(5).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(5).UNITAMT))
    END IF
  ELSE
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(5).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(5).UNITAMT))
    GOTO GotTAmt
  END IF
  
  'break 6
  IF LastTblCnt >= 7 THEN
    IF MeterConsump& >= RateTbl.TblBreaks(6).UNITS AND MeterConsump& <= RateTbl.TblBreaks(7).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(6).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(6).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(7).UNITS - RateTbl.TblBreaks(6).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(6).UNITAMT))
    END IF
  ELSE
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(6).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(6).UNITAMT))
    GOTO GotTAmt
  END IF
  
  'break 7
  IF LastTblCnt >= 8 THEN
    IF MeterConsump& >= RateTbl.TblBreaks(7).UNITS AND MeterConsump& <= RateTbl.TblBreaks(8).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(7).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(7).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(8).UNITS - RateTbl.TblBreaks(7).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(7).UNITAMT))
    END IF
  ELSE
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(7).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(7).UNITAMT))
    GOTO GotTAmt
  END IF
  
  'break 8
  IF LastTblCnt >= 9 THEN
    IF MeterConsump& >= RateTbl.TblBreaks(8).UNITS AND MeterConsump& <= RateTbl.TblBreaks(9).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(8).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(8).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(9).UNITS - RateTbl.TblBreaks(8).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(8).UNITAMT))
    END IF
  ELSE
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(8).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(8).UNITAMT))
    GOTO GotTAmt
  END IF
  
  'break 9
  IF LastTblCnt >= 10 THEN
    IF MeterConsump& >= RateTbl.TblBreaks(9).UNITS AND MeterConsump& <= RateTbl.TblBreaks(10).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(9).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(9).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(10).UNITS - RateTbl.TblBreaks(9).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(9).UNITAMT))
    END IF
  ELSE
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(9).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(9).UNITAMT))
    GOTO GotTAmt
  END IF
  
GotTAmt:
  GetRevCharge# = Round#(MinBillAmt# + TAmt#)
  
END FUNCTION

SUB MakeBillFile (AbortFlag%, FuelAdjAmt#, ThisCycle%, ThisBook%)
  
  BlockClear
  ShowProcessingScrn "Calculating Utility Charges."
  
  UBLog "IN: MakeBillFile."
  UBLog "MBF: Calculating charges."
  
  REDIM ProrateServ(1 TO 15) AS INTEGER

  REDIM UBSetUp(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUp(), UBSetUpLen
  
  IF INSTR(UBSetUp(1).UTILNAME, "PRINCETON") > 0 THEN
    PrinceFlag = True
  END IF

  IF INSTR(UBSetUp(1).UTILNAME, "YADKIN") > 0 THEN
    YadkinFlag = True
  END IF
  
'working here
  IF INSTR(UBSetUp(1).UTILNAME, "WADE") > 0 THEN 'OR INSTR(UBSetUp(1).UTILNAME, "WADE") THEN
    WadeFlag = True
  END IF

  IF INSTR(UBSetUp(1).UTILNAME, "ELKTON") > 0 THEN
    ElkFlag = True
  END IF

  IF INSTR(UBSetUp(1).UTILNAME, "SCOTTSBURG") > 0 THEN
    ScottFlag = True
  END IF
  
  IF UBSetUp(1).SkipInactive = "Y" THEN
    SkipInactive = True
  END IF
  
  IF UBSetUp(1).PreByBook = "Y" AND ThisBook > 0 THEN
    BookFlag = True
  ELSEIF UBSetUp(1).BILLCYCL = "Y" THEN
    CycleFlag = True
  END IF
  
  'find the electric revenue position
  FOR ThisRevCnt = 1 TO 15
    IF INSTR(UBSetUp(1).Revenues(ThisRevCnt).REVNAME, "ELECTRIC") THEN
      ElecRev = ThisRevCnt
      EXIT FOR
    END IF
  NEXT

  FOR ThisRevCnt = 1 TO 15
    IF UBSetUp(1).Revenues(ThisRevCnt).ProRate = "Y" THEN
      ProrateServ(ThisRevCnt) = True
    END IF
  NEXT
  
  REDIM UBBillRec(1) AS UBTransRecType
  REDIM UBCustRec(1) AS NewUBCustRecType
  
  UBBillRecLen = LEN(UBBillRec(1))
  UBCustRecLen = LEN(UBCustRec(1))
  
  NumOfRates = GetNumRateRecs%
  
  REDIM UBRateTbls(1 TO NumOfRates) AS UBRateTblRecType
  UBRateTblRecLen = LEN(UBRateTbls(1))
  
  RateFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS RateFile LEN = UBRateTblRecLen
  FOR Cnt = 1 TO NumOfRates
    GET RateFile, Cnt, UBRateTbls(Cnt)
  NEXT
  CLOSE RateFile
  
  NumCustRec& = FileSize&("UBCUST.DAT") \ UBCustRecLen
  
  IF Exist(UBBillsFile) THEN
    KillFile UBBillsFile
  END IF
  
  BillFile = FREEFILE
  OPEN UBBillsFile FOR RANDOM SHARED AS BillFile LEN = UBBillRecLen
  
  CustFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CustFile LEN = UBCustRecLen
  
  BillCnt = 0
  NumOfRevs = GetNumOfRevs%

  FOR LCnt& = 1 TO NumCustRec&
    REDIM UBBillRec(1) AS UBTransRecType        'clear bill rec for this customer
    GET CustFile, LCnt&, UBCustRec(1)
'
    IF UBCustRec(1).DelFlag <> 0 THEN
      UBBillRec(1).TransAmt = 0
      UBBillRec(1).ActiveFlag = False
      UBBillRec(1).CustAcctNo = LCnt&
      GOTO MSkipEm
    END IF
    
    IF SkipInactive AND UBCustRec(1).Status = "I" THEN
      UBBillRec(1).TransAmt = 0
      UBBillRec(1).ActiveFlag = False
      UBBillRec(1).CustAcctNo = LCnt&
      GOTO MSkipEm
    END IF
    
    IF BookFlag THEN
      IF QPValL(UBCustRec(1).Book) <> ThisBook THEN
        UBBillRec(1).TransAmt = 0
        FOR RCnt = 1 TO NumOfRevs
          UBBillRec(1).RevAmt(RCnt) = 0
          UBBillRec(1).TaxAmt(RCnt) = 0
        NEXT
        FOR zz = 1 TO 7
          UBBillRec(1).CurRead(zz) = 0
          UBBillRec(1).PrevRead(zz) = 0
        NEXT
        UBBillRec(1).ActiveFlag = False
        GOTO MSkipEm
      END IF
    END IF
    
    IF CycleFlag THEN
      IF UBCustRec(1).BILLCYCL <> ThisCycle THEN
        UBBillRec(1).TransAmt = 0
        FOR RCnt = 1 TO NumOfRevs
          UBBillRec(1).RevAmt(RCnt) = 0
          UBBillRec(1).TaxAmt(RCnt) = 0
        NEXT
        FOR zz = 1 TO 7
          UBBillRec(1).CurRead(zz) = UBCustRec(1).LocMeters(zz).CurRead
          UBBillRec(1).PrevRead(zz) = UBCustRec(1).LocMeters(zz).PrevRead
          UBBillRec(1).MtrTypes(zz) = GetCustMeterType(UBCustRec(), zz)
        NEXT
        UBBillRec(1).ActiveFlag = False
        GOTO MSkipEm
      END IF
    END IF
    
    IF UBCustRec(1).Status <> "A" THEN
      UBBillRec(1).TransAmt = 0
      FOR RCnt = 1 TO NumOfRevs
        UBBillRec(1).RevAmt(RCnt) = 0
        UBBillRec(1).TaxAmt(RCnt) = 0
      NEXT
      FOR zz = 1 TO 7
        UBBillRec(1).CurRead(zz) = UBCustRec(1).LocMeters(zz).CurRead
        UBBillRec(1).PrevRead(zz) = UBCustRec(1).LocMeters(zz).PrevRead
        UBBillRec(1).MtrTypes(zz) = GetCustMeterType(UBCustRec(), zz)
      NEXT
      UBBillRec(1).ActiveFlag = False
      GOTO MSkipEm
    END IF
    '052698 Added tax exempt flag to bill rec
    UBBillRec(1).TaxExempt = UBCustRec(1).TAXEXPT

'111398 Prorate
    ProrateFlag = False
    ProPct# = 100
    IF UBCustRec(1).ProRatePCT < 100 AND UBCustRec(1).ProRatePCT > 0 THEN
      UBBillRec(1).ProRatePCT = UBCustRec(1).ProRatePCT
      UBLog "MBF: Prorate Account No:" + STR$(LCnt&) + "  @" + QPTrim$(STR$(UBBillRec(1).ProRatePCT)) + "%"
      ProPct# = Round#(UBBillRec(1).ProRatePCT * .01)
      ProrateFlag = True
    ELSE
      UBBillRec(1).ProRatePCT = 100
    END IF

    MeterConsp& = 0
    TMeterConsp& = 0

    FOR FRCnt = 1 TO 4
      WhatService = UBCustRec(1).FlatRates(FRCnt).REVSRC
      IF WhatService > NumOfRevs THEN
        IF ErrorScrn(6, LCnt&) THEN
          AbortFlag = True
          GOTO AbortExit
        END IF
      END IF

      IF UBCustRec(1).FlatRates(FRCnt).FRAMT <> 0 AND WhatService > 0 THEN
        '11/19/96 Fixed Rev. amt. to add to current rev amt
        '01-09-97 Fixed Multiplier bug in flat rates
        IF UBCustRec(1).FlatRates(FRCnt).FRAMT < -1000000 THEN
          IF ErrorScrn(6, LCnt&) THEN
            AbortFlag = True
            GOTO AbortExit
          END IF
        END IF

        Multi = UBCustRec(1).FlatRates(FRCnt).NumMin
        IF Multi < 1 THEN Multi = 1
        FlatAmt# = Round#(UBCustRec(1).FlatRates(FRCnt).FRAMT * Multi)
'111398 Prorate
        IF ProrateFlag AND ProrateServ(WhatService) THEN
          FlatAmt# = Round#(FlatAmt# * ProPct#)
        END IF
        UBBillRec(1).RevAmt(WhatService) = Round#(UBBillRec(1).RevAmt(WhatService) + FlatAmt#)
        UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + FlatAmt#)
        IF UBSetUp(1).Revenues(WhatService).TAXRATE > 0 AND UBCustRec(1).TAXEXPT <> "Y" THEN
          TaxAmt# = Round#(UBBillRec(1).RevAmt(WhatService) * UBSetUp(1).Revenues(WhatService).TAXRATE)
          UBBillRec(1).TaxAmt(WhatService) = TaxAmt#
          UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + UBBillRec(1).TaxAmt(WhatService))
        END IF
      END IF
    NEXT
    
    '12-6-96  Monthly Billed amounts
    FOR MRCnt = 1 TO 2
      WhatService = UBCustRec(1).Monthly(MRCnt).RevSource
      IF WhatService > NumOfRevs OR WhatService < 0 THEN
        'IF ErrorScrn(7, LCnt&) THEN
        '  AbortFlag = True
        '  GOTO AbortExit
        'END IF
      END IF

      IF UBCustRec(1).Monthly(MRCnt).PayAmt > 0 AND WhatService > 0 THEN
        TestAmt# = Round#(UBCustRec(1).Monthly(MRCnt).TotAmtPD + UBCustRec(1).Monthly(MRCnt).PayAmt)
        IF TestAmt# > UBCustRec(1).Monthly(MRCnt).AMTOWED THEN
          HowMuch# = Round#(UBCustRec(1).Monthly(MRCnt).AMTOWED - UBCustRec(1).Monthly(MRCnt).TotAmtPD)
        ELSE
          HowMuch# = UBCustRec(1).Monthly(MRCnt).PayAmt
        END IF
        UBBillRec(1).RevAmt(WhatService) = Round#(UBBillRec(1).RevAmt(WhatService) + HowMuch#)
        UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + HowMuch#)
        IF UBSetUp(1).Revenues(WhatService).TAXRATE > 0 AND UBCustRec(1).TAXEXPT <> "Y" THEN
          TaxAmt# = Round#(HowMuch# * UBSetUp(1).Revenues(WhatService).TAXRATE)
          UBBillRec(1).TaxAmt(WhatService) = Round#(UBBillRec(1).TaxAmt(WhatService) + TaxAmt#)
          UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + TaxAmt#)
        END IF
      END IF
    NEXT
    
    FOR RCnt = 1 TO NumOfRevs   'look at each rev line
      MeterConsp& = 0
      TMeterConsp& = 0
      GOSUB GetWhatRateTable
      IF WhatTbl THEN
        IF UBSetUp(1).Revenues(RCnt).UseMtr = "N" THEN
          'if this is a non-metered service
          '02-05-97 added fix add to current rev amt
          IF UBRateTbls(WhatTbl).MINAMT > -1000000 THEN
            NonMAmt# = UBRateTbls(WhatTbl).MINAMT
            IF ProrateFlag AND ProrateServ(RCnt) THEN
              NonMAmt# = Round#(NonMAmt# * ProPct#)
            END IF
            UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + NonMAmt#)
            UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + NonMAmt#)
          ELSE
            RateCodeErrScrn UBRateTbls(WhatTbl).RATECODE
            AbortFlag = True
            GOTO AbortExit
          END IF
          GOTO GotAmt
        END IF
        'it's metered
        MeterType$ = UBCustRec(1).Serv(RCnt).RMtrType
        MeterLocNum = 0
        FOR MCnt = 1 TO 7
          IF MeterType$ = UBCustRec(1).LocMeters(MCnt).MTRType THEN
            MeterLocNum = MCnt
            UBBillRec(1).CurRead(MCnt) = UBCustRec(1).LocMeters(MCnt).CurRead
            UBBillRec(1).PrevRead(MCnt) = UBCustRec(1).LocMeters(MCnt).PrevRead
            UBBillRec(1).MtrTypes(MCnt) = GetCustMeterType(UBCustRec(), MCnt)
            'Found correct meter
            '052797 Added to stop overflow error.
            IF (UBCustRec(1).LocMeters(MCnt).CurRead < 0) OR (UBCustRec(1).LocMeters(MCnt).PrevRead < 0) THEN
              IF ErrorScrn(1, LCnt&) THEN
                AbortFlag = True
                GOTO AbortExit
              END IF
              MeterConsp& = 0
            ELSE
              MeterConsp& = UBCustRec(1).LocMeters(MCnt).CurRead - UBCustRec(1).LocMeters(MCnt).PrevRead
            END IF
            IF MeterConsp& < 0 THEN
              MaxMeterAmt& = 10& ^ (LEN(STR$(UBCustRec(1).LocMeters(MCnt).PrevRead)) - 1)
              MeterConsp& = (MaxMeterAmt& - UBCustRec(1).LocMeters(MCnt).PrevRead) + UBCustRec(1).LocMeters(MCnt).CurRead
            END IF
            IF UBCustRec(1).LocMeters(MCnt).MTRMulti > 0 THEN
              ThisMeterConsp# = (0# + MeterConsp&) * UBCustRec(1).LocMeters(MCnt).MTRMulti
              '                   ^This forces basic to convert to a Double
              '                    before calculation, traps overflow errors
              IF ThisMeterConsp# > 2147483647 THEN
                '                  ^Max long integer value
                IF ErrorScrn(1, LCnt&) THEN
                  AbortFlag = True
                  GOTO AbortExit
                END IF
              END IF
              MeterConsp& = ThisMeterConsp#
            END IF
            IF (UBBillRec(1).MtrTypes(MCnt) = 1 OR UBBillRec(1).MtrTypes(MCnt) = 2 OR UBBillRec(1).MtrTypes(MCnt) = 3) AND UBCustRec(1).LocMeters(MCnt).MTRUnit = "C" THEN
              MeterConsp& = MeterConsp& * 7.481
              'convert units from cubic feet to gallons here
            END IF
            TMeterConsp& = TMeterConsp& + MeterConsp&
          END IF
        NEXT
        IF MeterLocNum = 0 THEN
          IF ErrorScrn(4, LCnt&) THEN
            AbortFlag = True
            GOTO AbortExit
          END IF
        END IF
        AddRevAmt# = 0
        TMaxAmt# = 0
        IF UBRateTbls(WhatTbl).MaxAmt > 0 THEN
          TMaxAmt# = UBRateTbls(WhatTbl).MaxAmt
        END IF
        IF UBCustRec(1).LocMeters(MeterLocNum).NumUser > 1 THEN
          TMaxAmt# = Round#(UBRateTbls(WhatTbl).MaxAmt * UBCustRec(1).LocMeters(MeterLocNum).NumUser)
          'adjust min consumption for calc below
          NumUser& = UBCustRec(1).LocMeters(MeterLocNum).NumUser - 1
          AddRevAmt# = NumUser& * UBRateTbls(WhatTbl).MINAMT
          MinimumConsp& = NumUser& * UBRateTbls(WhatTbl).MINUNITS
          TMeterConsp& = TMeterConsp& - MinimumConsp&
          IF (TMeterConsp& - UBRateTbls(WhatTbl).MINUNITS) <= 0 THEN
            '062697 fix for min consump test to actual (NumUsers * MINUNITS)
            UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + (AddRevAmt# + UBRateTbls(WhatTbl).MINAMT))
            UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + (AddRevAmt# + UBRateTbls(WhatTbl).MINAMT))
            GOTO GotAmt
          END IF
        ELSE
          NumUser& = 1
          IF TMaxAmt# > 0 THEN
            TMaxAmt# = UBRateTbls(WhatTbl).MaxAmt
          END IF
          '033198 Added code to Calc correctly for Conway...
          IF ConwayFlag THEN
            IF TMeterConsp& MOD 1000 THEN
              TMeterConsp& = (INT(TMeterConsp& / 1000) + 1)
            ELSE
              TMeterConsp& = INT(TMeterConsp& / 1000)
            END IF
          END IF
          '033198 Conway *********
          '052998 Added code to calc correctly for Princeton
          IF PrinceFlag OR WadeFlag OR ScottFlag THEN
            IF TMeterConsp& MOD 1000 THEN
              TMeterConsp& = (INT(TMeterConsp& / 1000) + 1)
            ELSE
              TMeterConsp& = INT(TMeterConsp& / 1000)
            END IF
            TMeterConsp& = TMeterConsp& * 1000
          ELSEIF YadkinFlag THEN
            IF TMeterConsp& MOD 1000 THEN
              TMeterConsp& = TMeterConsp& / 1000
              TMeterConsp& = TMeterConsp& * 1000
            END IF
          END IF
          'Princeton*****
          IF TMeterConsp& <= UBRateTbls(WhatTbl).MINUNITS THEN
            'if we bill the minium
            IF UBRateTbls(WhatTbl).MINAMT > -1000000 THEN
              RevAmt# = NumUser& * UBRateTbls(WhatTbl).MINAMT
              IF ProrateFlag AND ProrateServ(RCnt) THEN
                RevAmt# = Round#(RevAmt# * ProPct#)
              END IF

              UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + RevAmt#)
              UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + RevAmt#)
              GOTO GotAmt
            END IF
          END IF
        END IF
'01-20-97 Added Fix for minium units test for equal to also.
'04-23-1997 'Fixed to ADD TO TOTAL
        RevAmt# = GetRevCharge#(UBRateTbls(WhatTbl), TMeterConsp&, MeterMulti&)
        RevAmt# = RevAmt# + AddRevAmt#
'111398 Prorate
        IF ProrateFlag AND ProrateServ(RCnt) THEN
          RevAmt# = Round#(RevAmt# * ProPct#)
        END IF
        IF TMaxAmt# > 0 THEN
          IF RevAmt# > TMaxAmt# THEN
            RevAmt# = TMaxAmt#
          END IF
        END IF
        UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + RevAmt#)
        IF RCnt = ElecRev THEN
          FuelAddAmt# = Round#(FuelAdjAmt# * TMeterConsp&)
          UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + FuelAddAmt#)
          UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + FuelAddAmt#)
        END IF
        UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + RevAmt#)
GotAmt:
        IF UBSetUp(1).Revenues(RCnt).TAXRATE > 0 AND UBCustRec(1).TAXEXPT <> "Y" THEN
          TaxAmt# = Round#(UBBillRec(1).RevAmt(RCnt) * UBSetUp(1).Revenues(RCnt).TAXRATE)
          UBBillRec(1).TaxAmt(RCnt) = Round#(UBBillRec(1).TaxAmt(RCnt) + TaxAmt#)
          UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + TaxAmt#)
        END IF
      ELSE
        IF LEN(QPTrim$(UBCustRec(1).Serv(RCnt).RMtrType)) > 0 THEN
          IF ErrorScrn(3, LCnt&) THEN
            AbortFlag = True
            GOTO AbortExit
          END IF
        END IF
      END IF
    NEXT        'loop through all revenue sources
    IF UBCustRec(1).Status = "I" AND UBBillRec(1).TransAmt > 0 THEN
      UBBillRec(1).TransAmt = 0
      FOR RCnt = 1 TO NumOfRevs
        UBBillRec(1).RevAmt(RCnt) = 0
      NEXT
      UBBillRec(1).ActiveFlag = False
      UBBillRec(1).CustAcctNo = LCnt&
    ELSE
      UBBillRec(1).ActiveFlag = False
    END IF
    
    'Mod for cleveland***
    IF UBCustRec(1).CustType = "NON" THEN
      UBBillRec(1).CustAcctNo = LCnt&
      UBBillRec(1).NONProfit = "Y"
    END IF
    '********************
    
    IF UBBillRec(1).TransAmt > 0 THEN
      BillCnt = BillCnt + 1
      UBBillRec(1).ActiveFlag = True
      UBBillRec(1).CustAcctNo = LCnt&
    END IF
    
'072799 Added NEW trap for a meter defined with no rate code.
'    FOR MTstCnt = 1 TO 7
'      IF LEN(QPTrim$(UBCustRec(1).LocMeters(MTstCnt).MTRType)) > 0 THEN
'        FOR MTCnt = 1 TO 7
'          IF UBBillRec(1).MtrTypes(MTCnt) > 0 THEN
'            GOTO ThereOK
'          END IF
'          IF ErrorScrn(8, LCnt&) THEN
'            AbortFlag = True
'            GOTO AbortExit
'          END IF
'        NEXT
'      END IF
'    NEXT

'04-07-99 Added special tax calc for elkton
    IF ElkFlag THEN
      IF UBBillRec(1).ActiveFlag THEN
        FOR TZCnt = 1 TO 15
          IF UBBillRec(1).TaxAmt(TZCnt) > 0 THEN
            CType$ = QPTrim$(UBCustRec(1).CustType)
            SELECT CASE CType$
            CASE "R"
              IF UBBillRec(1).TaxAmt(TZCnt) > 2 THEN
                Diff# = Round#(UBBillRec(1).TaxAmt(TZCnt) - 2)
                UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt - Diff#)
                UBBillRec(1).TaxAmt(TZCnt) = 2
              END IF
            CASE "C"
              IF UBBillRec(1).TaxAmt(TZCnt) > 20 THEN
                Diff# = Round#(UBBillRec(1).TaxAmt(TZCnt) - 20)
                UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt - Diff#)
                UBBillRec(1).TaxAmt(TZCnt) = 20
              END IF
            CASE ELSE
              IF ErrorScrn(9, LCnt&) THEN
                AbortFlag = True
                GOTO AbortExit
              END IF
            END SELECT
          END IF
        NEXT
      END IF
    END IF
'end of elkton section

MSkipEm:
   PUT BillFile, LCnt&, UBBillRec(1)
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
    ShowPctComp LCnt&, NumCustRec&
  NEXT
  
AbortExit:
  
  CLOSE BillFile, CustFile
  
  IF AbortFlag THEN
    UBLog "MBF: ABORTED!"
  ELSE
    UBLog "MBF: Finished calculations."
  END IF

  UBLog "OUT: MakeBillFile."
  ERASE UBBillRec, UBCustRec, UBSetUp, UBRateTbls
  EXIT SUB
'*******************************
  
GetWhatRateTable:
  WhatTbl = 0
  RATECODE$ = QPTrim$(UBCustRec(1).Serv(RCnt).RATECODE)
  IF LEN(RATECODE$) THEN        'if this rev has a rate code
    FOR TCnt = 1 TO NumOfRates  'find the right one
      IF RATECODE$ = QPTrim$(UBRateTbls(TCnt).RATECODE) THEN
        WhatTbl = TCnt
        EXIT FOR
      END IF
    NEXT
  END IF
RETURN
  
END SUB

SUB PostBillTrans
  
  UBLog "IN: Bill Posting."
  
  IF NOT Exist(UBBillsFile) THEN
    UBLog "ERROR: UBBILLS.DAT Calculation file NOT FOUND!"
    CursorOff
    BlockClear
    DisplayUBScrn "NON2POST"
    WaitForAction
    GOTO ExitBillPost:
  END IF
  
  IF NOT Exist("UBBILLS.PRN") THEN
    UBLog "ERROR: UBBILLS.PRN Print File NOT FOUND!"
    CursorOff
    BlockClear
    DisplayUBScrn "NOTPRNTD"
    WaitForAction
    GOTO ExitBillPost:
  END IF
  
  LibName$ = "UB"
  ScrnName$ = "PSTBILLS"
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen      'load setup file
  
  TownName$ = UBSetUpRec(1).UTILNAME
  
  'Section to check for customer modifications
  'Town of Lilesville Special Discount Situation
  
  IF INSTR(TownName$, "INDIAN TRAIL") THEN
    IndianFlag = True
  END IF
  
  '--define the multi-choice fields
  '--Initialize the form name array
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  Action = 1
  Frm(1).StayOnField = True
  
  '--Set screen number to one and display screen
  
  BlockClear
  
  DisplayUBScrn ScrnName$

  IF FileSize&("UBSNDEM.DAT") > 0 THEN
    FOR Cnt = 1 TO 3
      QPSound 1750, 2
      QPSound 1650, 2
    NEXT
  END IF
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      OKFlag = True
    CASE EscKey
      ExitFlag = True
    END SELECT
    
  LOOP UNTIL OKFlag OR ExitFlag
  
  IF ExitFlag THEN
    BlockClear
    CursorOff
    DisplayUBScrn "POSTCAN"
    WaitForAction
    UBLog "ABORTED:  Bill Posting"
    GOTO ExitBillPost
  END IF
  
  UBLog "START: Posting Transactions."
  
  REDIM UBBillRec(1) AS UBTransRecType
  REDIM UBCustRec(1) AS NewUBCustRecType
  
  UBBillRecLen = LEN(UBBillRec(1))
  UBCustRecLen = LEN(UBCustRec(1))
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  UBBill = FREEFILE
  OPEN UBBillsFile FOR RANDOM SHARED AS UBBill LEN = UBBillRecLen
  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBBillRecLen
  
  NumOfTranRecs& = LOF(UBTran) \ UBBillRecLen
  NumOfBillRecs = LOF(UBBill) \ UBBillRecLen
  
  ShowProcessingScrn "Posting Billing Transactions"
  FOR BillCnt = 1 TO NumOfBillRecs
    GET UBBill, BillCnt, UBBillRec(1)
    IF (UBBillRec(1).ActiveFlag AND UBBillRec(1).TransAmt > 0) OR (UBBillRec(1).NONProfit = "Y") THEN
      PostedCnt& = PostedCnt& + 1
      NumOfTranRecs& = NumOfTranRecs& + 1       'point to next trans to write
      GET UBCust, BillCnt, UBCustRec(1)
      EstFlag$ = QPTrim$(UBCustRec(1).EstFlag)
      FOR MRCnt = 1 TO 2
        WhatService = UBCustRec(1).Monthly(MRCnt).RevSource
        IF UBCustRec(1).Monthly(MRCnt).PayAmt > 0 AND WhatService > 0 THEN
          TestAmt# = Round#(UBCustRec(1).Monthly(MRCnt).TotAmtPD + UBCustRec(1).Monthly(MRCnt).PayAmt)
          IF TestAmt# > UBCustRec(1).Monthly(MRCnt).AMTOWED THEN
            HowMuch# = Round#(UBCustRec(1).Monthly(MRCnt).AMTOWED - UBCustRec(1).Monthly(MRCnt).TotAmtPD)
          ELSE
            HowMuch# = UBCustRec(1).Monthly(MRCnt).PayAmt
          END IF
          UBCustRec(1).Monthly(MRCnt).TotAmtPD = Round#(UBCustRec(1).Monthly(MRCnt).TotAmtPD + HowMuch#)
        END IF
      NEXT
      '062597 added removal of nonrecurring flat rates
      FRFlag = False
      FOR FRCnt = 1 TO 4        'Remove non-recurring flat rates
        IF UBCustRec(1).FlatRates(FRCnt).FRFREQ = "N" THEN
          UBCustRec(1).FlatRates(FRCnt).FRDESC = ""
          UBCustRec(1).FlatRates(FRCnt).FRAMT = 0
          UBCustRec(1).FlatRates(FRCnt).FRFREQ = ""
          UBCustRec(1).FlatRates(FRCnt).REVSRC = 0
          UBCustRec(1).FlatRates(FRCnt).NumMin = 0
          FRFlag = True
        END IF
      NEXT
      IF FRFlag THEN
        UBLog "BILL POST: Removed Flat Rate. Acct:" + STR$(BillCnt)
      END IF
'111698 Prorate
      IF UBBillRec(1).ProRatePCT < 100 THEN
        UBLog "BILL POST: Reset Prorate Acct:" + STR$(BillCnt) + " PCT:" + STR$(UBBillRec(1).ProRatePCT)
      END IF
      UBCustRec(1).ProRatePCT = 100
'*************
      UBCustRec(1).PrevBalance = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
      UBCustRec(1).CurrBalance = UBBillRec(1).TransAmt
      UBBillRec(1).RunBalance = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
      FOR RevCnt = 1 TO MaxRevsCnt
        UBCustRec(1).CurrRevAmts(RevCnt) = Round#(UBCustRec(1).CurrRevAmts(RevCnt) + UBBillRec(1).RevAmt(RevCnt) + UBBillRec(1).TaxAmt(RevCnt))
      NEXT
      UBBillRec(1).TransType = TranUtilityBill  'set transaction to Type 1
      UBBillRec(1).TransDesc = "Utility Billing"
      UBBillRec(1).TransDate = UBBillRec(1).BillDate

      FOR MtrCnt = 1 TO 7
        CubMtr = False
        IF UBCustRec(1).LocMeters(MtrCnt).CurRead >= 0 THEN
          IF LEN(EstFlag$) > 0 THEN
            UBBillRec(1).EstRead(MtrCnt) = "Y"
          END IF
          IF UBCustRec(1).LocMeters(MtrCnt).MTRUnit = "C" THEN
            CubMtr = True
          END IF
          ReadAmt& = UBBillRec(1).CurRead(MtrCnt) - UBBillRec(1).PrevRead(MtrCnt)
          IF ReadAmt& < 0 THEN  'Meter rolled over or, been misread
            MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(MtrCnt))) - 1)
            ReadAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(MtrCnt)) + UBBillRec(1).CurRead(MtrCnt)
          END IF
          IF CubMtr THEN
            ReadAmt& = ReadAmt& * 7.481
          END IF
          IF ReadAmt& < 1 THEN
            ReadAmt& = 1
          END IF
          IF UBCustRec(1).LocMeters(MtrCnt).AvgUse < 1 THEN
            UBCustRec(1).LocMeters(MtrCnt).AvgUse = 1
          END IF
          IF UBCustRec(1).LocMeters(MtrCnt).UseCnt < 1 THEN
            UBCustRec(1).LocMeters(MtrCnt).UseCnt = 1
          END IF
          TUse# = ReadAmt& + (UBCustRec(1).LocMeters(MtrCnt).AvgUse * UBCustRec(1).LocMeters(MtrCnt).UseCnt)
          UBCustRec(1).LocMeters(MtrCnt).UseCnt = UBCustRec(1).LocMeters(MtrCnt).UseCnt + 1
          UBCustRec(1).LocMeters(MtrCnt).AvgUse = TUse# / UBCustRec(1).LocMeters(MtrCnt).UseCnt
          UBCustRec(1).LocMeters(MtrCnt).ReadFlag = ""
        END IF
      NEXT
      PrevLastTrans& = UBCustRec(1).LastTrans
      UBBillRec(1).PrevTrans = PrevLastTrans&
      UBCustRec(1).LastTrans = NumOfTranRecs&
      IF IndianFlag THEN
        UBCustRec(1).USERCODE1 = ""
      END IF
'DALE
      PUT UBCust, BillCnt, UBCustRec(1)
      PUT UBTran, NumOfTranRecs&, UBBillRec(1)
'**************
    END IF
    ShowPctComp BillCnt, NumOfBillRecs
  NEXT
  CLOSE
  UBLog "  DONE: Posting Transactions."
  UBLog "POSTED:" + STR$(PostedCnt&) + " New BILL Transactions."
'DALE
  KillFile UBBillsFile
  KillFile "UBBILLS.PRN"
'**************
  UBLog "KILLED: UBBILLS.DAT & UBBILLS.PRN"

  ShowProcessingScrn "Activating Pending Accounts."

  UBLog "ACTIVATING ACCOUNTS:"

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  NumOfCust& = LOF(UBCust) / UBCustRecLen
  FOR Cnt = 1 TO NumOfCust&
    GET UBCust, Cnt, UBCustRec(1)
    IF UBCustRec(1).Status = "P" THEN
      UBCustRec(1).Status = "A"
      UBLog "ACTIVATED: " + STR$(Cnt) + "  " + UBCustRec(1).CustName
      Activated = Activated + 1
      PUT UBCust, Cnt, UBCustRec(1)
    END IF
    ShowPctComp Cnt, CINT(NumOfCust&)
  NEXT
  CLOSE
  UBLog "     DONE: Activating Accounts."
  UBLog "ACTIVATED:" + STR$(Activated) + " Pending Accounts."
  BlockClear
  DisplayUBScrn "UPDATEOK"
  WaitForAction
  
ExitBillPost:
  UBLog "OUT: Bill Posting." + CrLf$
END SUB

SUB PreBillReport
  
  UBLog "IN: Prebilling Report"

  Temp2$ = SPACE$(12)
  NumOfRevs = MaxRevsCnt        'assume max munber of revenue sources
  
  NumOfRates = GetNumRateRecs%
  REDIM UBRateTbls(1 TO NumOfRates) AS UBRateTblRecType
  UBRateTblRecLen = LEN(UBRateTbls(1))
  
  REDIM RateConsump(1 TO NumOfRates) AS DOUBLE
  
  RateFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS RateFile LEN = UBRateTblRecLen
  FOR Cnt = 1 TO NumOfRates
    GET RateFile, Cnt, UBRateTbls(Cnt)
  NEXT
  CLOSE
  
  SortT UBRateTbls(1), NumOfRates, 0, UBRateTblRecLen, 0, 4

  REDIM ProrateServ(1 TO 15) AS INTEGER

  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen
  
  TownName$ = UBSetUpRec(1).UTILNAME
  IF INSTR(TownName$, "MOWAS") > 0 THEN
    MowFlag = True
  END IF
  
  IF UBSetUpRec(1).DefState = "TN" THEN
    TennFlag = True
  END IF

  REDIM RevDesc(1 TO MaxRevsCnt) AS STRING * 12
  FOR Cnt = 1 TO MaxRevsCnt     'find last active revenue
    TempRev$ = QPTrim$(UBSetUpRec(1).Revenues(Cnt).REVNAME)
    IF LEN(TempRev$) = 0 THEN
      NumOfRevs = Cnt - 1       'set actual number of revenues
      EXIT FOR
    ELSE        'build revenue description lines
      LSET RevDesc(Cnt) = UCASE$(TempRev$)
      IF INSTR(RevDesc(Cnt), "ELECTRIC") THEN
        DoFuelAdjFlag = True
      END IF
    END IF
  NEXT

'111398 Prorate
  FOR Cnt = 1 TO MaxRevsCnt
    IF UBSetUpRec(1).Revenues(Cnt).ProRate = "Y" THEN
      ProrateServ(Cnt) = True
    END IF
  NEXT
  
  IF UBSetUpRec(1).SkipInactive = "Y" THEN
    SkipInactive = True
  END IF
  
  IF UBSetUpRec(1).SkipSeparator = "Y" THEN
    SkipSeparator = True
  END IF
  
  IF UBSetUpRec(1).PreByBook = "Y" THEN
    ThisBook = GetBillBook%
    IF ThisBook = -1 THEN
      BookFlag = False
    ELSEIF ThisBook <= 0 THEN
      GOTO ExitPreReport
    ELSE
      BookFlag = True
    END IF
  ELSEIF UBSetUpRec(1).BILLCYCL = "Y" THEN
    ThisCycle = GetBillCycle%
    IF ThisCycle <= 0 THEN
      GOTO ExitPreReport
    ELSE
      CycleFlag = True
    END IF
  END IF
  
  IF UBSetUpRec(1).UseSeq = "Y" THEN
    SeqFlag$ = "Y"
  END IF
  
Restart:
  
  GetPreBillOrder Choice, ExitFlag, SeqFlag$

  IF ExitFlag THEN GOTO ExitPreReport
  
  IF DoFuelAdjFlag THEN
    FuelAdjAmt# = GetAdjFactor#
    UBLog "Fuel adjustment factor:" + STR$(FuelAdjAmt#)
  ELSE
    FuelAdjAmt# = 0
  END IF
  
  IF FuelAdjAmt# = -10000 THEN GOTO Restart
  
  SELECT CASE Choice
  CASE 0
    ExitFlag = True
  CASE 1        'Name
    IndexName$ = NameIndexFile
    OKFlag = True
  CASE 2        'Acct
    IndexName$ = ""
    UsingAcct = True
    OKFlag = True
  CASE 3        'Location
    IndexName$ = BookIndexFile
    OKFlag = True
  CASE 4        'Postal Route
    IdxTypeText$ = "Postal Route"
    MakePostalIndex IdxTypeText$
    IndexName$ = TempIndexName
    OKFlag = True
  CASE 5        'ZipCode
    IdxTypeText$ = "Zip-Code"
    IF MowFlag THEN
      MakeMowZipCodeIndex IdxTypeText$
    ELSE
      MakeZipCodeIndex IdxTypeText$
    END IF
    IndexName$ = TempIndexName
    OKFlag = True
  CASE 6        'Sequence number
    IdxTypeText$ = "Sequence Number"
    MakeSequenceIndex IdxTypeText$
    IndexName$ = TempIndexName
    OKFlag = True
  END SELECT
  
  MakeBillFile AbortFlag, FuelAdjAmt#, ThisCycle, ThisBook
  
  IF AbortFlag THEN GOTO ExitPreReport
  
  MaxLines = 53
  
  REDIM Fmt$(0 TO 6)
  Fmt$(0) = STRING$(80, "-")
  Fmt$(1) = "#########.##"
  Fmt$(2) = "#########"
  Fmt$(3) = "######.##"
  Fmt$(4) = "###########"
  Fmt$(5) = "$$,#########.##"
  Fmt$(6) = "$$#######,.##"
  
  TheDate$ = "Date: " + DATE$
  
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBBillRec(1) AS UBTransRecType
  UBBillRecLen = LEN(UBBillRec(1))
  
  REDIM FlatTotals(1 TO NumOfRevs) AS DOUBLE
'021998 added flat revenue totals
  REDIM RevTotals(1 TO NumOfRevs) AS DOUBLE     'Revenue total amts
'052097 added tax by revenue totals
  REDIM TaxTotals(1 TO NumOfRevs) AS DOUBLE     'Tax total amts
  REDIM ConsumpTot(1 TO NumOfRevs, 1 TO 2) AS DOUBLE  'Consumption total amts
  REDIM RateConsump(1 TO NumOfRates) AS DOUBLE
'012698 Added bill count by rate code
  REDIM RateCount(1 TO NumOfRates) AS LONG
  REDIM RateTotals(1 TO NumOfRates) AS DOUBLE   'Rates total amts
'052097 added tax by rate code totals
  REDIM RTaxTot(1 TO NumOfRates) AS DOUBLE      'Rates Tax total amts
'052097 added tax by book totals to type def
  REDIM BookConsump(0 TO 1) AS BookConsumpType  'Consumption by book
  REDIM PumpConsump(0 TO 1) AS PumpConsumpType  'Consumption by pump code
  REDIM TaxExmp(0 TO NumOfRevs) AS DOUBLE
  
  TBooks = 0
  
  IF UsingAcct THEN
    NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
  ELSE          'load the index
    UBLog "Loading index file: " + IndexName$
    NumOfRecs = FileSize(IndexName$) \ 4
    REDIM IndexArray(1 TO NumOfRecs) AS UBCustIndexRecType
    FGetAH IndexName$, IndexArray(1), 4, NumOfRecs
  END IF
  
  UBBill = FREEFILE
  OPEN UBBillsFile FOR RANDOM SHARED AS UBBill LEN = UBBillRecLen
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  UBRpt = FREEFILE
  OPEN "UBPREBIL.RPT" FOR OUTPUT AS UBRpt
  
  BlockClear
  ShowProcessingScrn "Processing Pre-Billing Report"
  UBLog "Writing prebilling report to disk."
  
  GOSUB PrintPreHeader
  
  FOR Cnt = 1 TO NumOfRecs
    IF UsingAcct THEN
      ThisCustRec& = Cnt
    ELSE
      ThisCustRec& = IndexArray(Cnt).RecNum
    END IF

    GET UBCust, ThisCustRec&, UBCustRec(1)

    IF UBCustRec(1).DelFlag THEN
      GOTO SkipEM
    END IF
    
    IF SkipInactive AND UBCustRec(1).Status <> "A" THEN
      GOTO SkipEM
    ELSEIF UBCustRec(1).Status = "F" THEN       'skip over final's
      GOTO SkipEM
    ELSEIF UBCustRec(1).Status = "B" THEN       'skip over B-Status
      GOTO SkipEM
    END IF
    
    IF BookFlag THEN
      IF QPValL(UBCustRec(1).Book) <> ThisBook THEN
        GOTO SkipEM
      END IF
    END IF
    
    IF CycleFlag THEN
      IF UBCustRec(1).BILLCYCL <> ThisCycle THEN
        GOTO SkipEM
      END IF
    END IF
    
    GET UBBill, ThisCustRec&, UBBillRec(1)
    
    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, FF$
      GOSUB PrintPreHeader
    END IF
    
    IF UBBillRec(1).ActiveFlag <> 0 THEN
      IF UBCustRec(1).BillTo = "O" THEN
        BillTo$ = " O"
      ELSE
        BillTo$ = " C"
      END IF
      GOSUB GetWhatBook
      IF BadBookFlag THEN
        IF ErrorScrn(2, ThisCustRec&) THEN
          AbortFlag = True
          EXIT FOR
        END IF
      END IF
      BookConsump(WhatBook).CustCnt = BookConsump(WhatBook).CustCnt + 1
      PRINT #UBRpt, UBCustRec(1).Status; USING "  #####  "; ThisCustRec&;
      PRINT #UBRpt, UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB; "  "; LEFT$(UBCustRec(1).CustName, 25); " "; LEFT$(UBCustRec(1).SERVADDR, 22); " ";
      PRINT #UBRpt, USING "   ###%"; UBBillRec(1).ProRatePCT;
      PRINT #UBRpt, BillTo$
      LineCnt = LineCnt + 1
      FOR FRCnt = 1 TO 4
        WhatService = UBCustRec(1).FlatRates(FRCnt).REVSRC
        IF UBCustRec(1).FlatRates(FRCnt).FRAMT <> 0 AND WhatService > 0 THEN
          Multi = UBCustRec(1).FlatRates(FRCnt).NumMin
          IF Multi < 1 THEN Multi = 1
          FlatAmt# = Round#(UBCustRec(1).FlatRates(FRCnt).FRAMT * Multi)
          '021998 Added flat rate summaries
          FlatTotals(WhatService) = Round#(FlatTotals(WhatService) + FlatAmt#)
        END IF
      NEXT
      '102798 Added to skip accts that don't have a book/seq no. "J.R."
    ELSEIF LEN(QPTrim$(UBCustRec(1).Book)) = 0 AND LEN(QPTrim$(UBCustRec(1).SEQNUMB)) = 0 THEN
      GOTO SkipEM
    END IF
    WhatRate = 0
    DoneOne = False
    FOR TRevCnt = 1 TO NumOfRevs
      WhatRate = 0
      IF UBBillRec(1).RevAmt(TRevCnt) <> 0 THEN
        DoneOne = False
        PRINT #UBRpt, RevDesc(TRevCnt);
        '102198 Moved from meter loop, Stops multi meter tax report bug
        IF UBBillRec(1).TaxAmt(TRevCnt) > 0 THEN
          TaxTotals(TRevCnt) = Round#(TaxTotals(TRevCnt) + UBBillRec(1).TaxAmt(TRevCnt))
        END IF
        FOR TRateCnt = 1 TO NumOfRates
          IF UBRateTbls(TRateCnt).RATECODE = UBCustRec(1).Serv(TRevCnt).RATECODE THEN
            MINAMT& = UBRateTbls(TRateCnt).MINUNITS
            WhatRate = TRateCnt
            '102198 Moved from meter loop, Stops multi meter tax report bug
            RTaxTot(WhatRate) = Round#(RTaxTot(WhatRate) + UBBillRec(1).TaxAmt(TRevCnt))
            EXIT FOR
          END IF
        NEXT
        IF UBSetUpRec(1).Revenues(TRevCnt).UseMtr = "Y" THEN
          RevTotals(TRevCnt) = Round#(RevTotals(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
          BookConsump(WhatBook).RevAmt(TRevCnt) = Round#(BookConsump(WhatBook).RevAmt(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
          BookConsump(WhatBook).TaxAmt(TRevCnt) = Round#(BookConsump(WhatBook).TaxAmt(TRevCnt) + UBBillRec(1).TaxAmt(TRevCnt))
          '02-20-97 Add revenue totals by rate code
          IF WhatRate > 0 THEN
            RateTotals(WhatRate) = Round#(RateTotals(WhatRate) + UBBillRec(1).RevAmt(TRevCnt))
          END IF
          PrintedRevAmt = False
          FOR MCCnt = 1 TO 7
            CubMtr = False
            LocMeterType$ = QPTrim$(UBCustRec(1).LocMeters(MCCnt).MTRType)
            MeterMulti& = UBCustRec(1).LocMeters(MCCnt).MTRMulti
            '063098 Added adjustment for cubic meters in consumption totals
            IF UBCustRec(1).LocMeters(MCCnt).MTRUnit = "C" THEN
              CubMtr = True
            END IF
            IF MeterMulti& <= 0 THEN MeterMulti& = 1
            IF (LEN(LocMeterType$) > 0) AND (LocMeterType$ = UBCustRec(1).Serv(TRevCnt).RMtrType) THEN
              DoneOne = True
              MeterNum$ = QPTrim$(UBCustRec(1).Serv(TRevCnt).RATECODE)
              'use the Meternum$ to hold the rate code temporarily
              IF LEN(MeterNum$) > 0 THEN
                IF UBCustRec(1).LocMeters(MCCnt).NumUser > 1 THEN
                  MeterNum$ = MeterNum$ + "*" + QPTrim$(STR$(UBCustRec(1).LocMeters(MCCnt).NumUser))
                END IF
                RSET Temp2$ = MeterNum$
              END IF
              ReadAmt& = UBBillRec(1).CurRead(MCCnt) - UBBillRec(1).PrevRead(MCCnt)
              IF ReadAmt& < 0 THEN   'Meter rolled over or, been misread
                MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(MCCnt))) - 1)
                ReadAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(MCCnt)) + UBBillRec(1).CurRead(MCCnt)
              END IF
              IF CubMtr THEN
                ReadAmt& = ReadAmt& * 7.481
              END IF
              RateConsump(WhatRate) = RateConsump(WhatRate) + (ReadAmt& * MeterMulti&)
              RateCount(WhatRate) = RateCount(WhatRate) + 1
              BookConsump(WhatBook).Consump(TRevCnt) = BookConsump(WhatBook).Consump(TRevCnt) + (ReadAmt& * MeterMulti&)
              ConsumpTot(TRevCnt, 1) = ConsumpTot(TRevCnt, 1) + (ReadAmt& * MeterMulti&)
              Consump& = ReadAmt& * MeterMulti&
              ThisMeterUseCnt = UBCustRec(1).LocMeters(MCCnt).UseCnt
              IF ThisMeterUseCnt <= 0 THEN ThisMeterUseCnt = 1
              AvgUse& = UBCustRec(1).LocMeters(MCCnt).AvgUse
              IF AvgUse& > 0 THEN
                HiConsump& = Round#(AvgUse& * (UBSetUpRec(1).HighRead * .01))
                LowConsump& = Round#(AvgUse& * (UBSetUpRec(1).LowRead * .01))
              END IF
              PRINT #UBRpt, TAB(14); Temp2$; TAB(30); USING Fmt$(2); UBBillRec(1).CurRead(MCCnt); TAB(42); UBBillRec(1).PrevRead(MCCnt); TAB(54); ReadAmt& * MeterMulti&;
              IF UBCustRec(1).EstFlag = "E" THEN
                PRINT #UBRpt, " E";             'Est. Reading
              ELSEIF Consump& < LowConsump& THEN
                PRINT #UBRpt, " L";             'Low reading
              ELSEIF Consump& > HiConsump& THEN
                PRINT #UBRpt, " H";             'High Reading
              END IF
              IF Consump& < MINAMT& THEN
                PRINT #UBRpt, " M";             'Minium Usage
              END IF
              IF UBBillRec(1).RevAmt(TRevCnt) > 0 AND PrintedRevAmt = False THEN
                PrintedRevAmt = True
                PRINT #UBRpt, TAB(69); USING Fmt$(3); UBBillRec(1).RevAmt(TRevCnt);
                IF UBBillRec(1).ProRatePCT < 100 AND ProrateServ(TRevCnt) THEN
                  PRINT #UBRpt, "*";
                END IF
              END IF
              PRINT #UBRpt,
              LineCnt = LineCnt + 1
            END IF
          NEXT
'071197 Added this for mccormick. Has a sewer flat rate, Sewer is set up as
'       a metered service but no meter on a flat rate charge. Rev was added
'       to total, but didn't show on prebilling report.
          IF NOT DoneOne THEN
            DoneOne = True
            PRINT #UBRpt, TAB(69); USING Fmt$(3); UBBillRec(1).RevAmt(TRevCnt);
            IF UBBillRec(1).ProRatePCT < 100 AND ProrateServ(TRevCnt) THEN
              PRINT #UBRpt, "*";
            END IF
'THIS WAS REMARKED OUT DON'T KNOW WHY?
            PRINT #UBRpt,
''''''''''''''''''''''''''''''''''''''
            LineCnt = LineCnt + 1
          END IF
        ELSE    'it's a nonmetered service
          ConsumpTot(TRevCnt, 1) = ConsumpTot(TRevCnt, 1) + 1
          IF WhatRate > 0 THEN
            RateConsump(WhatRate) = RateConsump(WhatRate) + 1
            RateCount(WhatRate) = RateCount(WhatRate) + 1
            RateTotals(WhatRate) = Round#(RateTotals(WhatRate) + UBBillRec(1).RevAmt(TRevCnt))
          END IF
          BookConsump(WhatBook).Consump(TRevCnt) = BookConsump(WhatBook).Consump(TRevCnt) + 1
          BookConsump(WhatBook).RevAmt(TRevCnt) = Round#(BookConsump(WhatBook).RevAmt(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
          BookConsump(WhatBook).TaxAmt(TRevCnt) = Round#(BookConsump(WhatBook).TaxAmt(TRevCnt) + UBBillRec(1).TaxAmt(TRevCnt))
          RevTotals(TRevCnt) = Round#(RevTotals(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
          PRINT #UBRpt, TAB(69); USING Fmt$(3); UBBillRec(1).RevAmt(TRevCnt);
          IF UBBillRec(1).ProRatePCT < 100 AND ProrateServ(TRevCnt) THEN
            PRINT #UBRpt, "*";
          END IF
        END IF
        IF NOT DoneOne THEN
          PRINT #UBRpt,
          LineCnt = LineCnt + 1
        END IF
      END IF
      IF (TRevCnt = NumOfRevs) AND UBBillRec(1).TransAmt = 0 THEN
        IF UBBillRec(1).TransAmt = 0 THEN       'CONSUMPTION inactive account
          FOR TTRevCnt = 1 TO NumOfRevs
            FOR MCCnt = 1 TO 7
              LocMeterType$ = QPTrim$(UBCustRec(1).LocMeters(MCCnt).MTRType)
              IF (LEN(LocMeterType$) > 0) AND (LocMeterType$ = UBCustRec(1).Serv(TTRevCnt).RMtrType) THEN
                IF UBBillRec(1).CurRead(MCCnt) < 0 THEN
                  UBBillRec(1).CurRead(MCCnt) = 0
                END IF
                IF UBBillRec(1).PrevRead(MCCnt) < 0 THEN
                  UBBillRec(1).PrevRead(MCCnt) = 0
                END IF
                CurReadAmt& = UBBillRec(1).CurRead(MCCnt)
                PreReadAmt& = UBBillRec(1).PrevRead(MCCnt)
                IF CurReadAmt& <> PreReadAmt& THEN
                  IF NOT ConsumpFlag THEN
                    PRINT #UBRpt, UBCustRec(1).Status; USING "     #####   "; ThisCustRec&;
                    PRINT #UBRpt, UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB; "   "; LEFT$(UBCustRec(1).CustName, 25); "  "; LEFT$(UBCustRec(1).SERVADDR, 25)
                    LineCnt = LineCnt + 1
                  END IF
                  ConsumpFlag = True
                  MeterNum$ = QPTrim$(UBCustRec(1).Serv(TTRevCnt).RATECODE)
                  IF LEN(MeterNum$) > 0 THEN
                    IF UBCustRec(1).LocMeters(MCCnt).NumUser > 1 THEN
                      MeterNum$ = MeterNum$ + "*" + QPTrim$(STR$(UBCustRec(1).LocMeters(MCCnt).NumUser))
                    END IF
                    RSET Temp2$ = MeterNum$
                  END IF
                  ConsumpAmt& = CurReadAmt& - PreReadAmt&
                  '103098 Added meter roll over check to inactive consumption
                  IF ConsumpAmt& < 0 THEN       'Meter rolled over or, been misread
                    MaxMeterAmt& = 10& ^ (LEN(STR$(PreReadAmt&)) - 1)
                    ConsumpAmt& = (MaxMeterAmt& - PreReadAmt&) + CurReadAmt&
                  END IF
                  IF INSTR(UBCustRec(1).CustType, "NON") THEN
                    'For Nonprofits include consumption as normal   'cleveland
                    '040998 Made changes here
                    FOR NONRateCnt = 1 TO NumOfRates
                      IF UBRateTbls(NONRateCnt).RATECODE = UBCustRec(1).Serv(TTRevCnt).RATECODE THEN
                        NONRate = NONRateCnt
                        EXIT FOR
                      END IF
                    NEXT
                    IF NONRate > 0 THEN
                      RateConsump(NONRate) = RateConsump(NONRate) + ConsumpAmt&
                    END IF
                    ConsumpTot(TTRevCnt, 1) = ConsumpTot(TTRevCnt, 1) + ConsumpAmt&
                    BookConsump(WhatBook).Consump(TTRevCnt) = BookConsump(WhatBook).Consump(TTRevCnt) + ConsumpAmt&
                    '040998 Made changes here 'cleveland
                  ELSE          'add consumption to inactives
                    ConsumpTot(TTRevCnt, 2) = ConsumpTot(TTRevCnt, 2) + ConsumpAmt&
                  END IF
                  PRINT #UBRpt, RevDesc(TTRevCnt); TAB(14); Temp2$; TAB(30); USING Fmt$(2); CurReadAmt&; TAB(42); PreReadAmt&; TAB(54); ConsumpAmt&
                  LineCnt = LineCnt + 1
                END IF
              END IF
            NEXT
          NEXT
        END IF
        IF ConsumpFlag AND UBCustRec(1).Status <> "A" THEN
          ConsumpFlag = False
          PRINT #UBRpt, "**** Consumption Noted on an Inactive Account. ****"
          LineCnt = LineCnt + 1
          IF NOT SkipSeparator THEN
            PRINT #UBRpt, Fmt$(0)
            LineCnt = LineCnt + 1
          END IF
        ELSEIF ConsumpFlag THEN
          'Customer Status is "A"
          'This happens when a cust has consumption and there rate code
          'has a zero calc amount. "i.e. a Church or other nonprofit"
          IF INSTR(UBCustRec(1).CustType, "NON") THEN
            PRINT #UBRpt, "*** NON-PROFIT ***"
            LineCnt = LineCnt + 1
          END IF
          ConsumpFlag = False
          IF NOT SkipSeparator THEN
            PRINT #UBRpt, Fmt$(0)
            LineCnt = LineCnt + 1
          END IF
        END IF
      ELSEIF (TRevCnt = NumOfRevs) AND UBBillRec(1).TransAmt > 0 THEN
        '102998  Moved tax printing to here "now prints one tax line per customer"
        CTaxAmt# = 0
        FOR TXCnt = 1 TO 15
          IF UBBillRec(1).TaxAmt(TXCnt) > 0 THEN
            CTaxAmt# = Round#(CTaxAmt# + UBBillRec(1).TaxAmt(TXCnt))
          END IF
        NEXT
        IF CTaxAmt# > 0 THEN
          PRINT #UBRpt, " Tax"; TAB(69); USING Fmt$(3); CTaxAmt#
          LineCnt = LineCnt + 1
        END IF
        Bills2Print = Bills2Print + 1
        AcctBalance# = UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance
        PRINT #UBRpt, TAB(5); "Current:"; USING Fmt$(6); UBBillRec(1).TransAmt;
        IF AcctBalance# <> 0 THEN
          PRINT #UBRpt, TAB(30); "Previous:"; USING Fmt$(6); AcctBalance#;
          TAcctBalance# = Round#(TAcctBalance# + AcctBalance#)
        END IF
        PRINT #UBRpt, TAB(55); "Total:"; TAB(65); USING Fmt$(6); Round#(AcctBalance# + UBBillRec(1).TransAmt)
        LineCnt = LineCnt + 1
        IF NOT SkipSeparator THEN
          PRINT #UBRpt, Fmt$(0)
          LineCnt = LineCnt + 1
        END IF
      END IF
      
      IF UBBillRec(1).TaxExempt = "Y" THEN
        TaxExmp(TRevCnt) = Round#(TaxExmp(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
      END IF
    NEXT
'020199 Moved pump code processing to here. Stops bug in getting true
'       meter consumption figures.
    GOSUB GetWhatPump
    IF HasAPumpCode THEN
      PumpConsump(WhatPump).CustCnt = PumpConsump(WhatPump).CustCnt + 1
      FOR MPCnt = 1 TO 7
        PumpMtrOK = False
        CubMtr = False
        LocMeterType$ = QPTrim$(UBCustRec(1).LocMeters(MPCnt).MTRType)
        SELECT CASE LocMeterType$
        CASE "C", "S", "W"
          PumpMtrOK = True
        END SELECT
        IF PumpMtrOK THEN
          MeterMulti& = UBCustRec(1).LocMeters(MPCnt).MTRMulti
          IF UBCustRec(1).LocMeters(MPCnt).MTRUnit = "C" THEN
            CubMtr = True
          END IF
          IF MeterMulti& <= 0 THEN MeterMulti& = 1
          ReadAmt& = UBBillRec(1).CurRead(MPCnt) - UBBillRec(1).PrevRead(MPCnt)
          IF ReadAmt& < 0 THEN   'Meter rolled over or, been misread
            MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(MPCnt))) - 1)
            ReadAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(MPCnt)) + UBBillRec(1).CurRead(MPCnt)
          END IF
          IF CubMtr THEN
            ReadAmt& = ReadAmt& * 7.481
          END IF
          PumpConsump(WhatPump).Consump = PumpConsump(WhatPump).Consump + (ReadAmt& * MeterMulti&)
        END IF
      NEXT
    END IF
SkipEM:
    IF AskAbandonPrint% THEN
      UBLog "ABORTED: Prebilling report"
      UBLog "Closing files."
      CLOSE
      AbortFlag = True
      EXIT FOR
    END IF
    ShowPctComp Cnt, NumOfRecs
  NEXT
  
  IF AbortFlag THEN GOTO ExitPreReport
  
  PRINT #UBRpt, FF$
  
  GOSUB TitleLine
  PRINT #UBRpt, "Billing Grand Totals"
  IF TennFlag THEN
    PRINT #UBRpt, "                                Inactive          Taxed      NONTax     FlatRate"
    PRINT #UBRpt, "Revenue/Tax        Consump       Consump         Amount      Amount      Amount"
  ELSE
    PRINT #UBRpt, "                                 Inactive                             Flat Rate"
    PRINT #UBRpt, "Revenue/Tax    Consumption      Consumption            Amount           Amount"
  END IF
  PRINT #UBRpt, Fmt$(0)
  
  TotalFlatAmt# = 0
  TotalRevAmt# = 0
  TotalTaxAmt# = 0
  
  FOR RaCnt = 1 TO NumOfRevs
    IF TennFlag THEN
      PRINT #UBRpt, UBSetUpRec(1).Revenues(RaCnt).REVNAME; USING Fmt$(4); ConsumpTot(RaCnt, 1); TAB(30); ConsumpTot(RaCnt, 2);
      IF TaxTotals(RaCnt) > 0 THEN
        PRINT #UBRpt, TAB(44); USING Fmt$(1); Round#(RevTotals(RaCnt) - FlatTotals(RaCnt) - TaxExmp(RaCnt)); TAB(56); TaxExmp(RaCnt); TAB(68); FlatTotals(RaCnt)
      ELSE
        PRINT #UBRpt, TAB(44); USING Fmt$(1); Round#(RevTotals(RaCnt) - FlatTotals(RaCnt)); TAB(68); FlatTotals(RaCnt)
      END IF
    ELSE
      PRINT #UBRpt, UBSetUpRec(1).Revenues(RaCnt).REVNAME; USING Fmt$(4); ConsumpTot(RaCnt, 1); TAB(33); ConsumpTot(RaCnt, 2);
      PRINT #UBRpt, TAB(50); USING Fmt$(1); RevTotals(RaCnt) - FlatTotals(RaCnt); TAB(67); FlatTotals(RaCnt)
    END IF
    TotalFlatAmt# = Round#(TotalFlatAmt# + FlatTotals(RaCnt))
    TotalRevAmt# = Round#(TotalRevAmt# + RevTotals(RaCnt))
    IF TaxTotals(RaCnt) > 0 THEN
      IF TennFlag THEN
        PRINT #UBRpt, " Tax"; TAB(44); USING Fmt$(1); TaxTotals(RaCnt)
      ELSE
        PRINT #UBRpt, " Tax"; TAB(50); USING Fmt$(1); TaxTotals(RaCnt)
      END IF
      TotalTaxAmt# = Round#(TotalTaxAmt# + TaxTotals(RaCnt))
    END IF
  NEXT
  PRINT #UBRpt, Fmt$(0)
  PRINT #UBRpt, "  PREVIOUS: "; USING Fmt$(6); TAcctBalance#;
  PRINT #UBRpt, TAB(32); "REVENUE TOTAL: "; USING Fmt$(5); Round#(TotalRevAmt# - TotalFlatAmt#)
  PRINT #UBRpt, "BILL COUNT: "; USING Fmt$(2); Bills2Print;
  PRINT #UBRpt, TAB(32); "   FLAT TOTAL: "; USING Fmt$(5); TotalFlatAmt#
  PRINT #UBRpt, TAB(32); "    TAX TOTAL: "; USING Fmt$(5); TotalTaxAmt#
  PRINT #UBRpt, TAB(32); "BILLING TOTAL: "; USING Fmt$(5); Round#(TotalRevAmt# + TotalTaxAmt#)
  PRINT #UBRpt, FF$
  
  TotalRevAmt# = 0
  
  GOSUB RptTotRateHeader
  
  FOR RaCnt = 1 TO NumOfRates
    IF (RateTotals(RaCnt) <> 0) OR (RateConsump(RaCnt) <> 0) THEN
      IF LEN(QPTrim$(UBRateTbls(RaCnt).RATECODE)) > 0 THEN
        PRINT #UBRpt, UBRateTbls(RaCnt).RATECODE; "    "; UBRateTbls(RaCnt).RATEDESC; TAB(39); USING Fmt$(4); RateConsump(RaCnt);
        PRINT #UBRpt, TAB(55); USING Fmt$(1); RateTotals(RaCnt);
        PRINT #UBRpt, TAB(69); USING Fmt$(2); RateCount(RaCnt)
        LineCnt = LineCnt + 1
        TotalRevAmt# = Round#(TotalRevAmt# + RateTotals(RaCnt))
        IF RTaxTot(RaCnt) > 0 THEN
          PRINT #UBRpt, " Tax"; TAB(55); USING Fmt$(1); RTaxTot(RaCnt)
          LineCnt = LineCnt + 1
        END IF
        IF LineCnt > MaxLines THEN
          PRINT #UBRpt, FF$
          GOSUB RptTotRateHeader
        END IF
      END IF
    END IF
  NEXT
  
  PRINT #UBRpt, Fmt$(0)
  PRINT #UBRpt, TAB(36); "TAX TOTAL:"; TAB(52); USING Fmt$(5); TotalTaxAmt#
  PRINT #UBRpt, TAB(40); "TOTAL:"; TAB(52); USING Fmt$(5); Round#(TotalRevAmt# + TotalTaxAmt#)
  PRINT #UBRpt, FF$
  
  SortT BookConsump(1), TBooks, 0, LEN(BookConsump(1)), 0, -1
  
  GOSUB BookHeader
  
  FOR Cnt = 1 TO TBooks
    TestTot# = 0
    FOR ZCnt = 1 TO NumOfRevs
      TestTot# = Round#(TestTot# + BookConsump(Cnt).RevAmt(ZCnt))
    NEXT
    IF TestTot# <> 0 THEN
      IF BookConsump(Cnt).Book < 10 THEN
        Book$ = "0" + QPTrim$(STR$(BookConsump(Cnt).Book))
      ELSE
        Book$ = QPTrim$(STR$(BookConsump(Cnt).Book))
      END IF
      PRINT #UBRpt, "Book: "; Book$; "    Customers:"; BookConsump(Cnt).CustCnt
      TBookAmt# = 0
      TBTaxAmt# = 0
      FOR RCnt = 1 TO NumOfRevs
        PRINT #UBRpt, RevDesc(RCnt); TAB(30); USING Fmt$(4); BookConsump(Cnt).Consump(RCnt);
        PRINT #UBRpt, TAB(59); USING "##########.##"; BookConsump(Cnt).RevAmt(RCnt)
        TBookAmt# = Round#(TBookAmt# + BookConsump(Cnt).RevAmt(RCnt))
        IF BookConsump(Cnt).TaxAmt(RCnt) > 0 THEN
          PRINT #UBRpt, " Tax"; TAB(60); USING Fmt$(1); BookConsump(Cnt).TaxAmt(RCnt)
          TBTaxAmt# = Round#(TBTaxAmt# + BookConsump(Cnt).TaxAmt(RCnt))
          LineCnt = LineCnt + 1
        END IF
        LineCnt = LineCnt + 1
      NEXT
      TBookGTot# = Round#(TBookGTot# + TBookAmt# + TBTaxAmt#)
      PRINT #UBRpt, TAB(42); "Book Total:"; TAB(57); USING Fmt$(5); Round#(TBookAmt# + TBTaxAmt#)
      IF Cnt < TBooks THEN
        PRINT #UBRpt, Fmt$(0)
      END IF
      LineCnt = LineCnt + 1
    END IF
    IF LineCnt > MaxLines AND Cnt < TBooks THEN
      PRINT #UBRpt, FF$
      GOSUB BookHeader
    END IF
    
SkipThisBook:
  NEXT
  PRINT #UBRpt, Fmt$(0)
  PRINT #UBRpt, TAB(35); "Books GRAND Total:"; TAB(57); USING Fmt$(5); TBookGTot#
  PRINT #UBRpt, FF$

  IF TPumps > 0 THEN
    GOSUB PumpHeader
    TMMConsump# = 0
    FOR Cnt = 1 TO TPumps
      PRINT #UBRpt, PumpConsump(Cnt).PumpCode; TAB(30); USING "###########"; PumpConsump(Cnt).CustCnt; TAB(60); PumpConsump(Cnt).Consump
      TMMConsump# = TMMConsump# + PumpConsump(Cnt).Consump
    NEXT
    PRINT #UBRpt, Fmt$(0)
    PRINT #UBRpt, TAB(35); "Pump Code Total:"; TAB(60); USING "###########"; TMMConsump#
  END IF

  CLOSE
  
  UBLog "Finished writing prebilling report."
  
  SELECT CASE Choice
  CASE 1
    RptText$ = "(Customer"
  CASE 2
    RptText$ = "(Account"
  CASE 3
    RptText$ = "(Location"
  CASE 4
    RptText$ = "(Postal RT."
  CASE 5
    RptText$ = "(ZipCode"
  CASE 6
    RptText$ = "(Sequence"
  END SELECT
  RptText$ = RptText$ + " Order)"
  
  ERASE UBSetUpRec, RevDesc, UBRateTbls, RateConsump
  ERASE Fmt$, UBCustRec, UBBillRec, FlatTotals
  ERASE RevTotals, TaxTotals, ConsumpTot
  ERASE RateTotals, RTaxTot, BookConsump, IndexArray
  ERASE RateCount, ProrateServ
  ERASE PumpConsump, TaxExmp

  IF NOT AbortFlag THEN
    PrintRptFile "Pre-Billing Report " + RptText$, "UBPREBIL.RPT", LPTPort, RetCode, EntryPoint
    IF BookFlag THEN
      KILL UBBillsFile
    END IF
  END IF
  
  GOTO ExitPreReport
  
PrintPreHeader:
  GOSUB TitleLine
  PRINT #UBRpt, "Stat  Act.  Locat    Customer Name             Service Address       Prorate%"
  PRINT #UBRpt, "Revenue            R-Code     Cur Read    Pre Read     Consump        Charges"
  PRINT #UBRpt, Fmt$(0)
  LineCnt = 5
RETURN
  
GetWhatBook:
  BadBookFlag = False
  WhatBook = 0
  IF LEN(QPTrim$(UBCustRec(1).Book)) = 0 THEN
    IF UBCustRec(1).Status = "A" THEN
      BadBookFlag = True
      'testing vvv
      WhatBook = 0
    END IF
    GOTO ErrorBookExit
  END IF
  
  ThisBook = VAL(UBCustRec(1).Book)
  IF TBooks > 0 THEN
    FOR TBCnt = 1 TO TBooks
      IF BookConsump(TBCnt).Book = ThisBook THEN
        WhatBook = TBCnt
        EXIT FOR
      END IF
    NEXT
    IF WhatBook = 0 THEN
      TBooks = TBooks + 1
      REDIM PRESERVE BookConsump(0 TO TBooks) AS BookConsumpType
      BookConsump(TBooks).Book = ThisBook
      WhatBook = TBooks
    END IF
  ELSE
    TBooks = TBooks + 1
    BookConsump(TBooks).Book = ThisBook
    WhatBook = TBooks
  END IF
  
ErrorBookExit:
RETURN
  
GetWhatPump:
  HasAPumpCode = True   'assume they have a pump code
  WhatPump = 0
  IF LEN(QPTrim$(UBCustRec(1).PumpCode)) = 0 THEN
    IF UBCustRec(1).Status = "A" THEN
      HasAPumpCode = False  'no pump code
      WhatPump = 0
    END IF
    GOTO PumpCodeReturn
  END IF

  CustPump$ = UCASE$(QPTrim$(UBCustRec(1).PumpCode))
  IF LEN(CustPump$) > 0 THEN
    FOR TBCnt = 1 TO TPumps
      ThisPump$ = QPTrim$(PumpConsump(TBCnt).PumpCode)
      IF ThisPump$ = CustPump$ THEN
        WhatPump = TBCnt
        EXIT FOR
      END IF
    NEXT
    IF WhatPump = 0 THEN
      TPumps = TPumps + 1
      REDIM PRESERVE PumpConsump(0 TO TPumps) AS PumpConsumpType
      PumpConsump(TPumps).PumpCode = CustPump$
      WhatPump = TPumps
    END IF
  ELSE
    TPumps = TPumps + 1
    PumpConsump(TPumps).PumpCode = CustPump$
    WhatPump = TPumps
  END IF

PumpCodeReturn:
RETURN

RptTotRateHeader:
  GOSUB TitleLine
  PRINT #UBRpt,
  PRINT #UBRpt, "Report Totals by Rate Code"
  PRINT #UBRpt,
  PRINT #UBRpt, "Code      Rate Description            Consumption           Amount      Bills"
  PRINT #UBRpt, Fmt$(0)
  LineCnt = 5
RETURN
  
BookHeader:
  GOSUB TitleLine
  PRINT #UBRpt, "Report Totals by Book"
  PRINT #UBRpt,
  PRINT #UBRpt, "Book"
  PRINT #UBRpt, "Revenue                      Consumption                         Amount"
  PRINT #UBRpt, Fmt$(0)
  LineCnt = 7
RETURN

PumpHeader:
  GOSUB TitleLine
  PRINT #UBRpt, "Report Totals by Pump Code"
  PRINT #UBRpt,
  PRINT #UBRpt, "PumpCode                  Customer Count                    Consumption"
  PRINT #UBRpt, Fmt$(0)
  LineCnt = 6
RETURN

TitleLine:
  PageNo = PageNo + 1
  PRINT #UBRpt, "Utility Pre-Billing Report.  "; TownName$; TAB(70); "Page: "; PageNo
  PRINT #UBRpt, TheDate$
RETURN

ErrorAbortExit:
  CLOSE
  
ExitPreReport:
  UBLog "OUT: Prebilling Report" + CrLf$

END SUB

SUB RateCodeErrScrn (RATECODE$)
  REDIM TempArray(0) AS INTEGER
  SaveScrn TempArray()
  BlockClear
  DisplayUBScrn "ERRSCRN1"
  QPrintRC "RATE CODE:  " + QPTrim$(RATECODE$), 10, 22, -1
  QPrintRC "Has an INVALID entry!", 10, 39, -1
  QPrintRC "Correct and Print Again.", 12, 28, -1
  WaitForAction
  AbortFlag = True
  RestScrn TempArray()
  ERASE TempArray
END SUB

