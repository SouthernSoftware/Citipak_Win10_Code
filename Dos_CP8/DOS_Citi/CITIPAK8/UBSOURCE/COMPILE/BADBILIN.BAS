DEFINT A-Z
  
DECLARE SUB BillSummaryRpt ()
DECLARE FUNCTION GetBillBook% ()
DECLARE SUB UBLog (Text$)
DECLARE FUNCTION GetBillCycle% ()
DECLARE SUB RefundMenu ()
DECLARE SUB RefundDeposit ()
DECLARE FUNCTION GetCustMeterType% (UBLocatRec() AS ANY, ThisMeter%)
DECLARE SUB UtilAdjustments ()
DECLARE SUB RePrintUtilBills ()
DECLARE SUB PostBillTrans ()
DECLARE FUNCTION GetAdjFactor# ()
DECLARE FUNCTION IsDeleted% (CustAcctNo&)
DECLARE SUB ClearScrn ()
DECLARE SUB UBBillMenu ()
DECLARE SUB BillPrintMenu ()
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetupLen%)
DECLARE SUB PrintUtilBills ()
DECLARE SUB MakeSequenceIndex (IndexText$)
DECLARE FUNCTION AskAbandonPrint% ()
  
DECLARE FUNCTION FindRateTbl% (RATECODE$, NumOfRates%, UBRateTbls() AS ANY)
DECLARE FUNCTION GetNumOfRevs% ()
DECLARE FUNCTION GetRevCharge# (RateTbl AS ANY, TMeterConsp&, MeterMulti&)
DECLARE SUB MakeBillFile (AbortFlag%, FuelAdjAmt#, ThisCycle%, ThisBook%)
DECLARE FUNCTION Round# (N#)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB MakeZipCodeIndex ()
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB MakePostalIndex (IndexText$)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB MakeTempIndex (Indextype%)
DECLARE SUB PreBillReport ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB ManualMeterReading ()
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB EnterMeterReadingsMenu ()
DECLARE FUNCTION GetNumOfCust% ()
DECLARE FUNCTION GetNumRateRecs% ()
DECLARE FUNCTION Exist% (FileName$)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION AskSavePayment
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB HideCursor ()
DECLARE SUB KillFile (FileName$)
  
  '$INCLUDE: 'DefCnf.BI'
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FClose (Handle%)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB SortStr2 (BYVAL Address%, Size%, Direction%)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION QPValL& (LongNum$)
DECLARE SUB MPaintBox (TRow%, LCol%, BRow%, RCol%, Kolor%)
DECLARE SUB StuffBuf (Key2Stuff$)
DECLARE SUB ShowCustHistory (CustAcct&)
DECLARE SUB CursorOff ()
DECLARE SUB WaitForAction ()
DECLARE SUB BlockClear ()
DECLARE SUB SaveScrn (ScrnArray())
DECLARE SUB RestScrn (ScrnArray())
DECLARE SUB LookUp (CustAcct&, Text$, Parm1, Parm2, ActiveOnly%)
DECLARE FUNCTION PromptSaveData ()
  
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'ubdraft.BI'
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'PageInfo.bi'
  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubrate.BI'
  
  CONST False = 0, True = NOT False
  
  CrLf$ = CHR$(13) + CHR$(10)
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  
  IF INSTR(UBSetUpRec(1).DEFCITY, "MCCOR") > 0 THEN
    MccorFlag = True
  END IF
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 11)
  
  MChoice$(1) = "Meter Readings "
  MChoice$(2) = "Pre-Billing Report "
  MChoice$(3) = "Utility Bill Printing Menu "
  MChoice$(4) = "Post Billing Transactions"
  MChoice$(5) = "Bank Draft Processing"
  MChoice$(6) = "Apply Penalties "
  MChoice$(7) = "Refund a Customers Deposit "
  MChoice$(8) = "Utility Billing Adjustments "
  MChoice$(9) = "Print Late Notices "
  MChoice$(10) = "Exit to OS "
  
  '  MChoice$(4) = "Billing Summary Report"
  '  MChoice$(5) = "Post Billing Transactions"
  '  MChoice$(6) = "Bank Draft Processing"
  '  MChoice$(7) = "Apply Penalties "
  '  MChoice$(8) = "Refund a Customers Deposit "
  '  MChoice$(9) = "Utility Billing Adjustments "
  '  MChoice$(10) = "Print Late Notices "
  '  MChoice$(11) = "Exit to OS "
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 18   'limit the box length to go no lower than line 20
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
  Row = ((25 - (UBOUND(MChoice$))) \ 2)
  Col = ((80 - MaxLen) \ 2) - 2
  
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
      'RefundMenu        'unrem
    CASE 8
      UtilAdjustments
    CASE 9
      RUN "UBLNPRNT"
    CASE 10
      HideCursor
      ClearScrn
      END
      
    END SELECT
  LOOP UNTIL ExitFlag
  
  RUN "UBMENU"

FUNCTION FindRateTbl% (RATECODE$, NumOfRates, UBRateTbls() AS UBRateTblRecType) STATIC
  FindRateTbl = 0
  RATECODE$ = QPTrim$(RATECODE$)
  IF LEN(RATECODE$) THEN        'if this rev has a rate code
    FOR TCnt = 1 TO NumOfRates  'find the right one
      IF RATECODE$ = QPTrim$(UBRateTbls(TCnt).RATECODE) THEN
        FindRateTbl = TCnt
        EXIT FOR
      END IF
    NEXT
  END IF
END FUNCTION

FUNCTION GetAdjFactor#
  
  LibName$ = "UB"
  ScrnName$ = "FUELADJ"
  
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
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key, 13
      AdjFactor# = VAL(Form$(1, 0))
      ExitFlag = True
    CASE EscKey
      AdjFactor# = -10000
      ExitFlag = True
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 15   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 28 TO 40           '--Cancel button
          PressButton 27, 15, 28, 40
        CASE 41 TO 53           '--F5 button
          PressButton 13, 15, 41, 53
        END SELECT
      END SELECT                'row
    END IF
    
  LOOP UNTIL ExitFlag
  
  GetAdjFactor# = AdjFactor#
  ERASE Frm, Form$, Fld
  
END FUNCTION

FUNCTION GetBillBook%
  LibName$ = "UB"
  ScrnName$ = "BOOKGET"
  
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
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key, 13
      BookNum = VAL(Form$(1, 0))
      ExitFlag = True
    CASE EscKey
      BookNum = -10000
      ExitFlag = True
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 15   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 28 TO 40           '--Cancel button
          PressButton 27, 15, 28, 40
        CASE 41 TO 53           '--F5 button
          PressButton 13, 15, 41, 53
        END SELECT
      END SELECT                'row
    END IF
    
  LOOP UNTIL ExitFlag
  
  IF BookNum = 99 THEN
    BookNum = -1
  END IF
  
  GetBillBook% = BookNum
  ERASE Frm, Form$, Fld
  
END FUNCTION

FUNCTION GetBillCycle%
  LibName$ = "UB"
  ScrnName$ = "CYCLEGET"
  
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
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key, 13
      CycleNum = VAL(Form$(1, 0))
      ExitFlag = True
    CASE EscKey
      CycleNum = -10000
      ExitFlag = True
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 15   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 28 TO 40           '--Cancel button
          PressButton 27, 15, 28, 40
        CASE 41 TO 53           '--F5 button
          PressButton 13, 15, 41, 53
        END SELECT
      END SELECT                'row
    END IF
    
  LOOP UNTIL ExitFlag
  
  GetBillCycle% = CycleNum
  ERASE Frm, Form$, Fld
  
  
END FUNCTION

FUNCTION GetCustMeterType (UBCustRec() AS NewUBCustRecType, ThisMeter) STATIC
  
  'Meter Types
  'CONST MtrWaterOnly = 1
  'CONST MtrSewerOnly = 2
  'CONST MtrCombined = 3
  'CONST MtrElectric = 4
  'CONST MtrDemand = 5
  'CONST MtrGas = 6
  'CONST MtrTouchRead = 7
  
  LMtrType$ = QPTrim$(UBCustRec(1).LocMeters(ThisMeter).MTRType)
  LMtrTypeLen = LEN(LMtrType$)
  IF LMtrTypeLen > 0 THEN
    SELECT CASE LMtrType$
    CASE "W"
      LThisMeter = MtrWaterOnly
    CASE "S"
      LThisMeter = MtrSewerOnly
    CASE "C"
      LThisMeter = MtrCombined
    CASE "E"
      LThisMeter = MtrElectric
    CASE "D"
      LThisMeter = MtrDemand
    CASE "G"
      LThisMeter = MtrGas
    CASE "T"
      LThisMeter = MtrTouchRead
    CASE ELSE
      LThisMeter = True
    END SELECT
    GetCustMeterType = LThisMeter
  ELSE
    GetCustMeterType = 0
  END IF
  
END FUNCTION

FUNCTION GetNumOfRevs%
  NumOfRevs = 15
  REDIM UBSetUpRec(1) AS UBSetupRecType
  UBSetupLen = LEN(UBSetUpRec(1))
  FGetAH "UBSETUP.DAT", UBSetUpRec(1), UBSetupLen, 1            'load it
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
    IF MeterConsump& >= RateTbl.TblBreaks(1).UNITS AND MeterConsump& < RateTbl.TblBreaks(2).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(1).UNITS)
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
    IF MeterConsump& > RateTbl.TblBreaks(2).UNITS AND MeterConsump& < RateTbl.TblBreaks(3).UNITS THEN
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
    IF MeterConsump& >= RateTbl.TblBreaks(3).UNITS AND MeterConsump& < RateTbl.TblBreaks(4).UNITS THEN
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
    IF MeterConsump& >= RateTbl.TblBreaks(4).UNITS AND MeterConsump& < RateTbl.TblBreaks(5).UNITS THEN
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
    IF MeterConsump& >= RateTbl.TblBreaks(5).UNITS AND MeterConsump& < RateTbl.TblBreaks(6).UNITS THEN
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
    IF MeterConsump& >= RateTbl.TblBreaks(6).UNITS AND MeterConsump& < RateTbl.TblBreaks(7).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(6).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(6).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(6).UNITS - RateTbl.TblBreaks(6).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(6).UNITAMT))
    END IF
  ELSE
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(6).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(6).UNITAMT))
    GOTO GotTAmt
  END IF
  
  'break 7
  IF LastTblCnt >= 8 THEN
    IF MeterConsump& >= RateTbl.TblBreaks(7).UNITS AND MeterConsump& < RateTbl.TblBreaks(8).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(7).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(7).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(7).UNITS - RateTbl.TblBreaks(7).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(7).UNITAMT))
    END IF
  ELSE
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(7).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(7).UNITAMT))
    GOTO GotTAmt
  END IF
  
  'break 8
  IF LastTblCnt >= 9 THEN
    IF MeterConsump& >= RateTbl.TblBreaks(8).UNITS AND MeterConsump& < RateTbl.TblBreaks(9).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(8).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(8).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(8).UNITS - RateTbl.TblBreaks(8).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(8).UNITAMT))
    END IF
  ELSE
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(8).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(8).UNITAMT))
    GOTO GotTAmt
  END IF
  
  'break 9
  IF LastTblCnt >= 10 THEN
    IF MeterConsump& >= RateTbl.TblBreaks(9).UNITS AND MeterConsump& < RateTbl.TblBreaks(10).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(9).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(9).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(9).UNITS - RateTbl.TblBreaks(9).UNITS)
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
  
  REDIM UBSetUp(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUp(), UBSetupLen
  
  IF UBSetUp(1).PreByBook = "Y" AND ThisBook > 0 THEN
    BookFlag = True
  ELSEIF UBSetUp(1).BILLCYCL = "Y" THEN
    CycleFlag = True
  END IF
  
  'find the electric revenue position
  'IF FuelAdjAmt# <> 0 THEN
  FOR ThisRevCnt = 1 TO 15
    IF INSTR(UBSetUp(1).Revenues(ThisRevCnt).REVNAME, "ELECTRIC") THEN
      ElecRev = ThisRevCnt
      EXIT FOR
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
  CLOSE
  
  NumCustRec& = FileSize&("UBCUST.DAT") \ UBCustRecLen
  
  IF Exist(UBBillsFile) THEN
    KillFile UBBillsFile
  END IF
  
  IF Exist("UBBILLS.PRN") THEN
    KillFile "UBBILLS.PRN"
  END IF
  
  BillFile = FREEFILE
  OPEN UBBillsFile FOR RANDOM SHARED AS BillFile LEN = UBBillRecLen
  
  CustFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CustFile LEN = UBCustRecLen
  
  BillCnt = 0
  NumOfRevs = GetNumOfRevs%
  
  FOR LCnt = 1 TO NumCustRec&
    
    'clears bill rec ready for next customer
    REDIM UBBillRec(1) AS UBTransRecType
    
    GET CustFile, LCnt, UBCustRec(1)
    
    IF BookFlag THEN
      IF QPValL(UBCustRec(1).Book) <> ThisBook THEN
        UBBillRec(1).TransAmt = 0
        FOR RCnt = 1 TO NumOfRevs
          UBBillRec(1).RevAmt(RCnt) = 0
          UBBillRec(1).TaxAmt(RCnt) = 0
        NEXT
        FOR ZZ = 1 TO 7
          UBBillRec(1).CurRead(ZZ) = 0
          UBBillRec(1).PrevRead(ZZ) = 0
        NEXT
        UBBillRec(1).ActiveFlag = False
        GOTO MSkipem
      END IF
    END IF
    
    IF CycleFlag THEN
      IF UBCustRec(1).BILLCYCL <> ThisCycle THEN
        UBBillRec(1).TransAmt = 0
        FOR RCnt = 1 TO NumOfRevs
          UBBillRec(1).RevAmt(RCnt) = 0
          UBBillRec(1).TaxAmt(RCnt) = 0
        NEXT
        FOR ZZ = 1 TO 7
          UBBillRec(1).CurRead(ZZ) = UBCustRec(1).LocMeters(ZZ).CurRead
          UBBillRec(1).PrevRead(ZZ) = UBCustRec(1).LocMeters(ZZ).PrevRead
          UBBillRec(1).MtrTypes(ZZ) = GetCustMeterType(UBCustRec(), ZZ)
        NEXT
        UBBillRec(1).ActiveFlag = False
        GOTO MSkipem
      END IF
    END IF
    
    IF UBCustRec(1).STATUS <> "A" THEN
      UBBillRec(1).TransAmt = 0
      FOR RCnt = 1 TO NumOfRevs
        UBBillRec(1).RevAmt(RCnt) = 0
        UBBillRec(1).TaxAmt(RCnt) = 0
      NEXT
      FOR ZZ = 1 TO 7
        UBBillRec(1).CurRead(ZZ) = UBCustRec(1).LocMeters(ZZ).CurRead
        UBBillRec(1).PrevRead(ZZ) = UBCustRec(1).LocMeters(ZZ).PrevRead
        UBBillRec(1).MtrTypes(ZZ) = GetCustMeterType(UBCustRec(), ZZ)
      NEXT
      UBBillRec(1).ActiveFlag = False
      GOTO MSkipem
    END IF
    
    MeterConsp& = 0
    TMeterConsp& = 0
    
    FOR FRCnt = 1 TO 4
      WhatService = UBCustRec(1).FlatRates(FRCnt).REVSRC
      IF UBCustRec(1).FlatRates(FRCnt).FRAMT <> 0 AND WhatService > 0 THEN
        '11/19/96 Fixed Rev. amt. to add to current rev amt
        'UBBillRec(1).RevAmt(WhatService) = Round#(UBBillRec(1).RevAmt(WhatService) + UBCustRec(1).FlatRates(FRCnt).FRAmt)
        '01-09-97 Fixed Multiplier bug in flat rates
        Multi = UBCustRec(1).FlatRates(FRCnt).NumMin
        IF Multi < 1 THEN Multi = 1
        FlatAmt# = Round#(UBCustRec(1).FlatRates(FRCnt).FRAMT * Multi)
        UBBillRec(1).RevAmt(WhatService) = Round#(UBBillRec(1).RevAmt(WhatService) + FlatAmt#)
        UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + FlatAmt#)
        IF UBSetUp(1).Revenues(WhatService).TAXRATE > 0 AND UBCustRec(1).TAXEXPT <> "Y" THEN
          UBBillRec(1).TaxAmt(WhatService) = Round#(UBBillRec(1).RevAmt(WhatService) * UBSetUp(1).Revenues(WhatService).TAXRATE)
          UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + UBBillRec(1).TaxAmt(WhatService))
        END IF
      END IF
    NEXT
    
    '12-6-96  Monthly Billed amounts
    FOR MRCnt = 1 TO 2
      WhatService = UBCustRec(1).Monthly(MRCnt).RevSource
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
      WhatTbl = FindRateTbl(UBCustRec(1).Serv(RCnt).RATECODE, NumOfRates, UBRateTbls())
      
      IF WhatTbl THEN
        'if this is a non-metered service
        IF UBSetUp(1).Revenues(RCnt).UseMtr = "N" THEN
          '02-05-97 added fix add to current rev amt
          IF UBRateTbls(WhatTbl).MINAMT > -1000000 THEN
            UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + UBRateTbls(WhatTbl).MINAMT)
            UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + UBRateTbls(WhatTbl).MINAMT)
          ELSE
            REDIM TempArray(0) AS INTEGER
            SaveScrn TempArray()
            BlockClear
            DisplayUBScrn "ERRSCRN1"
            QPrintRC "RATE CODE:  " + QPTrim$(UBRateTbls(WhatTbl).RATECODE), 10, 22, -1
            QPrintRC "Has an INVALID entry!", 10, 39, -1
            QPrintRC "Correct and Print Again.", 12, 28, -1
            WaitForAction
            AbortFlag = True
            RestScrn TempArray()
            ERASE TempArray
            GOTO AbortExit
          END IF
          GOTO GotAmt
        END IF
        
        MeterType$ = UBCustRec(1).Serv(RCnt).RMTRTYPE
        
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
              MeterConsp& = 0
            ELSE
              MeterConsp& = UBCustRec(1).LocMeters(MCnt).CurRead - UBCustRec(1).LocMeters(MCnt).PrevRead
            END IF
            
            IF MeterConsp& < 0 THEN
              MaxMeterAmt& = 10& ^ (LEN(STR$(UBCustRec(1).LocMeters(MCnt).PrevRead)) - 1)
              MeterConsp& = (MaxMeterAmt& - UBCustRec(1).LocMeters(MCnt).PrevRead) + UBCustRec(1).LocMeters(MCnt).CurRead
            END IF
            'Meter multi??
            IF UBCustRec(1).LocMeters(MCnt).MTRMulti > 0 THEN
              MeterConsp& = MeterConsp& * UBCustRec(1).LocMeters(MCnt).MTRMulti
            END IF
            'PRINT UBCustRec(1).LocMeters(MCnt).NumUser
            'convert units here if necessary
            'To Convert From Cubic Feet to Gallons Mult by 7.481 Here
            IF (UBBillRec(1).MtrTypes(MCnt) = 1 OR UBBillRec(1).MtrTypes(MCnt) = 2 OR UBBillRec(1).MtrTypes(MCnt) = 3) AND UBCustRec(1).LocMeters(MCnt).MTRUnit = "C" THEN
              MeterConsp& = MeterConsp& * 7.481
            END IF
            TMeterConsp& = TMeterConsp& + MeterConsp&
          END IF
        NEXT
        
        IF MeterLocNum = 0 THEN
          REDIM TempArray(0) AS INTEGER
          SaveScrn TempArray()
          BlockClear
          DisplayUBScrn "ERRSCRN1"
          QPrintRC "ACCOUNT:" + STR$(LCnt), 10, 22, -1
          QPrintRC "Has Mismatched Meters!", 10, 37, -1
          QPrintRC "Correct and Print Again.", 12, 28, -1
          WaitForAction
          AbortFlag = True
          RestScrn TempArray()
          ERASE TempArray
          GOTO AbortExit
        END IF
        
        AddRevAmt# = 0
        
        IF UBCustRec(1).LocMeters(MeterLocNum).NumUser > 1 THEN
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
          'AddRevAmt# = 1 * UBRateTbls(WhatTbl).MINAMT
          IF TMeterConsp& <= UBRateTbls(WhatTbl).MINUNITS THEN
            'if we bill the minium
            IF UBRateTbls(WhatTbl).MINAMT > -1000000 THEN
              RevAmt# = NumUser& * UBRateTbls(WhatTbl).MINAMT
              UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + RevAmt#)
              UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + RevAmt#)
              GOTO GotAmt
            ELSE
              
            END IF
          END IF
        END IF
        
        '01-20-97 Add Fix for minium units test for equal to also.
        '04-23-1997 'Fixed to ADD TO TOTAL
        'AddRevAmt# = 0
        RevAmt# = GetRevCharge#(UBRateTbls(WhatTbl), TMeterConsp&, MeterMulti&)
        RevAmt# = RevAmt# + AddRevAmt#
        UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + RevAmt#)
        
        IF RCnt = ElecRev THEN
          FuelAddAmt# = Round#(FuelAdjAmt# * TMeterConsp&)
          UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + FuelAddAmt#)
          UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + FuelAddAmt#)
          'UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + Round#(UBBillRec(1).RevAmt(RCnt) * FuelAdjAmt#))
        END IF
        
        UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + RevAmt#)
GotAmt:
        IF UBSetUp(1).Revenues(RCnt).TAXRATE > 0 AND UBCustRec(1).TAXEXPT <> "Y" THEN
          UBBillRec(1).TaxAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) * UBSetUp(1).Revenues(RCnt).TAXRATE)
          UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + UBBillRec(1).TaxAmt(RCnt))
        END IF
      ELSE
        'find me
        IF LEN(QPTrim$(UBCustRec(1).Serv(RCnt).RMTRTYPE)) > 0 THEN
          '  DoDoCnt = DoDoCnt + 1
          '  LPRINT LCnt, UBSetUp(1).Revenues(RCnt).REVNAME
          '  find me
          '  REDIM TempArray(0) AS INTEGER
          '  SaveScrn TempArray()
          '  BlockClear
          '  DisplayUBScrn "ERRSCRN1"
          '  QPrintRC "ACCOUNT:" + STR$(LCnt), 10, 22, -1
          '  QPrintRC "Has a Defined Meter,", 10, 37, -1
          '  QPrintRC "with an INVALID, or BLANK ratecode!", 11, 22, -1
          '  QPrintRC "Correct and Print Again.", 13, 28, -1
          '  WaitForAction
          '  AbortFlag = True
          '  RestScrn TempArray()
          '  ERASE TempArray
          '  GOTO AbortExit
        END IF
      END IF
    NEXT        'loop through all revenue sources
    
    IF UBCustRec(1).STATUS = "I" AND UBBillRec(1).TransAmt > 0 THEN
      UBBillRec(1).TransAmt = 0
      FOR RCnt = 1 TO NumOfRevs
        UBBillRec(1).RevAmt(RCnt) = 0
      NEXT
      UBBillRec(1).ActiveFlag = False
      BadCnt = BadCnt + 1
      UBBillRec(1).CustAcctNo = LCnt
    ELSE
      UBBillRec(1).ActiveFlag = False
    END IF
    
    'Mod for cleveland***
    'IF UBCustRec(1).CUSTTYPE = "NON" THEN
    '  UBBillRec(1).CustAcctNo = LCnt
    '  UBBillRec(1).NONProfit = "Y"
    'END IF
    '***

    IF UBBillRec(1).TransAmt > 0 THEN
      BillCnt = BillCnt + 1
      UBBillRec(1).ActiveFlag = True
      UBBillRec(1).CustAcctNo = LCnt
    END IF
    
MSkipem:
    PUT BillFile, LCnt, UBBillRec(1)
    
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
    ShowPctComp LCnt, NumCustRec&
  NEXT
  
AbortExit:
  
  CLOSE
  'LPRINT DoDoCnt
  ERASE UBBillRec, UBCustRec, UBSetUp, UBRateTbls
  
END SUB

SUB PostBillTrans
  
  UBLog "IN: Bill Posting."
  
  IF NOT Exist("UBBILLS.DAT") THEN
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
  OPEN "UBBILLS.DAT" FOR RANDOM SHARED AS UBBill LEN = UBBillRecLen
  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBBillRecLen
  
  NumOfTranRecs& = LOF(UBTran) \ UBBillRecLen
  NumOfBillRecs = LOF(UBBill) \ UBBillRecLen
  
  ShowProcessingScrn "Posting Billing Transactions"
  'TransDate
  FOR BillCnt = 1 TO NumOfBillRecs
    GET UBBill, BillCnt, UBBillRec(1)
    'mod for cleveland***
    IF (UBBillRec(1).ActiveFlag AND UBBillRec(1).TransAmt > 0) OR (UBBillRec(1).NONProfit = "Y") THEN
      '***
      'IF UBBillRec(1).ActiveFlag AND UBBillRec(1).TransAmt > 0 THEN
      PostedCnt& = PostedCnt& + 1
      NumOfTranRecs& = NumOfTranRecs& + 1       'point to next trans to write
      GET UBCust, BillCnt, UBCustRec(1)
      
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
      
      '06-25-97 added removal of the nonrecurring flat rates
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
        IF UBCustRec(1).LocMeters(MtrCnt).CurRead > 0 THEN
          ReadAmt& = UBBillRec(1).CurRead(MtrCnt) - UBBillRec(1).PrevRead(MtrCnt)
          IF ReadAmt& < 0 THEN  'Meter rolled over or, been misread
            MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(MtrCnt))) - 1)
            ReadAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(MtrCnt)) + UBBillRec(1).CurRead(MtrCnt)
          END IF
          UBCustRec(1).LocMeters(MtrCnt).AvgUse = Round#(UBCustRec(1).LocMeters(MtrCnt).AvgUse + ReadAmt&)
          UBCustRec(1).LocMeters(MtrCnt).UseCnt = UBCustRec(1).LocMeters(MtrCnt).UseCnt + 1
          UBCustRec(1).LocMeters(MtrCnt).ReadFlag = ""
          'UBCustRec(1).LocMeters(MtrCnt).PrevRead = UBCustRec(1).LocMeters(MtrCnt).CurRead
        END IF
      NEXT
      PrevLastTrans& = UBCustRec(1).LastTrans
      UBBillRec(1).PrevTrans = PrevLastTrans&
      UBCustRec(1).LastTrans = NumOfTranRecs&
      PUT UBCust, BillCnt, UBCustRec(1)
      PUT UBTran, NumOfTranRecs&, UBBillRec(1)
    END IF
    ShowPctComp BillCnt, NumOfBillRecs
  NEXT
  CLOSE
  UBLog "FINISH: Posting Transactions."
  UBLog "POSTED:" + STR$(PostedCnt&) + " New BILL Transactions."
  
  KillFile "UBBILLS.DAT"
  KillFile "UBBILLS.PRN"
  
  UBLog "KILLED: UBBILLS.DAT & UBBILLS.PRN"
  BlockClear
  DisplayUBScrn "UPDATEOK"
  WaitForAction
  
ExitBillPost:
  UBLog "OUT: Bill Posting." + CrLf$
END SUB

SUB PreBillReport
  
  UBLog "IN: Prebilling Report"
  
  Dash80$ = STRING$(80, "-")
  Temp1$ = SPACE$(10)
  Temp2$ = SPACE$(12)
  
  NumOfRevs = MaxRevsCnt        'assume max munber of revenue sources
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  TownName$ = UBSetUpRec(1).UTILNAME
  
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
  
  NumOfRates = GetNumRateRecs%
  REDIM UBRateTbls(1 TO NumOfRates) AS UBRateTblRecType
  REDIM RateConsump(1 TO NumOfRates) AS LONG
  
  UBRateTblRecLen = LEN(UBRateTbls(1))
  
  RateFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS RateFile LEN = UBRateTblRecLen
  FOR Cnt = 1 TO NumOfRates
    GET RateFile, Cnt, UBRateTbls(Cnt)
  NEXT
  CLOSE
  
  IF UBSetUpRec(1).PreByBook = "Y" THEN
    ThisBook = GetBillBook%
    IF ThisBook = -1 THEN
      BookFlag = False
    ELSEIF ThisBook = 0 THEN
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
  
  MaxLines = 45
  
  REDIM MChoice$(1 TO 6)
  
  MChoice$(1) = "Customer Name Order"
  MChoice$(2) = "Account Number Order"
  MChoice$(3) = "Location Number Order"
  MChoice$(4) = "Postal Carrier Route Order"
  MChoice$(5) = "ZipCode Order"
  
  IF UBSetUpRec(1).UseSeq = "Y" THEN
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
  
ReStart:
  '--Center Menu within Screen
  Row = ((25 - (UBOUND(MChoice$))) \ 2)
  Col = ((80 - MaxLen) \ 2) - 2
  
  UsingAcct = False
  
  AbortFlag = False
  PageNo = 0
  
  DO            '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    TitleBox 3, Col, MaxLen + 3, "Pre-Billing Report ", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN
      Choice = 0
      ExitFlag = True
    END IF
    
    SELECT CASE Choice
    CASE 0
      ExitFlag = True
    CASE 1
      IndexName$ = NameIndexFile
      OKFlag = True
    CASE 2
      IndexName$ = ""
      UsingAcct = True
      OKFlag = True
    CASE 3
      IndexName$ = BookIndexFile
      OKFlag = True
    CASE 4, 5
      IF Choice = 4 THEN
        IdxTypeText$ = "Postal Route"
      ELSE
        IdxTypeText$ = "Zip-Code"
      END IF
      MakePostalIndex IdxTypeText$
      IndexName$ = TempIndexName
      OKFlag = True
    CASE 6      'Sequence number
      IdxTypeText$ = "Sequence Number"
      MakeSequenceIndex IdxTypeText$
      IndexName$ = TempIndexName
      OKFlag = True
    END SELECT
    
  LOOP UNTIL OKFlag OR ExitFlag
  
  IF ExitFlag THEN GOTO ExitPreReport
  
  IF DoFuelAdjFlag THEN
    FuelAdjAmt# = GetAdjFactor#
    UBLog "Fuel adjustment factor:" + STR$(FuelAdjAmt#)
  ELSE
    FuelAdjAmt# = 0
  END IF
  
  IF FuelAdjAmt# = -10000 THEN GOTO ReStart
  
  UBLog "Calculating utility charges."
  
  MakeBillFile AbortFlag, FuelAdjAmt#, ThisCycle, ThisBook
  
  IF AbortFlag THEN
    UBLog "ABORTED: CALCULATIONS"
    IF Exist(UBBillsFile) THEN
      KillFile UBBillsFile
    END IF
  ELSE
    UBLog "Finished utility calculations."
  END IF
  
  IF AbortFlag THEN GOTO ExitPreReport
  
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBBillRec(1) AS UBTransRecType
  UBBillRecLen = LEN(UBBillRec(1))
  
  REDIM RevTotals(1 TO NumOfRevs) AS DOUBLE     'holds revenues total amt
  '052097 added tax by revenue totals
  REDIM TaxTotals(1 TO NumOfRevs) AS DOUBLE     'holds revenues tax total amt
  REDIM ConsumpTot(1 TO NumOfRevs, 1 TO 2) AS DOUBLE            'holds each revenues total amt
  REDIM RateConsump(1 TO NumOfRates) AS LONG
  REDIM RateTotals(1 TO NumOfRates) AS DOUBLE   'holds each Rates $totals
  '052097 added tax by rate code totals
  REDIM RTaxTot(1 TO NumOfRates) AS DOUBLE      'holds each Rates Tax totals
  '052097 added tax by book totals to type def
  REDIM BookConsump(1 TO 1) AS BookConsumpType
  
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
  OPEN "UBBILLS.DAT" FOR RANDOM SHARED AS UBBill LEN = UBBillRecLen
  
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
    
    'IF ThisCustRec& = 11 THEN STOP
    
    IF UBCustRec(1).DelFlag THEN
      GOTO skipem
    END IF
    
    IF BookFlag THEN
      IF QPValL(UBCustRec(1).Book) <> ThisBook THEN
        GOTO skipem
      END IF
    END IF
    
    IF CycleFlag THEN
      IF UBCustRec(1).BILLCYCL <> ThisCycle THEN
        GOTO skipem
      END IF
    END IF
    
    'added to skip over final's
    IF UBCustRec(1).STATUS = "F" THEN
      GOTO skipem
    END IF
    
    GET UBBill, ThisCustRec&, UBBillRec(1)
    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, CHR$(12)
      GOSUB PrintPreHeader
    END IF
    
    GOSUB GetWhatBook
    
    IF UBBillRec(1).ActiveFlag <> 0 THEN
      BookConsump(WhatBook).CustCnt = BookConsump(WhatBook).CustCnt + 1
      PRINT #UBRpt, UBCustRec(1).STATUS; "  "; UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB; "  "; UBCustRec(1).CUSTNAME; "  "; QPTrim$(UBCustRec(1).SERVADDR)
    END IF
    
    WhatRate = 0
    
    DoneOne = False
    FOR TRevCnt = 1 TO NumOfRevs
      IF UBBillRec(1).RevAmt(TRevCnt) <> 0 THEN
        DoneOne = False
        PRINT #UBRpt, RevDesc(TRevCnt);
        FOR TRateCnt = 1 TO NumOfRates
          IF UBRateTbls(TRateCnt).RATECODE = UBCustRec(1).Serv(TRevCnt).RATECODE THEN
            MINAMT& = UBRateTbls(TRateCnt).MINUNITS
            WhatRate = TRateCnt
            EXIT FOR
          END IF
        NEXT
        
        IF UBSetUpRec(1).Revenues(TRevCnt).UseMtr = "Y" THEN
          RevTotals(TRevCnt) = Round#(RevTotals(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
          BookConsump(WhatBook).RevAmt(TRevCnt) = Round#(BookConsump(WhatBook).RevAmt(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
          BookConsump(WhatBook).TaxAmt(TRevCnt) = Round#(BookConsump(WhatBook).TaxAmt(TRevCnt) + UBBillRec(1).TaxAmt(TRevCnt))
          '02-20-97 Add new cash totals by rate
          
          IF WhatRate > 0 THEN
            RateTotals(WhatRate) = Round#(RateTotals(WhatRate) + UBBillRec(1).RevAmt(TRevCnt))
            RTaxTot(WhatRate) = Round#(RTaxTot(WhatRate) + UBBillRec(1).TaxAmt(TRevCnt))
          END IF
          
          FOR MCCnt = 1 TO 7
            LocMeterType$ = QPTrim$(UBCustRec(1).LocMeters(MCCnt).MTRType)
            MeterMulti& = UBCustRec(1).LocMeters(MCCnt).MTRMulti                'Meter Multi??
            IF MeterMulti& <= 0 THEN MeterMulti& = 1
            IF (LEN(LocMeterType$) > 0) AND (LocMeterType$ = UBCustRec(1).Serv(TRevCnt).RMTRTYPE) THEN
              DoneOne = True
              MeterNum$ = QPTrim$(UBCustRec(1).Serv(TRevCnt).RATECODE)
              'use the Meternum$ to hold the rate code temporarily
              IF LEN(MeterNum$) > 0 THEN
                IF UBCustRec(1).LocMeters(MCCnt).NumUser > 1 THEN
                  MeterNum$ = MeterNum$ + "*" + QPTrim$(STR$(UBCustRec(1).LocMeters(MCCnt).NumUser))
                END IF
                RSET Temp2$ = MeterNum$
                PRINT #UBRpt, TAB(14); Temp2$;
              ELSE
                RSET Temp2$ = "RATE ERROR"
                PRINT #UBRpt, TAB(14); Temp2$;
              END IF
              
              '052797 added a Read Error
              ReadErr = False
              IF UBBillRec(1).CurRead(MCCnt) < 0 THEN
                ReadErr = True
                'PRINT #UBRpt, TAB(30); USING "**#######"; 0;
                REDIM TempArray(0) AS INTEGER
                SaveScrn TempArray()
                BlockClear
                DisplayUBScrn "ERRSCRN1"
                QPrintRC "ACCOUNT:" + STR$(ThisCustRec&), 10, 22, -1
                QPrintRC "Has an INVALID Reading!", 10, 37, -1
                QPrintRC "Correct and Print Again.", 12, 28, -1
                WaitForAction
                AbortFlag = True
                GOTO ErrorAbortExit
                RestScrn TempArray()
                ERASE TempArray
                
              ELSE
                PRINT #UBRpt, TAB(30); USING "#########"; UBBillRec(1).CurRead(MCCnt);
              END IF
              
              IF UBBillRec(1).PrevRead(MCCnt) < 0 THEN
                ReadErr = True
                PRINT #UBRpt, TAB(42); USING "**#######"; 0;
              ELSE
                PRINT #UBRpt, TAB(42); USING "#########"; UBBillRec(1).PrevRead(MCCnt);
              END IF
              
              IF ReadErr THEN
                ReadAmt& = 0
              ELSE
                ReadAmt& = UBBillRec(1).CurRead(MCCnt) - UBBillRec(1).PrevRead(MCCnt)
              END IF
              
              IF ReadAmt& > 0 THEN
                RateConsump(WhatRate) = RateConsump(WhatRate) + (ReadAmt& * MeterMulti&)
                BookConsump(WhatBook).Consump(TRevCnt) = BookConsump(WhatBook).Consump(TRevCnt) + (ReadAmt& * MeterMulti&)
                ConsumpTot(TRevCnt, 1) = ConsumpTot(TRevCnt, 1) + (ReadAmt& * MeterMulti&)
              ELSEIF ReadAmt& < 0 THEN
                'Meter has rolled over or, been misread
                MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(MCCnt))) - 1)
                ReadAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(MCCnt)) + UBBillRec(1).CurRead(MCCnt)
                RateConsump(WhatRate) = RateConsump(WhatRate) + (ReadAmt& * MeterMulti&)
                BookConsump(WhatBook).Consump(TRevCnt) = BookConsump(WhatBook).Consump(TRevCnt) + (ReadAmt& * MeterMulti&)
                ConsumpTot(TRevCnt, 1) = ConsumpTot(TRevCnt, 1) + (ReadAmt& * MeterMulti&)
              END IF
              
              'ConsumpTot(TRevCnt, 1) = ConsumpTot(TRevCnt, 1) + ReadAmt&
              IF ReadErr THEN
                PRINT #UBRpt, TAB(56); USING "**#######"; ReadAmt& * MeterMulti&;
              ELSE
                PRINT #UBRpt, TAB(56); USING "#########"; ReadAmt& * MeterMulti&;
              END IF
              Consump& = ReadAmt& * MeterMulti&
              ThisMeterUseCnt = UBCustRec(1).LocMeters(MCCnt).UseCnt
              IF ThisMeterUseCnt <= 0 THEN ThisMeterUseCnt = 1
              '***
              AvgUse& = Round#((UBCustRec(1).LocMeters(MCCnt).AvgUse / ThisMeterUseCnt) + 0#)
              IF AvgUse& > 0 THEN
                HiConsump& = Round#(AvgUse& * (UBSetUpRec(1).LowRead * .01))
                LowConsump& = Round#(AvgUse& * (UBSetUpRec(1).HighRead * .01))
              END IF
              '***These read backwards
              IF UBCustRec(1).EstFlag = "E" THEN
                PRINT #UBRpt, " E";             'Est. Reading
              ELSEIF Consump& < MINAMT& THEN
                PRINT #UBRpt, " M";             'Minium Usage
              ELSEIF Consump& < LowConsump& THEN
                PRINT #UBRpt, " L";             'Low reading
              ELSEIF Consump& > HiConsump& THEN
                PRINT #UBRpt, " H";             'High Reading
              END IF
              IF UBBillRec(1).RevAmt(TRevCnt) > 0 THEN
                PRINT #UBRpt, TAB(69); USING "######.##"; UBBillRec(1).RevAmt(TRevCnt);
              END IF
              PRINT #UBRpt,
              LineCnt = LineCnt + 1
              IF UBBillRec(1).TaxAmt(TRevCnt) > 0 THEN
                TaxTotals(TRevCnt) = Round#(TaxTotals(TRevCnt) + UBBillRec(1).TaxAmt(TRevCnt))
                PRINT #UBRpt, " Tax"; TAB(69); USING "######.##"; UBBillRec(1).TaxAmt(TRevCnt)
                LineCnt = LineCnt + 1
              END IF
            END IF
          NEXT
          
          '071197 Added this for mccormick. Has a sewer flat rate, Sewer is set up as
          '       a metered service but no meter on a flat rate charge. Rev was added
          '       to total, but didn't show on prebilling report.
          IF NOT DoneOne THEN
            DoneOne = True
            PRINT #UBRpt, TAB(69); USING "######.##"; UBBillRec(1).RevAmt(TRevCnt)
            LineCnt = LineCnt + 1
          END IF
          '***********************************************************************
          
        ELSE    'it's a nonmetered service
          ConsumpTot(TRevCnt, 1) = ConsumpTot(TRevCnt, 1) + 1
          IF WhatRate > 0 THEN
            RateConsump(WhatRate) = RateConsump(WhatRate) + 1
            RateTotals(WhatRate) = Round#(RateTotals(WhatRate) + UBBillRec(1).RevAmt(TRevCnt))
            RTaxTot(WhatRate) = Round#(RTaxTot(WhatRate) + UBBillRec(1).TaxAmt(TRevCnt))
          ELSE
            'STOP
          END IF
          BookConsump(WhatBook).Consump(TRevCnt) = BookConsump(WhatBook).Consump(TRevCnt) + 1
          BookConsump(WhatBook).RevAmt(TRevCnt) = Round#(BookConsump(WhatBook).RevAmt(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
          BookConsump(WhatBook).TaxAmt(TRevCnt) = Round#(BookConsump(WhatBook).TaxAmt(TRevCnt) + UBBillRec(1).TaxAmt(TRevCnt))
          RevTotals(TRevCnt) = Round#(RevTotals(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
          PRINT #UBRpt, TAB(69); USING "######.##"; UBBillRec(1).RevAmt(TRevCnt);
          IF UBBillRec(1).TaxAmt(TRevCnt) > 0 THEN
            TaxTotals(TRevCnt) = Round#(TaxTotals(TRevCnt) + UBBillRec(1).TaxAmt(TRevCnt))
            PRINT #UBRpt,
            PRINT #UBRpt, " Tax"; TAB(69); USING "######.##"; UBBillRec(1).TaxAmt(TRevCnt);
            LineCnt = LineCnt + 1
          END IF
        END IF
        IF NOT DoneOne THEN
          PRINT #UBRpt,
          LineCnt = LineCnt + 1
        END IF
      END IF
      IF (TRevCnt = NumOfRevs) AND UBBillRec(1).TransAmt = 0 THEN
        'CONSUMPTION inactive account
        IF UBBillRec(1).TransAmt = 0 THEN
          FOR TTRevCnt = 1 TO NumOfRevs
            FOR MCCnt = 1 TO 7
              LocMeterType$ = QPTrim$(UBCustRec(1).LocMeters(MCCnt).MTRType)
              IF (LEN(LocMeterType$) > 0) AND (LocMeterType$ = UBCustRec(1).Serv(TTRevCnt).RMTRTYPE) THEN
                IF UBBillRec(1).CurRead(MCCnt) < 0 THEN
                  UBBillRec(1).CurRead(MCCnt) = 0
                END IF
                IF UBBillRec(1).PrevRead(MCCnt) < 0 THEN
                  UBBillRec(1).PrevRead(MCCnt) = 0
                END IF
                CurReadAmt& = UBBillRec(1).CurRead(MCCnt)
                PreReadAmt& = UBBillRec(1).PrevRead(MCCnt)
                IF CurReadAmt& <> PreReadAmt& THEN
                  '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
                  IF NOT ConsumpFlag THEN
                    PRINT #UBRpt, UBCustRec(1).STATUS; "  "; UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB; "  "; UBCustRec(1).CUSTNAME; "  "; QPTrim$(UBCustRec(1).SERVADDR)
                  END IF
                  '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
                  ConsumpFlag = True
                  PRINT #UBRpt, RevDesc(TTRevCnt);
                  'MeterNum$ = QPTrim$(UBCustRec(1).LocMeters(MCCnt).MTRNum)
                  MeterNum$ = QPTrim$(UBCustRec(1).Serv(MCCnt).RATECODE)
                  IF LEN(MeterNum$) > 0 THEN
                    IF UBCustRec(1).LocMeters(MCCnt).NumUser > 1 THEN
                      MeterNum$ = MeterNum$ + "*" + QPTrim$(STR$(UBCustRec(1).LocMeters(MCCnt).NumUser))
                    END IF
                    RSET Temp2$ = MeterNum$
                    PRINT #UBRpt, TAB(14); Temp2$;
                  ELSE
                    RSET Temp2$ = "RATE ERROR"
                    PRINT #UBRpt, TAB(14); Temp2$;
                  END IF
                  PRINT #UBRpt, TAB(30); USING "#########"; CurReadAmt&;
                  PRINT #UBRpt, TAB(42); USING "#########"; PreReadAmt&;
                  ConsumpAmt& = CurReadAmt& - PreReadAmt&
                  ConsumpTot(TTRevCnt, 2) = ConsumpTot(TTRevCnt, 2) + ConsumpAmt&
                  PRINT #UBRpt, TAB(56); USING "#########"; ConsumpAmt&
                  LineCnt = LineCnt + 1
                END IF
              END IF
            NEXT
          NEXT
        END IF
        IF ConsumpFlag AND UBCustRec(1).STATUS <> "A" THEN
          ConsumpFlag = False
          PRINT #UBRpt, "**** Consumption Noted on an Inactive Account. ****"
          PRINT #UBRpt, Dash80$
          LineCnt = LineCnt + 2
          NCCnt = NCCnt + 1
        ELSEIF ConsumpFlag THEN 'Customer Status is "A"
          'This can happen when a cust has consumption and there rate code
          'has a zero calc amount. "i.e. a Church or other nonprofit"
          ConsumpFlag = False
          PRINT #UBRpt, Dash80$
          LineCnt = LineCnt + 1
          NCCnt = NCCnt + 1
        END IF
      ELSEIF (TRevCnt = NumOfRevs) AND UBBillRec(1).TransAmt > 0 THEN
        AcctBalance# = UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance
        PRINT #UBRpt, TAB(5); "Current:";
        PRINT #UBRpt, USING "$$#######,.##"; UBBillRec(1).TransAmt;
        IF AcctBalance# <> 0 THEN
          PRINT #UBRpt, TAB(30); "Previous:";
          PRINT #UBRpt, USING "$$#######,.##"; AcctBalance#;
        END IF
        PRINT #UBRpt, TAB(55); "Total:";
        PRINT #UBRpt, TAB(65); USING "$$#######,.##"; Round#(AcctBalance# + UBBillRec(1).TransAmt)
        LineCnt = LineCnt + 1
        '-=-=-=-=-=-=-=-=-=-=-=-=-=
        PRINT #UBRpt, Dash80$
        LineCnt = LineCnt + 1
      END IF
    NEXT
    
skipem:
    IF AskAbandonPrint% THEN
      UBLog "ABORTED: Prebilling report"
      UBLog "Closing files."
      CLOSE
      AbortFlag = True
      EXIT FOR
    END IF
    
    ShowPctComp Cnt, NumOfRecs
  NEXT
  
  IF AbortFlag THEN
    IF Exist(UBBillsFile) THEN
      KillFile UBBillsFile
    END IF
    GOTO ExitPreReport
  END IF
  
  PRINT #UBRpt, CHR$(12)
  PageNo = PageNo + 1
  PRINT #UBRpt, "Utility Pre-Billing Report"; TAB(70); "Page: "; PageNo
  PRINT #UBRpt, "Date: "; DATE$
  PRINT #UBRpt, "Billing Grand Totals"
  PRINT #UBRpt, "                                        Inactive"
  PRINT #UBRpt, "Revenue/Tax         Consumption        Consumption             Amount"
  PRINT #UBRpt, Dash80$
  
  TotalRevAmt# = 0
  TotalTaxAmt# = 0
  FOR RaCnt = 1 TO NumOfRevs
    PRINT #UBRpt, UBSetUpRec(1).Revenues(RaCnt).REVNAME; TAB(20);
    PRINT #UBRpt, TAB(20); USING "###########"; ConsumpTot(RaCnt, 1);
    PRINT #UBRpt, TAB(40); USING "###########"; ConsumpTot(RaCnt, 2);
    PRINT #UBRpt, TAB(60); USING "#######.##"; RevTotals(RaCnt)
    TotalRevAmt# = Round#(TotalRevAmt# + RevTotals(RaCnt))
    IF TaxTotals(RaCnt) > 0 THEN
      PRINT #UBRpt, QPTrim$(UBSetUpRec(1).Revenues(RaCnt).REVNAME); " Tax"; TAB(20);
      PRINT #UBRpt, TAB(60); USING "#######.##"; TaxTotals(RaCnt)
      TotalTaxAmt# = Round#(TotalTaxAmt# + TaxTotals(RaCnt))
    END IF
  NEXT
  
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(31); "BILLING GRAND TOTAL:"; TAB(57); USING "$$,#######.##"; Round#(TotalTaxAmt# + TotalRevAmt#)
  PRINT #UBRpt, CHR$(12)
  PRINT #UBRpt,
  PRINT #UBRpt, "Report Totals by Rate Code"
  PRINT #UBRpt,
  PRINT #UBRpt, "Code      Rate Description            Consumption              Amount"
  PRINT #UBRpt, Dash80$
  
  TotalRevAmt# = 0
  
  FOR RaCnt = 1 TO NumOfRates
    PRINT #UBRpt, UBRateTbls(RaCnt).RATECODE; "    "; UBRateTbls(RaCnt).RATEDESC;
    PRINT #UBRpt, TAB(39); USING "###########"; RateConsump(RaCnt);
    PRINT #UBRpt, TAB(60); USING "#######.##"; RateTotals(RaCnt)
    TotalRevAmt# = Round#(TotalRevAmt# + RateTotals(RaCnt))
    IF RTaxTot(RaCnt) > 0 THEN
      PRINT #UBRpt, " Tax"; TAB(60); USING "#######.##"; RTaxTot(RaCnt)
    END IF
  NEXT
  
  PRINT #UBRpt,
  PRINT #UBRpt, "TAX TOTAL"; TAB(60); USING "#######.##"; TotalTaxAmt#
  PRINT #UBRpt, TAB(45); "TOTAL:"; TAB(57); USING "$$,#######.##"; Round#(TotalRevAmt# + TotalTaxAmt#)
  PRINT #UBRpt, CHR$(12)
  
  SortT BookConsump(1), TBooks, 0, 306, 0, -1
  
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
        PRINT #UBRpt, RevDesc(RCnt);
        PRINT #UBRpt, TAB(30); USING "#########"; BookConsump(Cnt).Consump(RCnt);
        PRINT #UBRpt, TAB(61); USING "######.##"; BookConsump(Cnt).RevAmt(RCnt)
        TBookAmt# = Round#(TBookAmt# + BookConsump(Cnt).RevAmt(RCnt))
        IF BookConsump(Cnt).TaxAmt(RCnt) > 0 THEN
          PRINT #UBRpt, " Tax"; TAB(61); USING "######.##"; BookConsump(Cnt).TaxAmt(RCnt)
          TBTaxAmt# = Round#(TBTaxAmt# + BookConsump(Cnt).TaxAmt(RCnt))
          LineCnt = LineCnt + 1
        END IF
        LineCnt = LineCnt + 1
      NEXT
      TBookGTot# = Round#(TBookGTot# + TBookAmt# + TBTaxAmt#)
      PRINT #UBRpt, TAB(42); "Book Total:"; TAB(57); USING "$$,#######.##"; Round#(TBookAmt# + TBTaxAmt#)
      IF Cnt < TBooks THEN
        PRINT #UBRpt,
      END IF
      LineCnt = LineCnt + 1
    END IF
    IF LineCnt > MaxLines AND Cnt < TBooks THEN
      PRINT #UBRpt, CHR$(12)
      GOSUB BookHeader
    END IF
SkipThisBook:
  NEXT
  
  PRINT #UBRpt, TAB(35); "Books GRAND Total:"; TAB(57); USING "$$,#######.##"; TBookGTot#
  PRINT #UBRpt, Dash80$
  PRINT #UBRpt, CHR$(12)
  CLOSE
  
  UBLog "Finished writing prebilling report."
  
  SELECT CASE Choice
  CASE 1
    RptText$ = " (Customer Order)"
  CASE 2
    RptText$ = "(Account Order)"
  CASE 3
    RptText$ = "(Location Order)"
  CASE 4
    RptText$ = "(Postal RT. Order)"
  CASE 5
    RptText$ = "(ZipCode Order)"
  CASE 6
    RptText$ = "(Sequence Order)"
  END SELECT
  
  ERASE UBCustRec, UBBillRec, RevTotals, TaxTotals
  ERASE ConsumpTot, RateConsump, BookConsump, UBRateTbls
  ERASE RateConsump, UBSetUpRec, RevDesc
  
  IF NOT AbortFlag THEN
    PrintRptFile "Pre-Billing Report " + RptText$, "UBPREBIL.RPT", LPTPort, RetCode, EntryPoint
    IF BookFlag THEN
      KILL "UBBILLS.DAT"
    END IF
  END IF
  
  '  IF NOT ExitFlag THEN GOTO ReStart
  
  GOTO ExitPreReport
  
PrintPreHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, "Utility Pre-Billing Report.  "; TownName$; TAB(70); "Page: "; PageNo
  PRINT #UBRpt, "Date: "; DATE$
  PRINT #UBRpt, "Location              Customer Name                Service Address"
  PRINT #UBRpt, "Status  No.          Rate      Current     Past"
  PRINT #UBRpt, "Revenue Source       Code      Reading     Reading    Consumption      Charges"
  PRINT #UBRpt, Dash80$
  LineCnt = 6
  RETURN
  
GetWhatBook:
  
  WhatBook = 0
  
  IF LEN(QPTrim$(UBCustRec(1).Book)) = 0 THEN RETURN
  
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
      REDIM PRESERVE BookConsump(1 TO TBooks) AS BookConsumpType
      BookConsump(TBooks).Book = ThisBook
      WhatBook = TBooks
    END IF
  ELSE
    TBooks = TBooks + 1
    BookConsump(TBooks).Book = ThisBook
    WhatBook = TBooks
  END IF
  RETURN
  
BookHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, "Utility Pre-Billing Report"; TAB(70); "Page: "; PageNo
  PRINT #UBRpt, "Date: "; DATE$
  PRINT #UBRpt, "Report Totals by Book"
  PRINT #UBRpt,
  PRINT #UBRpt, "Book"
  PRINT #UBRpt, "Revenue               Consumption                              Amount"
  PRINT #UBRpt, Dash80$
  LineCnt = 6
  RETURN
  
ErrorAbortExit:
  CLOSE
ExitPreReport:
  UBLog "OUT: Prebilling Report" + CrLf$
  
END SUB

SUB UtilAdjustments
  
  SHARED Choice$()
  
  AdjDate = Date2Num(DATE$)
  
  UBLog " IN: UTIL ADJUST"
  
  '**********************************************************
  REDIM TempScrn(0)
  REDIM UBCustRec(1)   AS NewUBCustRecType
  
  UBCustRecLen = LEN(UBCustRec(1))
  
  NumOfRevs = MaxRevsCnt
  
  REDIM RevText$(1 TO MaxRevsCnt)
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  
  TownName$ = UBSetUpRec(1).UTILNAME
  
  FOR RevCnt = 1 TO MaxRevsCnt
    RevText$(RevCnt) = LEFT$(QPTrim$(UBSetUpRec(1).Revenues(RevCnt).REVNAME), 14)
    IF LEN(RevText$(RevCnt)) = 0 THEN
      NumOfRevs = RevCnt - 1
      EXIT FOR
    END IF
  NEXT
  
  IF NumOfRevs < MaxRevsCnt THEN
    REDIM PRESERVE RevText$(1 TO NumOfRevs)
  END IF
  
  LibName$ = "UB"
  ScrnName$ = "UBADJUST"
  
  '--Get the total number of fields from all pages
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  Frm(1).StayOnField = True
  Frm(1).PageDisable = True
  
  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  CustAcctFld = FldNum%("CUSTACCT", Fld())
  AdjTypeFld = FldNum%("ADJTYPE", Fld())
  AdjDateFld = FldNum%("ADJDATE", Fld())
  TotAmtFld = FldNum%("TADJAMT", Fld())
  NoteFld = FldNum%("NOTE", Fld())
  
  REDIM ReadFlds(1 TO 7)
  REDIM CashFlds(1 TO MaxRevsCnt)
  
  FOR Cnt = 1 TO 7
    ReadFlds(Cnt) = FldNum%("READADJ" + LTRIM$(STR$(Cnt)), Fld())
  NEXT
  
  FOR Cnt = 1 TO MaxRevsCnt
    CashFlds(Cnt) = FldNum%("AMTADJ" + LTRIM$(STR$(Cnt)), Fld())
  NEXT
  
  FOR Cnt = NumOfRevs + 1 TO MaxRevsCnt
    Fld(CashFlds(Cnt)).Protected = True
  NEXT
  
  '--define the multi-choice fields
  
  REDIM Choice$(0 TO 2, 0 TO 0)
  
  Choice$(0, 0) = QPTrim$(STR$(AdjTypeFld))
  Choice$(1, 0) = "Upward"
  Choice$(2, 0) = "Downward"
  
  'GOSUB ClearForm
  
  Action = 1
  FirstTime = True
  
  DisplayUBScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF Frm(1).Edited AND Frm(1).PrevFld <> Frm(1).FldNo THEN
      BeenEditedFlag = True     'if the form has been edited
    END IF      'set the edited flag
    
    IF FirstTime THEN
      FirstTime = False         'if this is the first time
      GOSUB ShowRevSources      '
      Action = 2
    END IF
    
    IF Frm(1).FldNo > CustAcctFld AND Frm(1).PrevFld = CustAcctFld THEN
      CustAcct& = QPValL(Form$(CustAcctFld, 0))
      GOSUB GetCustInfo
    ELSEIF Frm(1).FldNo = CustAcctFld AND Frm(1).PrevFld <> CustAcctFld THEN
      MPaintBox 22, 37, 22, 41, 112
      MPaintBox 22, 35, 22, 36, 126
    END IF
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE EscKey
      IF BeenEditedFlag THEN
        SaveFlag = PromptSaveData
        SELECT CASE SaveFlag
        CASE True               'user wants to save
          EscapeFlag = True
          StuffBuf CHR$(0) + CHR$(ABS(F10Key))
        CASE False              'user wants to abandon
          ExitFlag = True
        CASE ELSE               'continue editing
        END SELECT
        Action = 1
      ELSE
        ExitFlag = True
      END IF
      
    CASE F4KEY  'Customer History
      IF CustAcct& > 0 THEN
        ShowCustHistory CustAcct&
        Action = 2
      END IF
      
    CASE F7KEY  'Lookup Customer
      IF Frm(1).FldNo = 1 THEN  'if user is on the Customer field
        SaveScrn TempScrn()     'and F7key then do lookup routine
        MPaintBox 4, 5, 22, 75, 8
        LookUp CustAcct&, "Payment", 2, False, False
        RestScrn TempScrn()
        IF CustAcct& > 0 THEN   'if this is a valid customer
          GOSUB GetCustInfo     'go get customer info
          UBLog "ADJUST:" + STR$(CustAcct&)
          Frm(1).FldNo = 4
          Action = 1
        ELSE
          GOSUB ClearForm
          Frm(1).FldNo = 1
          Action = 1
        END IF
      END IF
      
    CASE F10Key 'Save
      GOSUB CheckAdjustInfo
      IF AdjustOKFlag THEN
        OFldNo = Frm(1).FldNo
        OK = MsgBox%("UB.QSL", "OK2ADJ")        'Ask if they are sure
        SELECT CASE OK
        CASE 1  'Yes do the adjustment
          GOSUB SaveTransaction 'do the save routine
          GOSUB ClearForm
          Frm(1).FldNo = 1
          Action = 1
        CASE ELSE               'Nope continue editing
          Frm(1).FldNo = OFldNo - 1
          Action = 1
        END SELECT
        IF EscapeFlag THEN
          IF CustAcct& > 0 THEN
            UBLog "ABORT ADJUST:" + STR$(CustAcct&)
          END IF
          ExitFlag = True
        END IF
      ELSE
        Action = 2
      END IF
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 22   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 5 TO 13            'f4 History
          PressButton F4KEY, 22, 5, 13
        CASE 34 TO 42           'f7 Look-Up
          PressButton F7KEY, 22, 34, 42
        CASE 54 TO 63           'f10 Save
          PressButton F10Key, 22, 54, 63
        CASE 65 TO 75           '--cancel button
          PressButton EscKey, 22, 65, 75
        END SELECT
      END SELECT                'row
    END IF
  LOOP UNTIL ExitFlag
  
  ERASE TempScrn, UBCustRec, RevText$, UBSetUpRec
  
  HideCursor
  
ExitAdjustments:
  UBLog "OUT: UTIL ADJUST" + CrLf$
  EXIT SUB
  
GetCustInfo:
  NumOfCustRecs& = FileSize("UBCUST.DAT") \ UBCustRecLen
  IF CustAcct& > NumOfCustRecs& OR CustAcct& <= 0 THEN
    CustAcct& = 0
    OK = MsgBox%("UB.QSL", "BADACCTN")
    Frm(1).FldNo = 1
    GOSUB ClearForm
    GOTO SkipCustInfo
  END IF
  
  IF IsDeleted(CustAcct&) THEN
    CustAcct& = 0
    OK = MsgBox%("UB.QSL", "DELACCTN")
    Frm(1).FldNo = 1
    GOSUB ClearForm
    GOTO SkipCustInfo
  END IF
  
  GOSUB ClearForm
  
  CustFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CustFile LEN = UBCustRecLen
  GET CustFile, CustAcct&, UBCustRec(1)
  CLOSE CustFile
  
  FOR Cnt = 1 TO 7
    IF LEN(QPTrim$(UBCustRec(1).LocMeters(Cnt).MTRType)) = 0 THEN
      Fld(ReadFlds(Cnt)).Protected = True
    ELSE
      Fld(ReadFlds(Cnt)).Protected = False
    END IF
    QPrintRC UBCustRec(1).LocMeters(Cnt).MTRNUM, Cnt + 13, 8, -1
    QPrintRC UBCustRec(1).LocMeters(Cnt).MTRType, Cnt + 13, 24, -1
  NEXT
  
  LSET Form$(CustAcctFld, 0) = STR$(CustAcct&)
  LSET Form$(2, 0) = UBCustRec(1).CUSTNAME
  LSET Form$(3, 0) = UBCustRec(1).ADDR1
  SaveField 0, Form$(), Fld(), BadField
  CustAcct& = QPValL(Form$(CustAcctFld, 0))
  
  MPaintBox 22, 35, 22, 41, 115
  FirstTime = True
  
SkipCustInfo:
  Action = 1
  RETURN
  
ShowRevSources:
  FOR RevCnt = 1 TO NumOfRevs
    QPrintRC RevText$(RevCnt), RevCnt + 3, 42, -1
  NEXT
  RETURN
  
ClearForm:
  TempDate = Date2Num(Form$(AdjDateFld, 0))
  
  IF TempDate > 0 AND TempDate <> AdjDate THEN
    AdjDate = TempDate
  END IF
  
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""       '--Clear all fields
  NEXT
  
  LSET Form$(AdjDateFld, 0) = Num2Date(AdjDate)
  
  BeenEditedFlag = False        'clear the edited flag
  FOR Cnt = 1 TO 7
    QPrintRC "            ", Cnt + 13, 8, -1
    QPrintRC " ", Cnt + 13, 24, -1
  NEXT
  
  RETURN
  
SaveTransaction:
  
  SaveScrn TempScrn()
  DisplayUBScrn "UPDATDSK"
  REDIM UBTransRec(1) AS UBTransRecType
  UBTransRecLen = LEN(UBTransRec(1))
  
  TransDate = Date2Num(Form$(AdjDateFld, 0))
  TransAmt# = Value#(Form$(TotAmtFld, 0), ECode)
  
  UBTransRec(1).TransDate = TransDate
  UBTransRec(1).CustLocation = CustAcct&
  UBTransRec(1).CustStatus = UBCustRec(1).STATUS
  UBTransRec(1).CustAcctNo = CustAcct&
  UBTransRec(1).BillMsg = QPTrim$(Form$(NoteFld, 0))
  UBTransRec(1).TransAmt = TransAmt#
  
  FOR Cnt = 1 TO NumOfRevs
    IF LEN(QPTrim$(Form$(CashFlds(Cnt), 0))) > 0 THEN
      UBTransRec(1).RevAmt(Cnt) = Value#(Form$(CashFlds(Cnt), 0), ECode)
    END IF
  NEXT
  
  FOR Cnt = 1 TO 7
    IF LEN(QPTrim$(Form$(ReadFlds(Cnt), 0))) > 0 THEN
      UBTransRec(1).CurRead(Cnt) = Value#(Form$(ReadFlds(Cnt), 0), ECode)
    END IF
  NEXT
  
  AdjType$ = UCASE$(LEFT$(Form$(AdjTypeFld, 0), 1))
  
  SELECT CASE AdjType$
    
  CASE "U"
    UBTransRec(1).TransDesc = "Upward Adjustment"
    FOR RevCnt = 1 TO MaxRevsCnt
      UBCustRec(1).CurrRevAmts(RevCnt) = Round#(UBCustRec(1).CurrRevAmts(RevCnt) + UBTransRec(1).RevAmt(RevCnt))
      UBCustRec(1).CurrBalance = Round#(UBCustRec(1).CurrBalance + UBTransRec(1).RevAmt(RevCnt))
    NEXT
    UBTransRec(1).TransType = TranUpwardAdjustment
    UBLog "UP ADJUSTED CUST:" + STR$(CustAcct&)
  CASE "D"
    UBTransRec(1).TransDesc = "Downward Adjustment"
    FOR RevCnt = 1 TO MaxRevsCnt
      UBCustRec(1).CurrRevAmts(RevCnt) = Round#(UBCustRec(1).CurrRevAmts(RevCnt) - UBTransRec(1).RevAmt(RevCnt))
      UBCustRec(1).CurrBalance = Round#(UBCustRec(1).CurrBalance - UBTransRec(1).RevAmt(RevCnt))
    NEXT
    UBTransRec(1).TransType = TranDownwardAdjustment
    UBLog "DOWN ADJUSTED CUST:" + STR$(CustAcct&)
  END SELECT
  
  UBTransRec(1).RunBalance = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTransRecLen
  
  NextTranRecs& = (LOF(UBTran) \ UBTransRecLen) + 1
  PrevLastTrans& = UBCustRec(1).LastTrans
  UBTransRec(1).PrevTrans = PrevLastTrans&
  UBCustRec(1).LastTrans = NextTranRecs&
  
  IF Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance) = 0 THEN
    IF UBCustRec(1).STATUS = "B" THEN
      CustChCnt = CustChCnt + 1
      UBLog "ADJUST: SET CUST STATUS to I. Acct:" + STR$(UBTransRec(1).CustAcctNo)
      UBCustRec(1).STATUS = "I"
    END IF
  END IF
  
  PUT UBCust, UBTransRec(1).CustAcctNo, UBCustRec(1)
  PUT UBTran, NextTranRecs&, UBTransRec(1)
  CLOSE UBTran, UBCust
  
  UBLog "ADJUSTED CUST:" + STR$(CustAcct&) + "  TRANS:" + STR$(NextTranRecs&)
  
  RestScrn TempScrn()
  DisplayUBScrn "UPDATEOK"
  WaitForAction
  RestScrn TempScrn()
  FirstTime = True
  RETURN
  
CheckAdjustInfo:
  CurrFld = Frm(1).FldNo
  
  AdjustOKFlag = True
  ConsAdjFlag = False
  BillAdjFlag = False
  
  TestDate = Date2Num(Form$(AdjDateFld, 0))
  IF TestDate < 0 THEN
    OK = MsgBox%("UB.QSL", "UBBADTDT")
    Action = 2
    AdjustOKFlag = False
    Frm(1).FldNo = 4
    GOTO BadAdjustment
  END IF
  
  IF LEN(QPTrim$(Form$(AdjTypeFld, 0))) = 0 THEN
    OK = MsgBox%("UB.QSL", "UBBADTYP")
    Action = 2
    AdjustOKFlag = False
    Frm(1).FldNo = 5
    GOTO BadAdjustment
  END IF
  
  FOR Cnt = 1 TO 7
    IF LEN(QPTrim$(Form$(6 + Cnt, 0))) > 0 THEN
      ConsAdjFlag = True
      EXIT FOR
    END IF
  NEXT
  
  FOR Cnt = 1 TO MaxRevsCnt
    IF LEN(QPTrim$(Form$(13 + Cnt, 0))) > 0 THEN
      BillAdjFlag = True
      EXIT FOR
    END IF
  NEXT
  
  IF (NOT ConsAdjFlag AND NOT BillAdjFlag) THEN
    OK = MsgBox%("UB.QSL", "UBBADADJ")
    Action = 2
    AdjustOKFlag = False
    Frm(1).FldNo = CurrFld - 1
    GOTO BadAdjustment
  END IF
  
  
BadAdjustment:
  RETURN
  
END SUB

