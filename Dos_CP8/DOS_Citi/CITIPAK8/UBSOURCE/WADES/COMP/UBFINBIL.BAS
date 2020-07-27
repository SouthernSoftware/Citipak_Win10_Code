
DEFINT A-Z
DECLARE SUB BankDraftRegister ()
DECLARE SUB SetCustFinal ()
DECLARE SUB PrintFinalMenu ()
DECLARE FUNCTION GetUseDeposit% (Text$)
DECLARE SUB PostFinalBillTrans ()
DECLARE SUB MakeFinalBillFile (AbortFlag%, FuelAdjAmt#)
DECLARE SUB ShowCustHistory (CustRec&)
DECLARE SUB FinalMeterReading ()
DECLARE SUB RePrintUtilBills ()
DECLARE SUB PostBillTrans ()
DECLARE FUNCTION GetAdjFactor# ()
DECLARE SUB ClearScrn ()
DECLARE SUB UBBillMenu ()
DECLARE SUB BillPrintMenu ()
DECLARE SUB LoadUBSetUpFile (UBSetupRec() AS ANY, UBSetUpLen%)
DECLARE SUB PrintUtilBills ()
DECLARE SUB MakeSequenceIndex (IndexText$)
DECLARE FUNCTION AskAbandonPrint% ()
  
DECLARE FUNCTION FindRateTbl% (RATECODE$, NumOfRates%, UBRateTbls() AS ANY)
DECLARE FUNCTION GetNumOfRevs% ()
DECLARE FUNCTION GetRevCharge# (RateTbl AS ANY, TMeterConsp&, MeterMulti&)
DECLARE SUB MakeBillFile (AbortFlag%, FuelAdjAmt#)
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
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION WEnvTest% ()
DECLARE SUB ShowCursor ()
DECLARE SUB HideCursor ()
DECLARE SUB KillFile (FileName$)
DECLARE FUNCTION GetNumOfCust% ()
  
'$INCLUDE: 'DefCnf.BI'
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FClose (Handle%)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB SortStr2 (BYVAL Address%, Size%, Direction%)
DECLARE SUB StuffBuf (Key2Stuff$)
DECLARE FUNCTION PromptSaveData% ()
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION Num2Date$ (DateNum%)
  
DECLARE SUB CursorOff ()
DECLARE SUB WaitForAction ()
DECLARE SUB BlockClear ()
DECLARE SUB SaveScrn (ScrnArray())
DECLARE SUB RestScrn (ScrnArray())
DECLARE SUB LookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, LocationFlag%)
  
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
  '$INCLUDE: 'ubsetbal.BI'
  
  CONST False = 0, True = NOT False
  
  REDIM MChoice$(1 TO 7)
  
  MChoice$(1) = "Set a Customer to Final    "
  MChoice$(2) = "Enter Final Meter Readings "
  MChoice$(3) = "Pre-Billing Report (Final) "
  MChoice$(4) = "Print Final Utility Bills  "
  MChoice$(5) = "Bank Draft Register Report"
  MChoice$(6) = "Post Final Billing Transactions"
  MChoice$(7) = "Exit to OS "
  
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
  
  '--Center Menu within Screen
  Row = ((25 - (UBOUND(MChoice$))) \ 2) - 1
  Col = ((80 - MaxLen) \ 2) - 2
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    
    TitleBox 3, Col, MaxLen + 3, "Final Bill Processing", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN
      Choice = 0
      ExitFlag = True
      EXIT DO
    END IF
    
    SELECT CASE Choice
    CASE 1
      SetCustFinal
    CASE 2
      FinalMeterReading
    CASE 3
      PreBillReport
    CASE 4
      PrintFinalMenu
    CASE 5
      BankDraftRegister
    CASE 6
      PostFinalBillTrans
    CASE 7
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP UNTIL ExitFlag
  
  IF WEnvTest% AND NOT INSTR(COMMAND$, "TEST") = 0 THEN
    HideCursor
    ClearScrn
  ELSE
    RUN "UBMENU"
  END IF

SUB FinalMeterReading
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  
  REDIM UBSetupRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetupRec(), UBSetUpLen
  
  REDIM FinalRec(0) AS LONG
  CustRecLen = LEN(UBCustRec(1))
  
  NumOfCust& = GetNumOfCust%
  
  InfoLine$ = SPACE$(10)
  
  IdxRecLen = 4 'we are using a long integer
  IdxFileSize& = FileSize("UBCUSTNM.IDX")
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen

  IF IdxNumOfRecs < 1 THEN EXIT SUB

  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  FGetAH "UBCUSTNM.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it
  
  BlockClear
  ShowProcessingScrn "Scanning Customer Accounts."
  
  FinalRecCnt = 0
  FOpenS "UBCUST.DAT", CustFile
'  FOpenS "UBLOCA.DAT", LocaFile
'halt
  FOR Cnt& = 1 TO IdxNumOfRecs
    FGetRTA CustFile, UBCustRec(1), IdxBuff(Cnt&).RecNum, CustRecLen
    IF UBCustRec(1).STATUS = "F" THEN
      FinalRecCnt = FinalRecCnt + 1
      REDIM PRESERVE FinalRec(0 TO FinalRecCnt) AS LONG
      FinalRec(FinalRecCnt) = IdxBuff(Cnt&).RecNum
      DidAMeterFlag = False
      FOR MtrCnt = 1 TO 7
        IF UBCustRec(1).LocMeters(MtrCnt).ReadFlag <> "F" THEN
          IF UBCustRec(1).LocMeters(MtrCnt).ReadFlag <> "Y" THEN
            MTRType$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MTRType)
            IF LEN(MTRType$) > 0 THEN
              DidAMeterFlag = True
              UBCustRec(1).LocMeters(MtrCnt).PrevRead = UBCustRec(1).LocMeters(MtrCnt).CurRead
              'UBCustRec(1).LocMeters(MtrCnt).CurRead = 0
              UBCustRec(1).LocMeters(MtrCnt).ReadFlag = "F"
            END IF
          END IF
        END IF
      NEXT
      IF DidAMeterFlag THEN
        FPutRTA CustFile, UBCustRec(1), IdxBuff(Cnt&).RecNum, CustRecLen
      END IF
    END IF
    ShowPctComp Cnt&, NumOfCust&
  NEXT
  FClose CustFile
  
  IF FinalRecCnt <= 0 THEN
    BlockClear
    Ok = MsgBox%("UB", "NOFINALS")
    GOTO ExitFinalBillProcessing
  END IF
  
  TotalFinalRec = FinalRecCnt
  
  '*************
  LibName$ = "UB"
  ScrnName$ = "UBFINMTR"
  
  '--define the multi-choice fields
  
  NumFlds = -1
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  MActionRow = 22
  Frm(1).FldNo = 1
  Action = 1
  
  '--Set screen number to one and display screen
  Scr = 1
  DisplayUBScrn ScrnName$
  
  FinalRecCnt = 1
  FirstTime = True
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF Frm(1).Edited AND WasEditedFlag = False THEN
      WasEditedFlag = True
    END IF
    
    IF NOT HasAMeterFlag AND Frm(1).FldNo <> 3 THEN
      Frm(1).FldNo = 3
    END IF
    
    IF FirstTime THEN
      FirstTime = False
      GOSUB GetCustAcctInfo
      GOSUB UpDateDisplay
      Action = 1
    END IF
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE PgUpKey
      IF FinalRecCnt > 1 THEN
        GOSUB ChkWasEdited
        IF OK2Proceed THEN
          OK2Proceed = False
          FinalRecCnt = FinalRecCnt - 1
          GOSUB GetCustAcctInfo
          GOSUB UpDateDisplay
          Frm(1).FldNo = 1
          Action = 1
        END IF
      END IF
      
    CASE PgDnKey
      IF FinalRecCnt < TotalFinalRec THEN
        GOSUB ChkWasEdited
        IF OK2Proceed THEN
          OK2Proceed = False
          FinalRecCnt = FinalRecCnt + 1
          GOSUB GetCustAcctInfo
          GOSUB UpDateDisplay
          Frm(1).FldNo = 1
          Action = 1
        END IF
      END IF
    CASE F4KEY
      ShowCustHistory WhatCust&
      Action = 1
    CASE F10Key
      
      GOSUB SaveFinalData
      Action = 1
      
    CASE EscKey
      GOSUB ChkWasEdited
      IF OK2Proceed THEN
        OK2Proceed = False
        ExitFlag = True
      END IF
    END SELECT
    
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE MActionRow           'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 7 TO 18            '--Cancel button
          PressButton 27, MActionRow, 7, 18
        CASE 20 TO 31           '--F10 button
          PressButton F4KEY, MActionRow, 20, 31
        CASE 41 TO 50           '--F10 button
          PressButton F10Key, MActionRow, 41, 50
        CASE 52 TO 62           '--F5 button
          PressButton PgUpKey, MActionRow, 52, 62
        CASE 64 TO 74           '--F5 button
          PressButton PgDnKey, MActionRow, 64, 74
        END SELECT
      END SELECT                'row
    END IF
    
  LOOP UNTIL ExitFlag = True
  
  GOTO ExitFinalBillProcessing
  
GetCustAcctInfo:
  WhatCust& = FinalRec(FinalRecCnt)
  FOpenS "UBCUST.DAT", CustFile
  FGetRTA CustFile, UBCustRec(1), WhatCust&, CustRecLen
  FClose CustFile
  RETURN
  
UpDateDisplay:
  '--Clear all fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT
  
  LSET InfoLine$ = STR$(FinalRecCnt) + " of" + STR$(TotalFinalRec)
  LSET Form$(1, 0) = STR$(WhatCust&)
  LSET Form$(2, 0) = UBCustRec(1).CustName
  IF UBCustRec(1).LocMeters(1).ReadFlag <> "Y" THEN
    LSET Form$(3, 0) = DATE$
  ELSE
    LSET Form$(3, 0) = Num2Date(UBCustRec(1).LocMeters(1).CurDate)
  END IF
  LSET Form$(4, 0) = UBCustRec(1).Book
  LSET Form$(5, 0) = UBCustRec(1).SEQNUMB
  
  HasAMeterFlag = False
  
  FOR MCnt = 1 TO 7
    LSET Form$((5 * MCnt) + 1, 0) = UBCustRec(1).LocMeters(MCnt).MTRNUM
    IF LEN(QPTrim$(UBCustRec(1).LocMeters(MCnt).MTRType)) > 0 THEN
      HasAMeterFlag = True
      SELECT CASE UBCustRec(1).LocMeters(MCnt).MTRType
      CASE "W"
        LSET Form$((5 * MCnt) + 2, 0) = "Water Only"
      CASE "S"
        LSET Form$((5 * MCnt) + 2, 0) = "Sewer Only"
      CASE "C"
        LSET Form$((5 * MCnt) + 2, 0) = "Combined Water & Sewer"
      CASE "E"
        LSET Form$((5 * MCnt) + 2, 0) = "Electric Meter"
      CASE "D"
        LSET Form$((5 * MCnt) + 2, 0) = "Demand Electric"
      CASE "G"
        LSET Form$((5 * MCnt) + 2, 0) = "Gas Meter"
      CASE "T"
        LSET Form$((5 * MCnt) + 2, 0) = "Touch Read Meter"
        'CASE ELSE
        '  LSET Form$((5 * MCnt) + 2, 0) = "UNKNOWN"
      END SELECT
      SELECT CASE UBCustRec(1).LocMeters(MCnt).MTRUnit
      CASE "G"
        LSET Form$((5 * MCnt) + 3, 0) = "Gallons"
      CASE "C"
        LSET Form$((5 * MCnt) + 3, 0) = "Cubic Feet"
      CASE "K"
        LSET Form$((5 * MCnt) + 3, 0) = "Kilowatts"
      END SELECT
      
      IF UBCustRec(1).LocMeters(MCnt).CurRead > 0 THEN
        LSET Form$((5 * MCnt) + 4, 0) = STR$(UBCustRec(1).LocMeters(MCnt).CurRead)
      ELSE
        LSET Form$((5 * MCnt) + 4, 0) = STR$(0)
      END IF
      IF UBCustRec(1).LocMeters(MCnt).PrevRead > 0 THEN
        LSET Form$((5 * MCnt) + 5, 0) = STR$(UBCustRec(1).LocMeters(MCnt).PrevRead)
      ELSE
        LSET Form$((5 * MCnt) + 5, 0) = STR$(0)
      END IF
      
    END IF
  NEXT
  
  IF NOT HasAMeterFlag THEN
    LSET Form$(6, 0) = "No Metered"
    LSET Form$(11, 0) = "Services"
  END IF
  
  QPrintRC InfoLine$, 5, 65, -1
  RETURN
  
ChkWasEdited:
  IF WasEditedFlag THEN
    SaveFlag = PromptSaveData
    SELECT CASE SaveFlag
    CASE True   'user wants to save
      
      GOSUB SaveFinalData
      WasEditedFlag = False
      OK2Proceed = True
    CASE False  'user wants to abandon
      WasEditedFlag = False
      OK2Proceed = True
    CASE ELSE
      'continue editing
    END SELECT
    Action = 1
  ELSE
    OK2Proceed = True
  END IF
  RETURN
  
SaveFinalData:
  FinalReadDate = Date2Num(Form$(3, 0))
  FOpenS "UBCUST.DAT", CustFile
  FGetRTA CustFile, UBCustRec(1), WhatCust&, CustRecLen
  FOR MCnt = 1 TO 7
    UBCustRec(1).LocMeters(MCnt).CurRead = Value(Form$((5 * MCnt) + 4, 0), ECode)
    UBCustRec(1).LocMeters(MCnt).PrevRead = Value(Form$((5 * MCnt) + 5, 0), ECode)
    UBCustRec(1).LocMeters(MCnt).ReadFlag = "Y"
    UBCustRec(1).LocMeters(MCnt).CurDate = FinalReadDate
  NEXT
  FPutRTA CustFile, UBCustRec(1), WhatCust&, CustRecLen
  FClose CustFile
  
  WasEditedFlag = False
  OK2Proceed = True
  Frm(1).FldNo = 1
  RETURN
  
  '***************
ExitFinalBillProcessing:
  ERASE UBCustRec
  ERASE UBSetupRec, FinalRec, IdxBuff
END SUB

FUNCTION FindRateTbl% (RATECODE$, NumOfRates, UBRateTbls() AS UBRateTblRecType)

  RATECODE$ = QPTrim$(RATECODE$)
  IF LEN(RATECODE$) THEN        'if this rev has a rate code
    FOR TCnt = 1 TO NumOfRates  'find the right one
      IF RATECODE$ = QPTrim$(UBRateTbls(TCnt).RATECODE) THEN
        EXIT FOR
      END IF
    NEXT
  END IF

  IF TCnt > NumOfRates THEN
    FindRateTbl = 0
  ELSE
    FindRateTbl = TCnt
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
  
END FUNCTION

FUNCTION GetNumOfRevs%
  NumOfRevs = 15
  REDIM UBSetupRec(1) AS UBSetupRecType
  UBSetUpLen = LEN(UBSetupRec(1))
  FGetAH "UBSETUP.DAT", UBSetupRec(1), UBSetUpLen, 1            'load it
  FOR RevCnt = 1 TO 15
    TempRev$ = QPTrim$(UBSetupRec(1).Revenues(RevCnt).REVNAME)
    IF LEN(TempRev$) = 0 THEN
      NumOfRevs = RevCnt - 1
      EXIT FOR
    END IF
  NEXT
  GetNumOfRevs = NumOfRevs
  
END FUNCTION

FUNCTION GetRevCharge# (RateTbl AS UBRateTblRecType, TMeterConsp&, MeterMulti&)
  
  MinBillAmt# = RateTbl.MINAmt
  
  
  LastTblCnt = 10
  FOR BCnt = 1 TO 10
    IF RateTbl.TblBreaks(BCnt).UNITAMT <= 0 THEN
      LastTblCnt = BCnt - 1
      EXIT FOR
    END IF
  NEXT
  
  MeterConsump& = TMeterConsp&  ' - RateTbl.MINUNITS
  
  'IF MeterConsump& > 500000 THEN STOP
  
  TAmt# = 0
  
  'Break 1
  '  IF LastTblCnt = 1 THEN
  '    UNITS& = Round(MeterConsump& - RateTbl.TblBreaks(1).UNITS)' / 1000
  '    TAmt# = Round(TAmt# + (UNITS& * RateTbl.TblBreaks(1).UNITAMT))
  '    GOTO GotTAmt:
  '  END IF
  
  IF LastTblCnt >= 2 THEN
    IF MeterConsump& >= RateTbl.TblBreaks(1).UNITS AND MeterConsump& < RateTbl.TblBreaks(2).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(1).UNITS)     ' / 1000
      TAmt# = Round(TAmt# + (UNITS& * RateTbl.TblBreaks(1).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(2).UNITS - RateTbl.TblBreaks(1).UNITS)        ' / 1000
      TAmt# = Round(TAmt# + (UNITS& * RateTbl.TblBreaks(1).UNITAMT))
    END IF
  ELSE          'no other rate breaks
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(1).UNITS)       ' / 1000
    TAmt# = Round(TAmt# + (UNITS& * RateTbl.TblBreaks(1).UNITAMT))
    GOTO GotTAmt
  END IF
  
  'Break 2
  IF LastTblCnt >= 3 THEN
    IF MeterConsump& > RateTbl.TblBreaks(2).UNITS AND MeterConsump& < RateTbl.TblBreaks(3).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(2).UNITS)     ' / 1000
      TAmt# = Round(TAmt# + (UNITS& * RateTbl.TblBreaks(2).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(3).UNITS - RateTbl.TblBreaks(2).UNITS)        ' / 1000
      TAmt# = Round(TAmt# + (UNITS& * RateTbl.TblBreaks(2).UNITAMT))
    END IF
  ELSE
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(2).UNITS)       ' / 1000
    TAmt# = Round(TAmt# + (UNITS& * RateTbl.TblBreaks(2).UNITAMT))
    GOTO GotTAmt
  END IF
  
  'Break 3
  IF LastTblCnt >= 4 THEN
    IF MeterConsump& > RateTbl.TblBreaks(3).UNITS AND MeterConsump& < RateTbl.TblBreaks(4).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(3).UNITS)     ' / 1000
      TAmt# = Round(TAmt# + (UNITS& * RateTbl.TblBreaks(3).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(4).UNITS - RateTbl.TblBreaks(3).UNITS)        ' / 1000
      TAmt# = Round(TAmt# + (UNITS& * RateTbl.TblBreaks(3).UNITAMT))
    END IF
  ELSE
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(3).UNITS)       ' / 1000
    TAmt# = Round(TAmt# + (UNITS& * RateTbl.TblBreaks(3).UNITAMT))
    GOTO GotTAmt
  END IF
  
  'Break 4
  IF LastTblCnt >= 5 THEN
    IF MeterConsump& > RateTbl.TblBreaks(4).UNITS AND MeterConsump& < RateTbl.TblBreaks(5).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(4).UNITS)     ' / 1000
      TAmt# = Round(TAmt# + (UNITS& * RateTbl.TblBreaks(4).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(5).UNITS - RateTbl.TblBreaks(4).UNITS)        ' / 1000
      TAmt# = Round(TAmt# + (UNITS& * RateTbl.TblBreaks(4).UNITAMT))
    END IF
  ELSE
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(4).UNITS)       ' / 1000
    TAmt# = Round(TAmt# + (UNITS& * RateTbl.TblBreaks(4).UNITAMT))
    GOTO GotTAmt
  END IF
  
  'break 5
  IF LastTblCnt >= 6 THEN
    IF MeterConsump& > RateTbl.TblBreaks(5).UNITS AND MeterConsump& < RateTbl.TblBreaks(6).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(5).UNITS)     ' / 1000
      TAmt# = Round(TAmt# + (UNITS& * RateTbl.TblBreaks(5).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(6).UNITS - RateTbl.TblBreaks(5).UNITS)        ' / 1000
      TAmt# = Round(TAmt# + (UNITS& * RateTbl.TblBreaks(5).UNITAMT))
    END IF
  ELSE
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(5).UNITS)       ' / 1000
    TAmt# = Round(TAmt# + (UNITS& * RateTbl.TblBreaks(5).UNITAMT))
    GOTO GotTAmt
  END IF
  
  'break 6
  IF LastTblCnt >= 7 THEN
    IF MeterConsump& > RateTbl.TblBreaks(6).UNITS AND MeterConsump& < RateTbl.TblBreaks(7).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(6).UNITS)     ' / 1000
      TAmt# = Round(TAmt# + (UNITS& * RateTbl.TblBreaks(6).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(6).UNITS - RateTbl.TblBreaks(6).UNITS)        ' / 1000
      TAmt# = Round(TAmt# + (UNITS& * RateTbl.TblBreaks(6).UNITAMT))
    END IF
  ELSE
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(6).UNITS)       ' / 1000
    TAmt# = Round(TAmt# + (UNITS& * RateTbl.TblBreaks(6).UNITAMT))
    GOTO GotTAmt
  END IF
  
  'break 7
  IF LastTblCnt >= 8 THEN
    IF MeterConsump& > RateTbl.TblBreaks(7).UNITS AND MeterConsump& < RateTbl.TblBreaks(8).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(7).UNITS)     ' / 1000
      TAmt# = Round(TAmt# + (UNITS& * RateTbl.TblBreaks(7).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(7).UNITS - RateTbl.TblBreaks(7).UNITS)        ' / 1000
      TAmt# = Round(TAmt# + (UNITS& * RateTbl.TblBreaks(7).UNITAMT))
    END IF
  ELSE
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(7).UNITS)       ' / 1000
    TAmt# = Round(TAmt# + (UNITS& * RateTbl.TblBreaks(7).UNITAMT))
    GOTO GotTAmt
  END IF
  
  'break 8
  IF LastTblCnt >= 9 THEN
    IF MeterConsump& > RateTbl.TblBreaks(8).UNITS AND MeterConsump& < RateTbl.TblBreaks(9).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(8).UNITS)     ' / 1000
      TAmt# = Round(TAmt# + (UNITS& * RateTbl.TblBreaks(8).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(8).UNITS - RateTbl.TblBreaks(8).UNITS)        ' / 1000
      TAmt# = Round(TAmt# + (UNITS& * RateTbl.TblBreaks(8).UNITAMT))
    END IF
  ELSE
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(8).UNITS)       ' / 1000
    TAmt# = Round(TAmt# + (UNITS& * RateTbl.TblBreaks(8).UNITAMT))
    GOTO GotTAmt
  END IF
  
  'break 9
  IF LastTblCnt >= 10 THEN
    IF MeterConsump& > RateTbl.TblBreaks(9).UNITS AND MeterConsump& < RateTbl.TblBreaks(10).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(9).UNITS)     ' / 1000
      TAmt# = Round(TAmt# + (UNITS& * RateTbl.TblBreaks(9).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(9).UNITS - RateTbl.TblBreaks(9).UNITS)        ' / 1000
      TAmt# = Round(TAmt# + (UNITS& * RateTbl.TblBreaks(9).UNITAMT))
    END IF
  ELSE
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(9).UNITS)       ' / 1000
    TAmt# = Round(TAmt# + (UNITS& * RateTbl.TblBreaks(9).UNITAMT))
    GOTO GotTAmt
  END IF
  
  'Else consumption is greater than break 9 so bill the remainder 10th rate
  '  UNITS& = (MeterConsump& - RateTbl.TblBreaks(10).UNITS)  ' / 1000
  '  TAmt# = Round(TAmt# + (UNITS& * RateTbl.TblBreaks(10).UNITAmt))
  
  
GotTAmt:
  GetRevCharge# = Round(MinBillAmt# + TAmt#)
  
  'EXIT FUNCTION
  
  
END FUNCTION

FUNCTION GetUseDeposit% (Text$)
  
  LibName$ = "UB"
  ScrnName$ = "ASKUSEDP"
  
  '--define the multi-choice fields
  
  NumFlds = -1
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  UseDepositFlag = -100
  
  MActionRow = 15
  Frm(1).FldNo = 1
  Action = 1
  
  '--Set screen number to one and display screen
  BlockClear
  DisplayUBScrn ScrnName$
  
  StartPos = 25
  Text$ = QPTrim$(Text$)
  BracketPos = 25 + LEN(Text$)
  QPrintRC Text$, 8, StartPos, 14
  QPrintRC " ]", 8, BracketPos, 10
  
  FinalRecCnt = 1
  FirstTime = True
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      FirstTime = False
      LSET Form$(1, 0) = "N"
      Action = 1
    END IF
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key, EnterKey
      IF Form$(1, 0) = "Y" THEN
        GetUseDeposit% = True
        UseDepositFlag = True
      ELSE
        GetUseDeposit% = False
        UseDepositFlag = False
      END IF
      ExitFlag = True
    CASE EscKey
      GetUseDeposit% = -2
      ExitFlag = True
    END SELECT
    
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE MActionRow           'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 27 TO 38           '--Cancel button
          PressButton 27, MActionRow, 27, 38
        CASE 40 TO 52           '--F10 button
          PressButton F10Key, MActionRow, 40, 52
        END SELECT
      END SELECT                'row
    END IF
    
  LOOP UNTIL ExitFlag = True
  
  SELECT CASE UseDepositFlag
  CASE True
    OPEN "UBDEPFLG.DAT" FOR RANDOM SHARED AS #1 LEN = 2
    PUT #1, , UseDepositFlag
    CLOSE #1
  CASE False
    OPEN "UBDEPFLG.DAT" FOR RANDOM SHARED AS #1 LEN = 2
    PUT #1, , UseDepositFlag
    CLOSE #1
  CASE -100
  END SELECT
  
  
END FUNCTION

SUB MakeFinalBillFile (AbortFlag%, FuelAdjAmt#)
  
  BlockClear
  ShowProcessingScrn "Calculating Utility Charges."
  
  REDIM UBSetUp(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUp(), UBSetUpLen
  
  'find the electric revenue position
  IF FuelAdjAmt# <> 0 THEN
    FOR ThisRevCnt = 1 TO 15
      IF INSTR(UBSetUp(1).Revenues(ThisRevCnt).REVNAME, "ELECTRIC") THEN
        ElecRev = ThisRevCnt
        EXIT FOR
      END IF
    NEXT
  ELSE
    ElecRev = -1
  END IF
  
  REDIM UBBillRec(1) AS UBTransRecType
  'REDIM UBLocatRec(1) AS UBLocationRecType
  REDIM UBCustRec(1) AS NewUBCustRecType
  
  UBBillRecLen = LEN(UBBillRec(1))
  UBCustRecLen = LEN(UBCustRec(1))
  ' UBLocatRecLen = LEN(UBLocatRec(1))
  
  NumOfRates = GetNumRateRecs%
  
  REDIM UBRateTbls(1 TO NumOfRates) AS UBRateTblRecType
  UBRateTblRecLen = LEN(UBRateTbls(1))
  
  RateFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS RateFile LEN = UBRateTblRecLen
  FOR Cnt = 1 TO NumOfRates
    GET RateFile, Cnt, UBRateTbls(Cnt)
  NEXT
  CLOSE
  
  NumOfCustRecs& = FileSize&("UBCUST.DAT") \ UBCustRecLen
  
  IF Exist(UBFinBillsFile) THEN
    KILL UBFinBillsFile
  END IF
  
  BillFile = FREEFILE
  OPEN UBFinBillsFile FOR RANDOM SHARED AS BillFile LEN = UBBillRecLen
  CustFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CustFile LEN = UBCustRecLen
  
  BillCnt = 0
  NumOfRevs = GetNumOfRevs%
  
  FOR LCnt = 1 TO NumOfCustRecs&
    GET CustFile, LCnt, UBCustRec(1)

    'IF LCnt = 3207 THEN STOP

    REDIM UBBillRec(1) AS UBTransRecType
    IF UBCustRec(1).STATUS = "F" THEN
      GotCustFlag = True
    ELSE
      GotCustFlag = False
    END IF
    MeterConsp& = 0
    TMeterConsp& = 0
    
    IF NOT GotCustFlag THEN
      UBBillRec(1).TransAmt = 0
      FOR RCnt = 1 TO NumOfRevs
        UBBillRec(1).RevAmt(RCnt) = 0
      NEXT
      UBBillRec(1).ActiveFlag = False
      GOTO NotAFinal
    END IF

    'IF INSTR(UBCustRec(1).CustName, "B P") THEN STOP

    'Process Flat Rate Amounts
    FOR FRCnt = 1 TO 4
      WhatService = UBCustRec(1).FlatRates(FRCnt).REVSRC
      IF UBCustRec(1).FlatRates(FRCnt).FRAMT > 0 AND WhatService > 0 THEN
        '11/19/96 Fixed Rev. amt. to add to current amt
        'UBBillRec(1).RevAmt(WhatService) = Round#(UBBillRec(1).RevAmt(WhatService) + UBCustRec(1).FlatRates(FRCnt).FRAmt)
        '01-09-97 Fixed Multiplier problem in flat rates
        Multi = UBCustRec(1).FlatRates(FRCnt).NumMin
        IF Multi < 1 THEN Multi = 1
        UBBillRec(1).RevAmt(WhatService) = Round#(UBCustRec(1).FlatRates(FRCnt).FRAMT * Multi)
        UBBillRec(1).TransAmt = Round(UBBillRec(1).TransAmt + UBBillRec(1).RevAmt(WhatService))
        IF UBSetUp(1).Revenues(WhatService).TAXRATE > 0 AND UBCustRec(1).TAXEXPT <> "Y" THEN
          UBBillRec(1).TaxAmt(WhatService) = Round(UBBillRec(1).RevAmt(WhatService) * UBSetUp(1).Revenues(WhatService).TAXRATE)
          UBBillRec(1).TransAmt = Round(UBBillRec(1).TransAmt + UBBillRec(1).TaxAmt(WhatService))
        END IF
      END IF
    NEXT


'************
    FOR RCnt = 1 TO NumOfRevs   'look at each rev line
      MeterConsp& = 0
      TMeterConsp& = 0
      WhatTbl = FindRateTbl(UBCustRec(1).Serv(RCnt).RATECODE, NumOfRates, UBRateTbls())
      IF WhatTbl THEN           'if this is a non-metered service
        IF UBSetUp(1).Revenues(RCnt).UseMtr = "N" THEN
          '02-05-97 added fix add to current rev amt
          'UBBillRec(1).RevAmt(RCnt) = UBRateTbls(WhatTbl).MINAMT
          UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + UBRateTbls(WhatTbl).MINAmt)
          UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + UBRateTbls(WhatTbl).MINAmt)
          'UBBillRec(1).ActiveFlag = True
          GOTO GotAmt
        END IF

        MeterType$ = UBCustRec(1).Serv(RCnt).RMTRTYPE

        WasRead = False
        FOR MCnt = 1 TO 7
          IF MeterType$ = UBCustRec(1).LocMeters(MCnt).MTRType THEN
            'Found correct meter
            IF (INSTR(UBCustRec(1).CUSTTYPE, "MON") > 0) THEN
              MeterConsp& = UBCustRec(1).LocMeters(MCnt).CurRead - UBCustRec(1).LocMeters(MCnt).PrevRead
              IF MeterConsp& < 0 THEN
                MaxMeterAmt& = 10& ^ (LEN(STR$(UBCustRec(1).LocMeters(MCnt).PrevRead)) - 1)
                MeterConsp& = (MaxMeterAmt& - UBCustRec(1).LocMeters(MCnt).PrevRead) + UBCustRec(1).LocMeters(MCnt).CurRead
              END IF
            ELSEIF (UBCustRec(1).LocMeters(MCnt).ReadFlag = "Y") THEN

              GOSUB ChkForPrevBill

              IF HasPrevBill THEN
                WasRead = True
              ELSE
                'LPRINT LCnt, UBCustRec(1).CustName
                'This customer is new and is being read THIS month
                'so we will not set the WasRead flag here.
                'this will force the calc below not adjust for the est part!!
              END IF
              
              EstAvgUse& = UBCustRec(1).LocMeters(MCnt).AvgUse
              MeterConsp& = UBCustRec(1).LocMeters(MCnt).CurRead - UBCustRec(1).LocMeters(MCnt).PrevRead
            
            ELSE                'they are estimated
              IF UBCustRec(1).LocMeters(MCnt).AvgUse > 0 THEN
                UBCustRec(1).LocMeters(MCnt).CurRead = UBCustRec(1).LocMeters(MCnt).PrevRead + UBCustRec(1).LocMeters(MCnt).AvgUse
                MeterConsp& = UBCustRec(1).LocMeters(MCnt).CurRead - UBCustRec(1).LocMeters(MCnt).PrevRead
              END IF
              UBBillRec(1).ESTREAD(MCnt) = "Y"
            END IF

            'Meter multi??
            IF UBCustRec(1).LocMeters(MCnt).MTRMulti > 1 THEN
              MeterConsp& = MeterConsp& * UBCustRec(1).LocMeters(MCnt).MTRMulti
            END IF
            UBBillRec(1).CurRead(MCnt) = UBCustRec(1).LocMeters(MCnt).CurRead
            UBBillRec(1).PrevRead(MCnt) = UBCustRec(1).LocMeters(MCnt).PrevRead
            'UBBillRec(1).MtrTypes(MCnt) = GetCustMeterType(UBCustRec(), MCnt)
            'convert units here if necessary
            TMeterConsp& = TMeterConsp& + MeterConsp&
          END IF
        NEXT

        IF UBCustRec(1).LocMeters(RCnt).NumUser > 1 THEN
          'adjust consumption to amt for one user
          NumUser& = UBCustRec(1).LocMeters(RCnt).NumUser
          TMeterConsp& = TMeterConsp& / NumUser&
        ELSE
          NumUser& = 1
        END IF

'STOP
'------------------------------------
        IF WasRead THEN
          CurUsage& = TMeterConsp&
          IF EstAvgUse& <= 0 THEN EstAvgUse& = 0
          EstUsage& = EstAvgUse& * PrevBillCnt
          Usage& = CurUsage& - EstUsage&

          IF Usage& < 0 THEN
            AvgMthUsage& = CurUsage& / (PrevBillCnt + 1)
            IF AvgMthUsage& <= UBRateTbls(WhatTbl).MinUnits THEN
              MthAmtDue# = UBRateTbls(WhatTbl).MINAmt
            ELSE
              MthAmtDue# = GetRevCharge#(UBRateTbls(WhatTbl), AvgMthUsage&, MeterMulti&)
            END IF
            IF EstAvgUse& <= UBRateTbls(WhatTbl).MinUnits THEN
              EstAmtPaid# = UBRateTbls(WhatTbl).MINAmt
            ELSE
              EstAmtPaid# = GetRevCharge#(UBRateTbls(WhatTbl), EstAvgUse&, MeterMulti&)
            END IF
            TotalActAmtDue# = Round#(MthAmtDue# * (PrevBillCnt + 1))
            TotalEstAmtPaid# = Round#(EstAmtPaid# * PrevBillCnt)
            CurrentAmtToCharge# = Round#(TotalActAmtDue# - TotalEstAmtPaid#)
            UBBillRec(1).RevAmt(RCnt) = CurrentAmtToCharge#
            UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + UBBillRec(1).RevAmt(RCnt))
          ELSE
            IF Usage& <= UBRateTbls(WhatTbl).MinUnits THEN
              'if we bill the minium
              UBBillRec(1).RevAmt(RCnt) = NumUser& * UBRateTbls(WhatTbl).MINAmt
              UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + UBBillRec(1).RevAmt(RCnt))
              GOTO GotAmt
            END IF
            UBBillRec(1).RevAmt(RCnt) = GetRevCharge#(UBRateTbls(WhatTbl), Usage&, MeterMulti&)
            UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + UBBillRec(1).RevAmt(RCnt))
          END IF

'------------------------------------

'******************************
'        IF WasRead THEN
'          CurUsage& = TMeterConsp&
'          EstUsage& = EstAvgUse& * 2
'          Usage& = CurUsage& - EstUsage&
'          IF Usage& < 0 THEN
'            AvgMthUsage& = CurUsage& / 3
'            IF AvgMthUsage& <= UBRateTbls(WhatTbl).MinUnits THEN
'              MthAmtDue# = UBRateTbls(WhatTbl).MINAmt
'            ELSE
'              MthAmtDue# = GetRevCharge#(UBRateTbls(WhatTbl), AvgMthUsage&, MeterMulti&)
'            END IF
'            IF EstAvgUse& <= UBRateTbls(WhatTbl).MinUnits THEN
'              EstAmtPaid# = UBRateTbls(WhatTbl).MINAmt
'            ELSE
'              EstAmtPaid# = GetRevCharge#(UBRateTbls(WhatTbl), EstAvgUse&, MeterMulti&)
'            END IF
'            TotalActAmtDue# = Round#(MthAmtDue# * 3)
'            TotalEstAmtPaid# = Round#(EstAmtPaid# * 2)
'            CurrentAmtToCharge# = Round#(TotalActAmtDue# - TotalEstAmtPaid#)
'            UBBillRec(1).RevAmt(RCnt) = CurrentAmtToCharge#
'            UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + UBBillRec(1).RevAmt(RCnt))
'          ELSE
'            IF Usage& <= UBRateTbls(WhatTbl).MinUnits THEN
'              'if we bill the minium
'              UBBillRec(1).RevAmt(RCnt) = NumUser& * UBRateTbls(WhatTbl).MINAmt
'              UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + UBBillRec(1).RevAmt(RCnt))
'              GOTO GotAmt
'            END IF
'            UBBillRec(1).RevAmt(RCnt) = GetRevCharge#(UBRateTbls(WhatTbl), Usage&, MeterMulti&)
'            UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + UBBillRec(1).RevAmt(RCnt))
'          END IF
'****************


        ELSE
          '01-20-97 Add Fix for minium units test for equal to also.
          IF TMeterConsp& <= UBRateTbls(WhatTbl).MinUnits THEN
            'if we bill the minium
            UBBillRec(1).RevAmt(RCnt) = NumUser& * UBRateTbls(WhatTbl).MINAmt
            UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + UBBillRec(1).RevAmt(RCnt))
            GOTO GotAmt
          END IF
          UBBillRec(1).RevAmt(RCnt) = GetRevCharge#(UBRateTbls(WhatTbl), TMeterConsp&, MeterMulti&)
          UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + UBBillRec(1).RevAmt(RCnt))
        END IF

        IF RCnt = ElecRev THEN
          UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + Round#(UBBillRec(1).RevAmt(RCnt) * FuelAdjAmt#))
        END IF

GotAmt:
        IF UBSetUp(1).Revenues(RCnt).TAXRATE > 0 AND UBCustRec(1).TAXEXPT <> "Y" THEN
          UBBillRec(1).TaxAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) * UBSetUp(1).Revenues(RCnt).TAXRATE)
          UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + UBBillRec(1).TaxAmt(RCnt))
        END IF
      END IF
    NEXT        'loop through all revenue sources

    
    IF UBCustRec(1).STATUS = "I" AND UBBillRec(1).TransAmt > 0 THEN
      UBBillRec(1).TransAmt = 0
      FOR RCnt = 1 TO NumOfRevs
        UBBillRec(1).RevAmt(RCnt) = 0
      NEXT
      UBBillRec(1).ActiveFlag = True
    ELSE
      UBBillRec(1).ActiveFlag = False
    END IF
    
    'IF UBBillRec(1).TransAmt > 0 THEN
    BillCnt = BillCnt + 1
    UBBillRec(1).ActiveFlag = True
    UBBillRec(1).CustAcctNo = LCnt

    UBBillRec(1).CustCode = MKI$(PrevBillCnt)
    
NotAFinal:
    PUT BillFile, LCnt, UBBillRec(1)
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
    ShowPctComp LCnt, NumOfCustRecs&
  NEXT
  
  CLOSE

  ERASE UBBillRec, UBCustRec, UBSetUp, UBRateTbls
  EXIT SUB

ChkForPrevBill:

  'IF LCnt = 3207 THEN STOP
  
  HasPrevBill = False
  PrevBillCnt = 0
  EBillCnt = 0
  ThisTrans& = UBCustRec(1).LastTrans
  IF ThisTrans& > 0 THEN
    REDIM UBTransRec(1) AS UBTransRecType
    TranFile = FREEFILE
    OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS TranFile LEN = UBBillRecLen
    DO WHILE ThisTrans& > 0
      GET #TranFile, ThisTrans&, UBTransRec(1)
      IF UBTransRec(1).TransType = TranUtilityBill THEN
        HasPrevBill = True
        PrevBillCnt = PrevBillCnt + 1

        FOR EECnt = 1 TO 7
          IF UBTransRec(1).ESTREAD(EECnt) = "Y" THEN
            EBillCnt = EBillCnt + 1
            EXIT FOR
          END IF
        NEXT

        IF PrevBillCnt <> EBillCnt THEN
          PrevBillCnt = EBillCnt
          EXIT DO
        END IF

        IF PrevBillCnt >= 2 THEN
          PrevBillCnt = EBillCnt
          EXIT DO
        END IF
      END IF
      ThisTrans& = UBTransRec(1).PrevTrans
    LOOP
    CLOSE TranFile
  END IF

RETURN
  
END SUB

SUB PostFinalBillTrans
  
  IF NOT Exist(UBFinBillsFile) THEN
    CursorOff
    BlockClear
    DisplayUBScrn "NON2POST"
    WaitForAction
    GOTO ExitBillPost:
  END IF
  
  IF NOT Exist("UBFBILLS.PRN") THEN
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
    GOTO ExitBillPost
  END IF
  
  REDIM UBBillRec(1) AS UBTransRecType
  REDIM UBCustRec(1) AS NewUBCustRecType
  
  UBBillRecLen = LEN(UBBillRec(1))
  UBCustRecLen = LEN(UBCustRec(1))
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  UBBill = FREEFILE
  OPEN UBFinBillsFile FOR RANDOM SHARED AS UBBill LEN = UBBillRecLen
  
  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBBillRecLen
  
  NumOfBillRecs = LOF(UBBill) \ UBBillRecLen
  ShowProcessingScrn "Posting Final Billing Transactions"
  '*****************************************************

  FOR BillCnt = 1 TO NumOfBillRecs
    GET UBBill, BillCnt, UBBillRec(1)
    IF UBBillRec(1).ActiveFlag THEN             'AND UBBillRec(1).TransAmt > 0 THEN
      GET UBCust, UBBillRec(1).CustAcctNo, UBCustRec(1)
      UBCustRec(1).STATUS = "B"
      UBCustRec(1).PrevBalance = Round(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
      UBCustRec(1).CurrBalance = UBBillRec(1).TransAmt
      UBBillRec(1).RunBalance = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
      
      IF UBBillRec(1).ApplyDepFlag = "Y" THEN
        DepAppliedFlag = False
        DepositAmt# = UBCustRec(1).DepositAmt
        DepTranAmt# = -UBCustRec(1).DepositAmt
        FOR RevCnt = 1 TO MaxRevsCnt - 1
          UBCustRec(1).CurrRevAmts(RevCnt) = Round(UBCustRec(1).CurrRevAmts(RevCnt) + UBBillRec(1).RevAmt(RevCnt) + UBBillRec(1).TaxAmt(RevCnt))
          IF DepositAmt# > 0 THEN
            DepAppliedFlag = True
            IF UBCustRec(1).CurrRevAmts(RevCnt) < DepositAmt# THEN
              DepositAmt# = Round(DepositAmt# - UBCustRec(1).CurrRevAmts(RevCnt))
              UBCustRec(1).CurrRevAmts(RevCnt) = 0
            ELSEIF UBCustRec(1).CurrRevAmts(RevCnt) > DepositAmt# THEN
              UBCustRec(1).CurrRevAmts(RevCnt) = Round(UBCustRec(1).CurrRevAmts(RevCnt) - DepositAmt#)
              DepositAmt# = 0
            ELSE                'the deposit and the revenue are equal
              UBCustRec(1).CurrRevAmts(RevCnt) = 0
              DepositAmt# = 0
            END IF
          END IF
        NEXT
        'If there was any deposit left after applying to the cust rev totals
        IF DepositAmt# > 0 THEN
          UBCustRec(1).CurrRevAmts(1) = -DepositAmt#
        END IF
        UBCustRec(1).DepositAmt = 0
      ELSE
        FOR RevCnt = 1 TO MaxRevsCnt
          UBCustRec(1).CurrRevAmts(RevCnt) = Round(UBCustRec(1).CurrRevAmts(RevCnt) + UBBillRec(1).RevAmt(RevCnt) + UBBillRec(1).TaxAmt(RevCnt))
        NEXT
      END IF
      UBBillRec(1).TransType = TranUtilityBill  'set transaction to Type 1
      FOR MtrCnt = 1 TO 7
        IF UBCustRec(1).LocMeters(MtrCnt).CurRead > 0 THEN
          ReadAmt& = UBBillRec(1).CurRead(MtrCnt) - UBBillRec(1).PrevRead(MtrCnt)
          IF ReadAmt& < 0 THEN  'Meter rolled over or, misread
            MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(MtrCnt))) - 1)
            ReadAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(MtrCnt)) + UBBillRec(1).CurRead(MtrCnt)
          END IF
          UBCustRec(1).LocMeters(MtrCnt).AvgUse = Round(UBCustRec(1).LocMeters(MtrCnt).AvgUse + ReadAmt&)
          UBCustRec(1).LocMeters(MtrCnt).UseCnt = UBCustRec(1).LocMeters(MtrCnt).UseCnt + 1
          UBCustRec(1).LocMeters(MtrCnt).ReadFlag = ""
        END IF
      NEXT
      
      PrevLastTrans& = UBCustRec(1).LastTrans
      UBBillRec(1).PrevTrans = PrevLastTrans&
      NextTransRec& = (LOF(UBTran) \ UBBillRecLen) + 1          'point at next trans
      PUT UBTran, NextTransRec&, UBBillRec(1)
      UBCustRec(1).LastTrans = NextTransRec&
      PUT UBCust, UBBillRec(1).CustAcctNo, UBCustRec(1)
'040997 added Transaction to show customers applied deposit
      IF DepAppliedFlag THEN
        GOSUB MakeAppDepTrans
      END IF
    END IF
    ShowPctComp BillCnt, NumOfBillRecs
  NEXT
  
  CLOSE
  
  KillFile UBFinBillsFile
  KillFile "UBFBILLS.PRN"
  
  BlockClear
  DisplayUBScrn "UPDATEOK"
  WaitForAction
  
ExitBillPost:
  EXIT SUB

MakeAppDepTrans:
  REDIM UBTransRec(1) AS UBTransRecType
  UBTransRecLen = LEN(UBTransRec(1))

  UBTransRec(1).TransDate = UBBillRec(1).TransDate
  UBTransRec(1).CustLocation = UBBillRec(1).CustAcctNo
  UBTransRec(1).CustStatus = UBCustRec(1).STATUS
  UBTransRec(1).CustAcctNo = UBBillRec(1).CustAcctNo
  UBTransRec(1).TransAmt = DepTranAmt#

  FOR Cnt = 1 TO 15
    UBTransRec(1).RevAmt(Cnt) = 0
  NEXT

  UBTransRec(1).TransDesc = "Applied Deposit"
  UBTransRec(1).TransType = TranAppliedDeposit
  UBTransRec(1).RunBalance = Round#((UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance) - ABS(DepTranAmt#))
  UBCustRec(1).DepositAmt = 0
  UBCustRec(1).CurrBalance = Round#(UBCustRec(1).CurrBalance - ABS(DepTranAmt#))

  PrevLastTrans& = UBCustRec(1).LastTrans
  UBTransRec(1).PrevTrans = PrevLastTrans&
  NextTransRec& = (LOF(UBTran) \ UBTransRecLen) + 1          'point at next trans
  PUT UBTran, NextTransRec&, UBTransRec(1)
  UBCustRec(1).LastTrans = NextTransRec&
  PUT UBCust, UBTransRec(1).CustAcctNo, UBCustRec(1)

RETURN

END SUB

SUB PreBillReport
  
  Dash80$ = STRING$(80, "-")
  Temp1$ = SPACE$(10)
  Temp2$ = SPACE$(12)
  
  NumOfRevs = MaxRevsCnt        'assume max munber of revenue sources
  
  'load setup file
  REDIM UBSetupRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetupRec(), UBSetUpLen
  
  REDIM RevDesc(1 TO MaxRevsCnt) AS STRING * 12
  FOR Cnt = 1 TO MaxRevsCnt     'find last active revenue
    TempRev$ = QPTrim$(UBSetupRec(1).Revenues(Cnt).REVNAME)
    IF LEN(TempRev$) = 0 THEN
      NumOfRevs = Cnt - 1       'set actual number of revenues
      EXIT FOR
    ELSE
      LSET RevDesc(Cnt) = UCASE$(TempRev$)
      IF INSTR(RevDesc(Cnt), "ELECTRIC") THEN
        DoFuelAdjFlag = True
      END IF
      'build revenue description line
    END IF
  NEXT
  
  NumOfRates = GetNumRateRecs%
  IF NumOfRates = 0 THEN EXIT SUB

  REDIM UBRateTbls(1 TO NumOfRates) AS UBRateTblRecType
  REDIM RateConsump(1 TO NumOfRates) AS LONG
  
  UBRateTblRecLen = LEN(UBRateTbls(1))
  
  RateFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS RateFile LEN = UBRateTblRecLen
  FOR Cnt = 1 TO NumOfRates
    GET RateFile, Cnt, UBRateTbls(Cnt)
  NEXT
  CLOSE
  
  MaxLines = 42
  
  REDIM MChoice$(1 TO 6)
  
  MChoice$(1) = "Customer Name Order"
  MChoice$(2) = "Account Number Order"
  MChoice$(3) = "Location Number Order"
  MChoice$(4) = "Postal Carrier Route Order"
  MChoice$(5) = "ZipCode Order"
  
  IF UBSetupRec(1).UseSeq = "Y" THEN
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
  
Restart:
  '--Center Menu within Screen
  Row = ((25 - (UBOUND(MChoice$))) \ 2)
  Col = ((80 - MaxLen) \ 2) - 2
  'Help$ = "Utility Billing Main Menu"
  
  UsingBook = False
  UsingAcct = False
  UsingName = False
  
  AbortFlag = False
  PageNo = 0
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    TitleBox 3, Col, MaxLen + 3, "Pre-Billing Report (Final)", Cnf
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
      UsingName = True
      OKFlag = True
    CASE 2
      IndexName$ = ""
      UsingAcct = True
      OKFlag = True
    CASE 3
      IndexName$ = BookIndexFile
      UsingBook = True
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
  
  DepositFlag = GetUseDeposit(MChoice$(Choice))

  'DepositFlag = False

  IF DepositFlag = -2 THEN GOTO Restart
  
  IF DoFuelAdjFlag THEN
    FuelAdjAmt# = GetAdjFactor#
  ELSE
    FuelAdjAmt# = 0
  END IF
  
  IF FuelAdjAmt# = -10000 THEN GOTO Restart
  
  MakeFinalBillFile AbortFlag, FuelAdjAmt#
  
  IF AbortFlag THEN GOTO ExitPreReport
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBBillRec(1) AS UBTransRecType
  UBBillRecLen = LEN(UBBillRec(1))
  
  REDIM RevTotals(1 TO NumOfRevs)  AS DOUBLE    'holds each revenues total amt
  REDIM ConsumpTot(1 TO NumOfRevs, 1 TO 2) AS DOUBLE            'holds each revenues total amt
  REDIM RateConsump(1 TO NumOfRates) AS LONG
  REDIM BookConsump(1 TO 1) AS BookConsumpType
  
  TBooks = 0
  'Stop$ = CHR$(27)
  
  IF UsingAcct THEN             'load the index
    NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
  ELSE
    NumOfRecs = FileSize(IndexName$) \ 4
    REDIM IndexArray(1 TO NumOfRecs) AS UBCustIndexRecType
    FGetAH IndexName$, IndexArray(1), 4, NumOfRecs
  END IF
  
  UBBill = FREEFILE
  OPEN UBFinBillsFile FOR RANDOM SHARED AS UBBill LEN = UBBillRecLen
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  UBRpt = FREEFILE
  OPEN UBFinPreRptFile FOR OUTPUT AS UBRpt
  
  BlockClear
  ShowProcessingScrn "Processing Pre-Billing Report"
  
  GOSUB PrintPreHeader
  
  FOR Cnt = 1 TO NumOfRecs
    IF UsingAcct THEN
      WhatRec& = Cnt
    ELSE
      WhatRec& = IndexArray(Cnt).RecNum
    END IF
    
    GET UBCust, WhatRec&, UBCustRec(1)
    
    IF UBCustRec(1).STATUS = "F" THEN
      
      GET UBBill, WhatRec&, UBBillRec(1)
      'get deposit amt 1
      IF DepositFlag THEN
        IF UBCustRec(1).DepositAmt > 0 THEN
          CustDepAmt# = UBCustRec(1).DepositAmt
        ELSE
          CustDepAmt# = 0
        END IF
      ELSE
        CustDepAmt# = 0
      END IF
      UBBillRec(1).TaxAmt(15) = CustDepAmt#
      PUT UBBill, WhatRec&, UBBillRec(1)
      '*****************
      IF LineCnt > MaxLines THEN
        PRINT #UBRpt, CHR$(12)
        GOSUB PrintPreHeader
      END IF
      
      GOSUB GetWhatBook
      BookConsump(WhatBook).CustCnt = BookConsump(WhatBook).CustCnt + 1
      PRINT #UBRpt, UBCustRec(1).STATUS; "  "; UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB; "  "; UBCustRec(1).CustName; "  "; UBCustRec(1).SERVADDR
      
      DoneOne = False
      FOR TRevCnt = 1 TO NumOfRevs
        IF UBBillRec(1).RevAmt(TRevCnt) <> 0 THEN
          DoneOne = False
          PRINT #UBRpt, RevDesc(TRevCnt);
          
          FOR TRateCnt = 1 TO NumOfRates
            IF UBRateTbls(TRateCnt).RATECODE = UBCustRec(1).Serv(TRevCnt).RATECODE THEN
              MINAmt& = UBRateTbls(TRateCnt).MinUnits
              WhatRate = TRateCnt
              EXIT FOR
            END IF
          NEXT
          
          IF UBSetupRec(1).Revenues(TRevCnt).UseMtr = "Y" THEN
            RevTotals(TRevCnt) = Round#(RevTotals(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
            BookConsump(WhatBook).RevAmt(TRevCnt) = Round#(BookConsump(WhatBook).RevAmt(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
            
            FOR MCCnt = 1 TO 7
              LocMeterType$ = QPTrim$(UBCustRec(1).LocMeters(MCCnt).MTRType)
              MeterMulti& = UBCustRec(1).LocMeters(MCCnt).MTRMulti
              IF MeterMulti& <= 0 THEN MeterMulti& = 1
              IF (LEN(LocMeterType$) > 0) AND (LocMeterType$ = UBCustRec(1).Serv(TRevCnt).RMTRTYPE) THEN
                DoneOne = True
                MeterNum$ = QPTrim$(UBCustRec(1).LocMeters(MCCnt).MTRNUM)
                IF LEN(MeterNum$) > 0 THEN
                  RSET Temp2$ = MeterNum$
                  PRINT #UBRpt, TAB(14); Temp2$;
                END IF
                ReadAmt& = UBBillRec(1).CurRead(MCCnt)
                PRINT #UBRpt, TAB(30); USING "#########"; ReadAmt&;
                ReadAmt& = UBBillRec(1).PrevRead(MCCnt)
                PRINT #UBRpt, TAB(42); USING "#########"; ReadAmt&;
                ReadAmt& = UBBillRec(1).CurRead(MCCnt) - UBBillRec(1).PrevRead(MCCnt)
                IF ReadAmt& > 0 THEN
                  RateConsump(WhatRate) = RateConsump(WhatRate) + ReadAmt&
                  BookConsump(WhatBook).Consump(TRevCnt) = BookConsump(WhatBook).Consump(TRevCnt) + ReadAmt&
                ELSEIF ReadAmt& < 0 THEN
                  'Meter has rolled over or, been misread
'REMED THIS
                  'MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(MCCnt))) - 1)

                  ReadAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(MCCnt)) + UBBillRec(1).CurRead(MCCnt)
                  RateConsump(WhatRate) = RateConsump(WhatRate) + ReadAmt&
                END IF
                ConsumpTot(TRevCnt, 1) = ConsumpTot(TRevCnt, 1) + ReadAmt&
                PRINT #UBRpt, TAB(56); USING "#########"; ReadAmt& * MeterMulti&;
                Consump& = ReadAmt& * MeterMulti&
                '***091996
                ' LowConsump& = Round(UBCustRec(1).LocMeters(MCCnt).AvgUse * (UBSetUpRec(1).LOWREAD * .01))
                ' HiConsump& = Round(UBCustRec(1).LocMeters(MCCnt).AvgUse * (UBSetUpRec(1).HIGHREAD * .01))
                ThisMeterUseCnt = UBCustRec(1).LocMeters(MCCnt).UseCnt
                IF ThisMeterUseCnt <= 0 THEN ThisMeterUseCnt = 1
                '***
                AvgUse& = Round((UBCustRec(1).LocMeters(MCCnt).AvgUse / ThisMeterUseCnt) + 0#)
                IF AvgUse& > 0 THEN
                  HiConsump& = Round(AvgUse& * (UBSetupRec(1).LowRead * .01))
                  LowConsump& = Round(AvgUse& * (UBSetupRec(1).HighRead * .01))
                END IF
                '***These read backwards
                IF Consump& < MINAmt& THEN
                  PRINT #UBRpt, " M";
                ELSEIF Consump& < LowConsump& THEN
                  PRINT #UBRpt, " L";
                ELSEIF Consump& > HiConsump& THEN
                  PRINT #UBRpt, " H";
                END IF
                
                'IF UBBillRec(1).RevAmt(TRevCnt) > 0 THEN
                IF UBBillRec(1).RevAmt(TRevCnt) <> 0 THEN
                  PRINT #UBRpt, TAB(67); USING "$$######.##"; UBBillRec(1).RevAmt(TRevCnt);
                END IF
                PRINT #UBRpt,
                LineCnt = LineCnt + 1
                IF UBBillRec(1).TaxAmt(TRevCnt) > 0 THEN
                  PRINT #UBRpt, "TAX"; TAB(67); USING "$$######.##"; UBBillRec(1).TaxAmt(TRevCnt)
                  LineCnt = LineCnt + 1
                END IF
              END IF
            NEXT
          ELSE
            ConsumpTot(TRevCnt, 1) = ConsumpTot(TRevCnt, 1) + 1
            RateConsump(WhatRate) = RateConsump(WhatRate) + 1
            BookConsump(WhatBook).Consump(TRevCnt) = BookConsump(WhatBook).Consump(TRevCnt) + 1
            BookConsump(WhatBook).RevAmt(TRevCnt) = Round(BookConsump(WhatBook).RevAmt(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
            RevTotals(TRevCnt) = Round(RevTotals(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
            PRINT #UBRpt, TAB(67); USING "$$######.##"; UBBillRec(1).RevAmt(TRevCnt);
            IF UBBillRec(1).TaxAmt(TRevCnt) > 0 THEN
              PRINT #UBRpt,
              PRINT #UBRpt, "TAX"; TAB(67); USING "$$######.##"; UBBillRec(1).TaxAmt(TRevCnt);
              LineCnt = LineCnt + 1
            END IF
          END IF
          IF NOT DoneOne THEN
            PRINT #UBRpt,
            LineCnt = LineCnt + 1
          END IF
        END IF
        
        IF (TRevCnt = NumOfRevs) THEN
          TotalAmt# = Round(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance + UBBillRec(1).TransAmt)
          PRINT #UBRpt, TAB(45); "        Current: ";
          PRINT #UBRpt, TAB(65); USING "$$#######,.##"; UBBillRec(1).TransAmt
          LineCnt = LineCnt + 1
          
          IF UBCustRec(1).PrevBalance <> 0 OR UBCustRec(1).CurrBalance <> 0 THEN
            PRINT #UBRpt, TAB(45); "Balance Forward: ";
            PRINT #UBRpt, TAB(65); USING "$$#######,.##"; Round(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
            LineCnt = LineCnt + 1
          END IF
          
          PRINT #UBRpt, TAB(50); "Bill Total: ";
          PRINT #UBRpt, TAB(65); USING "$$#######,.##"; TotalAmt#
          
          IF DepositFlag THEN
            PRINT #UBRpt, TAB(48); "Less Deposit:";
            PRINT #UBRpt, TAB(65); USING "$$#######,.##"; -UBCustRec(1).DepositAmt
            BalanceAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
            SELECT CASE SGN(BalanceAmt#)
            CASE -1
              PRINT #UBRpt, TAB(50); "Refund Due:";
              PRINT #UBRpt, TAB(65); USING "$$#######,.##"; ABS(BalanceAmt#)
            CASE 0
              PRINT #UBRpt, TAB(50); "       Due:";
              PRINT #UBRpt, TAB(65); USING "$$#######,.##"; 0
            CASE 1
              PRINT #UBRpt, TAB(49); "Balance Due:";
              PRINT #UBRpt, TAB(65); USING "$$#######,.##"; BalanceAmt#
            END SELECT
            LineCnt = LineCnt + 3
          ELSE
            'PRINT #UBRpt, TAB(55); "Total: ";
            'PRINT #UBRpt, TAB(65); USING "$$#######,.##"; UBBillRec(1).TransAmt
            LineCnt = LineCnt + 1
          END IF
        END IF
      NEXT
      PRINT #UBRpt, Dash80$
      LineCnt = LineCnt + 1
    END IF
    
    
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
    ShowPctComp Cnt, NumOfRecs
  NEXT
  
  PRINT #UBRpt, CHR$(12)
  PageNo = PageNo + 1
  PRINT #UBRpt, "Utility Pre-Billing Report"; TAB(70); "Page: "; PageNo
  PRINT #UBRpt, "Date: "; DATE$
  PRINT #UBRpt, "Report Totals"
  PRINT #UBRpt, "                                        Inactive"
  PRINT #UBRpt, "Revenue             Consumption        Consumption             Amount"
  PRINT #UBRpt, Dash80$
  
  FOR RaCnt = 1 TO NumOfRevs
    PRINT #UBRpt, UBSetupRec(1).Revenues(RaCnt).REVNAME; TAB(20);
    PRINT #UBRpt, TAB(20); USING "###########"; ConsumpTot(RaCnt, 1);
    PRINT #UBRpt, TAB(40); USING "###########"; ConsumpTot(RaCnt, 2);
    PRINT #UBRpt, TAB(60); USING "#######.##"; RevTotals(RaCnt)
  NEXT
  PRINT #UBRpt,
  PRINT #UBRpt, "Consumption Totals by Rate Code"
  PRINT #UBRpt, Dash80$
  
  FOR RaCnt = 1 TO NumOfRates
    PRINT #UBRpt, UBRateTbls(RaCnt).RATECODE; "    "; UBRateTbls(RaCnt).RATEDESC;
    PRINT #UBRpt, TAB(59); USING "###########"; RateConsump(RaCnt)
  NEXT
  PRINT #UBRpt, CHR$(12)
  
  SortT BookConsump(1), TBooks, 0, 186, 0, -1
  
  GOSUB BookHeader
  
  FOR Cnt = 1 TO TBooks
    IF BookConsump(Cnt).Book < 10 THEN
      Book$ = "0" + QPTrim$(STR$(BookConsump(Cnt).Book))
    ELSE
      Book$ = QPTrim$(STR$(BookConsump(Cnt).Book))
    END IF
    PRINT #UBRpt, "Book: "; Book$; "    Customers:"; BookConsump(Cnt).CustCnt
    FOR RCnt = 1 TO NumOfRevs
      PRINT #UBRpt, RevDesc(RCnt);
      PRINT #UBRpt, TAB(30); USING "#########"; BookConsump(Cnt).Consump(RCnt);
      PRINT #UBRpt, TAB(59); USING "$$######.##"; BookConsump(Cnt).RevAmt(RCnt)
      LineCnt = LineCnt + 1
    NEXT
    PRINT #UBRpt,
    LineCnt = LineCnt + 1
    IF LineCnt > MaxLines AND Cnt < TBooks THEN
      PRINT #UBRpt, CHR$(12)
      GOSUB BookHeader
    END IF
  NEXT
  PRINT #UBRpt, Dash80$
  PRINT #UBRpt, CHR$(12)
  CLOSE
  
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
  
  ERASE RevTotals, ConsumpTot, RateConsump, BookConsump, UBRateTbls
  ERASE RateConsump, UBSetupRec, UBCustRec, UBBillRec, IndexArray
  
  
  IF NOT AbortFlag THEN
    PrintRptFile "Pre-Billing Report " + RptText$, UBFinPreRptFile, LPTPort, RetCode, EntryPoint
  END IF
  
  'IF NOT ExitFlag THEN GOTO Restart
  EXIT SUB
  
PrintPreHeader:
  PageNo = PageNo + 1
  
  PRINT #UBRpt, "Utility Pre-Billing Report"; TAB(70); "Page: "; PageNo
  PRINT #UBRpt, "Date: "; DATE$
  PRINT #UBRpt, "Location              Customer Name                Service Address"
  PRINT #UBRpt, "Status  No.        Meter       Current     Past"
  PRINT #UBRpt, "Revenue Source     Number      Reading     Reading    Consumption      Charges"
  PRINT #UBRpt, Dash80$
  LineCnt = 6
  
  RETURN
  
GetWhatBook:
  
  WhatBook = 0
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
  
ExitPreReport:
  
END SUB

