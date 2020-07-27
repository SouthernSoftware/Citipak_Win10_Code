DEFINT A-Z
DECLARE SUB MakeMowZipCodeIndex (IndexText$)
DECLARE FUNCTION GetCustMeterType% (UBCustRec() AS ANY, ThisMeter%)
DECLARE SUB SetCustFinal ()
DECLARE SUB UBLog (Text$)
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
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetUpLen%)
DECLARE SUB PrintUtilBills ()
DECLARE SUB MakeSequenceIndex (IndexText$)
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE FUNCTION QPValL& (Long$)
DECLARE FUNCTION FindRateTbl% (RateCode$, NumOfRates%, UBRateTbls() AS ANY)
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
  
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
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
  
  '$INCLUDE: 'DefCnf.BI'
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
  
  REDIM MChoice$(1 TO 6)
  
  MChoice$(1) = "Set a Customer to Final    "
  MChoice$(2) = "Enter Final Meter Readings "
  MChoice$(3) = "Pre-Billing Report (Final) "
  MChoice$(4) = "Print Final Utility Bills  "
  MChoice$(5) = "Post Final Bill Transactions"
  MChoice$(6) = "Exit to OS "
  
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
  Row = ((25 - (UBOUND(MChoice$))) \ 2)
  Col = ((80 - MaxLen) \ 2) - 1
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    
    TitleBox 2, Col, MaxLen + 3, "Final Bill Processing", Cnf
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
      SetCustFinal              'unrem
    CASE 2
      FinalMeterReading
    CASE 3
      PreBillReport
    CASE 4
      'PrintFinalMenu
      RUN "UBPRNFIN"
    CASE 5
      PostFinalBillTrans
    CASE 6
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
  
  '062598
  'Corrected final meter reading entry to not move read date after they have
  'have allready been moved once.

SUB FinalMeterReading
  
  UBLog " IN: Final meter readings."
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen
  
  IF INSTR(UBSetUpRec(1).UTILNAME, "LANDIS") > 0 THEN
    LandisFlag = True
  END IF
  
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
  FOR Cnt& = 1 TO IdxNumOfRecs
    FGetRTA CustFile, UBCustRec(1), IdxBuff(Cnt&).RecNum, CustRecLen
    IF UBCustRec(1).Status = "F" THEN
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
  
  UBLog " FINAL READINGS: Total Count:" + STR$(TotalFinalRec)
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
  LSET Form$(2, 0) = UBCustRec(1).CUSTNAME
  IF UBCustRec(1).LocMeters(1).ReadFlag <> "Y" THEN
    LSET Form$(3, 0) = DATE$
  ELSE
    LSET Form$(3, 0) = Num2Date(UBCustRec(1).LocMeters(1).CurDate)
  END IF
  LSET Form$(4, 0) = UBCustRec(1).BOOK
  LSET Form$(5, 0) = UBCustRec(1).SEQNUMB
  
  QPrintRC "               ", 8, 58, 15
  QPrintRC "Acct:" + STR$(WhatCust&), 8, 58, 15
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
  QPrintRC UBCustRec(1).SERVADDR, 6, 23, -1
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
    '120897 Added
    '062598 Corrected!!
    'Was moving the read dates if they edited the customer again
    IF UBCustRec(1).LocMeters(MCnt).ReadFlag = "F" THEN
      UBCustRec(1).LocMeters(MCnt).PastDate = UBCustRec(1).LocMeters(MCnt).CurDate
      UBCustRec(1).LocMeters(MCnt).ReadFlag = "Y"
    END IF
    '******************
    UBCustRec(1).LocMeters(MCnt).CurDate = FinalReadDate
  NEXT
  FPutRTA CustFile, UBCustRec(1), WhatCust&, CustRecLen
  
  IF NOT LandisFlag THEN
    OldRec& = UBCustRec(1).OldRec
    IF OldRec& > 0 AND OldRec& <= NumOfCust& THEN
      FGetRTA CustFile, UBCustRec(1), OldRec&, CustRecLen
      FOR MCnt = 1 TO 7
        UBCustRec(1).LocMeters(MCnt).PrevRead = UBCustRec(1).LocMeters(MCnt).CurRead
        UBCustRec(1).LocMeters(MCnt).CurRead = Value(Form$((5 * MCnt) + 4, 0), ECode)
        'UBCustRec(1).LocMeters(MCnt).PrevRead = Value(Form$((5 * MCnt) + 5, 0), ECode)
        UBCustRec(1).LocMeters(MCnt).ReadFlag = "N"
        UBCustRec(1).LocMeters(MCnt).PastDate = UBCustRec(1).LocMeters(MCnt).CurDate
        UBCustRec(1).LocMeters(MCnt).CurDate = FinalReadDate
      NEXT
      FPutRTA CustFile, UBCustRec(1), OldRec&, CustRecLen
    END IF
  END IF
  
  FClose CustFile
  UBLog "UPDATED READINGS:" + STR$(WhatCust&) + " P: " + Num2Date$(UBCustRec(1).LocMeters(1).PastDate) + " C: " + Num2Date$(FinalReadDate)
  WasEditedFlag = False
  OK2Proceed = True
  Frm(1).FldNo = 1
  
  RETURN
  
  '***************
ExitFinalBillProcessing:
  ERASE UBCustRec
  ERASE UBSetUpRec, FinalRec, IdxBuff
  UBLog "OUT: Final meter readings."
END SUB

FUNCTION FindRateTbl% (RateCode$, NumOfRates, UBRateTbls() AS UBRateTblRecType)
  
  RateCode$ = QPTrim$(RateCode$)
  IF LEN(RateCode$) THEN        'if this rev has a rate code
    FOR TCnt = 1 TO NumOfRates  'find the right one
      IF RateCode$ = QPTrim$(UBRateTbls(TCnt).RateCode) THEN
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
  
END FUNCTION

FUNCTION GetRevCharge# (RateTbl AS UBRateTblRecType, TMeterConsp&, MeterMulti&)
  
  MinBillAmt# = RateTbl.MINAMT
  
  'IF LEFT$(RateTbl.RateDesc, 1) = "E" THEN STOP
  
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
    IF MeterConsump& >= RateTbl.TblBreaks(1).UNITS AND MeterConsump& <= RateTbl.TblBreaks(2).UNITS THEN
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
    IF MeterConsump& > RateTbl.TblBreaks(2).UNITS AND MeterConsump& <= RateTbl.TblBreaks(3).UNITS THEN
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
    IF MeterConsump& > RateTbl.TblBreaks(3).UNITS AND MeterConsump& <= RateTbl.TblBreaks(4).UNITS THEN
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
    IF MeterConsump& > RateTbl.TblBreaks(4).UNITS AND MeterConsump& <= RateTbl.TblBreaks(5).UNITS THEN
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
    IF MeterConsump& > RateTbl.TblBreaks(5).UNITS AND MeterConsump& <= RateTbl.TblBreaks(6).UNITS THEN
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
    IF MeterConsump& > RateTbl.TblBreaks(6).UNITS AND MeterConsump& <= RateTbl.TblBreaks(7).UNITS THEN
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
    IF MeterConsump& > RateTbl.TblBreaks(7).UNITS AND MeterConsump& <= RateTbl.TblBreaks(8).UNITS THEN
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
    IF MeterConsump& > RateTbl.TblBreaks(8).UNITS AND MeterConsump& <= RateTbl.TblBreaks(9).UNITS THEN
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
    IF MeterConsump& > RateTbl.TblBreaks(9).UNITS AND MeterConsump& <= RateTbl.TblBreaks(10).UNITS THEN
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
  
  IF INSTR(UBSetUp(1).UTILNAME, "PRINCETON") > 0 THEN
    PrinceFlag = True
  END IF
  
  IF INSTR(UBSetUp(1).UTILNAME, "WADE") > 0 THEN
    WadeFlag = True
  END IF
  
  IF INSTR(UBSetUp(1).UTILNAME, "YADKIN") > 0 THEN
    YadkinFlag = True
  END IF
  
  IF INSTR(UBSetUp(1).UTILNAME, "NORWOOD") > 0 THEN
    NorwoodFlag = True
  END IF
  
  IF INSTR(UBSetUp(1).UTILNAME, "CONWAY") > 0 THEN
    ConwayFlag = True
  END IF

  IF INSTR(UBSetUp(1).UTILNAME, "ELKTON") > 0 THEN
    ElkFlag = True
  END IF

  IF INSTR(UBSetUp(1).UTILNAME, "SCOTTSBURG") > 0 THEN
    ScottFlag = True
  END IF

  
  '111698 Prorate
  REDIM ProrateServ(1 TO 15) AS INTEGER
  
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
  
  '111698 Prorate
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

'STOP
'    IF LCnt = 10668 THEN
'      STOP
'    END IF
    
    REDIM UBBillRec(1) AS UBTransRecType
    IF UBCustRec(1).Status = "F" THEN
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
    
    '111698 Prorate
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
    
    FOR FRCnt = 1 TO 4
      WhatService = UBCustRec(1).FlatRates(FRCnt).REVSRC
      IF UBCustRec(1).FlatRates(FRCnt).FRAMT > 0 AND WhatService > 0 THEN
        '11/19/96 Fixed Rev. amt. to add to current amt
        '01-09-97 Fixed Multiplier problem in flat rates
        Multi = UBCustRec(1).FlatRates(FRCnt).NumMin
        IF Multi < 1 THEN Multi = 1
        FlatAmt# = Round#(UBCustRec(1).FlatRates(FRCnt).FRAMT * Multi)
        '111698 Prorate
        IF ProrateFlag AND ProrateServ(WhatService) THEN
          FlatAmt# = Round#(FlatAmt# * ProPct#)
        END IF
        UBBillRec(1).RevAmt(WhatService) = Round#(UBBillRec(1).RevAmt(WhatService) + FlatAmt#)
        UBBillRec(1).TransAmt = Round(UBBillRec(1).TransAmt + FlatAmt#)
        IF UBSetUp(1).Revenues(WhatService).TAXRATE > 0 AND UBCustRec(1).TAXEXPT <> "Y" THEN
          UBBillRec(1).TaxAmt(WhatService) = Round(UBBillRec(1).RevAmt(WhatService) * UBSetUp(1).Revenues(WhatService).TAXRATE)
          UBBillRec(1).TransAmt = Round(UBBillRec(1).TransAmt + UBBillRec(1).TaxAmt(WhatService))
        END IF
      END IF
    NEXT
    
    FOR RCnt = 1 TO NumOfRevs   'look at each rev line
      MeterConsp& = 0
      TMeterConsp& = 0
      WhatTbl = FindRateTbl(UBCustRec(1).Serv(RCnt).RateCode, NumOfRates, UBRateTbls())
      IF WhatTbl THEN
        IF UBSetUp(1).Revenues(RCnt).UseMtr = "N" THEN
          'if this is a non-metered service
          '111398 Prorate
          NonMAmt# = UBRateTbls(WhatTbl).MINAMT
          IF ProrateFlag AND ProrateServ(RCnt) THEN
            NonMAmt# = Round#(NonMAmt# * ProPct#)
          END IF
          UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + NonMAmt#)
          '11/19/96 Fixed Rev. amt. to add to current amt
          UBBillRec(1).TransAmt = Round(UBBillRec(1).TransAmt + NonMAmt#)
          GOTO GotAmt
        END IF
        
        MeterType$ = UBCustRec(1).Serv(RCnt).RMtrType
        
        FOR MCnt = 1 TO 7
          CubMtr = False
          IF MeterType$ = UBCustRec(1).LocMeters(MCnt).MTRType THEN
            IF UBCustRec(1).LocMeters(MCnt).MTRUnit = "C" THEN
              CubMtr = True
            END IF
            'Found correct meter
            MeterConsp& = UBCustRec(1).LocMeters(MCnt).CurRead - UBCustRec(1).LocMeters(MCnt).PrevRead
            IF MeterConsp& < 0 THEN
              MaxMeterAmt& = 10& ^ (LEN(STR$(UBCustRec(1).LocMeters(MCnt).PrevRead)) - 1)
              MeterConsp& = (MaxMeterAmt& - UBCustRec(1).LocMeters(MCnt).PrevRead) + UBCustRec(1).LocMeters(MCnt).CurRead
            END IF
            IF CubMtr THEN
              MeterConsp& = MeterConsp& * 7.481
            END IF
            IF UBCustRec(1).LocMeters(MCnt).MTRMulti > 0 THEN
              MeterConsp& = MeterConsp& * UBCustRec(1).LocMeters(MCnt).MTRMulti
            END IF
            UBBillRec(1).CurRead(MCnt) = UBCustRec(1).LocMeters(MCnt).CurRead
            UBBillRec(1).PrevRead(MCnt) = UBCustRec(1).LocMeters(MCnt).PrevRead
            UBBillRec(1).MtrTypes(MCnt) = GetCustMeterType(UBCustRec(), MCnt)
            'convert units here if necessary
            TMeterConsp& = TMeterConsp& + MeterConsp&
          END IF
        NEXT
        
        AddRevAmt# = 0
        IF UBCustRec(1).LocMeters(RCnt).NumUser > 1 THEN
          '100798 Corrected NumUser calc bug.  'Hillsville'
          'adjust min consumption for calc below
          NumUser& = UBCustRec(1).LocMeters(RCnt).NumUser - 1
          AddRevAmt# = NumUser& * UBRateTbls(WhatTbl).MINAMT
          MinimumConsp& = NumUser& * UBRateTbls(WhatTbl).MINUNITS
          TMeterConsp& = TMeterConsp& - MinimumConsp&
          IF (TMeterConsp& - UBRateTbls(WhatTbl).MINUNITS) <= 0 THEN
            UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + (AddRevAmt# + UBRateTbls(WhatTbl).MINAMT))
            UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + (AddRevAmt# + UBRateTbls(WhatTbl).MINAMT))
            GOTO GotAmt
          END IF
        ELSE
          NumUser& = 1
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
          RevAmt# = Round#(NumUser& * UBRateTbls(WhatTbl).MINAMT)
          'ADDED!!!!!!!!!!!!!
          IF ProrateFlag AND ProrateServ(RCnt) THEN
            RevAmt# = Round#(RevAmt# * ProPct#)
          END IF
          'ADDED!!!!!!!!!!!!!
          UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + RevAmt#)
          '11/19/96 Fixed Rev. amt. to add to current amt
          UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + RevAmt#)
          GOTO GotAmt
        END IF
        
        '04-22-97 Fixed to add to current rev amt
        '05-29-97 Refixed
        RevAmt# = GetRevCharge#(UBRateTbls(WhatTbl), TMeterConsp&, MeterMulti&)
        RevAmt# = RevAmt# + AddRevAmt#
        
        '111398 Prorate
        IF ProrateFlag AND ProrateServ(RCnt) THEN
          RevAmt# = Round#(RevAmt# * ProPct#)
        END IF
        UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + RevAmt#)
        
'NORWOOD Minimum Check START
        IF RCnt = 2 AND NorwoodFlag THEN
          IF LEFT$(UBCustRec(1).ZONE, 1) = "I" AND UBBillRec(1).RevAmt(RCnt) < 6.77 THEN
            UBBillRec(1).RevAmt(RCnt) = 6.77
            RevAmt# = 6.77
          END IF
        END IF
        IF RCnt = 2 AND NorwoodFlag THEN
          IF LEFT$(UBCustRec(1).ZONE, 1) = "O" AND UBBillRec(1).RevAmt(RCnt) < 13.54 THEN
            UBBillRec(1).RevAmt(RCnt) = 13.54
            RevAmt# = 13.54
          END IF
        END IF
'NORWOOD Minimum Check END
        
        UBBillRec(1).TransAmt = Round(UBBillRec(1).TransAmt + RevAmt#)
        IF RCnt = ElecRev THEN
          AdjRev# = Round(UBBillRec(1).RevAmt(RCnt) * FuelAdjAmt#)
          UBBillRec(1).RevAmt(RCnt) = Round(UBBillRec(1).RevAmt(RCnt) + AdjRev#)
          UBBillRec(1).TransAmt = Round(UBBillRec(1).TransAmt + AdjRev#)
        END IF
GotAmt:
        IF UBSetUp(1).Revenues(RCnt).TAXRATE > 0 AND UBCustRec(1).TAXEXPT <> "Y" THEN
          UBBillRec(1).TaxAmt(RCnt) = Round(UBBillRec(1).RevAmt(RCnt) * UBSetUp(1).Revenues(RCnt).TAXRATE)
          UBBillRec(1).TransAmt = Round(UBBillRec(1).TransAmt + UBBillRec(1).TaxAmt(RCnt))
        END IF
      END IF
    NEXT        'loop through all revenue sources
    
    BillCnt = BillCnt + 1
    UBBillRec(1).ActiveFlag = True
    UBBillRec(1).CustAcctNo = LCnt

'04-08-99  Elkton
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
              BlockClear
              DisplayUBScrn "ERRSCRN1"
              QPrintRC "Invalid Customer Type!", 10, 36, -1
              QPrintRC "ACCOUNT:" + STR$(LCnt), 10, 22, -1
              QPrintRC "Correct and Print Again.", 13, 28, -1
              WaitForAction
              AbortFlag = True
              GOTO AbortExit
            END SELECT
          END IF
        NEXT
      END IF
    END IF
'ELKton
NotAFinal:
    PUT BillFile, LCnt, UBBillRec(1)
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
    'NotAFinal:
    ShowPctComp LCnt, NumOfCustRecs&
  NEXT
AbortExit:
  CLOSE
  
END SUB

SUB PostFinalBillTrans
  
  UBLog "IN: POST FINAL"
  
  REDIM DepRev(1 TO 15) AS DOUBLE
  REDIM DepRevKept(1 TO 15) AS DOUBLE
  REDIM UBTempDepTran(1) AS UBTransRecType
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen

'STOP
  
  IF INSTR(UBSetUpRec(1).UTILNAME, "CLEVELAND") THEN
    CleveFlag = True
    UBLog "POST FINAL:  CLEVELAND Detected "
  END IF
  
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
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  
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

      IF UBBillRec(1).CustAcctNo = 13528 THEN STOP

      UBCustRec(1).Status = "B"
      UBCustRec(1).PrevBalance = Round(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
      UBCustRec(1).CurrBalance = UBBillRec(1).TransAmt
      UBBillRec(1).RunBalance = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
      
      IF UBBillRec(1).ApplyDepFlag = "Y" THEN
        IF CleveFlag THEN
          GOSUB ProcCleveDeposit
        ELSE
          GOSUB ProcCustDeposit
        END IF
      ELSE
        FOR RevCnt = 1 TO MaxRevsCnt
          UBCustRec(1).CurrRevAmts(RevCnt) = Round(UBCustRec(1).CurrRevAmts(RevCnt) + UBBillRec(1).RevAmt(RevCnt) + UBBillRec(1).TaxAmt(RevCnt))
        NEXT
      END IF

      UBBillRec(1).TransType = TranUtilityBill  'set transaction to Type 1
      FOR MtrCnt = 1 TO 7
        CubMtr = False
        IF UBCustRec(1).LocMeters(MtrCnt).CurRead > 0 THEN
          IF UBCustRec(1).LocMeters(MtrCnt).MTRUnit = "C" THEN
            CubMtr = True
          END IF
          ReadAmt& = UBBillRec(1).CurRead(MtrCnt) - UBBillRec(1).PrevRead(MtrCnt)
          IF ReadAmt& < 0 THEN  'Meter rolled over or, misread
            MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(MtrCnt))) - 1)
            ReadAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(MtrCnt)) + UBBillRec(1).CurRead(MtrCnt)
          END IF
          IF CubMtr THEN
            ReadAmt& = ReadAmt& * 7.481
          END IF
          UBCustRec(1).LocMeters(MtrCnt).AvgUse = Round(UBCustRec(1).LocMeters(MtrCnt).AvgUse + ReadAmt&)
          UBCustRec(1).LocMeters(MtrCnt).UseCnt = UBCustRec(1).LocMeters(MtrCnt).UseCnt + 1
          UBCustRec(1).LocMeters(MtrCnt).ReadFlag = ""
          '050697 Fixed current reading not being move to previous
          UBCustRec(1).LocMeters(MtrCnt).PrevRead = UBCustRec(1).LocMeters(MtrCnt).CurRead
        END IF
      NEXT
      
      PrevLastTrans& = UBCustRec(1).LastTrans
      UBBillRec(1).PrevTrans = PrevLastTrans&
      NextTransRec& = (LOF(UBTran) \ UBBillRecLen) + 1          'point at next trans

      PUT UBTran, NextTransRec&, UBBillRec(1)

      UBCustRec(1).LastTrans = NextTransRec&
      
      'detach the new vacant rec from this customer
      IF UBCustRec(1).OldRec > 0 THEN
        UBLog "POST FINAL: DETACHED OLD ACCT:" + STR$(UBCustRec(1).OldRec)
        UBCustRec(1).OldRec = 0
      END IF

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
  UBLog "OUT: POST FINAL"
  EXIT SUB
  
MakeAppDepTrans:
  REDIM UBTransRec(1) AS UBTransRecType
  UBTransRecLen = LEN(UBTransRec(1))
  UBTransRec(1).TransDate = UBBillRec(1).TransDate
  'UBTransRec(1)CustLocation = UBBillRec(1).CustAcctNo
  UBTransRec(1).CustStatus = UBCustRec(1).Status
  UBTransRec(1).CustAcctNo = UBBillRec(1).CustAcctNo
  UBTransRec(1).TransAmt = DepTranAmt#
  '091198 Changed to put original deposit amounts in revenue source
  FOR Cnt = 1 TO 15
    UBTransRec(1).RevAmt(Cnt) = DepRevKept(Cnt)
  NEXT
  
  UBTransRec(1).TransDesc = "Applied Deposit"
  UBTransRec(1).TransType = TranAppliedDeposit
  UBTransRec(1).RunBalance = Round#((UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance) - ABS(DepTranAmt#))
  UBCustRec(1).DepositAmt = 0
  UBCustRec(1).CurrBalance = Round#(UBCustRec(1).CurrBalance - ABS(DepTranAmt#))
  
  PrevLastTrans& = UBCustRec(1).LastTrans
  UBTransRec(1).PrevTrans = PrevLastTrans&
  
  IF Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance) = 0 THEN
    IF UBCustRec(1).Status = "B" THEN
      CustChCnt = CustChCnt + 1
      UBLog "POST FINAL: SET CUST STATUS=I ACCT:" + STR$(UBTransRec(1).CustAcctNo)
      UBCustRec(1).Status = "I"
    END IF
  END IF
  NextTransRec& = (LOF(UBTran) \ UBTransRecLen) + 1             'point at next trans

  PUT UBTran, NextTransRec&, UBTransRec(1)

  UBCustRec(1).LastTrans = NextTransRec&

  PUT UBCust, UBTransRec(1).CustAcctNo, UBCustRec(1)

  UBLog "POST FINAL: DEP APPLIED TRANS:" + STR$(NextTransRec&)
  
RETURN
  
ProcCleveDeposit:
  FOR LLCnt = 1 TO 15
    DepRev(LLCnt) = 0
  NEXT LLCnt

    DepAppliedFlag = False
    DepTranAmt# = -UBCustRec(1).DepositAmt
    DepositAmt# = UBCustRec(1).DepositAmt
    IF DepositAmt# = 0 THEN
      FOR RevCnt = 1 TO MaxRevsCnt
        UBCustRec(1).CurrRevAmts(RevCnt) = Round(UBCustRec(1).CurrRevAmts(RevCnt) + UBBillRec(1).RevAmt(RevCnt) + UBBillRec(1).TaxAmt(RevCnt))
      NEXT
      GOTO NoDepReturn
    END IF

    ThisTran& = UBCustRec(1).LastTrans
    DO WHILE ThisTran& > 0
      GET UBTran, ThisTran&, UBTempDepTran(1)
      IF UBTempDepTran(1).TransType = TranDepositPayment THEN
        FOR DZCnt = 1 TO 15
          DepRev(DZCnt) = Round#(DepRev(DZCnt) + UBTempDepTran(1).RevAmt(DZCnt))
        NEXT
      END IF
      ThisTran& = UBTempDepTran(1).PrevTrans
    LOOP
    
    FOR RevCnt = 1 TO MaxRevsCnt - 1
      UBCustRec(1).CurrRevAmts(RevCnt) = Round(UBCustRec(1).CurrRevAmts(RevCnt) + UBBillRec(1).RevAmt(RevCnt) + UBBillRec(1).TaxAmt(RevCnt))
      IF DepRev(RevCnt) > 0 THEN
        DepAppliedFlag = True
        IF UBCustRec(1).CurrRevAmts(RevCnt) < DepRev(RevCnt) THEN
          DepRev(RevCnt) = Round#(DepRev(RevCnt) - UBCustRec(1).CurrRevAmts(RevCnt))
          UBCustRec(1).CurrRevAmts(RevCnt) = 0
        ELSEIF UBCustRec(1).CurrRevAmts(RevCnt) > DepRev(RevCnt) THEN
          UBCustRec(1).CurrRevAmts(RevCnt) = Round(UBCustRec(1).CurrRevAmts(RevCnt) - DepRev(RevCnt))
          DepRev(RevCnt) = 0
        ELSE    'the deposit and the revenue are equal
          UBCustRec(1).CurrRevAmts(RevCnt) = 0
          DepRev(RevCnt) = 0
        END IF
      END IF
    NEXT
    
    'If there was any deposit left after applying to the cust rev totals
    FOR RevCnt = 1 TO MaxRevsCnt - 1
      IF DepRev(RevCnt) > 0 THEN
        UBCustRec(1).CurrRevAmts(RevCnt) = -DepRev(RevCnt)
      END IF
    NEXT
    UBCustRec(1).DepositAmt = 0
    
NoDepReturn:
    
    RETURN

    
ProcCustDeposit:
    
    FOR LLCnt = 1 TO 15
      DepRev(LLCnt) = 0
    NEXT
    
    DepAppliedFlag = False
    DepTranAmt# = -UBCustRec(1).DepositAmt
    DepositAmt# = UBCustRec(1).DepositAmt
    IF DepositAmt# = 0 THEN
'051799 added to correct rev problem with accounts that had no deposit but
'       apply deposit to final bill was selected
      FOR RevCnt = 1 TO MaxRevsCnt
        UBCustRec(1).CurrRevAmts(RevCnt) = Round(UBCustRec(1).CurrRevAmts(RevCnt) + UBBillRec(1).RevAmt(RevCnt) + UBBillRec(1).TaxAmt(RevCnt))
      NEXT
      GOTO NoDepReturn
    END IF
    ThisTran& = UBCustRec(1).LastTrans
    DO WHILE ThisTran& > 0
      GET UBTran, ThisTran&, UBTempDepTran(1)
      IF UBTempDepTran(1).TransType = TranDepositPayment THEN
        FOR DZCnt = 1 TO 15
          DepRev(DZCnt) = Round#(DepRev(DZCnt) + UBTempDepTran(1).RevAmt(DZCnt))
          DepRevKept(DZCnt) = DepRev(DZCnt)
        NEXT
      ELSEIF (UBTempDepTran(1).TransType = TranAppliedDeposit) OR (UBTempDepTran(1).TransType = TranRefundDeposit) THEN
        FOR DZCnt = 1 TO 15
          DepRev(DZCnt) = Round#(DepRev(DZCnt) - UBTempDepTran(1).RevAmt(DZCnt))
          DepRevKept(DZCnt) = DepRev(DZCnt)
        NEXT
      END IF
      ThisTran& = UBTempDepTran(1).PrevTrans
    LOOP
    
    FOR RevCnt = 1 TO MaxRevsCnt - 1
      UBCustRec(1).CurrRevAmts(RevCnt) = Round(UBCustRec(1).CurrRevAmts(RevCnt) + UBBillRec(1).RevAmt(RevCnt) + UBBillRec(1).TaxAmt(RevCnt))
      IF DepRev(RevCnt) > 0 THEN
        DepAppliedFlag = True
        UBCustRec(1).CurrRevAmts(RevCnt) = Round(UBCustRec(1).CurrRevAmts(RevCnt) - DepRev(RevCnt))
        DepRev(RevCnt) = 0
      END IF
    NEXT
    UBCustRec(1).DepositAmt = 0

RETURN

END SUB

SUB PreBillReport
    
    UBLog "IN: Prebilling Report (Final)"
    
    Dash80$ = STRING$(80, "-")
    Temp1$ = SPACE$(10)
    Temp2$ = SPACE$(12)
    
    NumOfRevs = MaxRevsCnt      'assume max munber of revenue sources
    '111698 Prorate
    REDIM ProrateServ(1 TO 15) AS INTEGER
    
    REDIM UBSetUpRec(1) AS UBSetupRecType
    LoadUBSetUpFile UBSetUpRec(), UBSetUpLen
    TownName$ = UBSetUpRec(1).UTILNAME
    
    IF INSTR(TownName$, "MOWAS") > 0 THEN
      MowFlag = True
    END IF
    
    REDIM RevDesc(1 TO MaxRevsCnt) AS STRING * 12
    FOR Cnt = 1 TO MaxRevsCnt   'find last active revenue
      TempRev$ = QPTrim$(UBSetUpRec(1).Revenues(Cnt).REVNAME)
      IF LEN(TempRev$) = 0 THEN
        NumOfRevs = Cnt - 1     'set actual number of revenues
        EXIT FOR
      ELSE      'build revenue description lines
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
    
    MaxLines = 45
    
    REDIM MChoice$(1 TO 5)
    
    MChoice$(1) = "Customer Name Order"
    MChoice$(2) = "Account Number Order"
    MChoice$(3) = "Postal Carrier Route Order"
    MChoice$(4) = "ZipCode Order"
    
    IF UBSetUpRec(1).UseSeq = "Y" THEN
      MChoice$(5) = "Sequence Number Order"
    END IF
    
    MaxLen = 0  'Set menu width to zero
    BoxBot = 17 'limit the box length to go no lower than line 20
    Action = 0  '0 means stay in the menu until they select something
    Choice = 1  'Pre-load choice to highlight
    
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
    
    DO          '--Set upper left corner of menu, turn off the cursor
      LOCATE Row, Col, 0
      BlockClear
      TitleBox 2, Col, MaxLen + 3, "Pre-Billing Report ", Cnf
      TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
      
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
      CASE 3, 4
        IF Choice = 4 THEN
          IdxTypeText$ = "Postal Route"
        ELSE
          IdxTypeText$ = "Zip-Code"
        END IF
        IF MowFlag THEN
          MakeMowZipCodeIndex IdxTypeText$
        ELSE
          MakePostalIndex IdxTypeText$
        END IF
        IndexName$ = TempIndexName
        OKFlag = True
      CASE 5    'Sequence number
        IdxTypeText$ = "Sequence Number"
        MakeSequenceIndex IdxTypeText$
        IndexName$ = TempIndexName
        OKFlag = True
      END SELECT
      
    LOOP UNTIL OKFlag OR ExitFlag
    IF ExitFlag THEN GOTO ExitPreReport
    
    DepositFlag = GetUseDeposit(MChoice$(Choice))
    
    IF DepositFlag = -2 THEN GOTO ReStart
    
    IF DoFuelAdjFlag THEN
      FuelAdjAmt# = GetAdjFactor#
      UBLog "Fuel adjustment factor:" + STR$(FuelAdjAmt#)
    ELSE
      FuelAdjAmt# = 0
    END IF
    
    IF FuelAdjAmt# = -10000 THEN GOTO ReStart
    
    UBLog "Calculating utility charges."
    
    MakeFinalBillFile AbortFlag%, FuelAdjAmt#
    
    IF AbortFlag THEN
      UBLog "ABORTED: CALCULATIONS"
    ELSE
      UBLog "Finished utility calculations."
    END IF
    
    IF AbortFlag THEN GOTO ExitPreReport
    
    REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
    UBCustRecLen = LEN(UBCustRec(1))
    
    REDIM UBBillRec(1) AS UBTransRecType
    UBBillRecLen = LEN(UBBillRec(1))
    
    REDIM RevTotals(1 TO NumOfRevs) AS DOUBLE   'holds revenues total amt
    '052097 added tax by revenue totals
    REDIM TaxTotals(1 TO NumOfRevs) AS DOUBLE   'holds revenues tax total amt
    REDIM ConsumpTot(1 TO NumOfRevs, 1 TO 2) AS DOUBLE          'holds each revenues total amt
    REDIM RateConsump(1 TO NumOfRates) AS LONG
    REDIM RateTotals(1 TO NumOfRates) AS DOUBLE 'holds each Rates $totals
    '052097 added tax by rate code totals
    REDIM RTaxTot(1 TO NumOfRates) AS DOUBLE    'holds each Rates Tax totals
    '052097 added tax by book totals to type def
    
    '012698 Added bill count by rate code
    REDIM RateCount(1 TO NumOfRates) AS LONG
    REDIM PumpConsump(0 TO 1) AS PumpConsumpType                'Consumption by pump code
    
    TBooks = 0
    
    IF UsingAcct THEN
      NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
    ELSE        'load the index
      UBLog "Loading index file: " + IndexName$
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
    
    UBLog "Writing prebilling report to disk."
    
    GOSUB PrintPreHeader
    
    FOR Cnt = 1 TO NumOfRecs
      IF UsingAcct THEN
        ThisCustRec& = Cnt
      ELSE
        ThisCustRec& = IndexArray(Cnt).RecNum
      END IF
      
      GET UBCust, ThisCustRec&, UBCustRec(1)
      IF UBCustRec(1).BillTo = "O" THEN
        BillTo$ = " O"
      ELSE
        BillTo$ = " C"
      END IF
      
      IF UBCustRec(1).DelFlag THEN
        GOTO skipem
      END IF
      
      IF UBCustRec(1).Status <> "F" THEN
        GOTO skipem
      END IF
      
      GET UBBill, ThisCustRec&, UBBillRec(1)
      
      'IF ThisCustRec& = 190 THEN
      '  STOP
      'END IF
      
      IF DepositFlag THEN
        IF UBCustRec(1).DepositAmt > 0 THEN
          CustDepAmt# = UBCustRec(1).DepositAmt
          UBBillRec(1).TaxAmt(15) = CustDepAmt#
          PUT UBBill, ThisCustRec&, UBBillRec(1)
        END IF
      END IF
      
      IF LineCnt > MaxLines THEN
        PRINT #UBRpt, CHR$(12)
        GOSUB PrintPreHeader
      END IF
      
      IF UBBillRec(1).ActiveFlag <> 0 THEN
        PRINT #UBRpt, UBCustRec(1).Status; TAB(4); USING "######"; ThisCustRec&;
        PRINT #UBRpt, "      "; LEFT$(QPTrim$(UBCustRec(1).CUSTNAME), 30); TAB(48); LEFT$(UBCustRec(1).SERVADDR, 23);
        PRINT #UBRpt, USING "   ###%"; UBBillRec(1).ProRatePCT;
        PRINT #UBRpt, BillTo$
      END IF
      
      WhatRate = 0
      
      DoneOne = False
      FOR TRevCnt = 1 TO NumOfRevs
        IF UBBillRec(1).RevAmt(TRevCnt) <> 0 THEN
          DoneOne = False
          PRINT #UBRpt, RevDesc(TRevCnt);
          FOR TRateCnt = 1 TO NumOfRates
            IF UBRateTbls(TRateCnt).RateCode = UBCustRec(1).Serv(TRevCnt).RateCode THEN
              MINAMT& = UBRateTbls(TRateCnt).MINUNITS
              WhatRate = TRateCnt
              EXIT FOR
            END IF
          NEXT
          
          IF UBSetUpRec(1).Revenues(TRevCnt).UseMtr = "Y" THEN
            RevTotals(TRevCnt) = Round#(RevTotals(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
            '02-20-97 Add new cash totals by rate
            IF WhatRate > 0 THEN
              RateCount(WhatRate) = RateCount(WhatRate) + 1
              RateTotals(WhatRate) = Round#(RateTotals(WhatRate) + UBBillRec(1).RevAmt(TRevCnt))
              RTaxTot(WhatRate) = Round#(RTaxTot(WhatRate) + UBBillRec(1).TaxAmt(TRevCnt))
            END IF
            
            FOR MCCnt = 1 TO 7
              CubMtr = False
              LocMeterType$ = QPTrim$(UBCustRec(1).LocMeters(MCCnt).MTRType)
              MeterMulti& = UBCustRec(1).LocMeters(MCCnt).MTRMulti              'Meter Multi??
              
              IF UBCustRec(1).LocMeters(MCCnt).MTRUnit = "C" THEN
                CubMtr = True
              END IF
              
              IF MeterMulti& <= 0 THEN MeterMulti& = 1
              IF (LEN(LocMeterType$) > 0) AND (LocMeterType$ = UBCustRec(1).Serv(TRevCnt).RMtrType) THEN
                DoneOne = True
                MeterNum$ = QPTrim$(UBCustRec(1).Serv(TRevCnt).RateCode)
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
                '052797  Read Error
                ReadErr = False
                IF UBBillRec(1).CurRead(MCCnt) < 0 THEN
                  ReadErr = True
                  PRINT #UBRpt, TAB(30); USING "**#######"; 0;
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
                IF ReadAmt& < 0 THEN
                  'Meter has rolled over or, been misread
                  MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(MCCnt))) - 1)
                  ReadAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(MCCnt)) + UBBillRec(1).CurRead(MCCnt)
                END IF
                IF CubMtr THEN
                  ReadAmt& = ReadAmt& * 7.481
                END IF
                
                RateConsump(WhatRate) = RateConsump(WhatRate) + (ReadAmt& * MeterMulti&)
                ConsumpTot(TRevCnt, 1) = ConsumpTot(TRevCnt, 1) + (ReadAmt& * MeterMulti&)
                
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
                  LowConsump& = Round#(AvgUse& * (UBSetUpRec(1).LowRead * .01))
                  HiConsump& = Round#(AvgUse& * (UBSetUpRec(1).HighRead * .01))
                END IF
                
                IF UBCustRec(1).EstFlag = "E" THEN
                  PRINT #UBRpt, " E";           'Est. Reading
                ELSEIF Consump& < MINAMT& THEN
                  PRINT #UBRpt, " M";           'Minium Usage
                ELSEIF Consump& < LowConsump& THEN
                  PRINT #UBRpt, " L";           'Low reading
                ELSEIF Consump& > HiConsump& THEN
                  PRINT #UBRpt, " H";           'High Reading
                END IF
                IF UBBillRec(1).RevAmt(TRevCnt) > 0 THEN
                  PRINT #UBRpt, TAB(69); USING "######.##"; UBBillRec(1).RevAmt(TRevCnt);
                  IF UBBillRec(1).ProRatePCT < 100 AND ProrateServ(TRevCnt) THEN
                    PRINT #UBRpt, "*";
                  END IF
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
              PRINT #UBRpt, TAB(69); USING "######.##"; UBBillRec(1).RevAmt(TRevCnt);
              IF UBBillRec(1).ProRatePCT < 100 AND ProrateServ(TRevCnt) THEN
                PRINT #UBRpt, "*";
              END IF
              PRINT #UBRpt,
              LineCnt = LineCnt + 1
            END IF
            '***********************************************************************
            
          ELSE  'it's a nonmetered service
            ConsumpTot(TRevCnt, 1) = ConsumpTot(TRevCnt, 1) + 1
            IF WhatRate > 0 THEN
              RateCount(WhatRate) = RateCount(WhatRate) + 1
              RateConsump(WhatRate) = RateConsump(WhatRate) + 1
              RateTotals(WhatRate) = Round#(RateTotals(WhatRate) + UBBillRec(1).RevAmt(TRevCnt))
              RTaxTot(WhatRate) = Round#(RTaxTot(WhatRate) + UBBillRec(1).TaxAmt(TRevCnt))
            ELSE
              'STOP
            END IF
            RevTotals(TRevCnt) = Round#(RevTotals(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
            PRINT #UBRpt, TAB(69); USING "######.##"; UBBillRec(1).RevAmt(TRevCnt);
            IF UBBillRec(1).ProRatePCT < 100 AND ProrateServ(TRevCnt) THEN
              PRINT #UBRpt, "*";
            END IF
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
                    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
                    IF NOT ConsumpFlag THEN
                      PRINT #UBRpt, UBCustRec(1).Status; "  "; UBCustRec(1).BOOK; "-"; UBCustRec(1).SEQNUMB; "  "; UBCustRec(1).CUSTNAME; "  "; QPTrim$(UBCustRec(1).SERVADDR)
                    END IF
                    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
                    ConsumpFlag = True
                    PRINT #UBRpt, RevDesc(TTRevCnt);
                    'MeterNum$ = QPTrim$(UBCustRec(1).LocMeters(MCCnt).MTRNum)
                    MeterNum$ = QPTrim$(UBCustRec(1).Serv(MCCnt).RateCode)
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
          IF ConsumpFlag AND UBCustRec(1).Status <> "A" THEN
            ConsumpFlag = False
            PRINT #UBRpt, "**** Consumption Noted on an Inactive Account. ****"
            PRINT #UBRpt, Dash80$
            LineCnt = LineCnt + 2
            NCCnt = NCCnt + 1
          ELSEIF ConsumpFlag THEN               'Customer Status is "A"
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
          
          IF DepositFlag THEN
            PRINT #UBRpt, TAB(48); "Less Deposit:";
            PRINT #UBRpt, TAB(65); USING "$$#######,.##"; -UBCustRec(1).DepositAmt
            BalanceAmt# = Round#(AcctBalance# + UBBillRec(1).TransAmt - UBCustRec(1).DepositAmt)
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
          END IF
          
          '-=-=-=-=-=-=-=-=-=-=-=-=-=
          PRINT #UBRpt, Dash80$
          LineCnt = LineCnt + 1
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
            IF ReadAmt& < 0 THEN                'Meter rolled over or, been misread
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
    
    IF AbortFlag THEN GOTO ExitPreReport
    
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
    PRINT #UBRpt, "Code      Rate Description            Consumption        Amount       Bills"
    PRINT #UBRpt, Dash80$
    
    TotalRevAmt# = 0
    
    FOR RaCnt = 1 TO NumOfRates
      IF RateConsump(RaCnt) <> 0 OR RateTotals(RaCnt) <> 0 THEN
        PRINT #UBRpt, UBRateTbls(RaCnt).RateCode; "   "; UBRateTbls(RaCnt).RATEDESC;
        PRINT #UBRpt, TAB(38); USING "###########"; RateConsump(RaCnt);
        PRINT #UBRpt, TAB(54); USING "#######.##"; RateTotals(RaCnt);
        PRINT #UBRpt, TAB(71); USING "#####"; RateCount(RaCnt)
        TotalRevAmt# = Round#(TotalRevAmt# + RateTotals(RaCnt))
        IF RTaxTot(RaCnt) > 0 THEN
          PRINT #UBRpt, " Tax"; TAB(60); USING "#######.##"; RTaxTot(RaCnt)
        END IF
      END IF
    NEXT
    
    PRINT #UBRpt,
    PRINT #UBRpt, "TAX TOTAL"; TAB(60); USING "#######.##"; TotalTaxAmt#
    PRINT #UBRpt, TAB(45); "TOTAL:"; TAB(57); USING "$$,#######.##"; Round#(TotalRevAmt# + TotalTaxAmt#)
    PRINT #UBRpt, CHR$(12)
    
    IF TPumps > 0 THEN
      GOSUB PumpHeader
      TMMConsump# = 0
      FOR Cnt = 1 TO TPumps
        PRINT #UBRpt, PumpConsump(Cnt).PumpCode; TAB(30); USING "###########"; PumpConsump(Cnt).CustCnt; TAB(60); PumpConsump(Cnt).Consump
        TMMConsump# = TMMConsump# + PumpConsump(Cnt).Consump
      NEXT
      PRINT #UBRpt, Dash80$
      PRINT #UBRpt, TAB(35); "Pump Code Total:"; TAB(60); USING "###########"; TMMConsump#
    END IF
    
    
    CLOSE
    
    UBLog "Finished writing prebilling report."
    
    SELECT CASE Choice
    CASE 1
      RptText$ = " (Customer Order)"
    CASE 2
      RptText$ = "(Account Order)"
    CASE 3
      RptText$ = "(Postal RT. Order)"
    CASE 4
      RptText$ = "(ZipCode Order)"
    CASE 5
      RptText$ = "(Sequence Order)"
    END SELECT
    
    ERASE UBCustRec, UBBillRec, RevTotals, TaxTotals
    ERASE ConsumpTot, RateConsump, UBRateTbls
    ERASE RateConsump, UBSetUpRec, RevDesc
    
    IF NOT AbortFlag THEN
      PrintRptFile "Pre-Billing Report " + RptText$, UBFinPreRptFile, LPTPort, RetCode, EntryPoint
    END IF
    
    GOTO ExitPreReport
    
PrintPreHeader:
    PageNo = PageNo + 1
    PRINT #UBRpt, "Utility Pre-Billing Report.  "; TownName$; TAB(70); "Page: "; PageNo
    PRINT #UBRpt, "Date: "; DATE$
    PRINT #UBRpt, "                Customer Name                Service Address"
    PRINT #UBRpt, "Status  AcctNo.      Rate      Current     Past                       Prorate%"
    PRINT #UBRpt, "Revenue Source       Code      Reading     Reading    Consumption      Charges"
    PRINT #UBRpt, Dash80$
    LineCnt = 6
    RETURN
    
GetWhatPump:
    HasAPumpCode = True         'assume they have a pump code
    WhatPump = 0
    IF LEN(QPTrim$(UBCustRec(1).PumpCode)) = 0 THEN
      IF UBCustRec(1).Status = "A" THEN
        HasAPumpCode = False    'no pump code
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
    
PumpHeader:
    PageNo = PageNo + 1
    PRINT #UBRpt, "Utility Pre-Billing Report.  "; TownName$; TAB(70); "Page: "; PageNo
    'PRINT #UBRpt, "Utility Pre-Billing Report"; TAB(70); "Page: "; PageNo
    PRINT #UBRpt, TheDate$
    PRINT #UBRpt, "Report Totals by Pump Code"
    PRINT #UBRpt,
    PRINT #UBRpt, "PumpCode                  Customer Count                    Consumption"
    PRINT #UBRpt, Dash80$
    LineCnt = 6
    RETURN
    
    
ExitPreReport:
    UBLog "OUT: Prebilling Report" + CrLf$
    
END SUB

