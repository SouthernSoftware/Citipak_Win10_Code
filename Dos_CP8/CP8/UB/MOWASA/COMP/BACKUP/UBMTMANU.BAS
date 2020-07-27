DEFINT A-Z
DECLARE SUB CursorOff ()
DECLARE SUB PrintHighLowReport ()
DECLARE SUB DisplayHHScrn (ScrnName$)
DECLARE FUNCTION Chk4BookSeqNum (BOOK$, SeqNum$)
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetUpLen%)
DECLARE SUB MakeSequenceIndex (IndexText$)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB ManualMeterReading ()
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB PrintHelp (H$)
DECLARE SUB ShowCursor ()
DECLARE SUB HideCursor ()
DECLARE SUB ClearScrn ()

DECLARE FUNCTION AskAbandonPrint% ()
DECLARE FUNCTION Round# (N#)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION FmtBook$ (BOOK$)
DECLARE FUNCTION FmtSeqN$ (SeqN$)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB KillFile (FileName$)
DECLARE SUB FClose (Handle%)
DECLARE FUNCTION FLof& (FileHand%)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB WaitForAction ()
DECLARE SUB BlockClear ()
DECLARE SUB SaveScrn (TempScrn())
DECLARE SUB RestScrn (TempScrn())
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION QPValI% (Numb$)
DECLARE FUNCTION QPValL& (Numb$)

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubdraft.BI'
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'Newcust.bi'
  ''$INCLUDE: 'ubgilsen.bi'

  CONST False = 0, True = NOT False

SUB ManualMeterReading
  
  REDIM TempScrn(0)
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  Help$ = "Meter Reading Input"
  LibName$ = "UB"
  ScrnName$ = "UBMTRIN2"
  DisplayUBScrn ScrnName$
  
  GOSUB GetRoute
  
  IF BookNumber <= 0 THEN EXIT SUB
  
  'BeginReading = True
  ' We Have the Book in variable=BookNumber
  ' We Have the Def Meter Read Date in variable=MtrReadDate$
  ' Next Get Utility Setup and Check if the order is by seq or by location
  ' Sequence Order ???? ******************************************************
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen
  
  IF INSTR(UBSetUpRec(1).UTILNAME, "BEECH") THEN
    BeechMtn = True
  END IF
  
  IF UBSetUpRec(1).SkipInactive = "Y" THEN
    NoInactive = True
  ELSE
    NoInactive = False
  END IF
  
  HighVar = UBSetUpRec(1).HighRead
  LowVar = UBSetUpRec(1).LowRead
  
  IF UBSetUpRec(1).UseSeq = "Y" THEN
    MakeSequenceIndex "Sequence Number"
    NumOfRecs = FileSize&("UBTEMP.IDX") \ 4
    REDIM IndexArray(1 TO NumOfRecs) AS UBCustIndexRecType
    FGetAH "UBTEMP.IDX", IndexArray(1), 4, NumOfRecs
    LibName$ = "UB"
    ScrnName$ = "UBMTRIN2"
    
    NumScrns = 1
    
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
    
    '--Clear all fields
    FOR F = 1 TO NumFlds
      LSET Form$(F, 0) = ""
    NEXT
    
    '--Set screen number to one and display screen
    Scr = 1
    
    IdxRecLen = 4               'we are using a integer
    IdxFileSize& = FileSize&("UBCUSTBK.IDX")
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    
    REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
    FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs  'load it
    
    Cnt = 1
    
    DO
      IF Cnt < 1 THEN Cnt = 1   ' Do Not Allow to Fall Below 1
      inputting = False         ' Set Edit Finish to No
      PRec& = IndexArray(Cnt).RecNum
      IF NOT (PRec&) = 0 THEN
        GOSUB GetCustRecNo
        IF process THEN
          GOSUB UpdateMeter
        END IF
      END IF
    LOOP UNTIL Cnt > NumOfRecs  'OR Cnt = 0
    
    EXIT SUB
  END IF
  
  LibName$ = "UB"
  ScrnName$ = "UBMTRIN2"
  
  NumScrns = 1
  
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
  
  '--Clear all fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT
  
  '--Set screen number to one and display screen
  Scr = 1
  
  IdxRecLen = 4 'we are using a integer
  IdxFileSize& = FileSize&("UBCUSTBK.IDX")
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  
  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it
  
  Cnt = 1
  
  DO
    IF Cnt < 1 THEN
      Cnt = 1
      PageUp = False            ' Do Not Allow to Fall Below 1
    END IF
    inputting = False
    PRec& = IdxBuff(Cnt).RecNum
    QPrintRC STR$(PRec&), 25, 40, 112
    IF NOT (PRec&) = 0 THEN
      GOSUB GetCustRecNo
      IF process THEN
        GOSUB UpdateMeter
      END IF
    END IF
  LOOP UNTIL Cnt > IdxNumOfRecs
  
  EXIT SUB
  'End of Location Order Processing *******************************************
  'Update the Meter Reading Here *********************************************
  
UpdateMeter:
  
  IF BeechMtn THEN
    Fld(11).Protected = False
  END IF
  
  Help$ = "Record #" + STR$(Cnt) + " of " + STR$(IdxNumOfRecs)
  
  FOR MtrCnt = 1 TO 7           'find last active meter
    
    inputting = False
    TempRev$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MTRType)
    
    IF LEN(TempRev$) <> 0 THEN
      
      DisplayUBScrn ScrnName$
      PrintHelp Help$
      
      REM set information as known
      Form$(1, 0) = RIGHT$(STR$(BookNumber), 2)
      Form$(2, 0) = UBCustRec(1).CUSTNAME
      Form$(3, 0) = UBCustRec(1).BOOK
      Form$(4, 0) = UBCustRec(1).SEQNUMB
      Form$(5, 0) = STR$(PRec&)
      Form$(6, 0) = UBCustRec(1).SERVADDR
      
      GOSUB GetMeterType
      
      Form$(7, 0) = MeterType$
      Form$(8, 0) = UBCustRec(1).LocMeters(MtrCnt).MTRNUM
      Form$(9, 0) = MtrReadDate$
      
      Form$(13, 0) = "N"
      
      IF UBCustRec(1).LocMeters(MtrCnt).ReadFlag = "Y" THEN
        Form$(10, 0) = STR$(UBCustRec(1).LocMeters(MtrCnt).CurRead)
        Form$(11, 0) = STR$(UBCustRec(1).LocMeters(MtrCnt).PrevRead)
        IF UBCustRec(1).EstFlag = "E" THEN
          Form$(13, 0) = "Y"
        END IF
      ELSE
        Form$(11, 0) = STR$(UBCustRec(1).LocMeters(MtrCnt).CurRead)
      END IF
      
      IF MeterType$ = "Demand Electric" THEN
        Form$(11, 0) = "0"      ' Demand Always = 0 for Previous Read
      END IF
      
      GOSUB GetMeterStatus
      Form$(19, 0) = MeterStatus$
      
      Average# = UBCustRec(1).LocMeters(MtrCnt).AvgUse
      AvgCnt! = UBCustRec(1).LocMeters(MtrCnt).UseCnt
      
      
      ' RESET AVERAGE IS OUT OF WHACK
      IF Average# > 99999999 OR Average# <= 0 THEN
        UBCustRec(1).LocMeters(MtrCnt).AvgUse = 0
        UBCustRec(1).LocMeters(MtrCnt).UseCnt = 0
        Average# = 0
        AvgCnt! = 0
      END IF
      
      LOCATE 16, 53
      COLOR 3
      PRINT "Avg Use: "; USING "#######,#"; Average#
      
      Frm(1).FldNo = 10
      Action = 1
      
      DO
        EditForm Form$(), Fld(), Frm(1), Cnf, Action
        
        IF NOT BeechMtn THEN
          IF Frm(1).PrevFld = 10 AND Frm(1).FldNo > 10 THEN
            GOSUB DoCalcUsage
          END IF
        END IF
        
        IF Frm(1).PrevFld = 11 AND Frm(1).FldNo > 11 THEN
          GOSUB DoCalcUsage
        END IF
        
        '--Check for Key presses
        SELECT CASE Frm(1).KeyCode
          
        CASE PgUpKey
          PageUp = True
          Cnt = Cnt - 1
          LSET Form$(10, 0) = ""
          RETURN
          
        CASE PgDnKey
          PageUp = False
          inputting = True
          LSET Form$(10, 0) = ""
          
        CASE F5KEY              'Close Down and Exit
          EXIT SUB
          
        CASE F7KEY
          OPrev# = VAL(Form$(11, 0))
          LSET Form$(10, 0) = QPTrim$(STR$(OPrev# + Average#))
          Frm(1).FldNo = Frm(1).FldNo + 1
          Action = 1
          LSET Form$(13, 0) = "Y"
          
        CASE F10Key
          'update current reading and flag it
          PageUp = False
          IF UBCustRec(1).LocMeters(MtrCnt).ReadFlag <> "Y" THEN
            UBCustRec(1).LocMeters(MtrCnt).CurRead = VAL(Form$(10, 0))
            UBCustRec(1).LocMeters(MtrCnt).PrevRead = VAL(Form$(11, 0))
            UBCustRec(1).LocMeters(MtrCnt).ReadFlag = "Y"
            UBCustRec(1).LocMeters(MtrCnt).PastDate = UBCustRec(1).LocMeters(MtrCnt).CurDate
            UBCustRec(1).LocMeters(MtrCnt).CurDate = Date2Num(Form$(9, 0))
          ELSE
            UBCustRec(1).LocMeters(MtrCnt).CurRead = VAL(Form$(10, 0))
            IF BeechMtn THEN
              UBCustRec(1).LocMeters(MtrCnt).PrevRead = VAL(Form$(11, 0))
            END IF
          END IF
          
          IF Form$(13, 0) = "Y" THEN
            UBCustRec(1).EstFlag = "E"
          ELSE
            UBCustRec(1).EstFlag = ""
          END IF
          
          IF MeterType$ = "Demand Electric" THEN
            UBCustRec(1).LocMeters(MtrCnt).PrevRead = 0         ' Demand Always = 0 for Previous Read
          END IF
          
          FOpenS "UBCUST.DAT", C2Handle         'open data file
          FPutRTA C2Handle, UBCustRec(1), PRec&, UBCustRecLen
          FClose C2Handle
          
          inputting = True
          
        CASE EscKey
          inputting = True
        END SELECT
        
      LOOP WHILE inputting = False
      
    END IF
    LSET Form$(10, 0) = ""
    LSET Form$(12, 0) = ""
  NEXT MtrCnt
  
  IF PageUp THEN
    Cnt = Cnt - 1
  ELSE
    Cnt = Cnt + 1
  END IF
  RETURN
  
  ' END OF METER INPUT KEYING ROUTINE *****************************************
  
  
GetCustRecNo:
  FOpenS "UBCUST.DAT", C2Handle 'open data file
  FGetRTA C2Handle, UBCustRec(1), PRec&, UBCustRecLen
  FClose C2Handle
  
  IF UBCustRec(1).Status <> "A" AND NoInactive THEN
    process = False
    IF PageUp THEN
      Cnt = Cnt - 1
      IF Cnt < 1 THEN
        PageUp = False
      END IF
    ELSE
      Cnt = Cnt + 1
    END IF
    GOTO NoInactRet
  END IF
  IF VAL(UBCustRec(1).BOOK) = BookNumber AND (UBCustRec(1).Status <> "F") THEN
    process = True
  ELSE
    process = False
    IF PageUp THEN
      IF Cnt <= 1 THEN
        PageUp = False
      ELSE
        Cnt = Cnt - 1
      END IF
    ELSE
      Cnt = Cnt + 1
    END IF
    RETURN
  END IF
  
  RecNo& = PRec&
  
  IF RecNo& <= 0 THEN
    process = False
    Cnt = Cnt + 1
    RETURN
  END IF
NoInactRet:
  RETURN
  
  
GetMeterType:
  SELECT CASE UBCustRec(1).LocMeters(MtrCnt).MTRType
  CASE "C"
    MeterType$ = "Water/Sewer"
  CASE "W"
    MeterType$ = "Water Only"
  CASE "S"
    MeterType$ = "Sewer Only"
  CASE "E"
    MeterType$ = "Electric"
  CASE "D"
    MeterType$ = "Demand Electric"
  CASE "T"
    MeterType$ = "Touch Read"
  CASE ELSE
    MeterType$ = "Undefined"
  END SELECT
  RETURN
  
GetMeterStatus:
  SELECT CASE UBCustRec(1).Status
  CASE "A"
    MeterStatus$ = "Active"
  CASE "F"
    MeterStatus$ = "Final"
  CASE "I"
    MeterStatus$ = "Vacant"
  CASE ELSE
    MeterStatus$ = "Undef."
  END SELECT
  RETURN
  
  
' GET ROUTE *******************************************************************
GetRoute:
  LibName$ = "UB"
  ScrnName$ = "UBMTRIN"
  
EditReadingsRec:
  NumScrns = 1
  
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
  
  '--Clear all fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT
  
  '--Set screen number to one and display screen
  Scr = 1
  DisplayUBScrn ScrnName$
  Action = 1
  
  FirstTime = True
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      FirstTime = False
      LSET Form$(2, 0) = DATE$
      Action = 1
    END IF
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      BookNumber% = VAL(Form$(1, 0))
      IF BookNumber% < 1 THEN
        LOCATE 13, 26
        COLOR 4
        PRINT "**** Invalid Book Number *****"; CHR$(7);
        WaitForAction
        GOTO GetRoute
      END IF
      
      MtrReadDate$ = Form$(2, 0)
      GOSUB CheckPostDate
      IF BadDate = True THEN
        LOCATE 13, 26
        COLOR 4
        PRINT "**** Improper Date Entry *****"; CHR$(7);
        WaitForAction
        GOTO GetRoute
      END IF
      LOCATE 13, 25
      COLOR 15
      PRINT "RETRIEVING BOOK NUMBER: "; BookNumber
      ExitFlag = True
    CASE EscKey
      EXIT SUB
    END SELECT
    
  LOOP UNTIL ExitFlag
  
RETURN
  ' END OF ROUTE GET ROUTINE ***************************************************
  
  
CheckPostDate:
  IF VAL(LEFT$(MtrReadDate$, 2)) < 1 OR VAL(LEFT$(MtrReadDate$, 2)) > 12 THEN BadDate = True: RETURN
  IF VAL(MID$(MtrReadDate$, 4, 2)) < 1 OR VAL(MID$(MtrReadDate$, 4, 2)) > 31 THEN BadDate = True: RETURN
  BadDate = False
RETURN
  
DoCalcUsage:
  ' do the calculation here
  NewRead# = VAL(Form$(10, 0))
  PastRead# = VAL(Form$(11, 0))
  IF NewRead# >= PastRead# THEN
    Calcusage# = NewRead# - PastRead#
  ELSE
    J# = LEN(STR$(PastRead#)) - 1: max# = 10 ^ J#
    Calcusage# = max# - PastRead# + NewRead#
  END IF
  Form$(12, 0) = STR$(Calcusage#)
  LOCATE 16, 53
  COLOR 3
  PRINT "Avg Use: "; USING "#######,#"; Average#
  LOCATE 17, 48
  PRINT "              "
  IF Calcusage# <> Average# THEN
    LSET Form$(13, 0) = "N"
  END IF
  IF AvgCnt! > 0 AND Average# > 0 THEN
    IF Calcusage# > (Average# * (HighVar / 100)) THEN
      HighRead = True
      COLOR 28
      LOCATE 17, 48
      PRINT "<-- HIGH READ";
      BEEP
    END IF
    IF Calcusage# < (Average# * (LowVar / 100)) THEN
      COLOR 28
      LOCATE 17, 48
      PRINT "<--LOW READ";
      LowRead = True
      BEEP
    END IF
  END IF
  Action = 1
  'if UBCustRec(1).EstFlag = "E"
  RETURN
  
END SUB

