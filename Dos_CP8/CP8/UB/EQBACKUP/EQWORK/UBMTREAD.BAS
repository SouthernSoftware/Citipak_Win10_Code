DEFINT A-Z
'need to add the Skip Inactive Flag Check.
DECLARE SUB CursorOff ()
DECLARE SUB ProcessItron ()
DECLARE SUB ProcessLogicon ()
DECLARE SUB PrintHighLowReport ()
DECLARE SUB ProcessTelxon ()
DECLARE SUB DisplayHHScrn (ScrnName$)
DECLARE SUB ProcessHusky ()
DECLARE SUB PrintMeterNotes ()
DECLARE FUNCTION Chk4BookSeqNum (BOOK$, SeqNum$)
DECLARE SUB PrintMeterSheets ()
DECLARE SUB EstMeterReading ()
DECLARE SUB PrintMeterReport ()
DECLARE SUB PrintMeterList ()
DECLARE SUB HandHeldProcessing ()
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetUpLen%)
DECLARE SUB PrintUtilBills ()
DECLARE SUB MakeSequenceIndex (IndexText$)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB MakeZipCodeIndex (IndexText$)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB MakePostalIndex (IndexText$)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB MakeTempIndex (Indextype%)
DECLARE SUB PreBillReport ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB ManualMeterReading ()
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB EnterMeterReadingsMenu ()
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE SUB ShowCursor ()
DECLARE SUB HideCursor ()
DECLARE SUB ClearScrn ()
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE FUNCTION FindRateTbl% (RateCode$, NumOfRates%, UBRateTbls() AS ANY)
DECLARE FUNCTION GetNumOfRevs% ()
DECLARE FUNCTION Round# (N#)
DECLARE FUNCTION GetNumOfCust% ()
DECLARE FUNCTION GetNumRateRecs% ()
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION FmtBook$ (BOOK$)
DECLARE FUNCTION FmtSeqN$ (SeqN$)
  '$INCLUDE: 'DefCnf.BI'
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB KILLFile (FileName$)
DECLARE SUB FClose (Handle%)
DECLARE FUNCTION FLof& (FileHand%)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB SortStr2 (BYVAL Address%, Size%, Direction%)
DECLARE SUB Search4LNumber (LNumber$, RecNunber&, CLSFlg%)
DECLARE SUB WaitForAction ()
DECLARE SUB BlockClear ()
DECLARE SUB SaveScrn (TempScrn())
DECLARE SUB RestScrn (TempScrn())
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION QPValI% (Numb$)
DECLARE FUNCTION QPValL& (Numb$)
  
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubdraft.BI'
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'Newcust.bi'
  '$INCLUDE: 'ubgilsen.bi'
  
  CONST False = 0, True = NOT False
  
  REDIM MChoice$(1 TO 9)
  
  MChoice$(1) = "Manual Meter Reading Entry "
  MChoice$(2) = "Hand-Held Meter Reading Entry"
  MChoice$(3) = "Estimated Meter Reading Entry"
  MChoice$(4) = "Print Meter Reading List"
  MChoice$(5) = "Print Reading Report"
  MChoice$(6) = "Print High/Low Report"
  MChoice$(7) = "Print Meter Reading Notes"
  MChoice$(8) = "Print Meter Reading Sheets"
  MChoice$(9) = "Exit to OS"
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 17   'limit the box length to go no lower than line 20
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  '--Find max menu width
  FOR Cnt! = 1 TO UBOUND(MChoice$)
    TLen = LEN(MChoice$(Cnt!))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT
  
  
  '--Center Menu within Screen
  Row = ((25 - (UBOUND(MChoice$))) \ 2) - 1
  Col = ((80 - MaxLen) \ 2)
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    
    LOCATE Row, Col, 0
    
    BlockClear
    TitleBox 2, Col, MaxLen + 3, "Meter Reading Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      ManualMeterReading
    CASE 2
      HandHeldProcessing
    CASE 3
      EstMeterReading
    CASE 4
      PrintMeterList
    CASE 5
      PrintMeterReport
    CASE 6
      PrintHighLowReport
    CASE 7
      PrintMeterNotes
    CASE 8
      PrintMeterSheets
    CASE 9
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP
  RUN "ubbillin"

SUB DisplayH2Scrn (ScrnName$)
  LibFile2Scrn "UBHH", ScrnName$, MonoCode%, Attribute%, ErrorCode%
END SUB

SUB DisplayHHScrn (ScrnName$)
  LibFile2Scrn "UBSETUP", ScrnName$, MonoCode%, Attribute%, ErrorCode%
END SUB

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
    LOOP UNTIL Cnt > NumOfRecs 'OR Cnt = 0
    
    EXIT SUB
  END IF
  ' End of Seq Order **********************************************************
  ' Location Order ??? ********************************************************
  
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
      
      'Average#
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

SUB PrintHelp (Help$)
  
  t$ = SPACE$(40)
  LSET t$ = Help$
  QPrintRC t$, 25, 2, -1
  t$ = ""
  
END SUB

SUB ProcessHusky
  
  IF Exist("UBHUSKY.CFG") THEN
    CFGFile = FREEFILE
    OPEN "UBHUSKY.CFG" FOR INPUT AS #CFGFile
    LINE INPUT #CFGFile, Port$
    CLOSE CFGFile
  ELSE
    CLS
    PRINT "CAN NOT FIND THE HUSKY CONFIGURATION FILE:  UBHUSKY.CFG"
    PRINT
    PRINT "Please call Southern Software for support. . ."
    PRINT
    PRINT "Press any key to EXIT.";
    WaitForAction
    EXIT SUB
  END IF
  
  q$ = CHR$(34)
  cb$ = SPACE$(45)
  SHARED Choice$()
  REDIM Route(100)
  
  LibName$ = "UBHH"
  ScrnName$ = "UBHUSKHH"
  'CursorOff
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  REDIM Choice$(5, 0)
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Send Info to HUSKY"
  Choice$(2, 0) = "Get Info From HUSKY"
  
  Action = 1
  
  DisplayHHScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      IF NOT (LEFT$(Form$(1, 0), 1) = "G" OR LEFT$(Form$(1, 0), 1) = "S") THEN
        COLOR 12: LOCATE 11, 40: PRINT "Invalid Operation": BEEP:
        WaitForAction
        COLOR 15
        LOCATE 11, 40
        PRINT STRING$(25, 32)
        Frm(1).FldNo = 1
        Action = 1
        Done = False
      ELSE
        Done = True
      END IF
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  
  SELECT CASE LEFT$(Form$(1, 0), 1)
  CASE "S"
    GOSUB SendHusky
  CASE "G"
    GOSUB GetHusky
  CASE ELSE
  END SELECT
  
  
  EXIT SUB
  
SendHusky:
  LibName$ = "UBHH"
  ScrnName$ = "UBHUSK1S"
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  'LOOK HERE
  
  REDIM Choice$(5, 0)
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Water/Sewer"
  Choice$(2, 0) = "Electric"
  Choice$(3, 0) = "Gas Meters"
  Choice$(4, 0) = "ALL METERS"
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  Action = 1
  NumberofRoutes = 0
  
  DisplayHHScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF Frm(1).PrevFld = 13 AND Frm(1).FldNo > 2 THEN
      IF VAL(Form$(2, 0)) <> 0 THEN
        NumberofRoutes = NumberofRoutes + 1
        IF NumberofRoutes <= 10 THEN
          Form$(NumberofRoutes + 2, 0) = Form$(2, 0)
          Frm(1).FldNo = 2
          Form$(2, 0) = ""
          Action = 1
        END IF
      END IF
    END IF
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      WhatTypes$ = LEFT$(Form$(1, 0), 1)
      ' Next Check Routes
      IF NumberofRoutes > 0 THEN
        QPrintRC "      Scanning Account:        ", 11, 21, -1
        ' Set Route Number Values
        FOR Cnt! = 1 TO NumberofRoutes
          Route(Cnt!) = VAL(Form$(Cnt! + 2, 0))
        NEXT Cnt!
        'Open Customer Data File
        REDIM UBCustRec(1) AS NewUBCustRecType
        UBCustRecLen = LEN(UBCustRec(1))
        UBFile = FREEFILE
        OPEN "UBCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBFile LEN = UBCustRecLen
        
        'Open and Initialize the PC3000 Read Information File
        REDIM UBPC3000RDRec(1) AS UBPC3000ReadRecType
        UBPC3000RdRecLen = LEN(UBPC3000RDRec(1))
        UBPC3000RdFile = FREEFILE
        OPEN "UBCUSTTR.DAT" FOR RANDOM SHARED AS UBPC3000RdFile LEN = UBPC3000RdRecLen
        CLOSE UBPC3000RdFile
        KILL "UBCUSTTR.DAT"
        UBPC3000RdFile = FREEFILE
        OPEN "UBCUSTTR.DAT" FOR RANDOM SHARED AS UBPC3000RdFile LEN = UBPC3000RdRecLen
        NumPC3000RdRecs = LOF(UBPC3000RdFile) / UBPC3000RdRecLen
        
        REDIM UBSetUpRec(1) AS UBSetupRecType
        LoadUBSetUpFile UBSetUpRec(), UBSetUpLen
        
        'Open the Correct Order for Reading
        IF UBSetUpRec(1).UseSeq = "Y" THEN
          IdxRecLen = 4         'we are using a integer
          MakeSequenceIndex "Sequence Number"
          IdxNumOfRecs = FileSize&("UBTEMP.IDX") \ 4
          REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
          FGetAH "UBTEMP.IDX", IdxBuff(1), 4, IdxNumOfRecs
        ELSE
          IdxRecLen = 4         'we are using a integer
          IdxFileSize& = FileSize&("UBCUSTBK.IDX")
          IdxNumOfRecs = IdxFileSize& \ IdxRecLen
          REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
          FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs            'load it
        END IF
        Cnt! = 1
        
        DO
          PRec& = IdxBuff(Cnt!).RecNum
          IF NOT (PRec&) = 0 THEN
            GET UBFile, PRec&, UBCustRec(1)
            IF INSTR(UBCustRec(1).HHMSG1, "NOREAD") > 0 THEN
              GOTO HWriteSkip
            END IF
            DisplayAcct$ = UBCustRec(1).BOOK + "-" + UBCustRec(1).SEQNUMB
            QPrintRC DisplayAcct$, 11, 45, -1
            FOR BookCnt! = 1 TO NumberofRoutes
              IF VAL(UBCustRec(1).BOOK) = Route(BookCnt!) AND (UBCustRec(1).Status <> "F") THEN
                GOSUB HuskyWriteRec
              END IF
            NEXT BookCnt!
          END IF
HWriteSkip:
          Cnt! = Cnt! + 1
        LOOP UNTIL Cnt! > IdxNumOfRecs
        CLOSE
        Ok = MsgBox%("UBSETUP", "UBHUSK1C")
        
        IF Ok = 1 THEN
          DisplayHHScrn "UBHUSK3C"
          QPrintRC "Transfering file to the HUSKY.", 10, 25, -1
          LOCATE 14, 40
          SHELL "cmdhcom " + Port$ + " /tx=ubcusttr.dat /noscreen /abort"
          'WaitForAction
          QPrintRC "    File Transfer Complete.   ", 10, 25, 15
          QPrintRC "  Press any key to continue.  ", 14, 25, 15
          BEEP
          CALL WaitForAction
          Done = True
        ELSE
          DisplayHHScrn "UBHUSK3C"
          QPrintRC "    File Transfer Aborted.    ", 10, 25, -1
          QPrintRC "  Press any key to continue.  ", 14, 25, -1
          BEEP
          CALL WaitForAction
          Done = True
          
        END IF
      END IF
      
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  
  RETURN
  
GetHusky:
  
  Ok = MsgBox%("UBSETUP", "UBHUSK1C")
  IF Ok = 1 THEN
    DisplayHHScrn "UBHUSK3C"
    QPrintRC "Transfering file from the HUSKY.", 10, 24, -1
    LOCATE 14, 40
    SHELL "cmdhcom " + Port$ + " /RX=ubcusttr.dat /noscreen /abort"
    'WaitForAction
    QPrintRC "     File Transfer Complete.    ", 10, 24, 15
    QPrintRC "   Press any key to continue.   ", 14, 24, 15
    BEEP
    CALL WaitForAction
    Done = True
  ELSE
    DisplayHHScrn "UBHUSK3C"
    QPrintRC "    File Transfer Aborted.    ", 10, 25, -1
    QPrintRC "  Press any key to continue.  ", 14, 25, -1
    BEEP
    CALL WaitForAction
    Done = True
    GOTO HuskyGetExit:
  END IF
  
  '    'Open Customer Data File
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBFile LEN = UBCustRecLen
  
  'Open and Initialize the PC3000 Read Information File
  REDIM UBPC3000GetRDRec(1) AS UBPC3000ReadRecType
  UBPC3000GetRdRecLen = LEN(UBPC3000GetRDRec(1))
  UBPC3000GetRdFile = FREEFILE
  FileNam$ = "UBCUSTTR.DAT"
  OPEN FileNam$ FOR RANDOM ACCESS READ WRITE SHARED AS UBPC3000GetRdFile LEN = UBPC3000GetRdRecLen
  NumPC3000GetRdRecs = LOF(UBPC3000GetRdFile) / UBPC3000GetRdRecLen
  
  IF NumPC3000GetRdRecs = 0 THEN
    CLOSE
    QPrintRC "                                ", 10, 24, 15
    QPrintRC "                                ", 14, 24, 15
    QPrintRC "   ERROR: NO READINGS FOUND   ", 9, 25, -1
    QPrintRC " Check the handheld connection", 11, 25, -1
    QPrintRC " and try the transfer again!!!", 13, 25, -1
    QPrintRC "   Press any key to continue. ", 15, 25, 14
    BEEP: BEEP: BEEP
    WaitForAction
    GOTO GetHusky
  END IF
  
  QPrintRC "  Processing Meter Readings.  ", 9, 25, 15
  QPrintRC "                              ", 10, 25, 11
  QPrintRC "    Updated Count:            ", 11, 25, 11
  QPrintRC "                              ", 14, 25, 11
  
  Cnt! = 1      ' Initialize File Counter to 1
  DO
    GET UBPC3000GetRdFile, Cnt!, UBPC3000GetRDRec(1)
    ' New Way Undelete When Our System Sends and Retrieves
    PRec& = VAL(QPTrim$(UBPC3000GetRDRec(1).Account))
    
    IF NOT (PRec&) = 0 THEN
      GET UBFile, PRec&, UBCustRec(1)
      GOSUB HuskyExtractRecord
    END IF
    Cnt! = Cnt! + 1
    
  LOOP UNTIL Cnt! > NumPC3000GetRdRecs
  
  CLOSE
  Done = True
  
  QPrintRC " Readings Updated Successfully. ", 13, 24, 15
  QPrintRC "  Press any key to continue.  ", 15, 25, 14
  BEEP
  WaitForAction
  Done = True
  
HuskyGetExit:
  RETURN
  
HuskyExtractRecord:
  UpdCnt! = UpdCnt! + 1
  QPrintRC STR$(UpdCnt!), 11, 44, 15
  'New Way To Activate
  MeterRecord = VAL(RIGHT$((QPTrim$(UBPC3000GetRDRec(1).Account)), 1))
  ' Check Meter Updated Flag
  ' Update Meter W/Reading
  CurReading# = UBPC3000GetRDRec(1).CurRead
  
  IF UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y" THEN
    UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
    UBCustRec(1).LocMeters(MeterRecord).CurDate = UBPC3000GetRDRec(1).ReadDate
  ELSE
    UBCustRec(1).LocMeters(MeterRecord).PrevRead = UBCustRec(1).LocMeters(MeterRecord).CurRead
    UBCustRec(1).LocMeters(MeterRecord).PastDate = UBCustRec(1).LocMeters(MeterRecord).CurDate
    UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y"
    UBCustRec(1).LocMeters(MeterRecord).CurDate = UBPC3000GetRDRec(1).ReadDate
    UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
  END IF
  'Modifed 04-28-97
  IF UBPC3000GetRDRec(1).NoteStatus = "P" THEN
    UBCustRec(1).HHMSG1 = UBPC3000GetRDRec(1).Note1
    UBCustRec(1).HHMSG2 = UBPC3000GetRDRec(1).Note2
    UBCustRec(1).HHMSG3 = UBPC3000GetRDRec(1).Note3
  END IF
  '   ^^^
  PUT UBFile, PRec&, UBCustRec(1)
  RETURN
  
  
HuskyWriteRec:
  'May Have Up to 7 Meters to Read
  FOR MtrCnt! = 1 TO 7
    MeterOK = False
    Account$ = STR$(PRec&)
    Account$ = LEFT$(Account$, 6) + "-" + RIGHT$(STR$(MtrCnt!), 1)
    
    IF ASC(UBCustRec(1).LocMeters(MtrCnt!).MTRType) > 32 THEN
      MTRType$ = UBCustRec(1).LocMeters(MtrCnt!).MTRType
      IF MTRType$ = "W" OR MTRType$ = "S" OR MTRType$ = "C" OR MTRType$ = "E" OR MTRType$ = "D" OR MTRType$ = "G" THEN
        SELECT CASE WhatTypes$
        CASE "W"                'water/sewer
          IF MTRType$ = "W" OR MTRType$ = "S" OR MTRType$ = "C" THEN
            MeterOK = True
          END IF
        CASE "E"                'electric & demand elec.
          IF MTRType$ = "E" OR MTRType$ = "D" THEN
            MeterOK = True
          END IF
        CASE "G"                'gas
          IF MTRType$ = "G" THEN
            MeterOK = True
          END IF
        CASE "A", " "           'all meters
          MeterOK = True
        END SELECT
        
        IF MeterOK = True THEN
          
          ' Determine High and Low Reading
          Average# = UBCustRec(1).LocMeters(MtrCnt!).AvgUse
          
          ILowRead$ = RIGHT$(STR$((UBCustRec(1).LocMeters(MtrCnt!).CurRead)), 8)
          LowRead# = FIX(ILowRead#)
          HighRead# = FIX(Average# * (HighVar / 100)) + UBCustRec(1).LocMeters(MtrCnt!).CurRead
          
          
          MeterID$ = LTRIM$(UBCustRec(1).LocMeters(MtrCnt!).MTRNUM)
          MeterID$ = RTRIM$(MeterID$)
          
          IF VAL(MeterID$) = 0 THEN
            MeterID$ = UBCustRec(1).BOOK + UBCustRec(1).SEQNUMB
          END IF
          IF LEN(MeterID$) < 8 THEN
            MeterID$ = STRING$(8 - LEN(MeterID$), "0") + MeterID$
          END IF
          MeterID$ = LEFT$(MeterID$, 8)
          
          'Set Record Fields and Put On Disk
          UBPC3000RDRec(1).CUSTNAME = UBCustRec(1).CUSTNAME
          UBPC3000RDRec(1).ServAddress = UBCustRec(1).SERVADDR
          UBPC3000RDRec(1).MeterID = MeterID$
          UBPC3000RDRec(1).LowRead = LowRead#
          UBPC3000RDRec(1).HighRead = HighRead#
          UBPC3000RDRec(1).Account = Account$
          UBPC3000RDRec(1).MeterType$ = UBCustRec(1).LocMeters(MtrCnt!).MTRType
          UBPC3000RDRec(1).BOOK = VAL(UBCustRec(1).BOOK)
          UBPC3000RDRec(1).CurRead = UBCustRec(1).LocMeters(MtrCnt!).CurRead
          UBPC3000RDRec(1).PastRead = 0
          UBPC3000RDRec(1).ReadFlag = "N"
          'Modifed 04-28-97
          UBPC3000RDRec(1).Note1 = UBCustRec(1).HHMSG1
          UBPC3000RDRec(1).Note2 = UBCustRec(1).HHMSG2
          UBPC3000RDRec(1).Note3 = UBCustRec(1).HHMSG3
          '         ^^^
          UBPC3000RDRec(1).NoteStatus = ""
          PUT UBPC3000RdFile, (LOF(UBPC3000RdFile) / UBPC3000RdRecLen) + 1, UBPC3000RDRec(1)
        END IF
      END IF
    END IF
  NEXT MtrCnt!
  RETURN
  
END SUB

SUB ProcessItron
  SHARED Choice$()
  DIM Route(10)
  
  REDIM UBItronRec(1) AS UBItronRecType
  REDIM UBItronFRec(1) AS ItronFType
  REDIM UBItronARec(1) AS ItronAType
  REDIM UBItronBRec(1) AS ItronBType
  REDIM UBItronHRec(1) AS ItronHType
  REDIM UBItronZRec(1) AS ItronZType


  UBItronRecLen = LEN(UBItronRec(1))
  UBItronFile = FREEFILE
  OPEN "UBITRON.DAT" FOR RANDOM SHARED AS UBItronFile LEN = UBItronRecLen
  NumItronRecs = LOF(UBItronFile) / UBItronRecLen
  IF NumItronRecs = 1 THEN
    GET UBItronFile, 1, UBItronRec(1)
    PathWay$ = UBItronRec(1).PathWay
  END IF
  
  LibName$ = "UBSETUP"
  ScrnName$ = "UBITRONH"
  Help$ = "Itron Hand Held Processing"
  CursorOff
  'LOCATE 1, 1, 0
  
  Help$ = "Choose HandHeld Operation"
  
  DisplayHHScrn ScrnName$
  PrintHelp Help$
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  REDIM Choice$(3, 0)
  
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Send Info to Itron"
  Choice$(2, 0) = "Get Info From Itron"
  
  IF NumItronRecs = 1 THEN
    Form$(2, 0) = PathWay$
    'Fld(2).Protected = True
  END IF
  
  Action = 1
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      IF NOT (LEFT$(Form$(1, 0), 1) = "G" OR LEFT$(Form$(1, 0), 1) = "S") THEN
        COLOR 12
        LOCATE 11, 40
        PRINT "Invalid Operation"
        BEEP
        WaitForAction
        COLOR 15
        LOCATE 11, 40
        PRINT STRING$(25, 32)
        
        Frm(1).FldNo = 1
        Action = 1
        Done = False
        
      ELSE
        PathWay$ = Form$(2, 0)
        UBItronRec(1).PathWay = PathWay$
        PUT UBItronFile, 1, UBItronRec(1)
        CLOSE UBItronFile
        IF LEFT$(Form$(1, 0), 1) = "S" THEN GOSUB SendItron
        IF LEFT$(Form$(1, 0), 1) = "G" THEN GOSUB GetItron
        Done = True
      END IF
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  
  EXIT SUB
  
SendItron:
  LibName$ = "UBSETUP"
  ScrnName$ = "UBITRONS"
  Help$ = "Itron Hand Held Processing"
  CursorOff
  'LOCATE 1, 1, 0
  
  Help$ = "DownLoad HandHeld Operation"
  DisplayHHScrn ScrnName$
  PrintHelp Help$
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  Action = 1
  NumberofRoutes = 0
  
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    SELECT CASE Frm(1).KeyCode
    CASE F10Key 'First Check For Valid Interrogator Number and Routes
      GOSUB FILENAM1Itron       'Set Route Number Values
      FOR Cnt! = 1 TO NumberofRoutes
        Route(Cnt!) = VAL(Form$(Cnt! + 2, 0))
      NEXT Cnt!
      'Open Customer Data File
      REDIM UBCustRec(1) AS NewUBCustRecType
      UBCustRecLen = LEN(UBCustRec(1))
      UBFile = FREEFILE
      OPEN "UBCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBFile LEN = UBCustRecLen
      
      'Open Itron Date File
      UBItronRdRecLen = 80      '80 byte records
      UBItronRdFile = FREEFILE
      OPEN FileName$ FOR RANDOM SHARED AS UBItronRdFile LEN = 80
      CLOSE UBItronRdFile: KILL FileName$
      UBItronRdFile = FREEFILE
      OPEN FileName$ FOR RANDOM SHARED AS UBItronRdFile LEN = 80
      NumItronRdRecs = 0
      
      'Open the Location Order for Reading
      IdxRecLen = 4             'we are using a integer
      IdxFileSize& = FileSize&("UBCUSTBK.IDX")
      IdxNumOfRecs = IdxFileSize& \ IdxRecLen
      REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
      FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs                'load it
      Cnt! = 1
      
      
      DO
        PRec& = IdxBuff(Cnt!).RecNum
        LOCATE 11, 42: COLOR 15: PRINT USING "######"; PRec&
        IF NOT (PRec&) = 0 THEN
          GET UBFile, PRec&, UBCustRec(1)
          IF VAL(UBCustRec(1).BOOK) <> 0 THEN
            'Take Care of First Record Sent First
            IF FirstOnRoute = 0 THEN
              FirstOnRoute = VAL(UBCustRec(1).BOOK)
              UBItronFRec(1).RecordCode = "F"
              UBItronFRec(1).Route = UBCustRec(1).BOOK + "000000"
              UBItronFRec(1).Message = ""
              UBItronFRec(1).Filler = ""
              UBItronFRec(1).CRLF = CHR$(13) + CHR$(10)
              PUT #UBItronRdFile, 1, UBItronFRec(1)
            END IF
            
            'Compare to Check For End of Book Run
            IF FirstOnRoute < VAL(UBCustRec(1).BOOK) THEN
              'Create Z Record Trailor
              Route$ = LTRIM$(STR$(FirstOnRoute))
              IF LEN(Route$) = 1 THEN Route$ = "0" + Route$ + "000000"
              
              NumberofAccounts$ = LTRIM$(STR$(NumberofAccounts))
              IF LEN(NumberofAccounts$) < 4 THEN NumberofAccounts$ = STRING$(4 - LEN(NumberofAccounts$), "0") + NumberofAccounts$
              
              UBItronZRec(1).RecordCode = "Z"
              UBItronZRec(1).Route = Route$
              UBItronZRec(1).NumberAccts = NumberofAccounts$
              UBItronZRec(1).NumberMeters = NumberofAccounts$
              UBItronZRec(1).Filler = ""
              UBItronZRec(1).CRLF = CHR$(13) + CHR$(10)
              PUT #UBItronRdFile, (LOF(UBItronRdFile) / 80) + 1, UBItronZRec(1)
              
              'Create Next Route Header F
              FirstOnRoute = VAL(UBCustRec(1).BOOK)
              NumberofAccounts = 0
              UBItronFRec(1).RecordCode = "F"
              UBItronFRec(1).Route = UBCustRec(1).BOOK + "000000"
              UBItronFRec(1).Message = ""
              UBItronFRec(1).Filler = ""
              UBItronFRec(1).CRLF = CHR$(13) + CHR$(10)
              PUT #UBItronRdFile, (LOF(UBItronRdFile) / 80) + 1, UBItronFRec(1)
            END IF
            
            GOSUB WriteRecordItron
          END IF
        END IF
        Cnt! = Cnt! + 1
        
      LOOP UNTIL Cnt! > IdxNumOfRecs
      
      'Create Z Record Trailor End of Records
      Route$ = LTRIM$(STR$(FirstOnRoute))
      IF LEN(Route$) = 1 THEN Route$ = "0" + Route$ + "000000"
      
      NumberofAccounts$ = LTRIM$(STR$(NumberofAccounts))
      IF LEN(NumberofAccounts$) < 4 THEN NumberofAccounts$ = STRING$(4 - LEN(NumberofAccounts$), "0") + NumberofAccounts$
      
      UBItronZRec(1).RecordCode = "Z"
      UBItronZRec(1).Route = Route$
      UBItronZRec(1).NumberAccts = NumberofAccounts$
      UBItronZRec(1).NumberMeters = NumberofAccounts$
      UBItronZRec(1).Filler = ""
      UBItronZRec(1).CRLF = CHR$(13) + CHR$(10)
      PUT #UBItronRdFile, (LOF(UBItronRdFile) / 80) + 1, UBItronZRec(1)
      
      
      CLOSE
      Done = True
      COLOR 15
      LOCATE 13, 20
      PRINT "DONE: FILE  FROMHOST.PPS  CREATED"
      PRINT CHR$(7);
      WaitForAction
      Done = True
      
      
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  RETURN
  
  
WriteRecordItron:
  'May Have Up to 1 Meters to Read Per Account
  MtrCnt! = 1
  
  'Set Account Number to 6 Byte String Here
  Account$ = LTRIM$(STR$(PRec&))
  IF LEN(Account$) < 6 THEN Account$ = STRING$(6 - LEN(Account$), "0") + Account$
  
  'Set Sequence Number
  SeqNumber$ = LTRIM$(STR$(UBCustRec(1).SEQ))
  IF LEN(SeqNumber$) < 5 THEN SeqNumber$ = STRING$(5 - LEN(SeqNumber$), "0") + SeqNumber$
  
  'Set Customer Type
  CustomerType$ = LEFT$(QPTrim$(UBCustRec(1).CUSTTYPE), 1)
  IF LEN(CustomerType$) = 0 THEN
    CustomerType$ = "R"
  END IF
  
  'Set Customer Name
  CustomerName$ = QPTrim$(UBCustRec(1).CUSTNAME)
  'Set Customer Address
  CustomerAddress$ = QPTrim$(UBCustRec(1).SERVADDR)
  
  
  IF (ASC(UBCustRec(1).LocMeters(MtrCnt!).MTRType) > 32) THEN
    IF UBCustRec(1).LocMeters(MtrCnt!).MTRType = "C" OR UBCustRec(1).LocMeters(MtrCnt!).MTRType = "W" OR UBCustRec(1).LocMeters(MtrCnt!).MTRType = "S" THEN
      
      IF LEN(QPTrim$(UBCustRec(1).USERCODE1)) = 0 THEN
        UBItronHRec(1).Dials = "4"
      ELSE
        UBItronHRec(1).Dials = QPTrim$(UBCustRec(1).USERCODE1)
      END IF
      DialVal# = Value#(UBItronHRec(1).Dials, ECode)
      
      'Determine High and Low Reading
      Average# = UBCustRec(1).LocMeters(MtrCnt!).AvgUse
      ILowRead$ = LTRIM$(STR$((UBCustRec(1).LocMeters(MtrCnt!).CurRead)))
      
      IF LEN(ILowRead$) > DialVal# THEN
        ILowRead$ = RIGHT$(ILowRead$, DialVal#)
      END IF
      
      IF LEN(ILowRead$) > 8 THEN ILowRead$ = RIGHT$(ILowRead$, 8)
      IF LEN(ILowRead$) < 8 THEN ILowRead$ = STRING$(8 - LEN(ILowRead$), "0") + ILowRead$
      
      HighRead# = FIX(Average# * (HighVar / 100)) + UBCustRec(1).LocMeters(MtrCnt!).CurRead
      IF Average# = 0 THEN
        HighRead# = VAL(ILowRead$) + 150        'Fudge Factor Until History Built
      END IF
      
      IHighRead$ = LTRIM$(STR$(HighRead#))
      
      IF LEN(IHighRead$) > DialVal# THEN
        IHighRead$ = RIGHT$(IHighRead$, DialVal#)
      END IF
      
      IHighRead$ = RIGHT$(IHighRead$, 8)
      IF LEN(IHighRead$) > 8 THEN IHighRead$ = RIGHT$(IHighRead$, 8)
      IF LEN(IHighRead$) < 8 THEN IHighRead$ = STRING$(8 - LEN(IHighRead$), "0") + IHighRead$
      
      MeterID$ = LTRIM$(UBCustRec(1).LocMeters(MtrCnt!).MTRNUM)
      MeterID$ = RTRIM$(MeterID$)
      
      IF LEN(MeterID$) < 12 THEN
        MeterID$ = STRING$(12 - LEN(MeterID$), " ") + MeterID$
      END IF
      MeterID$ = LEFT$(MeterID$, 12)
      
      'Update Counter of Accounts
      NumberofAccounts = NumberofAccounts + 1
      
      'Create A Record
      UBItronARec(1).RecordCode = "A"
      UBItronARec(1).Route = UBCustRec(1).BOOK + "000000"
      UBItronARec(1).AcctNumb = Account$
      UBItronARec(1).Geo = ""
      UBItronARec(1).SEQNUMB = SeqNumber$
      UBItronARec(1).Message = "0"
      UBItronARec(1).AType = CustomerType$
      UBItronARec(1).Filler = ""
      UBItronARec(1).CRLF = CHR$(13) + CHR$(10)
      PUT #UBItronRdFile, (LOF(UBItronRdFile) / 80) + 1, UBItronARec(1)
      'Create B Record
      UBItronBRec(1).RecordCode = "B"
      UBItronBRec(1).CUSTNAME = CustomerName$
      UBItronBRec(1).CustAddr = CustomerAddress$
      UBItronBRec(1).Filler = ""
      UBItronBRec(1).CRLF = CHR$(13) + CHR$(10)
      PUT #UBItronRdFile, (LOF(UBItronRdFile) / 80) + 1, UBItronBRec(1)
      'Create H Record
      UBItronHRec(1).RecordCode = "H"
      UBItronHRec(1).MeterNumb = MeterID$
      
      
      UBItronHRec(1).LowRead = ILowRead$
      UBItronHRec(1).HighRead = IHighRead$
      UBItronHRec(1).LastRead = ILowRead$
      
      
      UBItronHRec(1).MeterType = "W"
      UBItronHRec(1).Multiplier = "000000"
      UBItronHRec(1).NoMths = "0"
      
      MtMake$ = QPTrim$(UBCustRec(1).USERCODE2)
      
      IF LEN(MtMake$) = 0 THEN
        UBItronHRec(1).MtrMake = "00"
      ELSE
        MtMake$ = "00" + MtMake$
        UBItronHRec(1).MtrMake = RIGHT$(MtMake$, 2)
      END IF
      UBItronHRec(1).DispCode = "  "
      UBItronHRec(1).NumbDec = "0"
      UBItronHRec(1).MustRead = "1"
      UBItronHRec(1).Status = "A"
      UBItronHRec(1).Filler = ""
      UBItronHRec(1).CRLF = CHR$(13) + CHR$(10)
      PUT #UBItronRdFile, (LOF(UBItronRdFile) / 80) + 1, UBItronHRec(1)
      
    END IF
  END IF
  
SkipEmItron:
  RETURN
  
FILENAM1Itron:
  PathWay$ = QPTrim$(PathWay$)
  FileName$ = PathWay$ + "FROMHOST.PPS"
  RETURN
  
GetItron:
  LibName$ = "UBSETUP"
  ScrnName$ = "UBITRONG"
  Help$ = "Itron Hand Held Processing"
  CursorOff
  'LOCATE 1, 1, 0
  
  Help$ = "Upload HandHeld Operation"
  DisplayHHScrn ScrnName$
  PrintHelp Help$
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  Action = 1
  NumberofRoutes = 0
  
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
        GOSUB FileNamItron      ' Get Itron File

        'Open Customer Data File
        REDIM UBCustRec(1) AS NewUBCustRecType
        UBCustRecLen = LEN(UBCustRec(1))
        UBFile = FREEFILE
        OPEN "UBCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBFile LEN = UBCustRecLen
        
        
        UBLogGetRdFile = FREEFILE
        OPEN FileName$ FOR INPUT AS UBLogGetRdFile
        
        IF LOF(UBLogGetRdFile) = 0 THEN
          LOCATE 12, 20: COLOR 12
          PRINT "            [  E R R O R  ]"
          LOCATE 13, 20: COLOR 11
          PRINT "     Nothing Found From Itron "
          BEEP
          WaitForAction
          CLOSE
          GOTO GetItron
        END IF

          
          DO
            LINE INPUT #UBLogGetRdFile, a$
            'Process C Line
            IF LEFT$(a$, 1) = "C" THEN
              ReadDate$ = MID$(a$, 25, 4)
              PRec& = VAL(MID$(a$, 10, 10))         'Grabs Acct#
             'Process D Line
              LINE INPUT #UBLogGetRdFile, a$
             'Process E Line
              LINE INPUT #UBLogGetRdFile, a$
              CurReading# = VAL(MID$(a$, 3, 8))
              ReadTime$ = MID$(a$, 12, 6)
             'Update the Account Now
              IF NOT (PRec& = 0 OR CurReading# = 0) THEN
               GET UBFile, PRec&, UBCustRec(1)
               GOSUB ExtractRecordItron
              END IF
            END IF
NextLoop:
        
        LOOP UNTIL EOF(UBLogGetRdFile)
        
        CLOSE
        Done = True
        FOR Itron1 = 8 TO 10
          LOCATE Itron1, 16
          PRINT STRING$(40, 32)
        NEXT Itron1
        LOCATE 8, 18
        COLOR 11
        PRINT "Itron Unit Number "; RouteID$
        LOCATE 13, 18
        COLOR 11
        PRINT "Reading Updated From Itron System Successfully"
        PRINT CHR$(7);
        WaitForAction
      
      Done = True
      
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  RETURN
  
  
ExtractRecordItron:
  UpdCnt! = UpdCnt! + 1
  LOCATE 13, 20: COLOR 11
  PRINT " Updated Count:  "; : COLOR 15: PRINT UpdCnt!
  COLOR 11
  




  

  MeterRecord = 1               'Only 1 meter per account for Itron

  MeterReadDate$ = LEFT$(ReadDate$, 2) + "/" + MID$(ReadDate$, 3, 2) + "/" + RIGHT$(DATE$, 4)
  
  IF UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y" THEN
    UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
    UBCustRec(1).LocMeters(MeterRecord).CurDate = Date2Num(MeterReadDate$)
  ELSE
    UBCustRec(1).LocMeters(MeterRecord).PrevRead = UBCustRec(1).LocMeters(MeterRecord).CurRead
    UBCustRec(1).LocMeters(MeterRecord).PastDate = UBCustRec(1).LocMeters(MeterRecord).CurDate
    UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y"
    UBCustRec(1).LocMeters(MeterRecord).CurDate = Date2Num(MeterReadDate$)
    UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
  END IF
  PUT UBFile, PRec&, UBCustRec(1)
  RETURN
  
FileNamItron:
  PathWay$ = QPTrim$(PathWay$)
  FileName$ = PathWay$ + "TOHOST.PPS"
  RETURN
  
END SUB

SUB ProcessLogicon
  SHARED Choice$()
  DIM Route(10)
  REDIM UBLogiconRec(1) AS UBLogiconRecType
  UBLogiconRecLen = LEN(UBLogiconRec(1))
  UBLogiconFile = FREEFILE
  OPEN "UBLogcon.DAT" FOR RANDOM SHARED AS UBLogiconFile LEN = UBLogiconRecLen
  NumLogiconRecs = LOF(UBLogiconFile) / UBLogiconRecLen
  IF NumLogiconRecs = 1 THEN
    GET UBLogiconFile, 1, UBLogiconRec(1)
    PathWay$ = UBLogiconRec(1).PathWay
  END IF
  
  LibName$ = "UBSETUP"
  ScrnName$ = "UBLOGHH"
  Help$ = "Logicon Hand Held Processing"
  CursorOff
  'LOCATE 1, 1, 0
  
  Help$ = "Choose HandHeld Operation"
  
  DisplayHHScrn ScrnName$
  PrintHelp Help$
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  REDIM Choice$(3, 0)
  
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Send Info to Logicon"
  Choice$(2, 0) = "Get Info From Logicon"
  
  IF NumLogiconRecs = 1 THEN
    Form$(2, 0) = PathWay$
    'Fld(2).Protected = True
  END IF
  
  Action = 1
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      IF NOT (LEFT$(Form$(1, 0), 1) = "G" OR LEFT$(Form$(1, 0), 1) = "S") THEN
        COLOR 12
        LOCATE 11, 40
        PRINT "Invalid Operation"
        BEEP
        WaitForAction
        COLOR 15
        LOCATE 11, 40
        PRINT STRING$(25, 32)
        
        Frm(1).FldNo = 1
        Action = 1
        Done = False
        
      ELSE
        PathWay$ = Form$(2, 0)
        UBLogiconRec(1).PathWay = PathWay$
        PUT UBLogiconFile, 1, UBLogiconRec(1)
        CLOSE UBLogiconFile
        IF LEFT$(Form$(1, 0), 1) = "S" THEN GOSUB SendLogicon
        IF LEFT$(Form$(1, 0), 1) = "G" THEN GOSUB GetLogicon
        Done = True
      END IF
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  
  EXIT SUB
  
SendLogicon:
  LibName$ = "UBSETUP"
  ScrnName$ = "UBLOG1S"
  Help$ = "Logicon Hand Held Processing"
  CursorOff
  'LOCATE 1, 1, 0
  
  Help$ = "DownLoad HandHeld Operation"
  DisplayHHScrn ScrnName$
  PrintHelp Help$
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  Action = 1
  NumberofRoutes = 0
  
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF Frm(1).PrevFld = 13 AND Frm(1).FldNo > 2 THEN
      IF VAL(Form$(2, 0)) <> 0 THEN
        NumberofRoutes = NumberofRoutes + 1
        IF NumberofRoutes <= 10 THEN
          Form$(NumberofRoutes + 2, 0) = Form$(2, 0)
          Frm(1).FldNo = 2
          Form$(2, 0) = ""
          Action = 1
        END IF
      END IF
    END IF
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      ' First Check For Valid Interrogator Number and Routes
      IF VAL(Form$(1, 0)) < 1 THEN
        Frm(1).FldNo = 1: Action = 1: Done = False
        
      ELSE
        
        IF NumberofRoutes > 0 THEN
          RouteID$ = Form$(1, 0)
          COLOR 11
          LOCATE 12, 18
          PRINT "LOADING THESE ROUTES !!!                      "
          GOSUB FILENAM1LC
          'Set Route Number Values
          FOR Cnt! = 1 TO NumberofRoutes
            Route(Cnt!) = VAL(Form$(Cnt! + 2, 0))
          NEXT Cnt!
          
          'Open Customer Data File
          REDIM UBCustRec(1) AS NewUBCustRecType
          UBCustRecLen = LEN(UBCustRec(1))
          UBFile = FREEFILE
          OPEN "UBCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBFile LEN = UBCustRecLen
          
          'Open Logicon Date File
          REDIM UBLogRdRec(1) AS UBLogiconReadRecType
          UBLogRdRecLen = LEN(UBLogRdRec(1))
          UBLogRdFile = FREEFILE
          OPEN FileName$ FOR RANDOM SHARED AS UBLogRdFile LEN = UBLogRdRecLen
          CLOSE UBLogRdFile: KILL FileName$
          UBLogRdFile = FREEFILE
          OPEN FileName$ FOR RANDOM SHARED AS UBLogRdFile LEN = UBLogRdRecLen
          NumLogRdRecs = LOF(UBLogRdFile) / UBLogRdRecLen
          
          'Write First Record With Route Information
          UBLogRdRec(1).RecType = "H"
          UBLogRdRec(1).RouteNo = RouteID$
          UBLogRdRec(1).AcctNo = ""
          UBLogRdRec(1).RecName = ""
          UBLogRdRec(1).ServAddress = ""
          UBLogRdRec(1).ReadDate = ""
          UBLogRdRec(1).ReadTime = ""
          UBLogRdRec(1).Consumption = ""
          UBLogRdRec(1).PrevRead = ""
          UBLogRdRec(1).CurRead = ""
          UBLogRdRec(1).LowRead = ""
          UBLogRdRec(1).HighRead = ""
          UBLogRdRec(1).MtrNumb = ""
          UBLogRdRec(1).CountChg = ""
          UBLogRdRec(1).ForceFlag = ""
          UBLogRdRec(1).ReportCode = ""
          UBLogRdRec(1).Remark = ""
          UBLogRdRec(1).Label = ""
          UBLogRdRec(1).PrintFlag = ""
          UBLogRdRec(1).MessageOut = ""
          UBLogRdRec(1).BOOK = ""
          UBLogRdRec(1).Future = ""
          UBLogRdRec(1).Recend = "X"
          UBLogRdRec(1).CRLF = CHR$(13) + CHR$(10)
          PUT UBLogRdFile, (LOF(UBLogRdFile) / UBLogRdRecLen) + 1, UBLogRdRec(1)
          
          'Open the Location Order for Reading
          IdxRecLen = 4         'we are using a integer
          IdxFileSize& = FileSize&("UBCUSTBK.IDX")
          IdxNumOfRecs = IdxFileSize& \ IdxRecLen
          REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
          FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs            'load it
          
          Cnt! = 1
          
          DO
            
            PRec& = IdxBuff(Cnt!).RecNum
            IF NOT (PRec&) = 0 THEN
              GET UBFile, PRec&, UBCustRec(1)
              FOR BookCnt! = 1 TO NumberofRoutes
                IF VAL(UBCustRec(1).BOOK) = Route(BookCnt!) AND (UBCustRec(1).Status <> "F") THEN
                  GOSUB WriteRecordLC
                END IF
              NEXT BookCnt!
            END IF
            Cnt! = Cnt! + 1
            
          LOOP UNTIL Cnt! > IdxNumOfRecs
          
          CLOSE
          Done = True
          FOR lc = 8 TO 10
            LOCATE lc, 16
            PRINT STRING$(40, 32)
          NEXT lc
          LOCATE 8, 18
          COLOR 11
          PRINT "Route Number "; RouteID$
          LOCATE 9, 18
          COLOR 11
          PRINT "Information Sent to Logicon System Successfully"
          LOCATE 12, 18: PRINT "THESE ROUTES LOADED!!!                      "
          PRINT CHR$(7);
          WaitForAction
        END IF
        Done = True
      END IF
      
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  RETURN
  
  
WriteRecordLC:
  'May Have Up to 10 Meters to Read
  MtrCnt! = 1
  Account$ = STR$(PRec&)
  Account$ = LEFT$(Account$, 6) + "-" + RIGHT$(STR$(MtrCnt!), 1)
  
  WHILE MtrCnt! < 8
    
    IF (ASC(UBCustRec(1).LocMeters(MtrCnt!).MTRType) > 32) THEN
      IF UBCustRec(1).LocMeters(MtrCnt!).MTRType = "C" OR UBCustRec(1).LocMeters(MtrCnt!).MTRType = "W" OR UBCustRec(1).LocMeters(MtrCnt!).MTRType = "T" OR UBCustRec(1).LocMeters(MtrCnt!).MTRType = "S" THEN
        
        ' Determine High and Low Reading
        Average# = UBCustRec(1).LocMeters(MtrCnt!).AvgUse
        ILowRead$ = RIGHT$(STR$((UBCustRec(1).LocMeters(MtrCnt!).CurRead)), 8)
        ILowRead# = VAL(ILowRead$)
        HighRead# = FIX(Average# * (HighVar / 100)) + UBCustRec(1).LocMeters(MtrCnt!).CurRead
        IF FIX(HighRead#) = ILowRead# THEN HighRead# = HighRead# + 12000
        IHighRead$ = STR$(HighRead#)
        IHighRead$ = RIGHT$(IHighRead$, 8)
        MeterID$ = LTRIM$(UBCustRec(1).LocMeters(MtrCnt!).MTRNUM)
        
        MeterID$ = RTRIM$(MeterID$)
        IF VAL(MeterID$) = 0 THEN
          MeterID$ = UBCustRec(1).BOOK + UBCustRec(1).SEQNUMB
        END IF
        IF LEN(MeterID$) < 8 THEN
          MeterID$ = STRING$(8 - LEN(MeterID$), "0") + MeterID$
        END IF
        MeterID$ = LEFT$(MeterID$, 8)
        
        UBLogRdRec(1).RecType = "A"
        UBLogRdRec(1).RouteNo = RouteID$
        UBLogRdRec(1).AcctNo = Account$
        UBLogRdRec(1).RecName = UBCustRec(1).CUSTNAME
        UBLogRdRec(1).ServAddress = UBCustRec(1).SERVADDR
        UBLogRdRec(1).ReadDate = "      "
        UBLogRdRec(1).ReadTime = "      "
        UBLogRdRec(1).Consumption = "        "
        UBLogRdRec(1).PrevRead = ILowRead$
        UBLogRdRec(1).CurRead = "XXXXXXXX"
        UBLogRdRec(1).LowRead = ILowRead$
        UBLogRdRec(1).HighRead = IHighRead$
        UBLogRdRec(1).MtrNumb = MeterID$
        UBLogRdRec(1).CountChg = "0"
        UBLogRdRec(1).ForceFlag = " "
        UBLogRdRec(1).ReportCode = "--"
        UBLogRdRec(1).Remark = ""
        UBLogRdRec(1).Label = ""
        UBLogRdRec(1).PrintFlag = ""
        UBLogRdRec(1).MessageOut = "OutPut to Meter Reader"
        UBLogRdRec(1).BOOK = UBCustRec(1).BOOK
        UBLogRdRec(1).Future = ""
        UBLogRdRec(1).Recend = "X"
        UBLogRdRec(1).CRLF = CHR$(13) + CHR$(10)
        PUT UBLogRdFile, (LOF(UBLogRdFile) / UBLogRdRecLen) + 1, UBLogRdRec(1)
      END IF
    END IF
    
SkipEmLC:
    MtrCnt! = MtrCnt! + 1
  WEND
  RETURN
  
FILENAM1LC:
  PathWay$ = QPTrim$(PathWay$)
  FileName$ = PathWay$ + "UBLOGNO" + LTRIM$(RouteID$) + ".DAT"
  RETURN
  
GetLogicon:
  LibName$ = "UBSETUP"
  ScrnName$ = "UBLOG1G"
  Help$ = "Logicon Hand Held Processing"
  CursorOff
  'LOCATE 1, 1, 0
  
  Help$ = "Upload HandHeld Operation"
  DisplayHHScrn ScrnName$
  PrintHelp Help$
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  Action = 1
  NumberofRoutes = 0
  
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      IF VAL(Form$(1, 0)) < 1 OR VAL(Form$(2, 0)) < 1 THEN
        LOCATE 12, 20: COLOR 12
        PRINT "            [  E R R O R  ]"
        LOCATE 13, 20: COLOR 11
        PRINT "     Invalid Route ID Number "
        LOCATE 14, 20: COLOR 11
        PRINT "      or Invalid Reading Date "
        BEEP
        WaitForAction
        CLOSE
        GOTO GetLogicon
        Frm(1).FldNo = 1
        Action = 1
        GOTO GetLogicon
      ELSE
        RouteID$ = Form$(1, 0)
        GOSUB FileNamLC         ' Get Logicon File
        'Open Customer Data File
        REDIM UBCustRec(1) AS NewUBCustRecType
        UBCustRecLen = LEN(UBCustRec(1))
        UBFile = FREEFILE
        OPEN "UBCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBFile LEN = UBCustRecLen
        
        REDIM UBLogGetRdRec(1) AS UBLogiconGetReadRecType
        UBLogGetRdRecLen = LEN(UBLogGetRdRec(1))
        
        UBLogGetRdFile = FREEFILE
        OPEN FileName$ FOR RANDOM ACCESS READ WRITE SHARED AS UBLogGetRdFile LEN = UBLogGetRdRecLen
        NumLogGetRdRecs = LOF(UBLogGetRdFile) / UBLogGetRdRecLen
        
        IF NumLogGetRdRecs = 0 THEN
          LOCATE 12, 20: COLOR 12
          PRINT "            [  E R R O R  ]"
          LOCATE 13, 20: COLOR 11
          PRINT "     Nothing Found From Logicon "
          LOCATE 14, 20: COLOR 11
          PRINT "   Please Reenter Route ID Number "
          BEEP
          WaitForAction
          CLOSE
          GOTO GetLogicon
        END IF
        
        Cnt! = 1                ' Initialize File Counter to 1
        DO
          GET UBLogGetRdFile, Cnt!, UBLogGetRdRec(1)
          PRec& = VAL(QPTrim$(UBLogGetRdRec(1).AcctNo))
          
          IF LEFT$(UBLogGetRdRec(1).CurRead, 1) <> "X" THEN
            IF NOT (PRec&) = 0 THEN
              GET UBFile, PRec&, UBCustRec(1)
              GOSUB ExtractRecordLC
            END IF
          END IF
          Cnt! = Cnt! + 1
          
        LOOP UNTIL Cnt! > NumLogGetRdRecs
        
        CLOSE
        Done = True
        FOR lc = 8 TO 10
          LOCATE lc, 16
          PRINT STRING$(40, 32)
        NEXT lc
        LOCATE 8, 18
        COLOR 11
        PRINT "Logicon Unit Number "; RouteID$
        LOCATE 13, 18
        COLOR 11
        PRINT "Reading Updated From Logicon System Successfully"
        PRINT CHR$(7);
        WaitForAction
      END IF
      Done = True
      
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  RETURN
  
  
ExtractRecordLC:
  UpdCnt! = UpdCnt! + 1
  LOCATE 13, 20: COLOR 11
  PRINT " Updated Count:  "; : COLOR 15: PRINT UpdCnt!
  COLOR 11
  
  
  'New Way To Activate
  MeterRecord = 1               'Only 1 meter per account for logicon
  ' Update Meter W/Reading
  CurReading# = VAL(UBLogGetRdRec(1).CurRead)
  MeterReadDate$ = MID$(UBLogGetRdRec(1).ReadDate, 3, 2) + "/" + MID$(UBLogGetRdRec(1).ReadDate, 5, 2) + "/" + RIGHT$(DATE$, 4)
  
  IF UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y" THEN
    UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
    UBCustRec(1).LocMeters(MeterRecord).CurDate = Date2Num(MeterReadDate$)
  ELSE
    UBCustRec(1).LocMeters(MeterRecord).PrevRead = UBCustRec(1).LocMeters(MeterRecord).CurRead
    UBCustRec(1).LocMeters(MeterRecord).PastDate = UBCustRec(1).LocMeters(MeterRecord).CurDate
    UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y"
    UBCustRec(1).LocMeters(MeterRecord).CurDate = Date2Num(MeterReadDate$)
    UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
  END IF
  PUT UBFile, PRec&, UBCustRec(1)
  RETURN
  
FileNamLC:
  PathWay$ = QPTrim$(PathWay$)
  FileName$ = PathWay$ + "UBLOGNO" + LTRIM$(RouteID$) + ".DAT"
  RETURN
  
END SUB

SUB ProcessTelxon
  
  q$ = CHR$(34)
  cb$ = SPACE$(45)
  SHARED Choice$()
  REDIM Route(100)
  
  LibName$ = "UBSetup"
  ScrnName$ = "UBTELXHH"
  'CursorOff
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  REDIM Choice$(5, 0)
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Send Info to Telxon"
  Choice$(2, 0) = "Get Info From Telxon"
  
  Action = 1
  
  DisplayHHScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      IF NOT (LEFT$(Form$(1, 0), 1) = "G" OR LEFT$(Form$(1, 0), 1) = "S") THEN
        COLOR 12: LOCATE 11, 40: PRINT "Invalid Operation": BEEP:
        WaitForAction
        COLOR 15
        LOCATE 11, 40
        PRINT STRING$(25, 32)
        
        Frm(1).FldNo = 1
        Action = 1
        Done = False
      ELSE
        Done = True
      END IF
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  
  SELECT CASE LEFT$(Form$(1, 0), 1)
  CASE "S"
    GOSUB SendTelxon
  CASE "G"
    GOSUB GetTelxon
  CASE ELSE
  END SELECT
  
  CLOSE
  EXIT SUB
  
SendTelxon:
  LibName$ = "UBSetup"
  ScrnName$ = "UBTELX1S"
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  'LOOK HERE
  
  REDIM Choice$(5, 0)
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Water/Sewer"
  Choice$(2, 0) = "Electric"
  Choice$(3, 0) = "Gas Meters"
  Choice$(4, 0) = "ALL METERS"
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  Action = 1
  NumberofRoutes = 0
  
  DisplayHHScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF Frm(1).PrevFld = 13 AND Frm(1).FldNo > 2 THEN
      IF VAL(Form$(2, 0)) <> 0 THEN
        NumberofRoutes = NumberofRoutes + 1
        IF NumberofRoutes <= 10 THEN
          Form$(NumberofRoutes + 2, 0) = Form$(2, 0)
          Frm(1).FldNo = 2
          Form$(2, 0) = ""
          Action = 1
        END IF
      END IF
    END IF
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      WhatTypes$ = LEFT$(Form$(1, 0), 1)
      
      ' Next Check Routes
      IF NumberofRoutes > 0 THEN
        QPrintRC "      Scanning Account:        ", 11, 21, -1
        
        ' Set Route Number Values
        FOR Cnt! = 1 TO NumberofRoutes
          Route(Cnt!) = VAL(Form$(Cnt! + 2, 0))
        NEXT Cnt!
        
        'Open Customer Data File
        REDIM UBCustRec(1) AS NewUBCustRecType
        UBCustRecLen = LEN(UBCustRec(1))
        UBFile = FREEFILE
        OPEN "UBCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBFile LEN = UBCustRecLen
        
        'Open and Initialize the PC3000 Read Information File
        REDIM UBPC3000RDRec(1) AS UBPC3000ReadRecType
        UBPC3000RdRecLen = LEN(UBPC3000RDRec(1))
        UBPC3000RdFile = FREEFILE
        OPEN "UBCUSTTR.DAT" FOR RANDOM SHARED AS UBPC3000RdFile LEN = UBPC3000RdRecLen
        CLOSE UBPC3000RdFile
        KILL "UBCUSTTR.DAT"
        UBPC3000RdFile = FREEFILE
        OPEN "UBCUSTTR.DAT" FOR RANDOM SHARED AS UBPC3000RdFile LEN = UBPC3000RdRecLen
        NumPC3000RdRecs = LOF(UBPC3000RdFile) / UBPC3000RdRecLen
        
        REDIM UBSetUpRec(1) AS UBSetupRecType
        LoadUBSetUpFile UBSetUpRec(), UBSetUpLen
        
        
        'Open the Correct Order for Reading
        IF UBSetUpRec(1).UseSeq = "Y" THEN
          IdxRecLen = 4         'we are using a integer
          MakeSequenceIndex "Sequence Number"
          IdxNumOfRecs = FileSize&("UBTEMP.IDX") \ 4
          REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
          FGetAH "UBTEMP.IDX", IdxBuff(1), 4, IdxNumOfRecs
        ELSE
          IdxRecLen = 4         'we are using a integer
          IdxFileSize& = FileSize&("UBCUSTBK.IDX")
          IdxNumOfRecs = IdxFileSize& \ IdxRecLen
          REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
          FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs            'load it
        END IF
        Cnt! = 1
        
        DO
          
          PRec& = IdxBuff(Cnt!).RecNum
          IF NOT (PRec&) = 0 THEN
            GET UBFile, PRec&, UBCustRec(1)
            DisplayAcct$ = UBCustRec(1).BOOK + "-" + UBCustRec(1).SEQNUMB
            QPrintRC DisplayAcct$, 11, 45, -1
            FOR BookCnt! = 1 TO NumberofRoutes
              IF VAL(UBCustRec(1).BOOK) = Route(BookCnt!) AND (UBCustRec(1).Status <> "F") THEN
                GOSUB TelxonPutRec
              END IF
            NEXT BookCnt!
          END IF
          Cnt! = Cnt! + 1
          
        LOOP UNTIL Cnt! > IdxNumOfRecs
        CLOSE
        
        Ok = MsgBox%("UBSETUP", "UBTELX1C")
        
        IF Ok = 1 THEN
          KILLFile "TELXON.ERR"
          DisplayHHScrn "UBTELX3C"
          QPrintRC "Transfering file to the TELXON.", 10, 25, -1
          REDIM TScrn(0)
          SaveScrn TScrn()
          SHELL "procomm /FSend.cmd"
          RestScrn TScrn()
          GOSUB CheckSendRetFile
          WaitForAction
          Done = True
        ELSE
          DisplayHHScrn "UBTELX3C"
          QPrintRC "    File Transfer Aborted.    ", 10, 25, -1
          QPrintRC "  Press any key to continue.  ", 14, 25, -1
          BEEP
          CALL WaitForAction
          Done = True
        END IF
      END IF
      
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  
  RETURN
  
GetTelxon:
  
  Ok = MsgBox%("UBSETUP", "UBTELX2C")
  IF Ok = 1 THEN
    KILLFile "TELXON.ERR"
    KILLFile "UBCUSTTR.DAT"
    REDIM TScrn(0)
    SaveScrn TScrn()
    SHELL "procomm /FGet.cmd"
    RestScrn TScrn()
    DisplayHHScrn "UBTELX3C"
    GOSUB CheckSendRetFile
    IF Exist("TELXON.ERR") THEN
      WaitForAction
      GOTO TelxonGetExit
    END IF
    Done = True
  ELSE
    DisplayHHScrn "UBTELX3C"
    QPrintRC "    File Transfer Aborted.    ", 10, 25, -1
    QPrintRC "  Press any key to continue.  ", 14, 25, -1
    BEEP
    CALL WaitForAction
    Done = True
    GOTO TelxonGetExit
  END IF
  
  '    'Open Customer Data File
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBFile LEN = UBCustRecLen
  
  'Open and Initialize the PC3000 Read Information File
  REDIM UBPC3000GetRDRec(1) AS UBPC3000ReadRecType
  UBPC3000GetRdRecLen = LEN(UBPC3000GetRDRec(1))
  UBPC3000GetRdFile = FREEFILE
  FileNam$ = "UBCUSTTR.DAT"
  OPEN FileNam$ FOR RANDOM ACCESS READ WRITE SHARED AS UBPC3000GetRdFile LEN = UBPC3000GetRdRecLen
  NumPC3000GetRdRecs = LOF(UBPC3000GetRdFile) / UBPC3000GetRdRecLen
  
  IF NumPC3000GetRdRecs = 0 THEN
    CLOSE
    FOR zz = 8 TO 16
      QPrintRC SPACE$(49), zz, 16, 15
    NEXT
    
    DisplayHHScrn "UBTELX3C"
    QPrintRC "ERROR: NO READINGS FOUND", 9, 27, -1
    QPrintRC " Check the handheld connection", 10, 24, -1
    QPrintRC " and try the transfer again!!!", 11, 24, -1
    QPrintRC "   Press any key to continue. ", 13, 24, 14
    BEEP
    WaitForAction
    GOTO GetTelxon
  END IF
  DisplayHHScrn "UBTELX3C"
  
  QPrintRC "  Processing Meter Readings.  ", 9, 25, 15
  QPrintRC "    Updated Count:            ", 11, 25, 11
  
  Cnt! = 1      ' Initialize File Counter to 1
  DO
    GET UBPC3000GetRdFile, Cnt!, UBPC3000GetRDRec(1)
    ' New Way Undelete When Our System Sends and Retrieves
    PRec& = VAL(QPTrim$(UBPC3000GetRDRec(1).Account))
    
    IF NOT (PRec&) = 0 THEN
      GET UBFile, PRec&, UBCustRec(1)
      GOSUB TelxonExtRec
    END IF
    Cnt! = Cnt! + 1
    
  LOOP UNTIL Cnt! > NumPC3000GetRdRecs
  
  CLOSE
  Done = True
  
  QPrintRC " Readings Updated Successfully. ", 13, 24, 15
  QPrintRC "  Press any key to continue.  ", 15, 25, 14
  BEEP
  WaitForAction
  Done = True
  
TelxonGetExit:
  RETURN
  
TelxonExtRec:
  UpdCnt! = UpdCnt! + 1
  QPrintRC STR$(UpdCnt!), 11, 44, 15
  'New Way To Activate
  MeterRecord = VAL(RIGHT$((QPTrim$(UBPC3000GetRDRec(1).Account)), 1))
  ' Check Meter Updated Flag
  ' Update Meter W/Reading
  CurReading# = UBPC3000GetRDRec(1).CurRead
  
  IF UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y" THEN
    UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
    UBCustRec(1).LocMeters(MeterRecord).CurDate = UBPC3000GetRDRec(1).ReadDate
  ELSE
    UBCustRec(1).LocMeters(MeterRecord).PrevRead = UBCustRec(1).LocMeters(MeterRecord).CurRead
    UBCustRec(1).LocMeters(MeterRecord).PastDate = UBCustRec(1).LocMeters(MeterRecord).CurDate
    UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y"
    UBCustRec(1).LocMeters(MeterRecord).CurDate = UBPC3000GetRDRec(1).ReadDate
    UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
  END IF
  'Modifed 04-28-97
  IF UBPC3000GetRDRec(1).NoteStatus = "P" THEN
    UBCustRec(1).HHMSG1 = UBPC3000GetRDRec(1).Note1
    UBCustRec(1).HHMSG2 = UBPC3000GetRDRec(1).Note2
    UBCustRec(1).HHMSG3 = UBPC3000GetRDRec(1).Note3
  END IF
  '   ^^^
  PUT UBFile, PRec&, UBCustRec(1)
  RETURN
  
TelxonPutRec:
  'May Have Up to 7 Meters to Read
  FOR MtrCnt! = 1 TO 7
    MeterOK = False
    Account$ = STR$(PRec&)
    Account$ = LEFT$(Account$, 6) + "-" + RIGHT$(STR$(MtrCnt!), 1)
    
    IF ASC(UBCustRec(1).LocMeters(MtrCnt!).MTRType) > 32 THEN
      MTRType$ = UBCustRec(1).LocMeters(MtrCnt!).MTRType
      IF MTRType$ = "W" OR MTRType$ = "S" OR MTRType$ = "C" OR MTRType$ = "E" OR MTRType$ = "D" OR MTRType$ = "G" THEN
        SELECT CASE WhatTypes$
        CASE "W"                'water/sewer
          IF MTRType$ = "W" OR MTRType$ = "S" OR MTRType$ = "C" THEN
            MeterOK = True
          END IF
        CASE "E"                'electric & demand elec.
          IF MTRType$ = "E" OR MTRType$ = "D" THEN
            MeterOK = True
          END IF
        CASE "G"                'gas
          IF MTRType$ = "G" THEN
            MeterOK = True
          END IF
        CASE "A", " "           'all meters
          MeterOK = True
        END SELECT
        
        IF MeterOK = True THEN
          ' Determine High and Low Reading
          Average# = UBCustRec(1).LocMeters(MtrCnt!).AvgUse
          
          ILowRead$ = RIGHT$(STR$((UBCustRec(1).LocMeters(MtrCnt!).CurRead)), 8)
          LowRead# = FIX(ILowRead#)
          HighRead# = FIX(Average# * (HighVar / 100)) + UBCustRec(1).LocMeters(MtrCnt!).CurRead
          MeterID$ = LTRIM$(UBCustRec(1).LocMeters(MtrCnt!).MTRNUM)
          MeterID$ = RTRIM$(MeterID$)
          
          IF VAL(MeterID$) = 0 THEN
            MeterID$ = UBCustRec(1).BOOK + UBCustRec(1).SEQNUMB
          END IF
          IF LEN(MeterID$) < 8 THEN
            MeterID$ = STRING$(8 - LEN(MeterID$), "0") + MeterID$
          END IF
          MeterID$ = LEFT$(MeterID$, 8)
          
          'Set Record Fields and Put On Disk
          UBPC3000RDRec(1).CUSTNAME = UBCustRec(1).CUSTNAME
          UBPC3000RDRec(1).ServAddress = UBCustRec(1).SERVADDR
          UBPC3000RDRec(1).MeterID = MeterID$
          UBPC3000RDRec(1).LowRead = LowRead#
          UBPC3000RDRec(1).HighRead = HighRead#
          UBPC3000RDRec(1).Account = Account$
          UBPC3000RDRec(1).MeterType$ = UBCustRec(1).LocMeters(MtrCnt!).MTRType
          UBPC3000RDRec(1).BOOK = VAL(UBCustRec(1).BOOK)
          UBPC3000RDRec(1).CurRead = UBCustRec(1).LocMeters(MtrCnt!).CurRead
          UBPC3000RDRec(1).PastRead = 0
          UBPC3000RDRec(1).ReadFlag = "N"
          'Modifed 04-28-97
          UBPC3000RDRec(1).Note1 = UBCustRec(1).HHMSG1
          UBPC3000RDRec(1).Note2 = UBCustRec(1).HHMSG2
          UBPC3000RDRec(1).Note3 = UBCustRec(1).HHMSG3
          '         ^^^
          UBPC3000RDRec(1).NoteStatus = ""
          PUT UBPC3000RdFile, (LOF(UBPC3000RdFile) / UBPC3000RdRecLen) + 1, UBPC3000RDRec(1)
        END IF
      END IF
    END IF
  NEXT MtrCnt!
  RETURN
  
CheckSendRetFile:
  IF Exist("TELXON.ERR") THEN
    QPrintRC "ERROR:                          ", 9, 23, 15
    QPrintRC "CHECK THE CABLE CONNECTION AND   ", 10, 23, 15
    QPrintRC "TRY TO TRANSFER THE FILE AGAIN. ", 11, 23, 15
  ELSE
    QPrintRC "     File Transfer Complete. . . ", 10, 23, 15
  END IF
  QPrintRC "  Press any key to continue.  ", 14, 25, 15
  RETURN
  
END SUB

