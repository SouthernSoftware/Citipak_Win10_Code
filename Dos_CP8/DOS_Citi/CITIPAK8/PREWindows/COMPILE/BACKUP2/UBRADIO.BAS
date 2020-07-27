DEFINT A-Z
DECLARE SUB GetRadioReads ()
DECLARE SUB RadioReadIF ()
DECLARE SUB InactiveConsReport ()
DECLARE SUB CursorOff ()
DECLARE SUB ProcessItron ()
DECLARE SUB ProcessLogicon ()
DECLARE SUB PrintHighLowReport ()
DECLARE SUB ProcessTelxon ()
DECLARE SUB DisplaySUScrn (ScrnName$)
DECLARE SUB ProcessHusky ()
DECLARE SUB PrintMeterNotes ()
DECLARE SUB PrintMeterSheets ()
DECLARE SUB EstMeterReading ()
DECLARE SUB PrintMeterReport ()
DECLARE SUB PrintMeterList ()
DECLARE SUB HandHeldProcessing ()
DECLARE SUB LoadUBSetUpFile (UBSetupRec() AS ANY, UBSetUpLen%)
DECLARE SUB MakeSequenceIndex (IndexText$)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB ManualMeterReading ()
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB EnterMeterReadingsMenu ()
DECLARE SUB PrintHelp (H$)
DECLARE SUB ShowCursor ()
DECLARE SUB HideCursor ()
DECLARE SUB ClearScrn ()
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB KillFile (FileName$)
DECLARE SUB FClose (Handle%)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB WaitForAction ()
DECLARE SUB BlockClear ()
DECLARE SUB SaveScrn (TempScrn())
DECLARE SUB RestScrn (TempScrn())
DECLARE FUNCTION Chk4BookSeqNum (BOOK$, SeqNum$)
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE FUNCTION Round# (N#)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION FmtBook$ (BOOK$)
DECLARE FUNCTION FmtSeqN$ (SeqN$)
DECLARE FUNCTION FLof& (FileHand%)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION QPValI% (Numb$)
DECLARE FUNCTION QPValL& (Numb$)
DECLARE FUNCTION GetBadgerPath$ ()
DECLARE FUNCTION GetAccount& (Account$)
DECLARE FUNCTION GetReading& (Reading$)

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubdraft.BI'
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'Newcust.bi'
  '$INCLUDE: 'ubradio.bi'
  
  CONST False = 0, True = NOT False
  
  REDIM MChoice$(1 TO 3)
  
  MChoice$(1) = " Radio Read Host Download "
  MChoice$(2) = " Radio Read Host Upload   "
  MChoice$(3) = " Exit to OS               "
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 17   'limit box to go no lower than line 17
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
  Col = ((80 - MaxLen) \ 2) - 1
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    
    LOCATE Row, Col, 0
    
    BlockClear
    
    TitleBox 2, Col, MaxLen + 3, "Badger Reading Menu", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      RadioReadIF
    CASE 2
      GetRadioReads
    CASE 3
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP
  RUN "ubmtread"

SUB DisplaySUScrn (ScrnName$)
  LibFile2Scrn "UBSETUP", ScrnName$, MonoCode%, Attribute%, ErrorCode%
END SUB

SUB GetRadioReads

  WIDTH LPRINT 255

  SHARED Choice$()

  REDIM Rdg(1) AS BADReadRecType
  RDGLen = LEN(Rdg(1))
  
  Path$ = GetBadgerPath$
  
  IF LEN(Path$) > 0 AND RIGHT$(Path$, 1) <> "\" THEN
    Path$ = Path$ + "\"
  END IF

  FileNam$ = Path$ + "UPLOAD.DAT"

  LibName$ = "UBSetup"
  ScrnName$ = "UBRADIOC"

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
  FirstTime = True

BadPathRestart:

  Action = 1
  Frm(1).FldNo = 1
  BlockClear
  DisplaySUScrn ScrnName$

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      LSET Form$(1, 0) = Path$
      Action = 1
    END IF

    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      GOSUB CheckBadgerPath
      IF InfoOK = False THEN
        GOTO BadPathRestart
      END IF
      GOSUB GetBadgerReadings
      Done = True
    EXIT DO
    CASE ESC
      Done = True
      GOTO AbortExit
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done

  CLOSE
  QPrintRC "Processing Complete!", 13, 30, 14
  QPrintRC "Press any key to continue.", 15, 27, 15
  WaitForAction

AbortExit:

  EXIT SUB

'**************************************************
GetBadgerReadings:

'Open Customer Data File
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  
  REDIM Rdg(1) AS BADReadRecType
  RDGLen = LEN(Rdg(1))
  ReadFile = FREEFILE
  OPEN FileNam$ FOR RANDOM SHARED AS ReadFile LEN = RDGLen
  NumReadRecs = LOF(ReadFile) / RDGLen

  IF NumReadRecs = 0 THEN
    CLOSE
    FOR zz = 8 TO 16
      QPrintRC SPACE$(49), zz, 16, 15
    NEXT
    DisplaySUScrn "BADMTRUP"
    QPrintRC " ERROR: NO READINGS FOUND", 9, 27, -1
    QPrintRC " Check the Badger Path and ", 10, 24, -1
    QPrintRC " try the transfer again!!!", 11, 24, -1
    QPrintRC " Press any key to continue. ", 13, 24, 14
    BEEP
    WaitForAction
  END IF

  BlockClear
  DisplaySUScrn "BADMTRUP"

  QPrintRC "  Processing Meter Readings.  ", 9, 25, 15
  QPrintRC "    Updated Count:            ", 11, 25, 11

  CustRec& = 0
  GotCust = False
  FOR Cnt& = 1 TO NumReadRecs
    GET ReadFile, Cnt&, Rdg(1)
    RecordID$ = LEFT$(Rdg(1).Fill, 3)
    SELECT CASE RecordID$
    CASE "CUS"
      CustRec& = GetAccount&(MID$(Rdg(1).Fill, 15, 20))
      CUSTNAME$ = QPTrim$(MID$(Rdg(1).Fill, 35, 20))
    CASE "RDG"
      GOSUB ProcessReading
      CustRec& = 0
    CASE ELSE
      'ignore all others
    END SELECT
  NEXT

  CLOSE

RETURN

ProcessReading:
  IF CustRec& <= 0 THEN
    GOTO BadCustRet
  END IF
  GET UBFile, CustRec&, UBCustRec(1)

  UpdCnt! = UpdCnt! + 1
  QPrintRC STR$(UpdCnt!), 11, 44, 15

'  MeterRecord = VAL(RIGHT$((QPTrim$(UBPC3000GetRDRec(1).Account)), 1))
  MeterRecord = 1
  ' Check Meter Updated Flag
  ' Update Meter W/Reading
  RDate$ = MID$(Rdg(1).Fill, 48, 8)
  ReadDate = Date2Num(LEFT$(RDate$, 2) + "/" + MID$(RDate$, 3, 2) + "/" + RIGHT$(RDate$, 4))

  CurReading# = GetReading&(MID$(Rdg(1).Fill, 34, 10))

'  CurReading# = GetReading&(MID$(Rdg(1).Fill, 32, 10))
 ' CurReading# = CurReading# * 100
  'IF CurReading# > 0 THEN STOP

  IF UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y" THEN
    UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
    UBCustRec(1).LocMeters(MeterRecord).CurDate = ReadDate
  ELSE
    UBCustRec(1).LocMeters(MeterRecord).PrevRead = UBCustRec(1).LocMeters(MeterRecord).CurRead
    UBCustRec(1).LocMeters(MeterRecord).PastDate = UBCustRec(1).LocMeters(MeterRecord).CurDate
    UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y"
    UBCustRec(1).LocMeters(MeterRecord).CurDate = ReadDate
    UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
  END IF
  
  PUT UBFile, CustRec&, UBCustRec(1)

BadCustRet:

RETURN

'**************************************************
CheckBadgerPath:
  InfoOK = True
  NewPath$ = QPTrim$(Form$(1, 0))
  Path$ = GetBadgerPath$

  IF RIGHT$(NewPath$, 1) <> "\" THEN
    NewPath$ = NewPath$ + "\"
  END IF

'if this is a change to the path to badger then update ubdata
  IF NewPath$ <> Path$ THEN
    PathFile = FREEFILE
    OPEN PathData$ FOR OUTPUT AS PathFile
    PRINT #PathFile, NewPath$
    CLOSE PathFile
  END IF
  Path$ = GetBadgerPath$

  IF RIGHT$(Path$, 1) <> "\" THEN
    Path$ = Path$ + "\"
  END IF

  FileNam$ = Path$ + "UPLOAD.DAT"
  IF INSTR(Path$, "BAD") > 0 THEN
    BlockClear
    DisplaySUScrn "Invalid"
    WaitForAction
    InfoOK = False
  END IF

RETURN

END SUB

SUB PrintHelp (Help$)
  
  t$ = SPACE$(40)
  LSET t$ = Help$
  QPrintRC t$, 25, 2, -1
  t$ = ""
  
END SUB

SUB RadioReadIF

  q$ = CHR$(34)
  cb$ = SPACE$(45)
  SHARED Choice$()

  'Open and Find Correct Read Order
  REDIM UBSetupRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetupRec(), UBSetUpLen

  IF INSTR(UBSetupRec(1).UTILNAME, "BROADWAY") > 0 THEN
    BroadFlag = True
  END IF

  'PathData$ = "UBBAGPTH.DAT"

  Path$ = GetBadgerPath$

  REDIM Route(100)
  REDIM FHD(1) AS FDRTYPE
  REDIM CHD(1) AS CDRTYPE
  REDIM RHD(1) AS RDRType
  REDIM CUS(1) AS CUSType
  REDIM MTR(1) AS MTRType
  REDIM Rdg(1) AS RDGType
  REDIM RFF(1) AS RFFType
  REDIM WRR(1) AS WRRType

  LibName$ = "UBSetup"
  ScrnName$ = "UBRADIOS"
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
  
BadDataRestart:
  Action = 1
  FirstTime = True
  BlockClear
  DisplaySUScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF FirstTime THEN
      FirstTime = False
      LSET Form$(4, 0) = Path$
      Action = 1
    END IF
    
    SELECT CASE Frm(1).KeyCode
     
    CASE F10Key
      GOSUB CheckInfo
      IF InfoOK THEN
        Cycle$ = QPTrim$(Form$(1, 0))
        IF LEN(Cycle$) = 1 THEN Cycle$ = "0" + Cycle$
        Route$ = QPTrim$(Form$(2, 0))
        IF LEN(Route$) = 1 THEN Route$ = "0" + Route$
        Route$ = "0000" + Route$
        GOSUB HostDownLoad
        MReading$ = STR$(TCnt!) '# of Readings RDG to Expect
        MReading$ = LTRIM$(MReading$)
        TReading$ = MReading$
        KReading$ = STR$(KCnt!) '# of Keyed Readings RDG to Expect
        KReading$ = LTRIM$(KReading$)
        RReading$ = STR$(RCnt!) '# of Radio Readings RDG to Expect
        RReading$ = LTRIM$(RReading$)
        IF LEN(MReading$) < 4 THEN
          MReading$ = STRING$(4 - LEN(MReading$), "0") + MReading$
        END IF
        IF LEN(TReading$) < 4 THEN
          TReading$ = STRING$(4 - LEN(TReading$), "0") + TReading$
        END IF
        IF LEN(KReading$) < 4 THEN
          KReading$ = STRING$(4 - LEN(KReading$), "0") + KReading$
        END IF
        IF LEN(RReading$) < 4 THEN
          RReading$ = STRING$(4 - LEN(RReading$), "0") + RReading$
        END IF

        'Check the len of the reading here
        '
        Reading$ = LTRIM$(STR$(TCnt!))
        IF VAL(Reading$) < 1000 THEN
          Reading$ = STRING$(4 - LEN(Reading$), "0") + Reading$
        END IF
        GOSUB BuildDownLoad
        Done = True
      ELSE
        GOTO BadDataRestart
      END IF
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  
  CLOSE
  EXIT SUB
  
HostDownLoad:
  
  ' Next Check Routes
  
  QPrintRC "      Scanning Account:        ", 15, 21, -1

  'Open and Initialize Work Files

  KillFile "UBHHCUS.DAT"
  KillFile "UBHHMTR.DAT"
  KillFile "UBHHRDG.DAT"
  KillFile "UBHHRFF.DAT"
  KillFile "UBHHWRR.DAT"

  'Open Customer Data File
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  
'  'Open and Find Correct Read Order
'  REDIM UBSetupRec(1) AS UBSetupRecType
'  LoadUBSetUpFile UBSetupRec(), UBSetUpLen

  HighReadPerc = UBSetupRec(1).HighRead
  LowReadPerc = UBSetupRec(1).LowRead
  
  'Open the Correct Order for Reading
  IF UBSetupRec(1).UseSeq = "Y" THEN
    IdxRecLen = 4               'we are using a integer
    MakeSequenceIndex "Sequence Number"
    IdxNumOfRecs = FileSize&("UBTEMP.IDX") \ 4
    REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
    FGetAH "UBTEMP.IDX", IdxBuff(1), 4, IdxNumOfRecs
  ELSE
    IdxRecLen = 4               'we are using a integer
    IdxFileSize& = FileSize&("UBCUSTBK.IDX")
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
    FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs  'load it
  END IF
  Cnt! = 1
  
  'Open HandHeld Work Files
  CUSFile = FREEFILE
  OPEN "UBHHCUS.DAT" FOR RANDOM SHARED AS CUSFile LEN = LEN(CUS(1))
  MTRFile = FREEFILE
  OPEN "UBHHMTR.DAT" FOR RANDOM SHARED AS MTRFile LEN = LEN(MTR(1))
  RDGFile = FREEFILE
  OPEN "UBHHRDG.DAT" FOR RANDOM SHARED AS RDGFile LEN = LEN(Rdg(1))
  RFFFile = FREEFILE
  OPEN "UBHHRFF.DAT" FOR RANDOM SHARED AS RFFFile LEN = LEN(RFF(1))
  WRRFile = FREEFILE
  OPEN "UBHHWRR.DAT" FOR RANDOM SHARED AS WRRFile LEN = LEN(WRR(1))
  
  DO
    
    Prec& = IdxBuff(Cnt!).RecNum
    RecNumber! = Prec&

    IF NOT (Prec&) = 0 THEN
      GET UBFile, Prec&, UBCustRec(1)
      DisplayAcct$ = UBCustRec(1).CUSTNAME
      QPrintRC DisplayAcct$, 15, 35, -1
      IF UBCustRec(1).BILLCYCL = VAL(Cycle$) AND VAL(UBCustRec(1).BOOK) = VAL(Route$) AND (UBCustRec(1).Status <> "F") THEN
        'Do SomeThing Here
        'Build Customer Meter Reading & Radio Read File
        TCnt! = TCnt! + 1
        'Build Customer First

        MeterNumber$ = UBCustRec(1).LocMeters(1).MTRNUM
        MNum2File$ = MeterNumber$
        IF LEFT$(MNum2File$, 2) = "WR" THEN
          MNum2File$ = MID$(MNum2File$, 3)
        END IF

        MeterSeq$ = STR$(Cnt! * 10)
        IF LEN(MeterSeq$) < 8 THEN
          MeterSeq$ = STRING$(8 - LEN(MeterSeq$), "0") + MeterSeq$
        END IF
        
        Dials = VAL(UBCustRec(1).USERCODE2)
        IF Dials = 0 THEN Dials = 7             'Default for Caldwell Cty
        Dials$ = "0" + RIGHT$(STR$(Dials), 1)
        Decimals$ = "00"
        
        IF LEFT$(MeterNumber$, 2) = "RR" THEN
          ERTID$ = MID$(MeterNumber$, 3, 8)
          ReadMethod$ = "R"
          RCnt! = RCnt! + 1
        ELSEIF LEFT$(MeterNumber$, 2) = "WR" THEN
          ERTID$ = MID$(MeterNumber$, 3, 8)
          ReadMethod$ = "W"
          RCnt! = RCnt! + 1
        ELSE
          ERTID$ = "00000000"
          ReadMethod$ = "K"
          KCnt! = KCnt! + 1
        END IF

'broadway W ^^

        PrevRead$ = LTRIM$(STR$(UBCustRec(1).LocMeters(1).CurRead))
        PrevRead$ = STRING$(10 - LEN(PrevRead$), "0") + PrevRead$

        'need to fix these
        
        IF UBCustRec(1).LocMeters(1).AvgUse < 0 THEN
          UBCustRec(1).LocMeters(1).AvgUse = 0
        END IF
       'HighRead# = FIX(((0# + UBCustRec(1).LocMeters(1).AvgUse) * (HighReadPerc / 100) + UBCustRec(1).LocMeters(1).CurRead))
        HighRead# = FIX((0# + UBCustRec(1).LocMeters(1).AvgUse) * (HighReadPerc / 100))
        IF HighRead# <= 1 THEN
          HighRead# = 1
        END IF

        HighRead$ = LTRIM$(STR$(HighRead#))
        IF LEN(HighRead$) < 10 THEN
          HighRead$ = STRING$(10 - LEN(HighRead$), "0") + HighRead$
        END IF

        LowRead# = 1 'UBCustRec(1).LocMeters(1).CurRead
        LowRead$ = LTRIM$(STR$(LowRead#))
        IF LEN(LowRead$) < 10 THEN
          LowRead$ = STRING$(10 - LEN(LowRead$), "0") + LowRead$
        END IF
        
        CUS(1).RecordID = "CUS"
        CUS(1).RouteNumber = Cycle$ + Route$
        CUS(1).Meters = "001"
        CUS(1).AccountNumber = LTRIM$(STR$(RecNumber!))
        CUS(1).Name = LTRIM$(LEFT$(UBCustRec(1).CUSTNAME, 20))
        CUS(1).Address1 = LTRIM$(LEFT$(UBCustRec(1).SERVADDR, 20))
        CUS(1).Address2 = STRING$(20, " ")
        CUS(1).Reserved = "  "
        CUS(1).Group = " "
        CUS(1).CustInfo = LTRIM$(LEFT$(UBCustRec(1).HHMSG1, 20))
        CUS(1).ExCustRec = "N"
        CUS(1).Segment = "    "
        CUS(1).UtilityID = "WS"
        CUS(1).PassThur = " "
        CUS(1).Pad = " "

        PUT CUSFile, TCnt!, CUS(1)
        
        MTR(1).RecordID = "MTR"
        MTR(1).RouteNumber = Cycle$ + Route$
        MTR(1).Reads = "001"
        MTR(1).Reserved = "00"
        MTR(1).Group = "0"
        MTR(1).Reserved1 = "0"
        MTR(1).Reserved2 = "0000000"
        MTR(1).Survey = "0"
        MTR(1).Reserved3 = "00"
        MTR(1).Survey2 = "0"
        IF VAL(UBCustRec(1).CUSTTYPE) = 1 THEN
          MTR(1).BillCode = "R"
        ELSE
          MTR(1).BillCode = "C"
        END IF
        MTR(1).MtrStatus = "A"
        MTR(1).OpticalProbe = STRING$(14, " ")
        MTR(1).MtrNumber = MNum2File$
        MTR(1).Reserved4 = "00"
        MTR(1).MeterType = "01" ' 1=Water 2=Elec 3=Gas
        MTR(1).MeterSeq = MeterSeq$
        MTR(1).MeterInfo = QPTrim$(UBCustRec(1).USERCODE1)
        MTR(1).Reserved5 = "0"
        MTR(1).Location = "00"
        MTR(1).Reserved6 = "0"
        MTR(1).ReadInst1 = "0"
        MTR(1).Reserved7 = "0"
        MTR(1).ReadInst2 = "00"
        MTR(1).SpecMessage = "0"
        MTR(1).Reserved8 = "0"
        MTR(1).SpecMessageYN = "N"
        MTR(1).MtrCat = "W"
        MTR(1).LocExtraMtr = "L"
        MTR(1).TimeCode = "000"
        MTR(1).MtrAudit1 = "  "
        MTR(1).MtrAudit2 = "  "
        MTR(1).MtrAudit3 = "  "
        MTR(1).MtrAudit4 = "  "
        MTR(1).Pad = STRING$(14, " ")

        PUT MTRFile, TCnt!, MTR(1)

        Rdg(1).RecordID = "RDG"
        Rdg(1).RouteNumber = Cycle$ + Route$
        Rdg(1).Text = "WATR"
        Rdg(1).Prompt = "Y"
        Rdg(1).ReadDir = "L"
        Rdg(1).Compare = "000"
        Rdg(1).Validation = "000"
        Rdg(1).Reserved = "0"
        Rdg(1).Channel = "00"
        Rdg(1).Dials = Dials$   ' STRING * 2
        Rdg(1).Decimals = Decimals$             ' STRING * 2
        Rdg(1).ReadMethod = ReadMethod$
        Rdg(1).PrevRead = PrevRead$
        Rdg(1).High1 = HighRead$                '    AS STRING * 10
        Rdg(1).Low1 = LowRead$  '    AS STRING * 10
        Rdg(1).MtrConstant = "000000"
        Rdg(1).ConstantFlag = "0"
        Rdg(1).HHFFlag = "0"
        Rdg(1).PosCreep = "00001"
        Rdg(1).Estimates = "0"
        Rdg(1).Reserved2 = "0"

        Rdg(1).ReadType = "01"

        'vvv Caldwell
        'Rdg(1).ReadType = Dials$

        'change to user2 field
        Rdg(1).MaxPercent = "000000"
        Rdg(1).MinPercent = "000000"
        Rdg(1).NegCreep = "00001"
        Rdg(1).Pad = STRING$(31, " ")

        PUT RDGFile, TCnt!, Rdg(1)
        
        RFF(1).RecordID = "RFF"
        RFF(1).RouteNumber = Cycle$ + Route$
        RFF(1).RFERT = ERTID$
        RFF(1).Reserved = "000000"
        RFF(1).RFProgram = "ERT "
        RFF(1).Reserved2 = "0000000"
        RFF(1).GeoArea = "00"
        RFF(1).RFFreq = "000956.25625" ' Hard Coded from Badger
        RFF(1).RFTone = "0012"         ' Hard Coded from Badger  'calddwell
        RFF(1).Reserved3 = "0000000000"
        RFF(1).Tamper = "17"
        RFF(1).ConcIndicator = "N"
        RFF(1).Pad = STRING$(59, " ")

        PUT RFFFile, TCnt!, RFF(1)

        WRR(1).RecordID = "WRR"
        WRR(1).RouteNumber = Cycle$ + Route$
        WRR(1).DeviceID = ERTID$
        LSET WRR(1).WandProg = "MTR2"
        LSET WRR(1).Resv1 = ""
        LSET WRR(1).Resv1 = ""
        WRR(1).NodeNumb = "00"
        WRR(1).NoTamper = "Y"
        LSET WRR(1).Fill1 = ""
        PUT WRRFile, TCnt!, WRR(1)
      END IF
    END IF
    Cnt! = Cnt! + 1
  LOOP UNTIL Cnt! > IdxNumOfRecs
  CLOSE
RETURN

BuildDownLoad:
  CUSFile = FREEFILE
  OPEN "UBHHCUS.DAT" FOR RANDOM SHARED AS CUSFile LEN = LEN(CUS(1))
  MTRFile = FREEFILE
  OPEN "UBHHMTR.DAT" FOR RANDOM SHARED AS MTRFile LEN = LEN(MTR(1))
  RDGFile = FREEFILE
  OPEN "UBHHRDG.DAT" FOR RANDOM SHARED AS RDGFile LEN = LEN(Rdg(1))
  RFFFile = FREEFILE
  OPEN "UBHHRFF.DAT" FOR RANDOM SHARED AS RFFFile LEN = LEN(RFF(1))
  
  'File Header
  FHD(1).RecordID = "FHD"
  FHD(1).TABLE = "N"
  
'  FHD(1).PROBE = "N"

'broadway
  FHD(1).PROBE = "Y"

  FHD(1).VERS = "01.00"
  FHD(1).Reserved = "   "
  FHD(1).CYCLES = "01"
  FHD(1).RADIOREADYN = "Y"
  

'  FHD(1).WANDYN = "N"

'broadway
  FHD(1).WANDYN = "Y"
  
  FHD(1).EXTFORMATYN = "N"
  FHD(1).Pad = STRING$(108, 32)
  FileHdrLine$ = FHD(1).RecordID + FHD(1).TABLE + FHD(1).PROBE + FHD(1).VERS + FHD(1).Reserved
  FileHdrLine$ = FileHdrLine$ + FHD(1).CYCLES + FHD(1).RADIOREADYN + FHD(1).WANDYN + FHD(1).EXTFORMATYN + FHD(1).Pad
  
  'Cycle Header
  CHD(1).RecordID = "CHD"
  CHD(1).CycleNumber = Cycle$
  CHD(1).NumberCycles = "0001"
  CHD(1).CycleDate = LEFT$(DATE$, 2) + MID$(DATE$, 4, 2) + MID$(DATE$, 7, 4)
  CHD(1).Pad = STRING$(109, 32)
  CycleLine$ = CHD(1).RecordID + CHD(1).CycleNumber + CHD(1).NumberCycles + CHD(1).CycleDate + CHD(1).Pad
  
  'Route Header
  RHD(1).RecordID = "RHD"
  RHD(1).RouteNumber = Cycle$ + Route$
  RHD(1).SurveyYN = "N"
  RHD(1).RouteMessYN = "N"
  RHD(1).Keys = "0000"
  RHD(1).Readings = Reading$
  RHD(1).Demands = "0000"
  RHD(1).Keyed = KReading$
  RHD(1).Probed = "0000"
  RHD(1).Radio = RReading$
  RHD(1).Customers = TReading$
  RHD(1).Meters = TReading$
  RHD(1).TimeAllowed = "000000"
  RHD(1).Gas = "0000"
  RHD(1).Water = TReading$
  RHD(1).Electric = "0000"
  RHD(1).Location = "0000"
  RHD(1).Extra = "0000"
  RHD(1).Region = "  "
  RHD(1).Zone = "  "
  RHD(1).Office = "  "
  
  'was this
  ' RHD(1).BillCycle = Cycle$
  
  RHD(1).BillCycle = ""
  
  RHD(1).DropCycle = "  "
  RHD(1).WandReads = "0000"
  RHD(1).AMR = " "
  RHD(1).Pad = STRING$(32, " ")
  RouteHeader$ = RHD(1).RecordID + RHD(1).RouteNumber + RHD(1).SurveyYN
  RouteHeader$ = RouteHeader$ + RHD(1).RouteMessYN + RHD(1).Keys + RHD(1).Readings
  RouteHeader$ = RouteHeader$ + RHD(1).Demands + RHD(1).Keyed + RHD(1).Probed
  RouteHeader$ = RouteHeader$ + RHD(1).Radio + RHD(1).Customers + RHD(1).Meters
  RouteHeader$ = RouteHeader$ + RHD(1).TimeAllowed + RHD(1).Gas + RHD(1).Water
  RouteHeader$ = RouteHeader$ + RHD(1).Electric + RHD(1).Location
  RouteHeader$ = RouteHeader$ + RHD(1).Extra + RHD(1).Region + RHD(1).Zone
  RouteHeader$ = RouteHeader$ + RHD(1).Office + RHD(1).BillCycle
  RouteHeader$ = RouteHeader$ + RHD(1).DropCycle + RHD(1).WandReads
  RouteHeader$ = RouteHeader$ + RHD(1).AMR + RHD(1).Pad
  
  'Start Build of Test File For Now

  OutFile = FREEFILE
  OPEN Path$ + "DOWNLOAD.DAT" FOR OUTPUT AS #OutFile
  PRINT #OutFile, FileHdrLine$
  PRINT #OutFile, CycleLine$
  PRINT #OutFile, RouteHeader$
  
  'Now Do Individual Customers, Readings , and Radio Read if necessary
  
  CUSFile = FREEFILE
  OPEN "UBHHCUS.DAT" FOR RANDOM SHARED AS CUSFile LEN = LEN(CUS(1))
  MTRFile = FREEFILE
  OPEN "UBHHMTR.DAT" FOR RANDOM SHARED AS MTRFile LEN = LEN(MTR(1))
  RDGFile = FREEFILE
  OPEN "UBHHRDG.DAT" FOR RANDOM SHARED AS RDGFile LEN = LEN(Rdg(1))
  RFFFile = FREEFILE
  OPEN "UBHHRFF.DAT" FOR RANDOM SHARED AS RFFFile LEN = LEN(RFF(1))
  WRRFile = FREEFILE
  OPEN "UBHHWRR.DAT" FOR RANDOM SHARED AS WRRFile LEN = LEN(WRR(1))
  
  FOR CusRecord = 1 TO TCnt!
    GET CUSFile, CusRecord, CUS(1)
    CusLine$ = CUS(1).RecordID + CUS(1).RouteNumber + CUS(1).Meters
    CusLine$ = CusLine$ + CUS(1).AccountNumber + CUS(1).Name
    CusLine$ = CusLine$ + CUS(1).Address1 + CUS(1).Address2 + CUS(1).Reserved
    CusLine$ = CusLine$ + CUS(1).Group + CUS(1).CustInfo + CUS(1).ExCustRec
    CusLine$ = CusLine$ + CUS(1).Segment + CUS(1).UtilityID + CUS(1).PassThur + CUS(1).Pad
    PRINT #OutFile, CusLine$
    GET MTRFile, CusRecord, MTR(1)
    MtrLine$ = MTR(1).RecordID + MTR(1).RouteNumber + MTR(1).Reads + MTR(1).Reserved
    MtrLine$ = MtrLine$ + MTR(1).Group + MTR(1).Reserved1 + MTR(1).Reserved2
    MtrLine$ = MtrLine$ + MTR(1).Survey + MTR(1).Reserved3 + MTR(1).Survey2
    MtrLine$ = MtrLine$ + MTR(1).BillCode + MTR(1).MtrStatus + MTR(1).OpticalProbe
    MtrLine$ = MtrLine$ + MTR(1).MtrNumber + MTR(1).Reserved4 + MTR(1).MeterType
    MtrLine$ = MtrLine$ + MTR(1).MeterSeq + MTR(1).MeterInfo + MTR(1).Reserved5
    MtrLine$ = MtrLine$ + MTR(1).Location + MTR(1).Reserved6 + MTR(1).ReadInst1
    MtrLine$ = MtrLine$ + MTR(1).Reserved7 + MTR(1).ReadInst2 + MTR(1).SpecMessage + MTR(1).Reserved8
    MtrLine$ = MtrLine$ + MTR(1).SpecMessageYN + MTR(1).MtrCat + MTR(1).LocExtraMtr + MTR(1).TimeCode
    MtrLine$ = MtrLine$ + MTR(1).MtrAudit1 + MTR(1).MtrAudit2 + MTR(1).MtrAudit3
    MtrLine$ = MtrLine$ + MTR(1).MtrAudit4 + MTR(1).Pad
    PRINT #OutFile, MtrLine$
    GET RDGFile, CusRecord, Rdg(1)
    RdgLine$ = Rdg(1).RecordID + Rdg(1).RouteNumber + Rdg(1).Text + Rdg(1).Prompt
    RdgLine$ = RdgLine$ + Rdg(1).ReadDir + Rdg(1).Compare + Rdg(1).Validation + Rdg(1).Reserved
    RdgLine$ = RdgLine$ + Rdg(1).Channel + Rdg(1).Dials + Rdg(1).Decimals
    RdgLine$ = RdgLine$ + Rdg(1).ReadMethod + Rdg(1).PrevRead + Rdg(1).High1
    RdgLine$ = RdgLine$ + Rdg(1).Low1 + Rdg(1).MtrConstant + Rdg(1).ConstantFlag
    RdgLine$ = RdgLine$ + Rdg(1).HHFFlag + Rdg(1).PosCreep + Rdg(1).Estimates
    RdgLine$ = RdgLine$ + Rdg(1).Reserved2 + Rdg(1).ReadType + Rdg(1).MaxPercent
    RdgLine$ = RdgLine$ + Rdg(1).MinPercent + Rdg(1).NegCreep + Rdg(1).Pad
    PRINT #OutFile, RdgLine$
    IF Rdg(1).ReadMethod = "R" THEN
      GET RFFFile, CusRecord, RFF(1)
      RffLine$ = RFF(1).RecordID + RFF(1).RouteNumber + RFF(1).RFERT
      RffLine$ = RffLine$ + RFF(1).Reserved + RFF(1).RFProgram + RFF(1).Reserved2
      RffLine$ = RffLine$ + RFF(1).GeoArea + RFF(1).RFFreq + RFF(1).RFTone
      RffLine$ = RffLine$ + RFF(1).Reserved3 + RFF(1).Tamper + RFF(1).ConcIndicator
      RffLine$ = RffLine$ + RFF(1).Pad
      PRINT #OutFile, RffLine$
    END IF
    IF Rdg(1).ReadMethod = "W" THEN
      GET WRRFile, CusRecord, WRR(1)
      WRRLine$ = WRR(1).RecordID + WRR(1).RouteNumber + WRR(1).DeviceID
      WRRLine$ = WRRLine$ + WRR(1).WandProg + WRR(1).Resv1 + WRR(1).Resv2
      WRRLine$ = WRRLine$ + WRR(1).NodeNumb + WRR(1).NoTamper + WRR(1).Fill1
      PRINT #OutFile, WRRLine$
    END IF

  NEXT CusRecord
  
  
  'File Trailor
  FHD(1).RecordID = "FTR"
  FHD(1).TABLE = "N"
  FHD(1).PROBE = "N"
  FHD(1).VERS = "01.00"
  FHD(1).Reserved = "   "
  FHD(1).CYCLES = "01"
  FHD(1).RADIOREADYN = "Y"
  FHD(1).WANDYN = "Y"
  FHD(1).EXTFORMATYN = "N"
  FHD(1).Pad = STRING$(108, 32)
  FileHdrLine$ = FHD(1).RecordID + FHD(1).TABLE + FHD(1).PROBE + FHD(1).VERS + FHD(1).Reserved
  FileHdrLine$ = FileHdrLine$ + FHD(1).CYCLES + FHD(1).RADIOREADYN + FHD(1).WANDYN + FHD(1).EXTFORMATYN + FHD(1).Pad
  
  'Cycle Trailor
  CHD(1).RecordID = "CTR"
  CHD(1).CycleNumber = Cycle$
  CHD(1).NumberCycles = "0001"
  CHD(1).CycleDate = LEFT$(DATE$, 2) + MID$(DATE$, 4, 2) + MID$(DATE$, 7, 4)
  CHD(1).Pad = STRING$(109, 32)
  CycleLine$ = CHD(1).RecordID + CHD(1).CycleNumber + CHD(1).NumberCycles + CHD(1).CycleDate + CHD(1).Pad
  
  'Route Trailor
  RHD(1).RecordID = "RTR"
  RHD(1).RouteNumber = Cycle$ + Route$
  RHD(1).SurveyYN = "N"
  RHD(1).RouteMessYN = "N"
  RHD(1).Keys = "0000"
  RHD(1).Readings = Reading$
  RHD(1).Demands = "0000"
  RHD(1).Keyed = KReading$
  RHD(1).Probed = "0000"
  RHD(1).Radio = RReading$
  RHD(1).Customers = TReading$
  RHD(1).Meters = TReading$
  RHD(1).TimeAllowed = "000000"
  RHD(1).Gas = "0000"
  RHD(1).Water = TReading$
  RHD(1).Electric = "0000"
  RHD(1).Location = "0000"
  RHD(1).Extra = "0000"
  RHD(1).Region = "  "
  RHD(1).Zone = "  "
  RHD(1).Office = "  "
  RHD(1).BillCycle = Cycle$
  RHD(1).DropCycle = "  "
  RHD(1).WandReads = "0000"
  RHD(1).AMR = " "
  RHD(1).Pad = STRING$(32, " ")
  RouteHeader$ = RHD(1).RecordID + RHD(1).RouteNumber + RHD(1).SurveyYN
  RouteHeader$ = RouteHeader$ + RHD(1).RouteMessYN + RHD(1).Keys + RHD(1).Readings
  RouteHeader$ = RouteHeader$ + RHD(1).Demands + RHD(1).Keyed + RHD(1).Probed
  RouteHeader$ = RouteHeader$ + RHD(1).Radio + RHD(1).Customers + RHD(1).Meters
  RouteHeader$ = RouteHeader$ + RHD(1).TimeAllowed + RHD(1).Gas + RHD(1).Water
  RouteHeader$ = RouteHeader$ + RHD(1).Electric + RHD(1).Location
  RouteHeader$ = RouteHeader$ + RHD(1).Extra + RHD(1).Region + RHD(1).Zone
  RouteHeader$ = RouteHeader$ + RHD(1).Office + RHD(1).BillCycle
  RouteHeader$ = RouteHeader$ + RHD(1).DropCycle + RHD(1).WandReads
  RouteHeader$ = RouteHeader$ + RHD(1).AMR + RHD(1).Pad
  PRINT #OutFile, RouteHeader$       'Actually Trailors
  PRINT #OutFile, CycleLine$
  PRINT #OutFile, FileHdrLine$
  CLOSE
RETURN


CheckInfo:

  InfoOK = True

  NewPath$ = QPTrim$(Form$(4, 0))
  Path$ = GetBadgerPath$

  IF RIGHT$(NewPath$, 1) <> "\" THEN
    NewPath$ = NewPath$ + "\"
  END IF

'if this is a chage to the path to badger then update ubdata
  IF NewPath$ <> Path$ THEN
    PathFile = FREEFILE
    OPEN PathData$ FOR OUTPUT AS PathFile
    PRINT #PathFile, NewPath$
    CLOSE PathFile
  END IF

  Path$ = GetBadgerPath$
  IF INSTR(Path$, "BAD") > 0 THEN
    DisplaySUScrn "Invalid"
    InfoOK = False
    GOTO NoGoodRet
  END IF
  
  IF NOT (VAL(LEFT$(Form$(1, 0), 2))) > 0 AND VAL(LEFT$(Form$(2, 0), 2)) > 0 THEN
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
    InfoOK = False
  END IF
  
NoGoodRet:
RETURN
  
END SUB

