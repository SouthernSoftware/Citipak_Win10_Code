DEFINT A-Z
DECLARE SUB MakeMOWSequenceIndex (Text$)
DECLARE SUB InactiveConsReport ()
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
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
'declare function

TYPE UBSequenceIndexTypeII
  SeqNumber    AS LONG
  RecNum       AS INTEGER
  Fill         AS STRING * 2
END TYPE

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubdraft.BI'
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'Newcust.bi'
  '$INCLUDE: 'ubgilsen.bi'
  '$INCLUDE: 'ubbadger.bi'
  


  CONST False = 0, True = NOT False
  
  REDIM MChoice$(1 TO 10)
  
  MChoice$(1) = "Manual Meter Reading Entry "
  MChoice$(2) = "Hand-Held Meter Reading Entry"
  MChoice$(3) = "Estimated Meter Reading Entry"
  MChoice$(4) = "Print Meter Reading List"
  MChoice$(5) = "Print Reading Report"
  MChoice$(6) = "Print High/Low Report"
  MChoice$(7) = "Inactive Consumption Report"
  MChoice$(8) = "Print Meter Reading Notes"
  MChoice$(9) = "Print Meter Reading Sheets"
  MChoice$(10) = "Exit to OS"
  
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
  Row = ((25 - (UBOUND(MChoice$))) \ 2)
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
      'ManualMeterReading             'unrem
    CASE 2
      HandHeldProcessing        'unrem
    CASE 3
      EstMeterReading           'unrem
    CASE 4
      RUN "ubmtrlst"
      'PrintMeterList            'unrem
    CASE 5
      PrintMeterReport          'unrem
    CASE 6
      'PrintHighLowReport             'unrem
    CASE 7
      InactiveConsReport        'unrem
    CASE 8
      PrintMeterNotes           'unrem
    CASE 9
      PrintMeterSheets          'unrem
    CASE 10
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP
  RUN "ubbillin"

SUB DisplayHHScrn (ScrnName$)
  LibFile2Scrn "UBSETUP", ScrnName$, MonoCode%, Attribute%, ErrorCode%
END SUB

SUB MakeMOWSequenceIndex (Text$)

  ShowProcessingScrn "Creating " + IndexText$ + " Index"
  QPrintRC "    Reading Location Records     ", 11, 25, -1

  REDIM UBCustRec(1) AS NewUBCustRecType
  CustRecLen = LEN(UBCustRec(1))

  CHandle = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CHandle LEN = CustRecLen

  NumCustRecs& = LOF(CHandle) / CustRecLen

  REDIM SequenceIndex(1 TO NumCustRecs&)  AS UBSequenceIndexTypeII
  IndexRecLen = LEN(SequenceIndex(1))

  FOR Cnt = 1 TO NumCustRecs&
    GET CHandle, Cnt, UBCustRec(1)
    SequenceIndex(Cnt).SeqNumber = UBCustRec(1).SEQ
    SequenceIndex(Cnt).RecNum = Cnt
    ShowPctComp Cnt, NumCustRecs&               'show user percentage complete
  NEXT

  CLOSE CHandle

  QPrintRC "         Sorting Index.        ", 11, 25, -1

  SortT SequenceIndex(1), CINT(NumCustRecs&), 0, 8, 0, -2

  QPrintRC "      Writing Index Records      ", 11, 25, -1

  FCreate TempIndexName

  IHandle = FREEFILE
  OPEN TempIndexName FOR RANDOM SHARED AS IHandle LEN = 4

  FOR Cnt = 1 TO NumCustRecs&
    Prec& = SequenceIndex(Cnt).RecNum
    PUT IHandle, Cnt, Prec&
    ShowPctComp Cnt, NumCustRecs&               'show user percentage complete
  NEXT
  CLOSE IHandle

  ERASE UBCustRec, SequenceIndex


END SUB

SUB PrintHelp (Help$)
  
  t$ = SPACE$(40)
  LSET t$ = Help$
  QPrintRC t$, 25, 2, -1
  t$ = ""
  
END SUB

SUB ProcessBadger
  
  SHARED Choice$()
  
  REDIM TempScrn(0)

  REDIM Route(100)
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  CrLf$ = CHR$(13) + CHR$(10)
  Z9$ = "000000000"

  BadgerFile$ = "BADGERMR.DAT"

  ' Check For Device Type and Run Appropriate Program
  REDIM UBSetupRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetupRec(), UBSetUpLen
  
  HighVar# = UBSetupRec(1).HighRead
  IF HighVar# > 0 THEN
    HighVar# = HighVar# / 100
  ELSE
    HighVar# = 1.5
  END IF
  
  LowVar# = UBSetupRec(1).LowRead
  IF LowVar# > 0 THEN
    LowVar# = LowVar# / 100
  ELSE
    LowVar# = .75
  END IF
  
  REDIM UBBadgerRec(1) AS UBSensusRecType
  UBBadgerRecLen = LEN(UBBadgerRec(1))
  BadgerCFGFile = FREEFILE
  OPEN "UBBADGER.DAT" FOR RANDOM SHARED AS BadgerCFGFile LEN = UBBadgerRecLen
  NumBadgerRec = LOF(BadgerCFGFile) / UBBadgerRecLen
  IF NumBadgerRec = 1 THEN
    GET BadgerCFGFile, 1, UBBadgerRec(1)
    PathWay$ = UBBadgerRec(1).PathWay
  END IF

  LibName$ = "UBSETUP"
  ScrnName$ = "UBBADGE1"
  CursorOff

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
  NEXT
  
  REDIM Choice$(3, 0)
  
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Send Info to Badger"
  Choice$(2, 0) = "Get Info From Badger"
  
  IF NumBadgerRec = 1 THEN
    Form$(2, 0) = PathWay$
  END IF
  Action = 1
  
  BlockClear
  
  DisplayHHScrn ScrnName$
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      PathWay$ = Form$(2, 0)
      UBBadgerRec(1).PathWay = PathWay$
      PUT BadgerCFGFile, 1, UBBadgerRec(1)
      CLOSE BadgerCFGFile

      Operation$ = LEFT$(QPTrim$(Form$(1, 0)), 1)
      SELECT CASE Operation$
      CASE "S"
        GOSUB BadgerSend
        Done = True
      CASE "G"
        GOSUB BadgerGet
        Done = True
      CASE ELSE
        OK = MsgBox(LibName$, "NOOPERAT")
        Frm(1).FldNo = 1
        Action = 1
        Done = False
      END SELECT
    CASE ESC
      Done = True
      EXIT SUB
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done

BadgerReadExit:
  EXIT SUB
  
BadgerSend:
  
  LibName$ = "UBSETUP"
  ScrnName$ = "UBBADGE2"
  
  CursorOff
  
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
  NEXT
  
  Action = 1
  NumberOfRoutes = 0
  
  DisplayHHScrn ScrnName$
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF Frm(1).PrevFld = 12 AND Frm(1).FldNo > 1 THEN
      IF VAL(Form$(1, 0)) <> 0 THEN
        NumberOfRoutes = NumberOfRoutes + 1
        IF NumberOfRoutes <= 10 THEN
          Form$(NumberOfRoutes + 1, 0) = Form$(1, 0)
          Frm(1).FldNo = 1
          Form$(1, 0) = ""
          Action = 1
        END IF
      END IF
    END IF
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      IF NumberOfRoutes > 0 THEN  ' Set Route Number Values
        FOR Cnt = 1 TO NumberOfRoutes
          Route(Cnt) = VAL(Form$(Cnt + 1, 0))
        NEXT

        GOSUB MakeOUTTFFileName
        GOSUB BadgerOpenCust    'Open Customer Data File
        
        REDIM UBBadRdRec(1) AS BadgerRecType
        UBBadRdRecLen = LEN(UBBadRdRec(1))
        
        KillFile FileName$
        UBBadRdFile = FREEFILE
        OPEN FileName$ FOR RANDOM SHARED AS UBBadRdFile LEN = UBBadRdRecLen
        NumSenRdRecs = LOF(UBBadRdFile) / UBBadRdRecLen
        '  IdxName$ = "UBCUSTBK.IDX"
        'Open the Location Order for Reading

        MakeMOWSequenceIndex ""

        IdxName$ = "UBTEMP.IDX"
        IdxRecLen = 4           'we are using a integer
        IdxFileSize& = FileSize&(IdxName$)
        IdxNumRecs& = IdxFileSize& \ IdxRecLen
        
        'IF (IdxNumRecs& * 4) > FRE(-1) THEN
        '  IdxNumRecs& = IdxNumRecs& \ 2
        'ELSE
        '  MakeMOWSequenceIndex ""
          'MakeSequenceIndex ""
        'END IF

        ShowProcessingScrn "Processing Handheld File."

        REDIM IdxBuff(1 TO IdxNumRecs&) AS UBCustIndexRecType
        IdxFile = FREEFILE
        OPEN IdxName$ FOR RANDOM AS IdxFile LEN = 4
        FOR Cnt& = 1 TO IdxNumRecs&
          GET #IdxFile, Cnt&, IdxBuff(Cnt&)
        NEXT
        CLOSE IdxFile

        'LOCATE 12, 1: PRINT "DONE LOADING INDEX"
        'WaitForAction

        DoneCnt& = 1
        DO
          Prec& = IdxBuff(DoneCnt&).RecNum
          IF Prec& > 0 THEN
            GET UBFile, Prec&, UBCustRec(1)
            FOR BookCnt = 1 TO NumberOfRoutes
              IF VAL(UBCustRec(1).BOOK) = Route(BookCnt) AND (UBCustRec(1).Status <> "F") THEN
                GOSUB BadgerPutRec
              END IF
            NEXT
          END IF
          DoneCnt& = DoneCnt& + 1
          ShowPctComp DoneCnt&, IdxNumRecs&
        LOOP UNTIL DoneCnt& > IdxNumRecs&
        CLOSE
        Done = True
        BlockClear
        DisplayUBScrn "UPDATEOK"
        WaitForAction
      END IF
      Done = True
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  EXIT SUB
RETURN

BadgerPutRec:
  
  'modifyed
  'May Have Up to 7 Meters to Read
  MtrCnt = 1
  Account$ = STR$(Prec&)
  WHILE MtrCnt < 8
    IF (ASC(UBCustRec(1).LocMeters(MtrCnt).MTRType) > 32) THEN
      SELECT CASE UBCustRec(1).LocMeters(MtrCnt).MTRType
      CASE "C", "S", "W", "T", "E", "D"         'here dale
        IF UBCustRec(1).LocMeters(MtrCnt).MtrMulti <= 0 THEN
          UBCustRec(1).LocMeters(MtrCnt).MtrMulti = 1
        END IF
        Multi& = UBCustRec(1).LocMeters(MtrCnt).MtrMulti
        
        Average# = UBCustRec(1).LocMeters(MtrCnt).AvgUse
        CRead# = UBCustRec(1).LocMeters(MtrCnt).CurRead
        'make sure we have valid average & current readings
        IF CRead# < 0 THEN
          CRead# = 0
        END IF
        IF Average# <= 0 THEN
          Average# = CRead#
        END IF
        
        IHiRead$ = RIGHT$((Z9$ + QPTrim$(STR$(CRead# + FIX(Average# * HighVar#)))), 9)
        ILoRead$ = RIGHT$((Z9$ + QPTrim$(STR$(CRead# + FIX(Average# * LowVar#)))), 9)
        CurRead$ = RIGHT$((Z9$ + QPTrim$(STR$(CRead#))), 9)
        
        Acct$ = UBCustRec(1).BOOK + UBCustRec(1).SeqNumb
        IF UBCustRec(1).SEQ < 0 THEN
          UBCustRec(1).SEQ = 0
        END IF
        
        MtCnt$ = QPTrim$(STR$(MtrCnt))
        SeqNumb$ = QPTrim$(STR$(UBCustRec(1).SEQ))
        OurInfo$ = QPTrim$(STR$(Prec&)) + "-" + MtCnt$

        MtNumb$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MTRNUM)
        'IF LEN(MtNumb$) > 7 THEN
        '  MtNumb$ = LEFT$(MtNumb$, 7)
        'END IF

        SELECT CASE Multi&
        CASE 1
          CircleC$ = "1"
        CASE 10
          CircleC$ = "2"
        CASE 100
          CircleC$ = "3"
        CASE 1000
          CircleC$ = "4"
        END SELECT
        
        LSET UBBadRdRec(1).Fill1 = ""
        LSET UBBadRdRec(1).CustName = QPTrim$(UBCustRec(1).CustName)
        LSET UBBadRdRec(1).ServAddr = QPTrim$(UBCustRec(1).ServAddr)
        LSET UBBadRdRec(1).MtrNum1 = MtNumb$
        LSET UBBadRdRec(1).Multi = ""           'Multi&
        LSET UBBadRdRec(1).Status = UBCustRec(1).Status
        LSET UBBadRdRec(1).ReadCode = MtCnt$
        LSET UBBadRdRec(1).ServFreq = MtCnt$
        LSET UBBadRdRec(1).DNI = ""
        LSET UBBadRdRec(1).MtrNum2 = UBBadRdRec(1).MtrNum1
        LSET UBBadRdRec(1).NumDials = ""
        RSET UBBadRdRec(1).HiRead = IHiRead$
        RSET UBBadRdRec(1).LoRead = ILoRead$
        LSET UBBadRdRec(1).CurrRead = ""        'changed per SCS 'CurRead$
        LSET UBBadRdRec(1).ReadTime = ""
        LSET UBBadRdRec(1).ReadCode2 = ""
        LSET UBBadRdRec(1).CmntCode = ""
        LSET UBBadRdRec(1).Fill2 = ""
        LSET UBBadRdRec(1).Account = Acct$
        LSET UBBadRdRec(1).ReadDate = ""
        LSET UBBadRdRec(1).DevCode = "P"
        LSET UBBadRdRec(1).MMILat = ""
        LSET UBBadRdRec(1).MMILong = ""
        LSET UBBadRdRec(1).MMIChanl = ""
        LSET UBBadRdRec(1).CircleCode = CircleC$                'changed per SCS
        'this is a code from above
        RSET UBBadRdRec(1).SeqNumb = SeqNumb$
        LSET UBBadRdRec(1).MfgModel = ""
        LSET UBBadRdRec(1).UserField = OurInfo$
        LSET UBBadRdRec(1).ReadID = ""
        LSET UBBadRdRec(1).ReadCo1 = ""
        LSET UBBadRdRec(1).ReadCo2 = ""
        LSET UBBadRdRec(1).ReadCo3 = ""
        LSET UBBadRdRec(1).MMIReadCode = ""
        LSET UBBadRdRec(1).Pad = ""
        LSET UBBadRdRec(1).CrLf = CrLf$
        PUT UBBadRdFile, , UBBadRdRec(1)
      END SELECT
    END IF
ESkipEm:
    MtrCnt = MtrCnt + 1
  WEND
  
RETURN
  
MakeOUTTFFileName:
  PathWay$ = QPTrim$(PathWay$)
  IF LEN(PathWay$) > 0 THEN
    IF RIGHT$(PathWay$, 1) <> "\" THEN
      PathWay$ = PathWay$ + "\"
    END IF
  END IF
  FileName$ = PathWay$ + "CONNECT.IN3"
RETURN

MakeINTFFileName:
  PathWay$ = QPTrim$(PathWay$)
  IF LEN(PathWay$) > 0 THEN
    IF RIGHT$(PathWay$, 1) <> "\" THEN
      PathWay$ = PathWay$ + "\"
    END IF
  END IF
  FileName$ = PathWay$ + "CONNECT.OT3"
RETURN

BadgerGet:
  
  LibName$ = "UBSETUP"
  ScrnName$ = "UBBADGE3"
  
  'Help$ = "Sensus Hand Held Processing"
  CursorOff
  
  'Help$ = "Upload HandHeld Operation"
  DisplayHHScrn ScrnName$
  
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
  NEXT
  
  Action = 1
  NumberOfRoutes = 0
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      OverDate = Date2Num(Form$(1, 0))
      IF OverDate > 0 THEN
        UseOverDate = True
      END IF
      GOSUB MakeINTFFileName      'Get Badger File
      GOSUB BadgerOpenCust      'Open Customer Data File

      REDIM UBBadRdRec(1) AS BadgerRecType
      UBBadRdRecLen = LEN(UBBadRdRec(1))

      'Open meter reading information File
      UBBadRdFile = FREEFILE
      OPEN FileName$ FOR RANDOM SHARED AS UBBadRdFile LEN = UBBadRdRecLen
      NumBadRdRecs& = LOF(UBBadRdFile) / UBBadRdRecLen
      
      IF NumBadRdRecs& = 0 THEN
        CLOSE
        BlockClear
        DisplayHHScrn "NOBGFile"
        QPrintRC FileName$, 10, 27, -1
        WaitForAction
        GOTO BadgerReadExit
      END IF
      
      BlockClear
      ShowProcessingScrn "Processing Meter Readings."
      
      DoneCnt& = 1              ' Initialize File Counter to 1
      DO
        GET UBBadRdFile, DoneCnt&, UBBadRdRec(1)
        Prec& = VAL(QPTrim$(UBBadRdRec(1).UserField))
        IF Prec& > 0 THEN
          GET UBFile, Prec&, UBCustRec(1)
          GOSUB EExtractRecord
        ELSE
          STOP
        END IF
        DoneCnt& = DoneCnt& + 1
        ShowPctComp DoneCnt&, NumBadRdRecs&
      LOOP UNTIL DoneCnt& > NumBadRdRecs&
      CLOSE
      Done = True
      SLEEP 2
      BlockClear
      DisplayUBScrn "UPDATEOK"
      WaitForAction
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  
RETURN
  
EExtractRecord:
'this extracts the reading & date

  UpdCnt& = UpdCnt& + 1
  QPrintRC " Updated Count:" + STR$(UpdCnt&), 11, 28, -1
  
  DashPos = INSTR(UBBadRdRec(1).UserField, "-")
  MT$ = MID$(UBBadRdRec(1).UserField, DashPos + 1)
  MeterRecord = VAL(MT$)
  
  IF MeterRecord = 0 THEN MeterRecord = 1
  ' Check Meter Updated Flag
  ' Update Meter W/Reading
  CurReading# = VAL(UBBadRdRec(1).CurrRead)

  IF UBCustRec(1).LocMeters(MeterRecord).MtrMulti = 10 THEN
    CurReading# = (CurReading# * .1)
  END IF

  IF UseOverDate THEN
    ReadingDate = OverDate          'if they want to overide the read date
  ELSE
    MeterReadDate$ = LEFT$(UBBadRdRec(1).ReadDate, 2) + "/" + MID$(UBBadRdRec(1).ReadDate, 3, 2) + "/" + RIGHT$(UBBadRdRec(1).ReadDate, 4)
    ReadingDate = Date2Num(MeterReadDate$)
    IF ReadingDate <= 0 THEN
      ReadingDate = Date2Num(DATE$) 'if the read date was bad then fix it
    END IF
  END IF
  
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
  'now update the customers record...
  PUT UBFile, Prec&, UBCustRec(1)
RETURN
  
  
BadgerOpenCust:
'  REDIM UBCustRec(1) AS NewUBCustRecType
'  UBCustRecLen = LEN(UBCustRec(1))
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
RETURN
  
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
  NumberOfRoutes = 0
  
  DisplayHHScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF Frm(1).PrevFld = 13 AND Frm(1).FldNo > 2 THEN
      IF VAL(Form$(2, 0)) <> 0 THEN
        NumberOfRoutes = NumberOfRoutes + 1
        IF NumberOfRoutes <= 10 THEN
          Form$(NumberOfRoutes + 2, 0) = Form$(2, 0)
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
      IF NumberOfRoutes > 0 THEN
        QPrintRC "      Scanning Account:        ", 11, 21, -1
        ' Set Route Number Values
        FOR Cnt! = 1 TO NumberOfRoutes
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
        
        REDIM UBSetupRec(1) AS UBSetupRecType
        LoadUBSetUpFile UBSetupRec(), UBSetUpLen
        
        'Open the Correct Order for Reading
        IF UBSetupRec(1).UseSeq = "Y" THEN
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
          Prec& = IdxBuff(Cnt!).RecNum
          IF NOT (Prec&) = 0 THEN
            GET UBFile, Prec&, UBCustRec(1)
            IF INSTR(UBCustRec(1).HHMSG1, "NOREAD") > 0 THEN
              GOTO HWriteSkip
            END IF
            DisplayAcct$ = UBCustRec(1).BOOK + "-" + UBCustRec(1).SeqNumb
            QPrintRC DisplayAcct$, 11, 45, -1
            FOR BookCnt! = 1 TO NumberOfRoutes
              IF VAL(UBCustRec(1).BOOK) = Route(BookCnt!) AND (UBCustRec(1).Status <> "F") THEN
                GOSUB HuskyWriteRec
              END IF
            NEXT BookCnt!
          END IF
HWriteSkip:
          Cnt! = Cnt! + 1
        LOOP UNTIL Cnt! > IdxNumOfRecs
        CLOSE
        OK = MsgBox%("UBSETUP", "UBHUSK1C")
        
        IF OK = 1 THEN
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
  
  OK = MsgBox%("UBSETUP", "UBHUSK1C")
  IF OK = 1 THEN
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
    Prec& = VAL(QPTrim$(UBPC3000GetRDRec(1).Account))
    
    IF NOT (Prec&) = 0 THEN
      GET UBFile, Prec&, UBCustRec(1)
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
    UBCustRec(1).NewNotes = True
    UBCustRec(1).HHMSG1 = UBPC3000GetRDRec(1).Note1
    UBCustRec(1).HHMSG2 = UBPC3000GetRDRec(1).Note2
    UBCustRec(1).HHMSG3 = UBPC3000GetRDRec(1).Note3
  END IF
  '   ^^^
  PUT UBFile, Prec&, UBCustRec(1)
  RETURN
  
  
HuskyWriteRec:
  'May Have Up to 7 Meters to Read
  FOR MtrCnt! = 1 TO 7
    MeterOK = False
    Account$ = STR$(Prec&)
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
            MeterID$ = UBCustRec(1).BOOK + UBCustRec(1).SeqNumb
          END IF
          IF LEN(MeterID$) < 8 THEN
            MeterID$ = STRING$(8 - LEN(MeterID$), "0") + MeterID$
          END IF
          MeterID$ = LEFT$(MeterID$, 8)
          
          'Set Record Fields and Put On Disk
          UBPC3000RDRec(1).CustName = UBCustRec(1).CustName
          UBPC3000RDRec(1).ServAddress = LEFT$(UBCustRec(1).ServAddr, 16)
          
          C$ = QPTrim$(UBCustRec(1).USERCODE1)
          IF LEN(C$) > 0 THEN
            MID$(UBPC3000RDRec(1).ServAddress, 19, 1) = LEFT$(C$, 1)
          END IF
          C$ = QPTrim$(UBCustRec(1).USERCODE2)
          IF LEN(C$) > 0 THEN
            MID$(UBPC3000RDRec(1).ServAddress, 20, 1) = LEFT$(C$, 1)
          END IF
          
          
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
  NumberOfRoutes = 0
  
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    SELECT CASE Frm(1).KeyCode
    CASE F10Key 'First Check For Valid Interrogator Number and Routes
      GOSUB FILENAM1Itron       'Set Route Number Values
      FOR Cnt! = 1 TO NumberOfRoutes
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
        Prec& = IdxBuff(Cnt!).RecNum
        LOCATE 11, 42: COLOR 15: PRINT USING "######"; Prec&
        IF NOT (Prec&) = 0 THEN
          GET UBFile, Prec&, UBCustRec(1)
          IF VAL(UBCustRec(1).BOOK) <> 0 THEN
            'Take Care of First Record Sent First
            IF FirstOnRoute = 0 THEN
              FirstOnRoute = VAL(UBCustRec(1).BOOK)
              UBItronFRec(1).RecordCode = "F"
              UBItronFRec(1).Route = UBCustRec(1).BOOK + "000000"
              UBItronFRec(1).Message = ""
              UBItronFRec(1).Filler = ""
              UBItronFRec(1).CrLf = CHR$(13) + CHR$(10)
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
              UBItronZRec(1).CrLf = CHR$(13) + CHR$(10)
              PUT #UBItronRdFile, (LOF(UBItronRdFile) / 80) + 1, UBItronZRec(1)
              
              'Create Next Route Header F
              FirstOnRoute = VAL(UBCustRec(1).BOOK)
              NumberofAccounts = 0
              UBItronFRec(1).RecordCode = "F"
              UBItronFRec(1).Route = UBCustRec(1).BOOK + "000000"
              UBItronFRec(1).Message = ""
              UBItronFRec(1).Filler = ""
              UBItronFRec(1).CrLf = CHR$(13) + CHR$(10)
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
      UBItronZRec(1).CrLf = CHR$(13) + CHR$(10)
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
  Account$ = LTRIM$(STR$(Prec&))
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
  CustomerName$ = QPTrim$(UBCustRec(1).CustName)
  'Set Customer Address
  CustomerAddress$ = QPTrim$(UBCustRec(1).ServAddr)
  
  
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
      IF Average# < 0 THEN Average# = 0
      ILowRead$ = LTRIM$(STR$((UBCustRec(1).LocMeters(MtrCnt!).CurRead)))
      
      IF LEN(ILowRead$) > DialVal# THEN
        ILowRead$ = RIGHT$(ILowRead$, DialVal#)
      END IF
      
      IF LEN(ILowRead$) > 8 THEN ILowRead$ = RIGHT$(ILowRead$, 8)
      IF LEN(ILowRead$) < 8 THEN ILowRead$ = STRING$(8 - LEN(ILowRead$), "0") + ILowRead$
      
      HighRead# = FIX(Average# * (150 / 100)) + UBCustRec(1).LocMeters(MtrCnt!).CurRead
      IF Average# = 0 THEN
        HighRead# = VAL(ILowRead$) + 150        'Fudge Factor Until History Built
      END IF
      LowRead# = VAL(ILowRead$)
      
      'IF HighRead# < LowRead# THEN STOP
      
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
      UBItronARec(1).SeqNumb = SeqNumber$
      UBItronARec(1).Message = "0"
      UBItronARec(1).AType = CustomerType$
      UBItronARec(1).Filler = ""
      UBItronARec(1).CrLf = CHR$(13) + CHR$(10)
      PUT #UBItronRdFile, (LOF(UBItronRdFile) / 80) + 1, UBItronARec(1)
      'Create B Record
      UBItronBRec(1).RecordCode = "B"
      UBItronBRec(1).CustName = CustomerName$
      UBItronBRec(1).CustAddr = CustomerAddress$
      UBItronBRec(1).Filler = ""
      UBItronBRec(1).CrLf = CHR$(13) + CHR$(10)
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
      UBItronHRec(1).CrLf = CHR$(13) + CHR$(10)
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
  NumberOfRoutes = 0
  
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      GOSUB FileNamItron        ' Get Itron File
      
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
          Prec& = VAL(MID$(a$, 10, 10))         'Grabs Acct#
          'Process D Line
          LINE INPUT #UBLogGetRdFile, a$
          'Process E Line
          LINE INPUT #UBLogGetRdFile, a$
          CurReading# = VAL(MID$(a$, 3, 8))
          ReadTime$ = MID$(a$, 12, 6)
          'Update the Account Now
          IF NOT (Prec& = 0 OR CurReading# = 0) THEN
            GET UBFile, Prec&, UBCustRec(1)
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
  PUT UBFile, Prec&, UBCustRec(1)
  RETURN
  
FileNamItron:
  PathWay$ = QPTrim$(PathWay$)
  FileName$ = PathWay$ + "TOHOST.PPS"
  RETURN
  
END SUB

