DECLARE SUB ProcessSysCon ()
DEFINT A-Z
DECLARE SUB DisplayH2Scrn (ScrnName$)
DECLARE SUB ProcessSchlum ()
DECLARE SUB ESensus ()
DECLARE SUB KillFile (FileName$)
DECLARE SUB DisplayHHScrn (ScrnName$)
DECLARE SUB ProcessItron ()
DECLARE SUB ProcessLogicon ()
DECLARE SUB ProcessTelxon ()
DECLARE SUB CursorOff ()
DECLARE SUB DisplayHuskyScrn (ScrnName$)
DECLARE SUB ProcessHusky ()
DECLARE SUB PrintMeterNotes ()
DECLARE FUNCTION Chk4BookSeqNum (BOOK$, SeqNum$)
DECLARE SUB PrintMeterSheets ()
DECLARE SUB EstMeterReading ()
DECLARE SUB PrintMeterReport ()
DECLARE SUB PrintMeterList ()
DECLARE SUB HandHeldProcessing ()
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetUpLen%)
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
  
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'Newcust.bi'
  '$INCLUDE: 'ubgilsen.bi'
  '$INCLUDE: 'ubschlum.bi'
  '$INCLUDE: 'ubsyscon.bi'

  CONST False = 0, True = NOT False

FUNCTION Chk4BookSeqNum (BOOK$, SeqNum$)
  
  Chk4BookSeqNum = False        'assume not found
  
  TBookSeq& = QPValL(BOOK$ + SeqNum$)
  
  REDIM UBBookSeq(1) AS BookSeqRecType
  BookSeqLen = LEN(UBBookSeq(1))
  
  IF FileSize("UBOOKSEQ.DAT") > 0 THEN
    FOpenS "UBOOKSEQ.DAT", Handle               'open data file
    NumBookSeq = FLof(Handle) \ BookSeqLen
    REDIM UBBookSeq(1 TO NumBookSeq) AS BookSeqRecType
    FGetRTA Handle, UBBookSeq(1), 1&, NumBookSeq * BookSeqLen
    FClose Handle
    
    FOR Cnt = 1 TO NumBookSeq
      IF UBBookSeq(Cnt).BookSeq = TBookSeq& THEN
        Chk4BookSeqNum = True   'found this book-seq
        EXIT FOR
      END IF
    NEXT
  END IF
  
END FUNCTION

SUB ESensus
  SHARED Choice$()
  
  REDIM TempScrn(0)
  REDIM Route(100)
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  ' Check For Device Type and Run Appropriate Program
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen
  
  'If Sensus we need to know the location path of your RMMS or STARS program
  'For LongView This is D:\RMMS\
  REDIM UBSensusRec(1) AS UBSensusRecType
  UBSensusRecLen = LEN(UBSensusRec(1))
  UBSensusFile = FREEFILE
  OPEN "UBSENSUS.DAT" FOR RANDOM SHARED AS UBSensusFile LEN = UBSensusRecLen
  NumSensusRecs = LOF(UBSensusFile) / UBSensusRecLen
  IF NumSensusRecs = 1 THEN
    GET UBSensusFile, 1, UBSensusRec(1)
    PathWay$ = UBSensusRec(1).PathWay
  END IF
  
  LibName$ = "UB"
  ScrnName$ = "UBSSHH"
  
  CursorOff
  
  Help$ = "Choose ESensus Operation"
  
  DisplayUBScrn ScrnName$
  
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
  NEXT
  
  REDIM Choice$(3, 0)
  
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Send Info to ESensus"
  Choice$(2, 0) = "Get Info From ESensus"
  
  IF NumSensusRecs = 1 THEN
    Form$(2, 0) = PathWay$
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
        UBSensusRec(1).PathWay = PathWay$
        PUT UBSensusFile, 1, UBSensusRec(1)
        CLOSE UBSensusFile
        IF LEFT$(Form$(1, 0), 1) = "S" THEN GOSUB ESendSensus
        IF LEFT$(Form$(1, 0), 1) = "G" THEN GOSUB EGetSensus
        Done = True
      END IF
    CASE ESC
      Done = True
      EXIT SUB
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  
ESendSensus:
  LibName$ = "UB"
  ScrnName$ = "UBSSHH1S"
  CursorOff
  
  Help$ = "DownLoad HandHeld Operation"
  DisplayUBScrn ScrnName$
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
  NEXT
  
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
        Frm(1).FldNo = 1
        Action = 1
        Done = False
      ELSE
        IF NumberofRoutes > 0 THEN
          INTERROGATOR = VAL(Form$(1, 0))
          COLOR 11
          LOCATE 12, 18
          PRINT "LOADING THESE ROUTES !!!                      "
          GOSUB EFILENAM1
          ' Set Route Number Values
          FOR Cnt = 1 TO NumberofRoutes
            Route(Cnt) = VAL(Form$(Cnt + 2, 0))
          NEXT
          
          GOSUB EOpenCustFile   'Open Customer Data File
          
          REDIM UBSenRdRec(1) AS UBGilSensusReadRecType
          UBSenRdRecLen = LEN(UBSenRdRec(1))
          
          KillFile FileName$
          UBSenRdFile = FREEFILE
          OPEN FileName$ FOR RANDOM SHARED AS UBSenRdFile LEN = UBSenRdRecLen
          NumSenRdRecs = LOF(UBSenRdFile) / UBSenRdRecLen
          
          IF PendFlag THEN
            SaveScrn TempScrn()
            MakeSequenceIndex ""
            IdxName$ = "UBTEMP.IDX"
            RestScrn TempScrn()
          ELSE
            IdxName$ = "UBCUSTBK.IDX"
          END IF
          'Open the Location Order for Reading
          IdxRecLen = 4         'we are using a integer
          IdxFileSize& = FileSize&(IdxName$)
          
          IdxNumOfRecs = IdxFileSize& \ IdxRecLen
          REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
          FGetAH IdxName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs  'load it
          
          Cnt! = 1
          
          DO
            PRec& = IdxBuff(Cnt!).RecNum
            IF PRec& > 0 THEN
              GET UBFile, PRec&, UBCustRec(1)
              FOR BookCnt = 1 TO NumberofRoutes
                IF VAL(UBCustRec(1).BOOK) = Route(BookCnt) AND (UBCustRec(1).Status <> "F") THEN
                  GOSUB EWriteRecord
                END IF
              NEXT
            END IF
            Cnt! = Cnt! + 1
          LOOP UNTIL Cnt! > IdxNumOfRecs
          CLOSE
          Done = True
          FOR lc = 8 TO 10
            LOCATE lc, 16
            PRINT Space40$
          NEXT
          LOCATE 8, 18
          COLOR 11
          PRINT "Interrogator Number "; INTERROGATOR
          LOCATE 9, 18
          COLOR 11
          PRINT "Information Sent to ESensus System Successfully"
          LOCATE 12, 18
          PRINT "THESE ROUTES LOADED!!!                      "
          BEEP
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
  EXIT SUB
  RETURN
  
  
EWriteRecord:
  'May Have Up to 10 Meters to Read
  MtrCnt! = 1
  
  Account$ = STR$(PRec&)
  WHILE MtrCnt! < 8
    IF (ASC(UBCustRec(1).LocMeters(MtrCnt!).MtrType) > 32) THEN
      'IF UBCustRec(1).LocMeters(MtrCnt!).MTRType = "C" OR UBCustRec(1).LocMeters(MtrCnt!).MTRType = "W" OR UBCustRec(1).LocMeters(MtrCnt!).MTRType = "T" OR UBCustRec(1).LocMeters(MtrCnt!).MTRType = "S" THEN
      SELECT CASE UBCustRec(1).LocMeters(MtrCnt!).MtrType
      CASE "C", "S", "W", "T", "E", "D"         'here dale
        IF (UBCustRec(1).LocMeters(MtrCnt!).MtrType = "T" AND VAL(QPTrim$(UBCustRec(1).LocMeters(MtrCnt!).MTRNUM)) = 0) THEN
          GOTO ESkipEm
        ELSE
          ' Determine Sensus Meter Type
          IF UBCustRec(1).LocMeters(MtrCnt!).MtrType = "T" THEN SensusType$ = "B" ELSE SensusType$ = "M"
          ' Determine High and Low Reading
          Average# = UBCustRec(1).LocMeters(MtrCnt!).AvgUse
          ILowRead$ = RIGHT$(STR$((UBCustRec(1).LocMeters(MtrCnt!).CurRead)), 8)
          ILowRead# = VAL(ILowRead$)
          HighRead# = FIX(Average# * (HighVar / 100)) + UBCustRec(1).LocMeters(MtrCnt!).CurRead
          IF FIX(HighRead#) = ILowRead# THEN HighRead# = HighRead# + 5
          IHighRead$ = STR$(HighRead#)
          IHighRead$ = RIGHT$(IHighRead$, 8)
          MeterID$ = LTRIM$(UBCustRec(1).LocMeters(MtrCnt!).MTRNUM)
          MeterID$ = RTRIM$(MeterID$)
          'Account$ = LEFT$(Account$, 6) + "-" + RIGHT$(STR$(MtrCnt!), 1)
          IF VAL(MeterID$) = 0 THEN
            MeterID$ = UBCustRec(1).BOOK + UBCustRec(1).SEQNUMB
          END IF
          IF LEN(MeterID$) < 8 THEN
            MeterID$ = STRING$(8 - LEN(MeterID$), "0") + MeterID$
          END IF
          
          MeterID$ = LEFT$(MeterID$, 8)
          
          'Set Record Fields and Put On Disk
          UBSenRdRec(1).CustLastName = ""
          UBSenRdRec(1).CustFirstName = ""
          UBSenRdRec(1).MeterID = ""
          UBSenRdRec(1).Account = ""
          UBSenRdRec(1).LowRead = ""
          UBSenRdRec(1).HighRead = ""
          UBSenRdRec(1).SensusType = ""
          UBSenRdRec(1).PastRead = ""
          UBSenRdRec(1).CurRead = ""
          UBSenRdRec(1).ServAddress = ""
          UBSenRdRec(1).LocationNumber = ""
          UBSenRdRec(1).Message = ""
          
          UBSenRdRec(1).ServAddress = QPTrim$(UBCustRec(1).SERVADDR)
          UBSenRdRec(1).MeterID = MeterID$
          UBSenRdRec(1).LowRead = ILowRead$
          UBSenRdRec(1).HighRead = IHighRead$
          UBSenRdRec(1).Account = Account$ + "-" + QPTrim$(STR$(MtrCnt!))
          UBSenRdRec(1).SensusType = SensusType$
          UBSenRdRec(1).CustLastName = QPTrim$(UBCustRec(1).CUSTNAME)
          UBSenRdRec(1).CustFirstName = ""
          UBSenRdRec(1).Message = QPTrim$(UBCustRec(1).HHMSG1)
          PUT UBSenRdFile, (LOF(UBSenRdFile) / UBSenRdRecLen) + 1, UBSenRdRec(1)
        END IF
      END SELECT
    END IF
ESkipEm:
    MtrCnt! = MtrCnt! + 1
  WEND
  
  RETURN
  
EFILENAM1:
  PathWay$ = QPTrim$(PathWay$)
  FileName$ = PathWay$ + "ssi00" + LTRIM$(STR$(INTERROGATOR)) + ".rte"
  RETURN
  
EGetSensus:
  LibName$ = "UB"
  ScrnName$ = "UBSSHH1G"
  Help$ = "Sensus Hand Held Processing"
  CursorOff
  
  Help$ = "Upload HandHeld Operation"
  DisplayUBScrn ScrnName$
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
  NEXT
  
  Action = 1
  NumberofRoutes = 0
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      IF VAL(Form$(1, 0)) < 1 OR VAL(Form$(2, 0)) < 1 THEN
        LOCATE 12, 20
        COLOR 12
        PRINT "            [  E R R O R  ]"
        LOCATE 13, 20
        COLOR 11
        PRINT "     Invalid Interrogator Number "
        LOCATE 14, 20
        COLOR 11
        PRINT "      or Invalid Reading Date "
        BEEP
        WaitForAction
        CLOSE
        GOTO EGetSensus
        Frm(1).FldNo = 1
        Action = 1
        GOTO EGetSensus
      ELSE
        INTERROGATOR = VAL(Form$(1, 0))
        GOSUB EFileNam          ' Get Sensus File
        GOSUB EOpenCustFile     'Open Customer Data File
        MRDate$ = QPTrim$(Form$(2, 0))
        'Open and Initialize the Sensus Read Information File
        REDIM UBSenGetRdRec(1) AS UBGilSensusGetReadRecType
        UBSenGetRdRecLen = LEN(UBSenGetRdRec(1))
        UBSenGetRdFile = FREEFILE
        OPEN FileName$ FOR RANDOM ACCESS READ WRITE SHARED AS UBSenGetRdFile LEN = UBSenGetRdRecLen
        NumSenGetRdRecs = LOF(UBSenGetRdFile) / UBSenGetRdRecLen
        
        IF NumSenGetRdRecs = 0 THEN
          LOCATE 12, 20: COLOR 12
          PRINT "            [  E R R O R  ]"
          LOCATE 13, 20: COLOR 11
          PRINT "     Nothing Found From Interrogator "
          LOCATE 14, 20: COLOR 11
          PRINT "       Please Reenter Unit Number "
          BEEP
          WaitForAction
          CLOSE
          GOTO EGetSensus
        END IF
        Cnt! = 1                ' Initialize File Counter to 1
        DO
          GET UBSenGetRdFile, Cnt!, UBSenGetRdRec(1)
          PRec& = VAL(QPTrim$(UBSenGetRdRec(1).Account))
          IF PRec& > 0 THEN
            GET UBFile, PRec&, UBCustRec(1)
            GOSUB EExtractRecord
          END IF
          Cnt! = Cnt! + 1
        LOOP UNTIL Cnt! > NumSenGetRdRecs
        CLOSE
        
        Done = True
        FOR lc = 8 TO 10
          LOCATE lc, 16
          PRINT Space40$
        NEXT
        LOCATE 8, 18
        COLOR 11
        PRINT "Interrogator Number "; INTERROGATOR
        LOCATE 13, 18
        COLOR 11
        PRINT "Reading Updated From Sensus System Successfully"
        BEEP
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
  
EExtractRecord:
  UpdCnt! = UpdCnt! + 1
  LOCATE 13, 20: COLOR 11
  PRINT " Updated Count:  "; : COLOR 15: PRINT UpdCnt!
  COLOR 11
  
  'New Way To Activate
  DashPos = INSTR(UBSenGetRdRec(1).Account, "-")
  
  MT$ = MID$(UBSenGetRdRec(1).Account, DashPos + 1)
  MeterRecord = VAL(MT$)
  
  'MeterRecord = VAL(RIGHT$((QPTrim$(UBSenGetRdRec(1).Account)), 1))
  
  IF MeterRecord = 0 THEN MeterRecord = 1
  
  ' Check Meter Updated Flag
  ' Update Meter W/Reading
  CurReading# = VAL(UBSenGetRdRec(1).CurRead)

  MeterReadDate$ = LEFT$(UBSenGetRdRec(1).DateRead, 2) + "/" + MID$(UBSenGetRdRec(1).DateRead, 3, 2) + "/" + RIGHT$(Form$(2, 0), 2)

  IF Date2Num(MeterReadDate$) < 0 THEN
    MeterReadDate$ = MRDate$
  END IF

  'IF UBCustRec(1).LocMeters(MeterRecord).MTRType = "T" THEN
  '  CurReading# = CurReading# * 1000
  'END IF
  
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
  
EFileNam:
  PathWay$ = QPTrim$(PathWay$)
  FileName$ = PathWay$ + "exssi00" + QPTrim$(STR$(INTERROGATOR)) + ".dat"
  RETURN
  
EOpenCustFile:
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  RETURN
  
END SUB

SUB EstMeterReading
  
  REDIM TempScrn(0)
  LibName$ = "UB"
  ScrnName$ = "BOOK2EST"
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen
  
  TownName$ = UBSetUpRec(1).UTILNAME
  
  IF INSTR(TownName$, "BEECH MOUNTAIN") THEN
    BeechFlag = True
  END IF

  IF INSTR(TownName$, "NORTH LENOIR") THEN
    NorthFlag = True
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
  
  FirstTime = True
  
  BlockClear
  DisplayUBScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      FirstTime = False
      LSET Form$(1, 0) = DATE$
      Action = 1
    END IF
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      GOSUB CheckEstInfo
      IF Ok2EST THEN
        ExitFlag = True
      ELSE
        SaveScrn TempScrn()
        Ok = MsgBox%("UB", "BADEINFO")
        RestScrn TempScrn()
        Action = 2
        Frm(1).FldNo = 1
      END IF
    CASE EscKey
      ExitFlag = True
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 14   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 29 TO 40           '--Cancel button
          PressButton 27, 14, 29, 40
        CASE 42 TO 53           '--F5 button
          PressButton 13, 14, 42, 53
        END SELECT
      END SELECT                'row
    END IF
    
  LOOP UNTIL ExitFlag
  
  ERASE Frm, Form$, Fld
  
  IF NOT Ok2EST THEN
    GOTO ExitEst
  END IF
  
  IF BOOK < 10 THEN
    BOOK$ = "0" + QPTrim$(STR$(BOOK))
  ELSE
    BOOK$ = QPTrim$(STR$(BOOK))
  END IF
  
  BlockClear
  ShowProcessingScrn "Estimating Book: " + BOOK$ + "  "
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  UBCustFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCustFile LEN = UBCustRecLen
  NumOfCust = LOF(UBCustFile) \ UBCustRecLen
  
  FOR LCnt& = 1 TO NumOfCust
    GET #UBCustFile, LCnt&, UBCustRec(1)
    DidEM = False
    IF UBCustRec(1).Status = "A" THEN
      IF UBCustRec(1).BOOK = BOOK$ THEN
        FOR zz = 1 TO 7
          IF LEN(QPTrim$(UBCustRec(1).LocMeters(zz).MtrType)) > 0 THEN
            IF UBCustRec(1).LocMeters(zz).UseCnt > 0 AND UBCustRec(1).LocMeters(zz).ReadFlag <> "Y" THEN
              DidEM = True
              UBCustRec(1).LocMeters(zz).PrevRead = UBCustRec(1).LocMeters(zz).CurRead
              UBCustRec(1).LocMeters(zz).PastDate = UBCustRec(1).LocMeters(zz).CurDate
              UBCustRec(1).LocMeters(zz).ReadFlag = "Y"
              IF BeechFlag THEN
                AvgUse& = 0
              ELSEIF NorthFlag THEN
                AvgUse& = UBCustRec(1).LocMeters(zz).AvgUse / 100
                AvgUse& = AvgUse& * 100
              ELSE
                AvgUse& = UBCustRec(1).LocMeters(zz).AvgUse
              END IF
              'AvgUse& = (UBCustRec(1).LocMeters(zz).AvgUse) / UBCustRec(1).LocMeters(zz).UseCnt)
              UBCustRec(1).LocMeters(zz).CurRead = UBCustRec(1).LocMeters(zz).CurRead + AvgUse&
              UBCustRec(1).LocMeters(zz).CurDate = ESTDate
            END IF
          END IF
        NEXT
        IF DidEM THEN
          DoneCnt = DoneCnt + 1
          UBCustRec(1).EstFlag = "E"
          PUT #UBCustFile, LCnt&, UBCustRec(1)
        END IF
      END IF
    END IF
    ShowPctComp LCnt&, NumOfCust
  NEXT
  CLOSE
  
  BlockClear
  DisplayUBScrn "UPDATEOK"
  WaitForAction
  
ExitEst:
  EXIT SUB
  
CheckEstInfo:
  
  ESTDate = Date2Num%(Form$(1, 0))
  BOOK = VAL(Form$(2, 0))
  IF ESTDate < 0 OR BOOK < 1 THEN
    Ok2EST = False
  ELSE
    Ok2EST = True
  END IF
  
  
  RETURN
END SUB

SUB HandHeldProcessing
  
  SHARED Choice$()
  
  REDIM TempScrn(0)
  REDIM Route(100)
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  ' Check For Device Type and Run Appropriate Program
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen
  
  Space40$ = STRING$(40, 32)
  
  TownName$ = UBSetUpRec(1).UTILNAME
  IF INSTR(TownName$, "HERTFORD") THEN
    HertFlag = True
  END IF
  
  IF INSTR(TownName$, "PENDELTON") THEN
    PendFlag = True
  END IF
  
  IF INSTR(TownName$, "RURAL") AND INSTR(TownName$, "DIST") THEN
    GeorgFlag = True
  END IF
  
  IF INSTR(TownName$, "JOHNSON") THEN
    GeorgFlag = True
  END IF
  
  IF INSTR(TownName$, "GRAYSON") THEN
    GrayFlag = True
  END IF

  IF INSTR(TownName$, "LONGVIEW") THEN
    LongFlag = True
  END IF
  IF INSTR(TownName$, "HILLSVILLE") THEN
    HillFlag = True
  END IF

  IF UBSetUpLen = 0 THEN
    BEEP
    CursorOff
    DisplayHHScrn "UBNOHH"
    WaitForAction
    CLOSE
    EXIT SUB
  END IF
  
  SELECT CASE UBSetUpRec(1).HHDEVICE
  CASE "H", "S", "E", "P", "C", "D", "M", "T", "L", "I", "Z"
    GoodHHDevice = True
  CASE ELSE
    BEEP
    CursorOff
    DisplayUBScrn "UBNOHH"
    WaitForAction
    CLOSE
    EXIT SUB
  END SELECT
  
  ' Valid HandHeld Type Found
  HighVar = UBSetUpRec(1).HighRead
  LowVar = UBSetUpRec(1).LowRead
  
  SELECT CASE UBSetUpRec(1).HHDEVICE
  CASE "S"
    GOSUB Sensus
  CASE "E"
    ESensus
  CASE "P"
    IF UBSetUpRec(1).HHDEVICE = "P" THEN
      PC2000flag$ = "N"
      GOSUB PC3000
    END IF
  CASE "D"
    GOSUB DataGeneral
  CASE "C"

    ProcessSysCon
  CASE "M"
    IF UBSetUpRec(1).HHDEVICE = "M" THEN
      PC2000flag$ = "Y"
      GOSUB PC3000
    END IF
  CASE "H"
    ProcessHusky
  CASE "T"
    ProcessTelxon
  CASE "L"
    ProcessLogicon
  CASE "I"
    ProcessItron
  CASE "Z"
    ProcessSchlum

  END SELECT
  
  EXIT SUB
  
Sensus:
  'If Sensus we need to know the location path of your RMMS or STARS program
  'For LongView This is D:\RMMS\
  REDIM UBSensusRec(1) AS UBSensusRecType
  UBSensusRecLen = LEN(UBSensusRec(1))
  UBSensusFile = FREEFILE
  OPEN "UBSENSUS.DAT" FOR RANDOM SHARED AS UBSensusFile LEN = UBSensusRecLen
  NumSensusRecs = LOF(UBSensusFile) / UBSensusRecLen
  IF NumSensusRecs = 1 THEN
    GET UBSensusFile, 1, UBSensusRec(1)
    PathWay$ = UBSensusRec(1).PathWay
  END IF
  
  LibName$ = "UB"
  ScrnName$ = "UBSSHH"
  Help$ = "Sensus Hand Held Processing"
  CursorOff
  
  Help$ = "Choose HandHeld Operation"
  DisplayUBScrn ScrnName$
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
  NEXT
  
  REDIM Choice$(3, 0)
  
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Send Info to Sensus"
  Choice$(2, 0) = "Get Info From Sensus"
  
  IF NumSensusRecs = 1 THEN
    Form$(2, 0) = PathWay$
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
        UBSensusRec(1).PathWay = PathWay$
        PUT UBSensusFile, 1, UBSensusRec(1)
        CLOSE UBSensusFile
        IF LEFT$(Form$(1, 0), 1) = "S" THEN GOSUB SendSensus
        IF LEFT$(Form$(1, 0), 1) = "G" THEN GOSUB GetSensus
        Done = True
      END IF
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  RETURN
  
SendSensus:
  LibName$ = "UB"
  ScrnName$ = "UBSSHH1S"
  CursorOff
  
  Help$ = "DownLoad HandHeld Operation"
  DisplayUBScrn ScrnName$
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
  NEXT
  
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
        Frm(1).FldNo = 1
        Action = 1
        Done = False
      ELSE
        IF NumberofRoutes > 0 THEN
          INTERROGATOR = VAL(Form$(1, 0))
          COLOR 11
          LOCATE 12, 18
          PRINT "LOADING THESE ROUTES !!!                      "
          GOSUB FILENAM1
          ' Set Route Number Values
          FOR Cnt = 1 TO NumberofRoutes
            Route(Cnt) = VAL(Form$(Cnt + 2, 0))
          NEXT
          
          GOSUB OpenCustFile    'Open Customer Data File
          
          'if Gilbert sc then use different include
          '"HILLSVILLE" CHILHOWIE
          
          IF INSTR(TownName$, "GILBERT") OR INSTR(TownName$, "HILLSVILLE") OR INSTR(TownName$, "CHILHOWIE") OR INSTR(TownName$, "SAWMILL") THEN
            REDIM UBGilSenRdRec(1) AS UBGilSensusReadRecType
            UBSenRdRecLen = LEN(UBGilSenRdRec(1))
          ELSE
            REDIM UBSenRdRec(1) AS UBSensusReadRecType
            UBSenRdRecLen = LEN(UBSenRdRec(1))
          END IF
          
          'UBSenRdFile = FREEFILE
          'OPEN FileName$ FOR RANDOM SHARED AS UBSenRdFile LEN = UBSenRdRecLen
          'CLOSE UBSenRdFile
          
          KillFile FileName$
          'KILL FileName$
          UBSenRdFile = FREEFILE
          OPEN FileName$ FOR RANDOM SHARED AS UBSenRdFile LEN = UBSenRdRecLen
          NumSenRdRecs = LOF(UBSenRdFile) / UBSenRdRecLen
          
          IF PendFlag THEN
            SaveScrn TempScrn()
            MakeSequenceIndex ""
            IdxName$ = "UBTEMP.IDX"
            RestScrn TempScrn()
          ELSE
            IdxName$ = "UBCUSTBK.IDX"
          END IF
          'Open the Location Order for Reading
          IdxRecLen = 4         'we are using a integer
          IdxFileSize& = FileSize&(IdxName$)
          
          IdxNumOfRecs = IdxFileSize& \ IdxRecLen
          REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
          FGetAH IdxName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs  'load it
          
          Cnt! = 1
          
          DO
            PRec& = IdxBuff(Cnt!).RecNum
            IF PRec& > 0 THEN
              GET UBFile, PRec&, UBCustRec(1)
              FOR BookCnt = 1 TO NumberofRoutes
                IF VAL(UBCustRec(1).BOOK) = Route(BookCnt) AND (UBCustRec(1).Status <> "F") THEN
                  IF HertFlag AND UBCustRec(1).Status = "I" THEN
                    'don't write this one; skip inactive
                  ELSE
                    GOSUB WriteRecord
                  END IF
                END IF
              NEXT
            END IF
            Cnt! = Cnt! + 1
          LOOP UNTIL Cnt! > IdxNumOfRecs
          CLOSE
          Done = True
          FOR lc = 8 TO 10
            LOCATE lc, 16
            PRINT Space40$
          NEXT
          LOCATE 8, 18
          COLOR 11
          PRINT "Interrogator Number "; INTERROGATOR
          LOCATE 9, 18
          COLOR 11
          PRINT "Information Sent to Sensus System Successfully"
          LOCATE 12, 18
          PRINT "THESE ROUTES LOADED!!!                      "
          BEEP
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
  
  
WriteRecord:
  'May Have Up to 10 Meters to Read
  MtrCnt! = 1
  Account$ = STR$(PRec&)
  Account$ = LEFT$(Account$, 6) + "-" + RIGHT$(STR$(MtrCnt!), 1)
  
  WHILE MtrCnt! < 8
    IF (ASC(UBCustRec(1).LocMeters(MtrCnt!).MtrType) > 32) THEN
      IF UBCustRec(1).LocMeters(MtrCnt!).MtrType = "C" OR UBCustRec(1).LocMeters(MtrCnt!).MtrType = "W" OR UBCustRec(1).LocMeters(MtrCnt!).MtrType = "T" OR UBCustRec(1).LocMeters(MtrCnt!).MtrType = "S" THEN
        'here dale
        IF (UBCustRec(1).LocMeters(MtrCnt!).MtrType = "T" AND VAL(QPTrim$(UBCustRec(1).LocMeters(MtrCnt!).MTRNUM)) = 0) THEN
          GOTO SkipEm
        ELSE
          ' Determine Sensus Meter Type
          IF UBCustRec(1).LocMeters(MtrCnt!).MtrType = "T" THEN SensusType$ = "B" ELSE SensusType$ = "M"
          ' Determine High and Low Reading
          Average# = UBCustRec(1).LocMeters(MtrCnt!).AvgUse
          ILowRead$ = RIGHT$(STR$((UBCustRec(1).LocMeters(MtrCnt!).CurRead)), 8)
          ILowRead# = VAL(ILowRead$)
          HighRead# = FIX(Average# * (HighVar / 100)) + UBCustRec(1).LocMeters(MtrCnt!).CurRead
          IF FIX(HighRead#) = ILowRead# THEN HighRead# = HighRead# + 5
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
          
          'Set Record Fields and Put On Disk
          IF INSTR(TownName$, "GILBERT") OR INSTR(TownName$, "HILLSVILLE") OR INSTR(TownName$, "CHILHOWIE") OR INSTR(TownName$, "SAWMILL") THEN
            UBGilSenRdRec(1).CustLastName = UBCustRec(1).CUSTNAME
            UBGilSenRdRec(1).CustFirstName = ""
            UBGilSenRdRec(1).MeterID = MeterID$
            UBGilSenRdRec(1).Account = Account$
            UBGilSenRdRec(1).LowRead = ILowRead$
            UBGilSenRdRec(1).HighRead = IHighRead$
            UBGilSenRdRec(1).SensusType = SensusType$           ' B=Touch Read : M=Manual
            UBGilSenRdRec(1).PastRead = ILowRead$
            UBGilSenRdRec(1).CurRead = ILowRead$
            UBGilSenRdRec(1).ServAddress = UBCustRec(1).SERVADDR
            UBGilSenRdRec(1).LocationNumber = UBCustRec(1).BOOK + "-" + UBCustRec(1).SEQNUMB
            UBGilSenRdRec(1).Message = UBCustRec(1).HHMSG1 + LEFT$(UBCustRec(1).HHMSG2, 10)
            PUT UBSenRdFile, (LOF(UBSenRdFile) / UBSenRdRecLen) + 1, UBGilSenRdRec(1)
          ELSE
            UBSenRdRec(1).ServAddress = UBCustRec(1).SERVADDR
            UBSenRdRec(1).MeterID = MeterID$
            UBSenRdRec(1).LowRead = ILowRead$
            UBSenRdRec(1).HighRead = IHighRead$
            UBSenRdRec(1).Account = Account$
            UBSenRdRec(1).SensusType = SensusType$
            UBSenRdRec(1).CUSTNAME = UBCustRec(1).CUSTNAME
            UBSenRdRec(1).SerialNumb = ""
            PUT UBSenRdFile, (LOF(UBSenRdFile) / UBSenRdRecLen) + 1, UBSenRdRec(1)
          END IF
        END IF
      END IF
    END IF
SkipEm:
    MtrCnt! = MtrCnt! + 1
  WEND
  
  RETURN
  
FILENAM1:
  PathWay$ = QPTrim$(PathWay$)
  FileName$ = PathWay$ + "ssi00" + LTRIM$(STR$(INTERROGATOR)) + ".rte"
  RETURN
  
GetSensus:
  LibName$ = "UB"
  ScrnName$ = "UBSSHH1G"
  Help$ = "Sensus Hand Held Processing"
  CursorOff
  
  Help$ = "Upload HandHeld Operation"
  DisplayUBScrn ScrnName$
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
  NEXT
  
  Action = 1
  NumberofRoutes = 0
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key

      IF VAL(Form$(1, 0)) < 1 OR VAL(Form$(2, 0)) < 1 THEN
        LOCATE 12, 20
        COLOR 12
        PRINT "            [  E R R O R  ]"
        LOCATE 13, 20
        COLOR 11
        PRINT "     Invalid Interrogator Number "
        LOCATE 14, 20
        COLOR 11
        PRINT "      or Invalid Reading Date "
        BEEP
        WaitForAction
        CLOSE
        GOTO GetSensus
        Frm(1).FldNo = 1
        Action = 1
        GOTO GetSensus
      ELSE
        INTERROGATOR = VAL(Form$(1, 0))
        GOSUB FileNam           ' Get Sensus File
        
        GOSUB OpenCustFile      'Open Customer Data File
        
        'Open and Initialize the Sensus Read Information File
        'if Gilbert sc then use different include
        IF INSTR(TownName$, "GILBERT") OR INSTR(TownName$, "HILLSVILLE") OR INSTR(TownName$, "CHILHOWIE") OR INSTR(TownName$, "SAWMILL") THEN
          'IF LEFT$(TownName$, 7) = "GILBERT" THEN
          REDIM UBGilSenGetRdRec(1) AS UBGilSensusGetReadRecType
          UBSenGetRdRecLen = LEN(UBGilSenGetRdRec(1))
        ELSE
          REDIM UBSenGetRdRec(1) AS UBSensusGetReadRecType
          UBSenGetRdRecLen = LEN(UBSenGetRdRec(1))
        END IF
        UBSenGetRdFile = FREEFILE
        OPEN FileName$ FOR RANDOM ACCESS READ WRITE SHARED AS UBSenGetRdFile LEN = UBSenGetRdRecLen
        NumSenGetRdRecs = LOF(UBSenGetRdFile) / UBSenGetRdRecLen
        IF NumSenGetRdRecs = 0 THEN
          LOCATE 12, 20: COLOR 12
          PRINT "            [  E R R O R  ]"
          LOCATE 13, 20: COLOR 11
          PRINT "     Nothing Found From Interrogator "
          LOCATE 14, 20: COLOR 11
          PRINT "       Please Reenter Unit Number "
          BEEP
          WaitForAction
          CLOSE
          GOTO GetSensus
        END IF
        Cnt! = 1                ' Initialize File Counter to 1
        DO
          IF INSTR(TownName$, "GILBERT") OR INSTR(TownName$, "HILLSVILLE") OR INSTR(TownName$, "CHILHOWIE") OR INSTR(TownName$, "SAWMILL") THEN
            GET UBSenGetRdFile, Cnt!, UBGilSenGetRdRec(1)
            PRec& = VAL(QPTrim$(UBGilSenGetRdRec(1).Account))
          ELSE
            GET UBSenGetRdFile, Cnt!, UBSenGetRdRec(1)
            PRec& = VAL(QPTrim$(UBSenGetRdRec(1).Account))
          END IF
          
          IF PRec& > 0 THEN
            GET UBFile, PRec&, UBCustRec(1)
            GOSUB ExtractRecord
          END IF
          Cnt! = Cnt! + 1
        LOOP UNTIL Cnt! > NumSenGetRdRecs
        CLOSE
        
        Done = True
        FOR lc = 8 TO 10
          LOCATE lc, 16
          PRINT Space40$
        NEXT
        LOCATE 8, 18
        COLOR 11
        PRINT "Interrogator Number "; INTERROGATOR
        LOCATE 13, 18
        COLOR 11
        PRINT "Reading Updated From Sensus System Successfully"
        BEEP
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
  
  
ExtractRecord:
  UpdCnt! = UpdCnt! + 1
  LOCATE 13, 20: COLOR 11
  PRINT " Updated Count:  "; : COLOR 15: PRINT UpdCnt!
  COLOR 11
  
  'New Way To Activate
  IF INSTR(TownName$, "GILBERT") OR INSTR(TownName$, "HILLSVILLE") OR INSTR(TownName$, "CHILHOWIE") OR INSTR(TownName$, "SAWMILL") THEN
    MeterRecord = VAL(RIGHT$((QPTrim$(UBGilSenGetRdRec(1).Account)), 1))
  ELSE
    MeterRecord = VAL(RIGHT$((QPTrim$(UBSenGetRdRec(1).Account)), 1))
  END IF
  
  IF MeterRecord = 0 THEN MeterRecord = 1
  
  ' Check Meter Updated Flag
  ' Update Meter W/Reading
  IF INSTR(TownName$, "GILBERT") OR INSTR(TownName$, "HILLSVILLE") OR INSTR(TownName$, "CHILHOWIE") OR INSTR(TownName$, "SAWMILL") THEN
    CurReading# = VAL(UBGilSenGetRdRec(1).CurRead)
    'MeterReadDate$ = LEFT$(DATE$, 6) + RIGHT$(DATE$, 2)
    MeterReadDate$ = Form$(2, 0)
  ELSE
    CurReading# = VAL(UBSenGetRdRec(1).Reading)
    'MeterReadDate$ = LEFT$(UBSenGetRdRec(1).DateRead, 2) + "/" + MID$(UBSenGetRdRec(1).DateRead, 3, 2) + "/" + RIGHT$(Form$(2, 0), 2)
    MeterReadDate$ = LEFT$(UBSenGetRdRec(1).DateRead, 2) + "/" + MID$(UBSenGetRdRec(1).DateRead, 3, 2) + "/" + MID$(Form$(2, 0), 7, 2) + RIGHT$(UBSenGetRdRec(1).DateRead, 2)
  END IF
  
  IF UBCustRec(1).LocMeters(MeterRecord).MtrType = "T" THEN

    IF GrayFlag THEN
      CurReading# = CurReading# * 100
    ELSEIF LongFlag OR HillFlag THEN
      'Don't do anything!
    ELSEIF NOT HertFlag THEN
      CurReading# = CurReading# * 1000
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
  PUT UBFile, PRec&, UBCustRec(1)
  RETURN
  
FileNam:
  PathWay$ = QPTrim$(PathWay$)
  FileName$ = PathWay$ + "exssi00" + QPTrim$(STR$(INTERROGATOR)) + ".dat"
RETURN
  
PC3000:
  REDIM UBPC3000Rec(1) AS UBPC3000RecType
  UBPC3000RecLen = LEN(UBPC3000Rec(1))
  UBPC3000File = FREEFILE
  OPEN "UBPC3000.DAT" FOR RANDOM SHARED AS UBPC3000File LEN = UBPC3000RecLen
  NumPC3000Recs = LOF(UBPC3000File) / UBPC3000RecLen
  IF NumPC3000Recs = 1 THEN
    GET UBPC3000File, 1, UBPC3000Rec(1)
    PathWay$ = UBPC3000Rec(1).PathWay
  END IF
  
  LibName$ = "UB"
  ScrnName$ = "UBPC30HH"
  Help$ = "MicroPlam Hand Held Processing"
  CursorOff
  
  Help$ = "Choose HandHeld Operation"
  DisplayUBScrn ScrnName$
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
  NEXT
  
  REDIM Choice$(5, 0)
  
  Choice$(0, 0) = "1"
  
  IF PC2000flag$ = "Y" THEN
    Choice$(1, 0) = "Send Info to PC2000"
    Choice$(2, 0) = "Get Info From PC2000"
  ELSE
    Choice$(1, 0) = "Send Info to PC3000"
    Choice$(2, 0) = "Get Info From PC3000"
  END IF
  
  IF NumPC3000Recs = 1 THEN
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
        UBPC3000Rec(1).PathWay = PathWay$
        PUT UBPC3000File, 1, UBPC3000Rec(1)
        CLOSE UBPC3000File
        IF LEFT$(Form$(1, 0), 1) = "S" THEN GOSUB SendPC3000
        IF LEFT$(Form$(1, 0), 1) = "G" THEN GOSUB GetPC3000
        Done = True
      END IF
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  RETURN
  
SendPC3000:
  LibName$ = "UB"
  ScrnName$ = "UBPCHH1S"
  Help$ = "MicroPalm Hand Held Processing"
  CursorOff
  
  Help$ = "DownLoad HandHeld Operation"
  DisplayUBScrn ScrnName$
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
  NEXT
  
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
      WhatTypes$ = LEFT$(Form$(1, 0), 1)
      ' Next Check Routes
      IF GeorgFlag THEN
        NumberofRoutes = 100
      END IF
      IF NumberofRoutes > 0 THEN
        COLOR 11
        LOCATE 12, 18
        PRINT "LOADING THESE ROUTES !!!                      "
        GOSUB FILENAM1
        ' Set Route Number Values
        IF GeorgFlag THEN
          FOR Cnt = 1 TO NumberofRoutes
            Route(Cnt) = Cnt - 1
          NEXT
        ELSE
          FOR Cnt = 1 TO NumberofRoutes
            Route(Cnt) = VAL(Form$(Cnt + 2, 0))
          NEXT
        END IF
        
        PathWay$ = RTRIM$(PathWay$)
        
        GOSUB OpenCustFile      'Open Customer Data File
        
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
          IdxRecLen = 4         'we are using an integer
          MakeSequenceIndex "Sequence Number"
          IdxNumOfRecs = FileSize&("UBTEMP.IDX") \ 4
          REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
          FGetAH "UBTEMP.IDX", IdxBuff(1), 4, IdxNumOfRecs
        ELSE
          IdxRecLen = 4         'we are using an integer
          IdxFileSize& = FileSize&("UBCUSTBK.IDX")
          IdxNumOfRecs = IdxFileSize& \ IdxRecLen
          REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
          FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs            'load it
        END IF
        Cnt! = 1
        
        DO
          
          PRec& = IdxBuff(Cnt!).RecNum
          IF PRec& > 0 THEN
            GET UBFile, PRec&, UBCustRec(1)
            DisplayAcct$ = "Processing Location #" + UBCustRec(1).BOOK + UBCustRec(1).SEQNUMB
            PrintHelp DisplayAcct$
            FOR BookCnt = 1 TO NumberofRoutes
              IF VAL(UBCustRec(1).BOOK) = Route(BookCnt) AND (UBCustRec(1).Status <> "F") THEN
                GOSUB PC3000WriteRecord
              END IF
            NEXT
          END IF
          Cnt! = Cnt! + 1
          
        LOOP UNTIL Cnt! > IdxNumOfRecs
        CLOSE
        CLS
        COLOR 11
        IF PC2000flag$ = "Y" THEN
          PRINT "PC 2000 Connect"
          COLOR 15
          PRINT "Make Sure the PC 2000 is in the Cradle or Connected by Data Cable"
          PRINT "At the C>  on the PC 2000, press  c and press <enter>"
          PRINT "This will run the remserve program on the PC 2000."
        ELSE
          PRINT "PC 3000 Connect"
          COLOR 15
          PRINT "Make Sure the PC 3000 is in the Cradle or Connected by Data Cable"
          PRINT "The Screen Should be at Serial Communications in Progress"
          PRINT "You get to the Communications Screen by Pressing the 'L' key when"
          PRINT "the directory is being displayed."
        END IF
        PRINT
        COLOR 11
        
        Destination$ = "copy ubcusttr.dat  " + PathWay$
        IF PC2000flag$ = "Y" THEN
          SHELL "REMDISK /com1"
          SHELL Destination$
          SHELL "REMDISK /u"
        ELSE
          SHELL "CONNECT"
          SHELL Destination$
        END IF
        
        PRINT
        PRINT "Press any key. . ."
        WaitForAction
        
        Done = True
        DisplayUBScrn ScrnName$
        FOR lc = 8 TO 10
          LOCATE lc, 16
          PRINT Space40$
        NEXT
        LOCATE 9, 18
        COLOR 11
        IF PC2000flag$ = "Y" THEN
          PRINT "Information Sent to PC2000 Successfully"
        ELSE
          PRINT "Information Sent to PC3000 Successfully"
        END IF
        LOCATE 12, 18: PRINT "THESE ROUTES LOADED!!!                      "
        BEEP
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
  
  
PC3000WriteRecord:
  'May Have Up to 7 Meters to Read
  FOR MtrCnt! = 1 TO 7
    MeterOK = False
    Account$ = STR$(PRec&)
    Account$ = LEFT$(Account$, 6) + "-" + RIGHT$(STR$(MtrCnt!), 1)
    
    IF ASC(UBCustRec(1).LocMeters(MtrCnt!).MtrType) > 32 THEN
      MtrType$ = UBCustRec(1).LocMeters(MtrCnt!).MtrType
      IF MtrType$ = "W" OR MtrType$ = "S" OR MtrType$ = "C" OR MtrType$ = "E" OR MtrType$ = "D" OR MtrType$ = "G" THEN
        SELECT CASE WhatTypes$
        CASE "W"                'water/sewer
          IF MtrType$ = "W" OR MtrType$ = "S" OR MtrType$ = "C" THEN
            MeterOK = True
          END IF
        CASE "E"                'electric & demand elec.
          IF MtrType$ = "E" OR MtrType$ = "D" THEN
            MeterOK = True
          END IF
        CASE "G"                'gas
          IF MtrType$ = "G" THEN
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
          UBPC3000RDRec(1).MeterType$ = UBCustRec(1).LocMeters(MtrCnt!).MtrType
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
  NEXT
  RETURN
  
GetPC3000:
  
  LibName$ = "UB"
  ScrnName$ = "UBPCHH1G"
  CursorOff
  
  Help$ = "Upload HandHeld Operation"
  DisplayUBScrn ScrnName$
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
  NEXT
  
  IF NumPC3000Recs = 1 THEN
    Form$(1, 0) = PathWay$
    'Fld(2).Protected = True
  END IF
  
  Action = 1
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      CLS : COLOR 11: PRINT "PC 3000 Connect"
      COLOR 15
      IF PC2000flag$ = "Y" THEN
        PRINT "Make Sure the PC 2000 is in the cradle.  At the C> type c and press"
        PRINT "the <ENTER> key to allow for data upload."
      ELSE
        PRINT "Make Sure the PC 3000 is in the Cradle or Connected by Data Cable"
        PRINT "The Screen Should be at Serial Communications in Progress"
        PRINT "You get to the Communications Screen by Pressing the 'L' key when"
        PRINT "the directory is being displayed."
      END IF
      PRINT
      COLOR 11
      IF PC2000flag$ = "Y" THEN
        SHELL "REMDISK /com1"
      ELSE
        SHELL "CONNECT"
      END IF
      
      DisplayUBScrn ScrnName$
      
      GOSUB OpenCustFile        'Open Customer Data File
      
      'Open and Initialize the PC3000 Read Information File
      REDIM UBPC3000GetRDRec(1) AS UBPC3000ReadRecType
      UBPC3000GetRdRecLen = LEN(UBPC3000GetRDRec(1))
      UBPC3000GetRdFile = FREEFILE
      FileNam$ = RTRIM$(PathWay$) + "UBCUSTTR.DAT"
      OPEN FileNam$ FOR RANDOM ACCESS READ WRITE SHARED AS UBPC3000GetRdFile LEN = UBPC3000GetRdRecLen
      NumPC3000GetRdRecs = LOF(UBPC3000GetRdFile) / UBPC3000GetRdRecLen
      IF NumPC3000GetRdRecs = 0 THEN
        FOR Cnt = 12 TO 16
          LOCATE Cnt, 20
          PRINT Space40$
        NEXT
        LOCATE 12, 20
        COLOR 12
        PRINT "      [  E R R O R  ]"
        LOCATE 13, 20
        COLOR 11
        PRINT " Nothing Found From MicroPalm "
        LOCATE 14, 20
        PRINT " Please Check Unit Connection "
        BEEP
        WaitForAction
        CLOSE
        GOTO GetPC3000
      END IF
      
      Cnt! = 1  ' Initialize File Counter to 1
      DO
        GET UBPC3000GetRdFile, Cnt!, UBPC3000GetRDRec(1)
        ' New Way Undelete When Our System Sends and Retrieves
        PRec& = VAL(QPTrim$(UBPC3000GetRDRec(1).Account))
        IF PRec& > 0 THEN
          GET UBFile, PRec&, UBCustRec(1)
          GOSUB PC3000ExtractRecord
        END IF
        Cnt! = Cnt! + 1
      LOOP UNTIL Cnt! > NumPC3000GetRdRecs
      CLOSE
      Done = True
      FOR lc = 12 TO 14
        LOCATE lc, 16
        PRINT Space40$
      NEXT
      LOCATE 12, 18
      COLOR 11
      PRINT "Readings Updated From PC3000 Successfully"
      BEEP
      WaitForAction
      Done = True
      IF PC2000flag$ = "Y" THEN SHELL "REMDISK /u"
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  RETURN
  
  
PC3000ExtractRecord:
  UpdCnt! = UpdCnt! + 1
  LOCATE 14, 16
  PRINT STRING$(30, 32);
  LOCATE 13, 16
  COLOR 11
  PRINT "   Updated Count:  ";
  COLOR 15
  PRINT UpdCnt!
  COLOR 11
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
  PUT UBFile, PRec&, UBCustRec(1)
  
  RETURN
  
DataGeneral:
  REDIM UBDGRec(1) AS UBDGRecType
  UBDGRecLen = LEN(UBDGRec(1))
  UBDGFile = FREEFILE
  OPEN "UBDG.DAT" FOR RANDOM SHARED AS UBDGFile LEN = UBDGRecLen
  NumDGRecs = LOF(UBDGFile) / UBDGRecLen
  IF NumDGRecs = 1 THEN
    GET UBDGFile, 1, UBDGRec(1)
    PathWay$ = UBDGRec(1).PathWay
  END IF
  
  LibName$ = "UB"
  ScrnName$ = "UBDGHH1"
  CursorOff
  Help$ = "Choose HandHeld Operation"
  
  DisplayUBScrn ScrnName$
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
  NEXT
  
  REDIM Choice$(5, 0)
  
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Send Info to Genie"
  Choice$(2, 0) = "Get Info From Genie"
  
  IF NumPC3000Recs = 1 THEN
    Form$(2, 0) = PathWay$
    'Fld(2).Protected = True
  END IF
  
  Action = 1
  
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
        PathWay$ = Form$(2, 0)
        UBDGRec(1).PathWay = PathWay$
        PUT UBDGFile, 1, UBDGRec(1)
        CLOSE UBDGFile
        SELECT CASE LEFT$(Form$(1, 0), 1)
        CASE "S"
          GOSUB SendDG
        CASE "G"
          GOSUB GetDG
        END SELECT
        'IF LEFT$(Form$(1, 0), 1) = "S" THEN GOSUB SendDG
        'IF LEFT$(Form$(1, 0), 1) = "G" THEN GOSUB GetDG
        Done = True
      END IF
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  RETURN
  
SendDG:
  LibName$ = "UB"
  ScrnName$ = "UBPCHH1S"
  CursorOff
  Help$ = "DownLoad HandHeld Operation"
  DisplayUBScrn ScrnName$
  COLOR 14
  LOCATE 7, 21
  PRINT " DG 'Genie' File Maker "
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
  NEXT
  
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
      WhatTypes$ = LEFT$(Form$(1, 0), 1)
      ' Next Check Routes
      IF GeorgFlag THEN
        NumberofRoutes = 100
      END IF
      
      IF NumberofRoutes > 0 THEN
        COLOR 11
        LOCATE 12, 18
        PRINT "LOADING THESE ROUTES !!!                      "
        GOSUB FILENAM1
        ' Set Route Number Values
        
        IF GeorgFlag THEN
          FOR Cnt = 1 TO NumberofRoutes
            Route(Cnt) = Cnt - 1
          NEXT
        ELSE
          FOR Cnt = 1 TO NumberofRoutes
            Route(Cnt) = VAL(Form$(Cnt + 2, 0))
          NEXT
        END IF
        
        PathWay$ = RTRIM$(PathWay$)
        
        GOSUB OpenCustFile      'Open Customer Data File
        
        'Open and Initialize the DG Genie Read Information File
        REDIM UBDGRdRec(1) AS UBDGProcRecType
        UBDGRdRecLen = LEN(UBDGRdRec(1))
        UBDGRdFile = FREEFILE
        OPEN "UBCUSTTR.DAT" FOR RANDOM SHARED AS UBDGRdFile LEN = UBDGRdRecLen
        CLOSE UBDGRdFile
        KILL "UBCUSTTR.DAT"
        UBDGRdFile = FREEFILE
        OPEN "UBCUSTTR.DAT" FOR RANDOM SHARED AS UBDGRdFile LEN = UBDGRdRecLen
        NumDGRdRecs = LOF(UBDGRdFile) / UBDGRdRecLen
        'Open the Location Order for Reading
        IdxRecLen = 4           'we are using a integer
        IdxFileSize& = FileSize&("UBCUSTBK.IDX")
        IdxNumOfRecs = IdxFileSize& \ IdxRecLen
        REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
        FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs              'load it
        Cnt! = 1
        DO
          PRec& = IdxBuff(Cnt!).RecNum
          IF PRec& > 0 THEN
            GET UBFile, PRec&, UBCustRec(1)
            DisplayAcct$ = "Processing Location #" + UBCustRec(1).BOOK + UBCustRec(1).SEQNUMB
            PrintHelp DisplayAcct$
            FOR BookCnt = 1 TO NumberofRoutes
              IF VAL(UBCustRec(1).BOOK) = Route(BookCnt) AND (UBCustRec(1).Status <> "F") THEN
                GOSUB DGWriteRecord
              END IF
            NEXT
          END IF
          Cnt! = Cnt! + 1
        LOOP UNTIL Cnt! > IdxNumOfRecs
        CLOSE
        FOR ll = 14 TO 16
          LOCATE ll, 17
          PRINT STRING$(45, 32)
        NEXT
        LOCATE 15, 22
        BEEP
        PRINT "Data General 'Genie' File Now Ready"
        LOCATE 18, 17
        PRINT STRING$(45, 32)
        LOCATE 18, 28
        PRINT "Press any key to Continue"
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
  
DGWriteRecord:
  'May Have Up to 7 Meters to Read
  FOR MtrCnt! = 1 TO 7
    MeterOK = False
    Account$ = STR$(PRec&)
    Account$ = LEFT$(Account$, 6) + "-" + RIGHT$(STR$(MtrCnt!), 1)
    IF ASC(UBCustRec(1).LocMeters(MtrCnt!).MtrType) > 32 THEN
      MtrType$ = UBCustRec(1).LocMeters(MtrCnt!).MtrType
      IF MtrType$ = "W" OR MtrType$ = "S" OR MtrType$ = "C" OR MtrType$ = "E" OR MtrType$ = "D" OR MtrType$ = "G" THEN
        SELECT CASE WhatTypes$
        CASE "W"                'water/sewer
          IF MtrType$ = "W" OR MtrType$ = "S" OR MtrType$ = "C" THEN
            MeterType$ = "W"
            MeterOK = True
          END IF
        CASE "E"                'electric & demand elec.
          IF MtrType$ = "E" OR MtrType$ = "D" THEN
            MeterOK = True
          END IF
        CASE "G"                'gas
          IF MtrType$ = "G" THEN
            MeterType$ = "G"
            MeterOK = True
          END IF
        CASE "A", " "           'all meters
          IF MtrType$ = "W" OR MtrType$ = "S" OR MtrType$ = "C" THEN
            MeterType$ = "W"
          END IF
          IF MtrType$ = "E" OR MtrType$ = "D" THEN
            MeterType$ = "E"
          END IF
          MeterOK = True
        END SELECT
        
        IF MeterOK = True THEN
          ' Determine High and Low Reading
          Average# = UBCustRec(1).LocMeters(MtrCnt!).AvgUse
          TimesRead! = UBCustRec(1).LocMeters(MtrCnt!).UseCnt
          ILowRead$ = RIGHT$(STR$((UBCustRec(1).LocMeters(MtrCnt!).CurRead)), 8)
          ILowRead# = VAL(ILowRead$)
          LowRead# = FIX(ILowRead#)
          HighRead# = FIX(Average# * (HighVar / 100)) + UBCustRec(1).LocMeters(MtrCnt!).CurRead
          IF HighRead# < 0 OR HighRead# > 99999999 THEN HighRead# = 0
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
          UBDGRdRec(1).RouteID = LTRIM$(STR$(UBCustRec(1).SEQ))
          UBDGRdRec(1).SvcTyp = MeterType$
          UBDGRdRec(1).CUSTNAME = UBCustRec(1).CUSTNAME
          UBDGRdRec(1).SvcLoc = UBCustRec(1).SERVADDR
          UBDGRdRec(1).MeterSN = MeterID$
          UBDGRdRec(1).MeterType = "C"
          UBDGRdRec(1).High = STR$(HighRead#)
          UBDGRdRec(1).Low = STR$(LowRead#)
          UBDGRdRec(1).Msg = UBCustRec(1).HHMSG1 + " " + UBCustRec(1).HHMSG2 + " " + UBCustRec(1).HHMSG3
          UBDGRdRec(1).Account = Account$
          UBDGRdRec(1).NewRdng = ""
          UBDGRdRec(1).NewDmnd = ""
          UBDGRdRec(1).Date = ""
          UBDGRdRec(1).Time = ""
          UBDGRdRec(1).NewAcctRte = ""
          PUT UBDGRdFile, (LOF(UBDGRdFile) / UBDGRdRecLen) + 1, UBDGRdRec(1)
        END IF
      END IF
    END IF
  NEXT
  RETURN
  
  
GetDG:
  DisplayUBScrn ScrnName$
  'OPEN "dodo.txt" FOR OUTPUT AS #30
  
  GOSUB OpenCustFile            'Open Customer Data File
  
  'Open and Initialize the Genie Information File
  REDIM UBDGGetRDRec(1) AS UBDGProcRecType
  UBDGGetRdRecLen = LEN(UBDGGetRDRec(1))
  UBDGGetRdFile = FREEFILE
  FileNam$ = RTRIM$(PathWay$) + "UBCUSTTR.DAT"
  OPEN FileNam$ FOR RANDOM ACCESS READ WRITE SHARED AS UBDGGetRdFile LEN = UBDGGetRdRecLen
  NumDGGetRdRecs = LOF(UBDGGetRdFile) / UBDGGetRdRecLen
  IF NumDGGetRdRecs = 0 THEN
    FOR Cnt = 12 TO 16
      LOCATE Cnt, 20
      PRINT Space40$
    NEXT
    LOCATE 12, 20
    COLOR 12
    PRINT "      [  E R R O R  ]"
    LOCATE 13, 20
    COLOR 11
    PRINT " Nothing Found From Genie "
    BEEP
    WaitForAction
    CLOSE
    GOTO DataGeneral
  END IF
  Cnt! = 1      ' Initialize File Counter to 1
  DO
    GET UBDGGetRdFile, Cnt!, UBDGGetRDRec(1)
    ' PRINT #30, UBDGGetRDRec(1).CustName, ;
    ' New Way Undelete When Our System Sends and Retrieves
    PRec& = VAL(QPTrim$(UBDGGetRDRec(1).Account))
    IF PRec& > 0 THEN
      GET UBFile, PRec&, UBCustRec(1)
      '    PRINT #30, UBCustRec(1).CustName
      GOSUB DGExtractRecord
    END IF
    Cnt! = Cnt! + 1
  LOOP UNTIL Cnt! > NumDGGetRdRecs
  CLOSE
  Done = True
  FOR lc = 12 TO 14
    LOCATE lc, 16
    PRINT Space40$
  NEXT
  LOCATE 12, 18
  COLOR 11
  PRINT "Readings Updated Successfully"
  BEEP
  WaitForAction
  RETURN
  
DGExtractRecord:
  UpdCnt! = UpdCnt! + 1
  LOCATE 14, 16
  PRINT STRING$(30, 32);
  LOCATE 13, 16
  COLOR 11
  PRINT "   Updated Count:  ";
  COLOR 15
  PRINT UpdCnt!
  COLOR 11
  'New Way To Activate
  IF LEN(QPTrim$(UBDGGetRDRec(1).NewRdng)) > 0 THEN
    MeterRecord = VAL(RIGHT$((QPTrim$(UBDGGetRDRec(1).Account)), 1))
    ' Check Meter Updated Flag
    ' Update Meter W/Reading
    CurReading# = VAL(UBDGGetRDRec(1).NewRdng)

    ReadDate$ = LEFT$(UBDGGetRDRec(1).Date, 2) + "/" + MID$(UBDGGetRDRec(1).Date, 3, 2) + "/20" + RIGHT$(UBDGGetRDRec(1).Date, 2)

    IF UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y" THEN
      UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
      UBCustRec(1).LocMeters(MeterRecord).CurDate = Date2Num%(ReadDate$)
    ELSE
      UBCustRec(1).LocMeters(MeterRecord).PrevRead = UBCustRec(1).LocMeters(MeterRecord).CurRead
      UBCustRec(1).LocMeters(MeterRecord).PastDate = UBCustRec(1).LocMeters(MeterRecord).CurDate
      UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y"
      UBCustRec(1).LocMeters(MeterRecord).CurDate = Date2Num%(ReadDate$)
      UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
    END IF
    PUT UBFile, PRec&, UBCustRec(1)
  END IF
  RETURN
  
SysCon:
  RETURN
  
OpenCustFile:
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
RETURN
  
END SUB

SUB ProcessSchlum

  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen

  HighVar = UBSetUpRec(1).HighRead
  LowVar = UBSetUpRec(1).LowRead

  CrLf$ = CHR$(13) + CHR$(10)
  q$ = CHR$(34)
  cb$ = SPACE$(45)
  SHARED Choice$()
  REDIM Route(100)

  LibName$ = "UBHH"
  ScrnName$ = "UBSCHUMI"
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
  Choice$(1, 0) = "Export Meter Readings"
  Choice$(2, 0) = "Import Meter Readings"

  Action = 1
  
  DisplayH2Scrn ScrnName$

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode

    CASE F10Key
      WhatOp$ = LEFT$(Form$(1, 0), 1)
      IF WhatOp$ <> "E" AND WhatOp$ <> "I" THEN
        QPrintRC "Invalid Operation", 12, 40, 12
        BEEP
        WaitForAction
        QPrintRC STRING$(25, 32), 12, 40, 15
        Frm(1).FldNo = 1
        Action = 1
        Done = False
      ELSE
        Done = True
      END IF
    CASE ESC
      EXIT SUB
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done

  SELECT CASE WhatOp$
  CASE "E"
    GOSUB ExportSchlum
  CASE "I"
    GOSUB ImportSchlum
  CASE ELSE
  END SELECT


  EXIT SUB

ExportSchlum:

  LibName$ = "UBHH"
  ScrnName$ = "UBSCHL1S"

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

  DisplayH2Scrn ScrnName$

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
      IF NumberofRoutes > 0 THEN       ' Next Check Routes
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
        REDIM UBSchlumHHRec(1) AS SchlumHHType
        UBSchlumHHRecLen = LEN(UBSchlumHHRec(1))
        KillFile "UBCUSTTR.DAT"
        UBSchlFile = FREEFILE
        OPEN "UBCUSTTR.DAT" FOR RANDOM SHARED AS UBSchlFile LEN = UBSchlumHHRecLen
        
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
        WriteCnt& = 0

        DO
          PRec& = IdxBuff(Cnt!).RecNum
          IF NOT PRec& = 0 THEN
            GET UBFile, PRec&, UBCustRec(1)
            IF INSTR(UBCustRec(1).HHMSG1, "NOREAD") > 0 THEN
              GOTO HWriteSkip
            END IF
            DisplayAcct$ = UBCustRec(1).BOOK + "-" + UBCustRec(1).SEQNUMB
            QPrintRC DisplayAcct$, 11, 45, -1
            FOR BookCnt! = 1 TO NumberofRoutes
              IF VAL(UBCustRec(1).BOOK) = Route(BookCnt!) AND (UBCustRec(1).Status <> "F") THEN
                GOSUB SchlumWriteRec
              END IF
            NEXT BookCnt!
          END IF
HWriteSkip:
          Cnt! = Cnt! + 1
        LOOP UNTIL Cnt! > IdxNumOfRecs
        CLOSE

        DisplayH2Scrn "EXPORTOK"
        WaitForAction
        Done = True
      END IF

    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done

RETURN

SchlumWriteRec:
  'May Have Up to 7 Meters to Read
  FOR MtrCnt! = 1 TO 7
    MeterOK = False
    Account$ = STR$(PRec&)
    Account$ = LEFT$(Account$, 6) + "-" + RIGHT$(STR$(MtrCnt!), 1)
    MtrType$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt!).MtrType)
    IF LEN(MtrType$) > 0 THEN
      IF MtrType$ = "W" OR MtrType$ = "S" OR MtrType$ = "C" OR MtrType$ = "E" OR MtrType$ = "D" OR MtrType$ = "G" THEN
        SELECT CASE WhatTypes$
        CASE "W"                'water/sewer
          IF MtrType$ = "W" OR MtrType$ = "S" OR MtrType$ = "C" THEN
            MeterOK = True
          END IF
        CASE "E"                'electric & demand elec.
          IF MtrType$ = "E" OR MtrType$ = "D" THEN
            MeterOK = True
          END IF
        CASE "G"                'gas
          IF MtrType$ = "G" THEN
            MeterOK = True
          END IF
        CASE "A", " "           'all meters
          MeterOK = True
        END SELECT

        IF MeterOK = True THEN   ' Determine High and Low Reading
          SRouteID$ = UBCustRec(1).BOOK + UBCustRec(1).SEQNUMB + "0" + QPTrim$(STR$(MtrCnt!))
          LSET UBSchlumHHRec(1).Route = SRouteID$
          WriteCnt& = WriteCnt& + 1
          WalkSeq$ = "0000" + QPTrim$(STR$(WriteCnt&))
          LSET UBSchlumHHRec(1).WalkSeq$ = RIGHT$(WalkSeq$, 4)

          UBSchlumHHRec(1).PageNum = UBSchlumHHRec(1).WalkSeq$
          UBSchlumHHRec(1).ReadSeq = UBSchlumHHRec(1).WalkSeq$
          UBSchlumHHRec(1).HHID = ""    'HH Number 'no information
          UBSchlumHHRec(1).ReadDir = "L"
          UBSchlumHHRec(1).NumDial = ""

          MeterID$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt!).MTRNUM)

          LSET UBSchlumHHRec(1).IDExpected = MeterID$
          LSET UBSchlumHHRec(1).IDCaptured = ""
          LSET UBSchlumHHRec(1).IDOverride = ""
          LSET UBSchlumHHRec(1).Decimals = ""
          LSET UBSchlumHHRec(1).MtrRead = ""    'current from HH
          LSET UBSchlumHHRec(1).ReadOVRide = ""
          
          Average# = UBCustRec(1).LocMeters(MtrCnt!).AvgUse
          ILowRead$ = QPTrim$(STR$(UBCustRec(1).LocMeters(MtrCnt!).CurRead))
          PrevRead$ = QPTrim$(STR$(UBCustRec(1).LocMeters(MtrCnt!).PrevRead))

          PrevDate$ = Num2Date(UBCustRec(1).LocMeters(MtrCnt!).PastDate)
          PrevDate$ = LEFT$(PrevDate$, 2) + MID$(PrevDate$, 4, 2) + RIGHT$(PrevDate$, 2)

          LowRead# = FIX(ILowRead#)

          HighRead# = FIX((Average# * (HighVar / 100) - Average#) + UBCustRec(1).LocMeters(MtrCnt!).CurRead)
          HighRead$ = QPTrim$(STR$(HighRead#))
          LSET UBSchlumHHRec(1).HighLimit = HighRead$
          LSET UBSchlumHHRec(1).LowLimit = ILowRead$

          LSET UBSchlumHHRec(1).Date2Read = ""
          LSET UBSchlumHHRec(1).Date2Exp = ""
          LSET UBSchlumHHRec(1).NoteCodes = ""
          LSET UBSchlumHHRec(1).LocatCode = ""
          LSET UBSchlumHHRec(1).MtrRCode = ""
          LSET UBSchlumHHRec(1).RecType = "EU"
          LSET UBSchlumHHRec(1).RecStatus = ""
          LSET UBSchlumHHRec(1).ReadDate = ""
          LSET UBSchlumHHRec(1).ReadTime = ""
          LSET UBSchlumHHRec(1).ReadType = ""
          LSET UBSchlumHHRec(1).NetNumb = ""
          LSET UBSchlumHHRec(1).ReadAtmpt = ""
          LSET UBSchlumHHRec(1).UserChar = ""
          LSET UBSchlumHHRec(1).HHManufac = ""
          LSET UBSchlumHHRec(1).ActStatus = UBCustRec(1).Status
          LSET UBSchlumHHRec(1).MtrType = ""
          LSET UBSchlumHHRec(1).ReadFailCode = ""
          LSET UBSchlumHHRec(1).PrevRead = PrevRead$
          LSET UBSchlumHHRec(1).PrevDate = PrevDate$

          LSET UBSchlumHHRec(1).HHDisp1 = QPTrim$(UBCustRec(1).SERVADDR)
          LSET UBSchlumHHRec(1).HHDisp2 = SRouteID$
          LSET UBSchlumHHRec(1).HHDisp3 = QPTrim$(UBCustRec(1).ADDR1)
          LSET UBSchlumHHRec(1).HHDisp4 = QPTrim$(UBCustRec(1).CUSTNAME)

          LSET UBSchlumHHRec(1).Notes1 = QPTrim$(UBCustRec(1).HHMSG1)
          LSET UBSchlumHHRec(1).Notes2 = QPTrim$(UBCustRec(1).HHMSG2)
          LSET UBSchlumHHRec(1).Notes3 = QPTrim$(UBCustRec(1).HHMSG3)
          LSET UBSchlumHHRec(1).Notes4 = ""
          LSET UBSchlumHHRec(1).Notes5 = ""
          LSET UBSchlumHHRec(1).Notes6 = ""
          LSET UBSchlumHHRec(1).Notes7 = ""
          LSET UBSchlumHHRec(1).Notes8 = ""
          LSET UBSchlumHHRec(1).OpCode = ""

          LSET UBSchlumHHRec(1).UBAcctNo = QPTrim$(STR$(PRec&))
          LSET UBSchlumHHRec(1).MtrSlot = QPTrim$(STR$(MtrCnt!))

          LSET UBSchlumHHRec(1).UtilFld = ""
          LSET UBSchlumHHRec(1).CrLf = CrLf$
          
          PUT UBSchlFile, (LOF(UBSchlFile) / UBSchlumHHRecLen) + 1, UBSchlumHHRec(1)
        END IF
      END IF
    END IF
  NEXT MtrCnt!

RETURN


ImportSchlum:

  Ok = MsgBox%("UBHH", "UBSCHL1C")

  DisplayHHScrn "UBHUSK3C"
  QPrintRC " Import Meter Readings ", 7, 21, -1
  'overwrite where husky is displayed in the box

  IF Ok = 1 THEN
    Done = True
  ELSE
    QPrintRC "   Import Readings Aborted.   ", 10, 25, -1
    QPrintRC "  Press any key to continue.  ", 14, 25, -1
    CursorOff
    CALL WaitForAction
    Done = True
    GOTO HuskyGetExit:
  END IF

  '    'Open Customer Data File

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  REDIM UBSchlumHHRec(1) AS SchlumHHType
  UBSchlumHHRecLen = LEN(UBSchlumHHRec(1))

  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBFile LEN = UBCustRecLen

  UBSchlFile = FREEFILE
  OPEN "UBCUSTTR.DAT" FOR RANDOM SHARED AS UBSchlFile LEN = UBSchlumHHRecLen

  NumTRGetRecs = LOF(UBSchlFile) / UBSchlumHHRecLen

  IF NumTRGetRecs = 0 THEN
    CLOSE
    QPrintRC "                                ", 10, 24, 15
    QPrintRC "                                ", 14, 24, 15
    QPrintRC " ERROR: IMPORT FILE NOT FOUND ", 9, 25, -1
    QPrintRC " Make sure that 'UBCUSTTR.DAT'", 11, 25, -1
    QPrintRC " is in the CITIPAK directory!!", 13, 25, -1
    QPrintRC "   Press any key to continue. ", 15, 25, 14
    BEEP
    WaitForAction
    EXIT SUB
  END IF

  QPrintRC "  Processing Meter Readings.  ", 9, 25, 15
  QPrintRC "                              ", 10, 25, 11
  QPrintRC "    Updated Count:            ", 11, 25, 11
  QPrintRC "                              ", 14, 25, 11

  Cnt! = 1      ' Initialize File Counter to 1
  DO
    GET UBSchlFile, Cnt!, UBSchlumHHRec(1)
    ' New Way Undelete When Our System Sends and Retrieves
    PRec& = VAL(QPTrim$(UBSchlumHHRec(1).UBAcctNo))

    IF NOT (PRec&) = 0 THEN
      GET UBFile, PRec&, UBCustRec(1)
      GOSUB SchlumExtractRecord
    END IF
    Cnt! = Cnt! + 1

  LOOP UNTIL Cnt! > NumTRGetRecs

  CLOSE
  Done = True

  QPrintRC " Readings Updated Successfully. ", 13, 24, 15
  QPrintRC "  Press any key to continue.  ", 15, 25, 14
  BEEP
  WaitForAction
  Done = True

HuskyGetExit:
  RETURN

SchlumExtractRecord:
  UpdCnt! = UpdCnt! + 1
  QPrintRC STR$(UpdCnt!), 11, 44, 15
  'New Way To Activate
  MeterSlot = VAL(QPTrim$(UBSchlumHHRec(1).MtrSlot))
  CurReading# = VAL(QPTrim$(UBSchlumHHRec(1).MtrRead))
  RYear$ = QPTrim$(RIGHT$(UBSchlumHHRec(1).ReadDate, 2))

  IF LEN(RYear$) < 2 THEN
    BadDate = True
  END IF
  WhatYear = VAL(RYear$)
  IF WhatYear < 65 THEN
    ReadYear$ = "/20" + RYear$
  ELSE
    ReadYear$ = "/19" + RYear$
  END IF
  DateRead$ = LEFT$(UBSchlumHHRec(1).ReadDate, 2) + "/" + MID$(UBSchlumHHRec(1).ReadDate, 3, 2) + ReadYear$
  ReadDate = Date2Num(DateRead$)
  
  IF UBCustRec(1).LocMeters(MeterSlot).ReadFlag = "Y" THEN
    UBCustRec(1).LocMeters(MeterSlot).CurRead = CurReading#
    UBCustRec(1).LocMeters(MeterSlot).CurDate = ReadDate
  ELSE
    UBCustRec(1).LocMeters(MeterSlot).PrevRead = UBCustRec(1).LocMeters(MeterSlot).CurRead
    UBCustRec(1).LocMeters(MeterSlot).PastDate = UBCustRec(1).LocMeters(MeterSlot).CurDate
    UBCustRec(1).LocMeters(MeterSlot).ReadFlag = "Y"
    UBCustRec(1).LocMeters(MeterSlot).CurDate = ReadDate
    UBCustRec(1).LocMeters(MeterSlot).CurRead = CurReading#
  END IF
  
  c1$ = QPTrim$(UBCustRec(1).HHMSG1)
  s1$ = QPTrim$(UBSchlumHHRec(1).Notes1)
  IF s1$ <> c1$ THEN
    UBCustRec(1).NewNotes = True
    UBCustRec(1).HHMSG1 = s1$
  END IF
  c1$ = QPTrim$(UBCustRec(1).HHMSG2)
  s1$ = QPTrim$(UBSchlumHHRec(1).Notes2)
  IF s1$ <> c1$ THEN
    UBCustRec(1).NewNotes = True
    UBCustRec(1).HHMSG2 = s1$
  END IF
  c1$ = QPTrim$(UBCustRec(1).HHMSG3)
  s1$ = QPTrim$(UBSchlumHHRec(1).Notes3)
  IF s1$ <> c1$ THEN
    UBCustRec(1).NewNotes = True
    UBCustRec(1).HHMSG3 = s1$
  END IF

  PUT UBFile, PRec&, UBCustRec(1)

RETURN



END SUB

