  DEFINT A-Z
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
  REDIM UBSetupRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetupRec(), UBSetUpLen
  
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
  NumberOfRoutes = 0
  
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
      ' First Check For Valid Interrogator Number and Routes
      IF VAL(Form$(1, 0)) < 1 THEN
        Frm(1).FldNo = 1
        Action = 1
        Done = False
      ELSE
        IF NumberOfRoutes > 0 THEN
          INTERROGATOR = VAL(Form$(1, 0))
          COLOR 11
          LOCATE 12, 18
          PRINT "LOADING THESE ROUTES !!!                      "
          GOSUB EFILENAM1
          ' Set Route Number Values
          FOR Cnt = 1 TO NumberOfRoutes
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
            Prec& = IdxBuff(Cnt!).RecNum
            IF Prec& > 0 THEN
              GET UBFile, Prec&, UBCustRec(1)
              FOR BookCnt = 1 TO NumberOfRoutes
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
  
  Account$ = STR$(Prec&)
  WHILE MtrCnt! < 8
    IF (ASC(UBCustRec(1).LocMeters(MtrCnt!).MTRType) > 32) THEN
      'IF UBCustRec(1).LocMeters(MtrCnt!).MTRType = "C" OR UBCustRec(1).LocMeters(MtrCnt!).MTRType = "W" OR UBCustRec(1).LocMeters(MtrCnt!).MTRType = "T" OR UBCustRec(1).LocMeters(MtrCnt!).MTRType = "S" THEN
      SELECT CASE UBCustRec(1).LocMeters(MtrCnt!).MTRType
      CASE "C", "S", "W", "T", "E", "D"         'here dale
        IF (UBCustRec(1).LocMeters(MtrCnt!).MTRType = "T" AND VAL(QPTrim$(UBCustRec(1).LocMeters(MtrCnt!).MTRNUM)) = 0) THEN
          GOTO ESkipEm
        ELSE
          ' Determine Sensus Meter Type
          IF UBCustRec(1).LocMeters(MtrCnt!).MTRType = "T" THEN SensusType$ = "B" ELSE SensusType$ = "M"
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
          UBSenRdRec(1).ServAddress = UBCustRec(1).SERVADDR
          UBSenRdRec(1).MeterID = MeterID$
          UBSenRdRec(1).LowRead = ILowRead$
          UBSenRdRec(1).HighRead = IHighRead$
          UBSenRdRec(1).Account = Account$ + "-" + QPTrim$(STR$(MtrCnt!))
          UBSenRdRec(1).SensusType = SensusType$
          UBSenRdRec(1).CustLastNAME = UBCustRec(1).CUSTNAME
          UBSenRdRec(1).Message = UBCustRec(1).HHMSG1
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
  NumberOfRoutes = 0
  
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
          Prec& = VAL(QPTrim$(UBSenGetRdRec(1).Account))
          IF Prec& > 0 THEN
            GET UBFile, Prec&, UBCustRec(1)
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
  PUT UBFile, Prec&, UBCustRec(1)
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
  
  REDIM UBSetupRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetupRec(), UBSetUpLen
  
  TownName$ = UBSetupRec(1).UTILNAME
  
  IF INSTR(TownName$, "BEECH MOUNTAIN") THEN
    BeechFlag = True
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
          IF LEN(QPTrim$(UBCustRec(1).LocMeters(zz).MTRType)) > 0 THEN
            IF UBCustRec(1).LocMeters(zz).UseCnt > 0 AND UBCustRec(1).LocMeters(zz).ReadFlag <> "Y" THEN
              DidEM = True
              UBCustRec(1).LocMeters(zz).PrevRead = UBCustRec(1).LocMeters(zz).CurRead
              UBCustRec(1).LocMeters(zz).PastDate = UBCustRec(1).LocMeters(zz).CurDate
              UBCustRec(1).LocMeters(zz).ReadFlag = "Y"
              IF BeechFlag THEN
                AvgUse& = 0
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
  REDIM UBSetupRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetupRec(), UBSetUpLen
  
  Space40$ = STRING$(40, 32)
  
  TownName$ = UBSetupRec(1).UTILNAME
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
  
  IF UBSetUpLen = 0 THEN
    BEEP
    CursorOff
    DisplayHHScrn "UBNOHH"
    WaitForAction
    CLOSE
    EXIT SUB
  END IF
  
  SELECT CASE UBSetupRec(1).HHDEVICE
  CASE "H", "S", "E", "P", "C", "D", "M", "T", "L", "I"
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
  HighVar = UBSetupRec(1).HighRead
  LowVar = UBSetupRec(1).LowRead
  
  SELECT CASE UBSetupRec(1).HHDEVICE
  CASE "S"
    GOSUB Sensus
  CASE "E"
    ESensus
  CASE "P"
    IF UBSetupRec(1).HHDEVICE = "P" THEN
      PC2000flag$ = "N"
      GOSUB PC3000
    END IF
  CASE "D"
    GOSUB DataGeneral
  CASE "C"
    GOSUB SysCom
  CASE "M"
    IF UBSetupRec(1).HHDEVICE = "M" THEN
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
  NumberOfRoutes = 0
  
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
      
      ' First Check For Valid Interrogator Number and Routes
      IF VAL(Form$(1, 0)) < 1 THEN
        Frm(1).FldNo = 1
        Action = 1
        Done = False
      ELSE
        IF NumberOfRoutes > 0 THEN
          INTERROGATOR = VAL(Form$(1, 0))
          COLOR 11
          LOCATE 12, 18
          PRINT "LOADING THESE ROUTES !!!                      "
          GOSUB FILENAM1
          ' Set Route Number Values
          FOR Cnt = 1 TO NumberOfRoutes
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
            Prec& = IdxBuff(Cnt!).RecNum
            IF Prec& > 0 THEN
              GET UBFile, Prec&, UBCustRec(1)
              FOR BookCnt = 1 TO NumberOfRoutes
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
  Account$ = STR$(Prec&)
  Account$ = LEFT$(Account$, 6) + "-" + RIGHT$(STR$(MtrCnt!), 1)
  
  WHILE MtrCnt! < 8
    IF (ASC(UBCustRec(1).LocMeters(MtrCnt!).MTRType) > 32) THEN
      IF UBCustRec(1).LocMeters(MtrCnt!).MTRType = "C" OR UBCustRec(1).LocMeters(MtrCnt!).MTRType = "W" OR UBCustRec(1).LocMeters(MtrCnt!).MTRType = "T" OR UBCustRec(1).LocMeters(MtrCnt!).MTRType = "S" THEN
        'here dale
        IF (UBCustRec(1).LocMeters(MtrCnt!).MTRType = "T" AND VAL(QPTrim$(UBCustRec(1).LocMeters(MtrCnt!).MTRNUM)) = 0) THEN
          GOTO SkipEm
        ELSE
          ' Determine Sensus Meter Type
          IF UBCustRec(1).LocMeters(MtrCnt!).MTRType = "T" THEN SensusType$ = "B" ELSE SensusType$ = "M"
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
            UBGilSenRdRec(1).CustLastNAME = UBCustRec(1).CUSTNAME
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
  NumberOfRoutes = 0
  
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
            Prec& = VAL(QPTrim$(UBGilSenGetRdRec(1).Account))
          ELSE
            GET UBSenGetRdFile, Cnt!, UBSenGetRdRec(1)
            Prec& = VAL(QPTrim$(UBSenGetRdRec(1).Account))
          END IF
          
          IF Prec& > 0 THEN
            GET UBFile, Prec&, UBCustRec(1)
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
    MeterReadDate$ = LEFT$(UBSenGetRdRec(1).DateRead, 2) + "/" + MID$(UBSenGetRdRec(1).DateRead, 3, 2) + "/" + RIGHT$(Form$(2, 0), 2)
  END IF
  
  IF UBCustRec(1).LocMeters(MeterRecord).MTRType = "T" THEN

    IF GrayFlag THEN
      CurReading# = CurReading# * 100
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
  PUT UBFile, Prec&, UBCustRec(1)
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
  NumberOfRoutes = 0
  
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
      IF GeorgFlag THEN
        NumberOfRoutes = 100
      END IF
      IF NumberOfRoutes > 0 THEN
        COLOR 11
        LOCATE 12, 18
        PRINT "LOADING THESE ROUTES !!!                      "
        GOSUB FILENAM1
        ' Set Route Number Values
        IF GeorgFlag THEN
          FOR Cnt = 1 TO NumberOfRoutes
            Route(Cnt) = Cnt - 1
          NEXT
        ELSE
          FOR Cnt = 1 TO NumberOfRoutes
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
          IF Prec& > 0 THEN
            GET UBFile, Prec&, UBCustRec(1)
            DisplayAcct$ = "Processing Location #" + UBCustRec(1).BOOK + UBCustRec(1).SEQNUMB
            PrintHelp DisplayAcct$
            FOR BookCnt = 1 TO NumberOfRoutes
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
        Prec& = VAL(QPTrim$(UBPC3000GetRDRec(1).Account))
        IF Prec& > 0 THEN
          GET UBFile, Prec&, UBCustRec(1)
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
  PUT UBFile, Prec&, UBCustRec(1)
  
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
  NumberOfRoutes = 0
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
      IF GeorgFlag THEN
        NumberOfRoutes = 100
      END IF
      
      IF NumberOfRoutes > 0 THEN
        COLOR 11
        LOCATE 12, 18
        PRINT "LOADING THESE ROUTES !!!                      "
        GOSUB FILENAM1
        ' Set Route Number Values
        
        IF GeorgFlag THEN
          FOR Cnt = 1 TO NumberOfRoutes
            Route(Cnt) = Cnt - 1
          NEXT
        ELSE
          FOR Cnt = 1 TO NumberOfRoutes
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
          Prec& = IdxBuff(Cnt!).RecNum
          IF Prec& > 0 THEN
            GET UBFile, Prec&, UBCustRec(1)
            DisplayAcct$ = "Processing Location #" + UBCustRec(1).BOOK + UBCustRec(1).SEQNUMB
            PrintHelp DisplayAcct$
            FOR BookCnt = 1 TO NumberOfRoutes
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
    Account$ = STR$(Prec&)
    Account$ = LEFT$(Account$, 6) + "-" + RIGHT$(STR$(MtrCnt!), 1)
    IF ASC(UBCustRec(1).LocMeters(MtrCnt!).MTRType) > 32 THEN
      MTRType$ = UBCustRec(1).LocMeters(MtrCnt!).MTRType
      IF MTRType$ = "W" OR MTRType$ = "S" OR MTRType$ = "C" OR MTRType$ = "E" OR MTRType$ = "D" OR MTRType$ = "G" THEN
        SELECT CASE WhatTypes$
        CASE "W"                'water/sewer
          IF MTRType$ = "W" OR MTRType$ = "S" OR MTRType$ = "C" THEN
            MeterType$ = "W"
            MeterOK = True
          END IF
        CASE "E"                'electric & demand elec.
          IF MTRType$ = "E" OR MTRType$ = "D" THEN
            MeterOK = True
          END IF
        CASE "G"                'gas
          IF MTRType$ = "G" THEN
            MeterType$ = "G"
            MeterOK = True
          END IF
        CASE "A", " "           'all meters
          IF MTRType$ = "W" OR MTRType$ = "S" OR MTRType$ = "C" THEN
            MeterType$ = "W"
          END IF
          IF MTRType$ = "E" OR MTRType$ = "D" THEN
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
    Prec& = VAL(QPTrim$(UBDGGetRDRec(1).Account))
    IF Prec& > 0 THEN
      GET UBFile, Prec&, UBCustRec(1)
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
    IF UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y" THEN
      UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
      UBCustRec(1).LocMeters(MeterRecord).CurDate = Date2Num%(UBDGGetRDRec(1).Date)
    ELSE
      UBCustRec(1).LocMeters(MeterRecord).PrevRead = UBCustRec(1).LocMeters(MeterRecord).CurRead
      UBCustRec(1).LocMeters(MeterRecord).PastDate = UBCustRec(1).LocMeters(MeterRecord).CurDate
      UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y"
      UBCustRec(1).LocMeters(MeterRecord).CurDate = Date2Num%(UBDGGetRDRec(1).Date)
      UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
    END IF
    PUT UBFile, Prec&, UBCustRec(1)
  END IF
  RETURN
  
SysCom:
  RETURN
  
OpenCustFile:
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
RETURN
  
END SUB

SUB InactiveConsReport
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  MaxLines = 52
  FF$ = CHR$(12)
  
  'Open Report File
  ReportFile$ = "UBINCONS.RPT"
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  'Open the Utility Setup File to Grab Meter List Order (Seq or Loc)
  REDIM UBSetupRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetupRec(), UBSetUpLen
  
  CUSTNAME$ = SPACE$(25)
  SAddr$ = SPACE$(25)
  
  UBCustFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCustFile LEN = UBCustRecLen
  NumOfCust& = LOF(UBCustFile) / UBCustRecLen
  'Cnt& = 1
  
  GOSUB PrintInactHeading
  ShowProcessingScrn "Reading Meter Information"
  
  FOR LCnt& = 1 TO NumOfCust&
    GET #UBCustFile, LCnt&, UBCustRec(1)
    IF UBCustRec(1).DelFlag = 0 THEN
      IF INSTR(UBCustRec(1).HHMSG1, "NOREAD") = 0 THEN
        LSET CUSTNAME$ = UBCustRec(1).CUSTNAME
        LSET SAddr$ = LEFT$(UBCustRec(1).SERVADDR, 30)
        IF UBCustRec(1).Status = "I" THEN '<> "A") AND (UBCustRec(1).Status <> "F") THEN
          DidOne = False
          FOR MtrCnt& = 1 TO 7  'find last active meter
            TempRev$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt&).MTRType)
            IF LEN(TempRev$) > 0 THEN
              GOSUB GetInactMtrTypePrint
              IF PrintMrtFlag THEN
                Multi# = UBCustRec(1).LocMeters(MtrCnt&).MtrMulti
                IF Multi# = 0 THEN Multi# = 1
                IF UBCustRec(1).LocMeters(MtrCnt&).CurRead < 0 OR UBCustRec(1).LocMeters(MtrCnt&).PrevRead < 0 THEN
                  MeterConsp# = 0
                ELSE
                  MeterConsp# = UBCustRec(1).LocMeters(MtrCnt&).CurRead - UBCustRec(1).LocMeters(MtrCnt&).PrevRead
                END IF
                IF MeterConsp# < 0 THEN
                  MaxMeterAmt& = 10& ^ (LEN(STR$(UBCustRec(1).LocMeters(MtrCnt&).PrevRead)) - 1)
                  MeterConsp# = (MaxMeterAmt& - UBCustRec(1).LocMeters(MtrCnt&).PrevRead) + UBCustRec(1).LocMeters(MtrCnt&).CurRead
                END IF
                MeterConsp# = Round#(MeterConsp# * Multi#)
                IF MeterConsp# > 0 THEN
                  IF NOT DidOne = True THEN
                    PRINT #RptHandle, " "; UBCustRec(1).Status; "   "; UBCustRec(1).BOOK; "-"; UBCustRec(1).SEQNUMB; USING "  #####"; LCnt&;
                    PRINT #RptHandle, TAB(25); CUSTNAME$; TAB(50); SAddr$
                    DidOne = True
                    PrintedOne = True
                  END IF
                  PRINT #RptHandle, UBCustRec(1).LocMeters(MtrCnt&).MTRNUM;
                  PRINT #RptHandle, TAB(14); MeterType$;
                  PRINT #RptHandle, TAB(24); USING "#####"; Multi#;
                  PRINT #RptHandle, TAB(31); USING "##########"; UBCustRec(1).LocMeters(MtrCnt&).CurRead;
                  PRINT #RptHandle, TAB(42); USING "##########"; UBCustRec(1).LocMeters(MtrCnt&).PrevRead;
                  PRINT #RptHandle, TAB(53); USING "##########"; MeterConsp#;
                  PRINT #RptHandle, TAB(67); Num2Date$(UBCustRec(1).LocMeters(MtrCnt&).CurDate)
                  LineCnt = LineCnt + 1
                END IF
              END IF
            END IF
          NEXT MtrCnt&
          DidOne = False
          IF PrintedOne THEN
            PRINT #RptHandle, STRING$(79, "-")
            LineCnt = LineCnt + 1
            PrintedOne = False
          END IF
        END IF
      END IF
    END IF
    IF LineCnt >= MaxLines THEN
      PRINT #RptHandle, FF$
      GOSUB PrintInactHeading
    END IF
    ShowPctComp LCnt&, NumOfCust&
  NEXT
  
  PRINT #RptHandle, FF$
  CLOSE
  
  Header$ = "Inactive Consumption Report"
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  EXIT SUB
  
PrintInactHeading:
  page = page + 1
  PRINT #RptHandle, TAB(30); "Inactive Consumption Report"
  PRINT #RptHandle, "Date: "; DATE$; TAB(70); "Page #"; page
  PRINT #RptHandle, ""
  PRINT #RptHandle, "Status  Loc.   Acount     Customer Name"; TAB(51); "Service Address"
  PRINT #RptHandle, " Mtr No.   Mtr Type    Multi     Current   Previous    Consump     Read Date"
  PRINT #RptHandle, STRING$(80, "=")
  LineCnt = 7
RETURN
  
GetInactMtrTypePrint:
  PrintMrtFlag = False
  SELECT CASE UBCustRec(1).LocMeters(MtrCnt&).MTRType
  CASE "C"
    MeterType$ = "Wat/Sew"
    PrintMrtFlag = True
  CASE "W"
    MeterType$ = "Water"
    PrintMrtFlag = True
  CASE "S"
    MeterType$ = "Sewer"
    PrintMrtFlag = True
  CASE "T"
    MeterType$ = "T-Read"
    PrintMrtFlag = True
  CASE "E", "D"
    MeterType$ = "Elec"
    PrintMrtFlag = True
  CASE "G"
    MeterType$ = "Gas"
    PrintMrtFlag = True
  CASE ELSE
    MeterType$ = "Undef"
    PrintMrtFlag = True
  END SELECT
RETURN

END SUB

SUB PrintMeterList
  
  REDIM TempScrn(0)
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  SHARED Choice$()
  DevFld = 3
  
ListMainBody:
  
  'Help$ = "Print Meter Reading List"
  LibName$ = "UB"
  ScrnName$ = "UBMTRLST"
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  FirstCust = True
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  REDIM Choice$(3, 0)
  
  Choice$(0, 0) = STR$(DevFld)  'Choices for Output Device
  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn#1"
  Choice$(3, 0) = "Prn#2"
  LineCnt = 0
  Action = 1
  
  BlockClear
  DisplayUBScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    '--Check for Key presses
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      ' Check for Proper Order of Route Questions
      IF VAL(Form$(1, 0)) = 0 THEN
        BegRoute = 1
        EndRoute = 99
        inputting = True
      ELSE
        BegRoute = VAL(Form$(1, 0))
        EndRoute = VAL(Form$(2, 0))
      END IF
      
      IF BegRoute > EndRoute THEN
        COLOR 12
        LOCATE 13, 25: PRINT "Invalid Route Order : Please Redo": BEEP
        WaitForAction
        LOCATE 13, 25: PRINT STRING$(35, " ")
        Form$(1, 0) = "": Form$(2, 0) = "": Frm(1).FldNo = 1
        Action = 1
      ELSE
        inputting = True
      END IF
      
    CASE EscKey
      EXIT SUB
      
    END SELECT
    
  LOOP WHILE inputting = False
  
  ' Print Report
  'COLOR 7
  'LOCATE 25, 42: PRINT STRING$(39, 219);
  
  IF LEN(LTRIM$(RTRIM$(Form$(DevFld, 0)))) = 0 THEN
    dev$ = "S"
    LPTNo = 1
  ELSE
    dev$ = LEFT$(Form$(DevFld, 0), 1)
    LPTNo = VAL(RIGHT$(RTRIM$(Form$(DevFld, 0)), 1))
  END IF
  
  'Free Up Some Memory
  ERASE Form$, Fld, Choice$, Frm
  
  ShowProcessingScrn "Processing Routes"
  
  'Open Report File
  ReportFile$ = "WBMTRLST.RPT"
  
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  'Open the Utility Setup File to Grab Meter List Order (Seq or Loc)
  REDIM UBSetupRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetupRec(), UBSetUpLen
  
  IF UBSetupRec(1).UseSeq = "Y" THEN
    SeqFlag = True
    MakeSequenceIndex "Sequence Number"
    IdxName$ = "UBTEMP.IDX"
  ELSE
    IdxName$ = "UBCUSTBK.IDX"
  END IF
  
  NumOfRecs = FileSize&(IdxName$) \ 4
  IdxNumOfRecs = NumOfRecs
  REDIM IndexArray(1 TO NumOfRecs) AS UBCustIndexRecType
  FGetAH IdxName$, IndexArray(1), 4, NumOfRecs
  
  DO
    IF Cnt < 1 THEN Cnt = 1     ' Do Not Allow to Fall Below 1
    inputting = False           ' Set Edit Finish to No
    Prec& = IndexArray(Cnt).RecNum
    IF NOT Prec& = 0 THEN
      GOSUB GetCustRecNoPrint
      IF process THEN
        GOSUB PrintLine
      END IF
    END IF
    Cnt = Cnt + 1
    ShowPctComp Cnt, NumOfRecs
  LOOP UNTIL Cnt > NumOfRecs
  
  CLOSE
  
  SELECT CASE dev$
  CASE "S"
    EntryPoint = 2
  CASE "P"
    EntryPoint = 5
  END SELECT
  
  Header$ = "Customer Meter Listing Report"
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  EXIT SUB
  
  'Print the Meter Reading Here *********************************************
  
PrintLine:
  IF LineCnt = 0 THEN GOSUB PrintHeading
  
  'Help$ = "Process Location Record #" + STR$(Cnt) + " of " + STR$(IdxNumOfRecs)
  'PrintHelp Help$
  
  MtrCnt = 0
  ValidCustomer = False
  DO
    MtrCnt = MtrCnt + 1         'Check For Meter This Customer
    TempRev$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MTRType)
    IF LEN(TempRev$) <> 0 THEN ValidCustomer = True: EXIT DO
  LOOP UNTIL MtrCnt = 7
  IF ValidCustomer = False THEN RETURN
  
  IF LineCnt >= 53 THEN
    PRINT #RptHandle, CHR$(12);
    GOSUB PrintHeading
  END IF
  
  GOSUB GetMeterStatusPrint
  
  PRINT #RptHandle, LEFT$(UBCustRec(1).CUSTNAME, 30);
  IF UBSetupRec(1).UseSeq = "Y" THEN
    IF UBCustRec(1).SEQ < 0 THEN UBCustRec(1).SEQ = 0
    PRINT #RptHandle, TAB(32); USING "######"; UBCustRec(1).SEQ;
  END IF
  PRINT #RptHandle, TAB(40); LEFT$(UBCustRec(1).SERVADDR, 28);
  PRINT #RptHandle, TAB(70); MeterStatus$
  LineCnt = LineCnt + 1
  
  FOR MtrCnt = 1 TO 7           'find last active meter
    TempRev$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MTRType)
    IF LEN(TempRev$) <> 0 THEN
      GOSUB GetMeterTypePrint
      PRINT #RptHandle, UBCustRec(1).BOOK; "-"; UBCustRec(1).SEQNUMB;
      PRINT #RptHandle, TAB(15); QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MTRNUM);
      PRINT #RptHandle, TAB(35); MeterType$;
      PRINT #RptHandle, TAB(55); USING "##########"; UBCustRec(1).LocMeters(MtrCnt).CurRead;
      PRINT #RptHandle, TAB(68); "___________"
      LineCnt = LineCnt + 1
    END IF
  NEXT MtrCnt
  PRINT #RptHandle, STRING$(79, "-"): LineCnt = LineCnt + 1
  RETURN
  ' END OF PRINT ROUTINE *****************************************
  
PrintHeading:
  page = page + 1
  PRINT #RptHandle, TAB(27); "Meter Reading Listing Report"; TAB(65); "Date: "; DATE$
  PRINT #RptHandle, "Beginning Route: "; BegRoute
  PRINT #RptHandle, "   Ending Route: "; EndRoute; TAB(65); "Page #"; page
  PRINT #RptHandle,
  PRINT #RptHandle, "Customer Name";
  IF UBSetupRec(1).UseSeq = "Y" THEN
    PRINT #RptHandle, TAB(32); "Seq #";
  END IF
  PRINT #RptHandle, TAB(40); "Service Address"; TAB(70); "Status"
  PRINT #RptHandle, "Location"; TAB(15); "Meter Number"; TAB(35); "Mtr Type"; TAB(55); "Cur Read"; TAB(70); "New Read"
  PRINT #RptHandle, STRING$(80, "=")
  LineCnt = 5
  RETURN
  
GetCustRecNoPrint:
  RecNo& = Prec&
  FOpenS "UBCUST.DAT", L2Handle 'open data file
  FGetRTA L2Handle, UBCustRec(1), RecNo&, UBCustRecLen
  FClose L2Handle
  IF FirstCust THEN
    FirstCust = False
    WhatBook = VAL(UBCustRec(1).BOOK)
  END IF
  
  IF (UBCustRec(1).DelFlag <> 0) OR INSTR(UBCustRec(1).HHMSG1, "NOREAD") > 0 THEN
    process = False
    RETURN
  END IF
  
  IF VAL(UBCustRec(1).BOOK) >= BegRoute AND VAL(UBCustRec(1).BOOK) <= EndRoute THEN
    IF NOT SeqFlag THEN
      IF VAL(UBCustRec(1).BOOK) <> WhatBook THEN
        PRINT #RptHandle, CHR$(12);
        DoHeaderFlag = True
        LineCnt = 0
        WhatBook = VAL(UBCustRec(1).BOOK)
      END IF
    END IF
    process = True
  ELSE
    process = False
    RETURN
  END IF
  
  IF RecNo& <= 0 THEN
    process = False
  END IF
  
  RETURN
  
GetMeterTypePrint:
  SELECT CASE UBCustRec(1).LocMeters(MtrCnt).MTRType
  CASE "C"
    MeterType$ = "Water/Sewer"
  CASE "W"
    MeterType$ = "Water Only"
  CASE "S"
    MeterType$ = "Sewer Only"
  CASE "T"
    MeterType$ = "Touch Read"
  CASE "E"
    MeterType$ = "Electric"
  CASE "D"
    MeterType$ = "Demand"
  CASE "G"
    MeterType$ = "Gas"
  CASE ELSE
    MeterType$ = "Undefined"
  END SELECT
  RETURN
  
GetMeterStatusPrint:
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
  
END SUB

SUB PrintMeterNotes
  REDIM TempScrn(0)
  
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  MaxLines = 40
  FF$ = CHR$(12)
  
  LibName$ = "UB"
  ScrnName$ = "UBMTRRPT"
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  Action = 1
  FirstTime = True
  BlockClear
  DisplayUBScrn ScrnName$
  QPrintRC "Notes  ", 6, 38, -1
  QPrintRC "New Notes Only:", 12, 24, 11
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    GOSUB CheckField3
    
    IF FirstTime THEN
      LSET Form$(1, 0) = "1"
      LSET Form$(2, 0) = "99"
      LSET Form$(3, 0) = "N"
      'Fld3Was$ = QPTrim$(Form$(3, 0))
      Fld(3).Protected = False
      FirstTime = False
      Action = 1
    END IF
    
    '--Check for Key presses
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      ' Check for Proper Order of Route Questions
      NewFlag = Form$(3, 0) = "Y"
      IF QPValI(Form$(1, 0)) = 0 THEN
        BegRoute = 1
        EndRoute = 99
        inputting = True
      ELSE
        BegRoute = QPValI(Form$(1, 0))
        EndRoute = QPValI(Form$(2, 0))
      END IF
      
      IF BegRoute > EndRoute THEN
        COLOR 12
        LOCATE 13, 25
        PRINT "Invalid Route Order : Please Redo"
        BEEP
        WaitForAction
        LOCATE 13, 25
        PRINT STRING$(35, " ")
        Form$(1, 0) = ""
        Form$(2, 0) = ""
        Frm(1).FldNo = 1
        Action = 1
      ELSE
        inputting = True
      END IF
      
    CASE EscKey
      EXIT SUB
      
    END SELECT
    
  LOOP WHILE inputting = False
  'Free Up Some Memory
  ERASE Form$, Fld, Frm
  
  'Open Report File
  ReportFile$ = "UBMTRNOT.RPT"
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  CUSTNAME$ = SPACE$(30)
  
  ' Location Order ********************************************************
  IdxRecLen = 4 'we are using a long integer
  IdxFileSize& = FileSize&("UBCUSTBK.IDX")
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  
  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it
  
  UBCustFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCustFile LEN = UBCustRecLen
  
  Cnt& = 1
  GOSUB PrintNotesHeading
  ShowProcessingScrn "Reading Meter Information"
  FOR LCnt& = 1 TO IdxNumOfRecs
    GET #UBCustFile, IdxBuff(LCnt&).RecNum, UBCustRec(1)
    BOOK = QPValI(UBCustRec(1).BOOK)
    IF BOOK >= BegRoute AND BOOK <= EndRoute THEN
      IF UBCustRec(1).DelFlag = 0 THEN
        IF NewFlag THEN
          IF UBCustRec(1).NewNotes <> True THEN
            GOTO SkipNotes
          ELSE
            UBCustRec(1).NewNotes = False
            PUT #UBCustFile, IdxBuff(LCnt&).RecNum, UBCustRec(1)
          END IF
        END IF
        LSET CUSTNAME$ = UBCustRec(1).CUSTNAME
        IF LEN(QPTrim$(UBCustRec(1).HHMSG1)) > 0 OR LEN(QPTrim$(UBCustRec(1).HHMSG2)) > 0 OR LEN(QPTrim$(UBCustRec(1).HHMSG3)) > 0 THEN
          PRINT #RptHandle, UBCustRec(1).BOOK; "-"; UBCustRec(1).SEQNUMB; "  "; CUSTNAME$; LEFT$(UBCustRec(1).SERVADDR, 30)
          PRINT #RptHandle, TAB(3); UBCustRec(1).HHMSG1; TAB(26); UBCustRec(1).HHMSG2; TAB(49); UBCustRec(1).HHMSG3
          PRINT #RptHandle, STRING$(79, "-")
          LineCnt = LineCnt + 3
        END IF
      END IF
    END IF
    IF LineCnt >= MaxLines THEN
      PRINT #RptHandle, FF$
      GOSUB PrintNotesHeading
    END IF
SkipNotes:
    ShowPctComp LCnt&, IdxNumOfRecs
  NEXT
  
  PRINT #RptHandle, FF$
  CLOSE
  
  Header$ = "Meter Reading Notes"
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  EXIT SUB
  
PrintNotesHeading:
  page = page + 1
  PRINT #RptHandle, TAB(30); "Meter Reading Notes"; TAB(65); "Date: "; DATE$
  PRINT #RptHandle, "Beginning Route: "; BegRoute
  PRINT #RptHandle, "   Ending Route: "; EndRoute; TAB(70); "Page #"; page
  PRINT #RptHandle, ""
  PRINT #RptHandle, "Location   Customer Name"; TAB(41); "Service Address"
  PRINT #RptHandle, TAB(3); "NOTE1"; TAB(26); "NOTE2"; TAB(49); "NOTE3"
  PRINT #RptHandle, STRING$(80, "=")
  LineCnt = 5
  RETURN
  
  
CheckField3:
  Fld3$ = QPTrim$(Form$(3, 0))
  IF LEN(Fld3$) > 0 THEN
    IF Fld3$ <> Fld3Was$ THEN
      Fld3Was$ = Fld3$
      SELECT CASE Fld3$
      CASE "Y"
        QPrintRC "  NOTE:  Will Clear New Notes Flag!  ", 13, 23, 14
      CASE ELSE
        QPrintRC "NOTE:  Will NOT Clear New Notes Flag!", 13, 23, 14
      END SELECT
    END IF
  END IF
  
  RETURN
  
END SUB

SUB PrintMeterReport
  
  REDIM TempScrn(0)
  
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  MaxLines = 52
  FF$ = CHR$(12)
  
  LibName$ = "UB"
  ScrnName$ = "UBMTRRPT"
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  'LineCnt& = 0
  
  Action = 1
  FirstTime = True
  BlockClear
  DisplayUBScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      LSET Form$(1, 0) = "1"
      LSET Form$(2, 0) = "99"
      FirstTime = False
      Action = 1
    END IF
    
    '--Check for Key presses
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key ' Check for Proper Order of Route Questions
      IF QPValI(Form$(1, 0)) = 0 THEN
        BegRoute = 1
        EndRoute = 99
        inputting = True
      ELSE
        BegRoute = QPValI(Form$(1, 0))
        EndRoute = QPValI(Form$(2, 0))
      END IF
      
      IF BegRoute > EndRoute THEN
        COLOR 12
        LOCATE 13, 25
        PRINT "Invalid Route Order : Please Redo"
        BEEP
        WaitForAction
        LOCATE 13, 25
        PRINT STRING$(35, " ")
        Form$(1, 0) = ""
        Form$(2, 0) = ""
        Frm(1).FldNo = 1
        Action = 1
      ELSE
        inputting = True
      END IF
      
    CASE EscKey
      EXIT SUB
      
    END SELECT
    
  LOOP WHILE inputting = False
  'Free Up Some Memory
  ERASE Form$, Fld, Frm
  
  'Open Report File
  ReportFile$ = "WBMTRLST.RPT"
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  'Open the Utility Setup File to Grab Meter List Order (Seq or Loc)
  REDIM UBSetupRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetupRec(), UBSetUpLen
  
  CUSTNAME$ = SPACE$(30)
  
  ' Location Order ********************************************************
  'if UBSetupRec(1).
  
  IdxRecLen = 4 'we are using a long integer
  IdxFileSize& = FileSize&("UBCUSTBK.IDX")
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  
  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it
  
  UBCustFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCustFile LEN = UBCustRecLen
  
  Cnt& = 1
  GOSUB PrintReadHeading
  ShowProcessingScrn "Reading Meter Information"
  FOR LCnt& = 1 TO IdxNumOfRecs
    GET #UBCustFile, IdxBuff(LCnt&).RecNum, UBCustRec(1)
    IF UBCustRec(1).DelFlag = 0 THEN
      IF INSTR(UBCustRec(1).HHMSG1, "NOREAD") = 0 THEN
        BOOK = QPValI(UBCustRec(1).BOOK)
        IF BOOK >= BegRoute AND BOOK <= EndRoute THEN
          LSET CUSTNAME$ = UBCustRec(1).CUSTNAME
          IF UBCustRec(1).Status > "" THEN
            PRINT #RptHandle, UBCustRec(1).BOOK; "-"; UBCustRec(1).SEQNUMB; "  "; CUSTNAME$; LEFT$(UBCustRec(1).SERVADDR, 30)
            'IF LEN(QPTrim$(UBCustRec(1).EstFlag)) > 0 THEN STOP
            LineCnt = LineCnt + 1
            FOR MtrCnt& = 1 TO 7                'find last active meter
              TempRev$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt&).MTRType)
              IF LEN(TempRev$) > 0 THEN
                GOSUB GetReadMeterTypePrint
                IF PrintMrtFlag THEN
                  '    IF MtrCnt& > 1 THEN
                  '      PRINT #RptHandle, "HERE DALE"
                  '    END IF
                  DidOne = True
                  PRINT #RptHandle, UBCustRec(1).LocMeters(MtrCnt&).MTRNUM;
                  PRINT #RptHandle, TAB(14); MeterType$;
                  Multi# = UBCustRec(1).LocMeters(MtrCnt&).MtrMulti
                  IF Multi# = 0 THEN Multi# = 1
                  PRINT #RptHandle, TAB(24); USING "#####"; Multi#;
                  PRINT #RptHandle, TAB(31); USING "##########"; UBCustRec(1).LocMeters(MtrCnt&).CurRead;
                  PRINT #RptHandle, TAB(42); USING "##########"; UBCustRec(1).LocMeters(MtrCnt&).PrevRead;
                  IF UBCustRec(1).LocMeters(MtrCnt&).CurRead < 0 OR UBCustRec(1).LocMeters(MtrCnt&).PrevRead < 0 THEN
                    MeterConsp# = 0
                  ELSE
                    MeterConsp# = UBCustRec(1).LocMeters(MtrCnt&).CurRead - UBCustRec(1).LocMeters(MtrCnt&).PrevRead
                  END IF
                  IF MeterConsp# < 0 THEN
                    MaxMeterAmt& = 10& ^ (LEN(STR$(UBCustRec(1).LocMeters(MtrCnt&).PrevRead)) - 1)
                    MeterConsp# = (MaxMeterAmt& - UBCustRec(1).LocMeters(MtrCnt&).PrevRead) + UBCustRec(1).LocMeters(MtrCnt&).CurRead
                  END IF
                  
                  MeterConsp# = Round#(MeterConsp# * Multi#)
                  
                  PRINT #RptHandle, TAB(53); USING "##########"; MeterConsp#;
                  
                  IF UBCustRec(1).LocMeters(MtrCnt&).ReadFlag <> "Y" THEN
                    PRINT #RptHandle, TAB(67); "UNREAD"
                  ELSE
                    PRINT #RptHandle, TAB(67); Num2Date$(UBCustRec(1).LocMeters(MtrCnt&).CurDate)
                  END IF
                  
                  'RINT #RptHandle, TAB(67); UBCustRec(1).LocMeters(MtrCnt&).MtrMulti
                  'PRINT #RptHandle, " "; UBCustRec(1).LocMeters(MtrCnt&).ReadFlag
                  LineCnt = LineCnt + 1
                END IF
              END IF
            NEXT MtrCnt&
            IF NOT DidOne THEN
              PRINT #RptHandle, TAB(14); "NO METERED SERVICE"
              LineCnt = LineCnt + 1
            END IF
            DidOne = False
            PRINT #RptHandle, STRING$(79, "-")
            LineCnt = LineCnt + 1
          END IF
        END IF
      END IF
    END IF
    IF LineCnt >= MaxLines THEN
      PRINT #RptHandle, FF$
      GOSUB PrintReadHeading
    END IF
    ShowPctComp LCnt&, IdxNumOfRecs
  NEXT
  
  PRINT #RptHandle, FF$
  CLOSE
  
  Header$ = "Customer Meter Reading Report"

  ERASE IdxBuff
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  EXIT SUB
  
PrintReadHeading:
  page = page + 1
  PRINT #RptHandle, TAB(30); "Meter Reading Report"; TAB(65); "Date: "; DATE$
  PRINT #RptHandle, "Beginning Route: "; BegRoute
  PRINT #RptHandle, "   Ending Route: "; EndRoute; TAB(70); "Page #"; page
  PRINT #RptHandle, ""
  PRINT #RptHandle, "Location   Customer Name"; TAB(41); "Service Address"
  PRINT #RptHandle, " Mtr No.   Mtr Type    Multi     Current   Previous    Consump     Read Date"
  PRINT #RptHandle, STRING$(80, "=")
  LineCnt = 7
  RETURN
  
GetReadMeterTypePrint:
  PrintMrtFlag = False
  SELECT CASE UBCustRec(1).LocMeters(MtrCnt&).MTRType
  CASE "C"
    MeterType$ = "Wat/Sew"
    PrintMrtFlag = True
  CASE "W"
    MeterType$ = "Water"
    PrintMrtFlag = True
  CASE "S"
    MeterType$ = "Sewer"
    PrintMrtFlag = True
  CASE "T"
    MeterType$ = "T-Read"
    PrintMrtFlag = True
  CASE "E", "D"
    MeterType$ = "Elec"
    PrintMrtFlag = True
  CASE "G"
    MeterType$ = "Gas"
    PrintMrtFlag = True
  CASE ELSE
    MeterType$ = "Undef"
    PrintMrtFlag = True
  END SELECT
  RETURN
  
  
END SUB

SUB PrintMeterSheets
  BlockClear
  
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBSetupRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetupRec(), UBSetUpLen
  
  IF UBSetupRec(1).UseSeq = "Y" THEN
    MakeSequenceIndex "Sequence Number"
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize&("UBTEMP.IDX")
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
    FGetAH "UBTEMP.IDX", IdxBuff(1), 4, IdxNumOfRecs
    UseSeq = True
  ELSE
    ShowProcessingScrn "Scanning Accounts"
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize&("UBCUSTBK.IDX")
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
    FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs  'load it
    UBCustFile = FREEFILE
    OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCustFile LEN = UBCustRecLen
    GET UBCustFile, IdxBuff(IdxNumOfRecs).RecNum, UBCustRec(1)
    Book2$ = UBCustRec(1).BOOK
    SeqNumb2$ = UBCustRec(1).SEQNUMB
    FOR Cnt& = 1 TO IdxNumOfRecs
      GET UBCustFile, IdxBuff(Cnt&).RecNum, UBCustRec(1)
      IF LEN(QPTrim$(UBCustRec(1).BOOK)) > 0 THEN
        Book1$ = UBCustRec(1).BOOK
        SeqNumb1$ = UBCustRec(1).SEQNUMB
        EXIT FOR
      END IF
      ShowPctComp Cnt&, IdxNumOfRecs
    NEXT
    CLOSE UBCustFile
    UseSeq = False
  END IF
  
  ReportFile$ = "UBMTRSHT.RPT"
  REDIM TempScrn(0)
  
  IF UseSeq THEN
    GOTO SeqJump
  END IF
  
  
  LibName$ = "UB"
  ScrnName$ = "UBMTRSHT"
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  Action = 1
  
  FirstTime = True
  BlockClear
  DisplayUBScrn ScrnName$
  
  'Frm(1).StayOnField = True
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      LSET Form$(1, 0) = Book1$
      LSET Form$(2, 0) = SeqNumb1$
      LSET Form$(3, 0) = Book2$
      LSET Form$(4, 0) = SeqNumb2$
      Action = 1
    END IF
    
    IF Frm(1).FldNo = 3 AND Frm(1).PrevFld < 3 THEN             '>2
      In1 = True
      GOSUB CheckBookSequence
    ELSEIF Frm(1).FldNo > 4 AND Frm(1).PrevFld = 4 THEN         '>2
      In1 = False
      GOSUB CheckBookSequence
    END IF
    
    SELECT CASE Frm(1).KeyCode  '--Check for Key presses
      
    CASE F10Key
      'Check for valid Order of Route Questions
      In1 = True
      GOSUB CheckBookSequence
      IF OkFlag THEN
        In1 = False
        GOSUB CheckBookSequence
      END IF
      IF OkFlag THEN
        Book1& = QPValL(Form$(1, 0))
        Sequ1& = QPValL(Form$(2, 0))
        Book2& = QPValL(Form$(3, 0))
        Sequ2& = QPValL(Form$(4, 0))
        Done = True
      END IF
      
    CASE F5KEY
      GOSUB DoSheetMask
      
    CASE EscKey
      GOTO ExitPrintSheets
      
    END SELECT
    
  LOOP UNTIL Done
  
  'Free Up Some Memory
  ERASE Form$, Fld, Frm
  
SeqJump:
  
  ' Location Order ********************************************************
  
  UBCustFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCustFile LEN = UBCustRecLen
  
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle     'Open Report File
  
  ShowProcessingScrn "Reading Meter Information"
  
  FOR LCnt& = 1 TO IdxNumOfRecs
    GET #UBCustFile, IdxBuff(LCnt&).RecNum, UBCustRec(1)
    IF UBCustRec(1).DelFlag <> 0 THEN
      GOTO SkipSheet
    END IF
    
    IF UseSeq = False THEN
      BOOK = QPValL(UBCustRec(1).BOOK)
      Sequ& = QPValL(UBCustRec(1).SEQNUMB)
      IF BOOK < Book1& OR BOOK > Book2& THEN
        GOTO SkipSheet
      END IF
      IF Sequ& < Sequ1& OR Sequ& > Sequ2& THEN
        GOTO SkipSheet
      END IF
      GOSUB PrintEm
    ELSE
      GOSUB PrintEm
    END IF
    
SkipSheet:
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
    ShowPctComp LCnt&, IdxNumOfRecs
  NEXT
  
  CLOSE
  
  IF AbortFlag GOTO ExitPrintSheets
  
  Header$ = "Meter Reading Sheets"
  
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  GOTO ExitPrintSheets
  
PrintEm:
  GOSUB GetMeterFlags
  GOSUB LookForSecLights
  GOSUB LookForTrash
  GOSUB PrintSkipHeader
  ZONE$ = QPTrim$(UBCustRec(1).ZONE)
  SELECT CASE LEFT$(QPTrim$(UBCustRec(1).CUSTTYPE), 1)
  CASE "B"
    CustT$ = "Commerical"
  CASE "R"
    CustT$ = "Residential"
  CASE ELSE
    CustT$ = "??????????"
  END SELECT
  IF EleFlag THEN
    PRINT #RptHandle, " Electric"; "  "; ECode$
  ELSE
    PRINT #RptHandle, ""
  END IF
  IF WatFlag THEN
    PRINT #RptHandle, " Water"; "  "; ZONE$; "  Min ="; WatMin
  ELSE
    PRINT #RptHandle, ""
  END IF
  IF SewFlag THEN
    PRINT #RptHandle, " Sewer"; "  "; ZONE$; "  Min ="; SewMin
  ELSE
    PRINT #RptHandle, ""
  END IF
  
  IF SecFlag THEN
    PRINT #RptHandle, " Security Lights   #"; SecCnt
  ELSE
    PRINT #RptHandle, ""
  END IF
  
  IF TrashFlag THEN
    PRINT #RptHandle, " Trash   "; TCode$
  ELSE
    PRINT #RptHandle, ""
  END IF
  
  PRINT #RptHandle, TAB(15); EleMin
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, EleSer$; TAB(24); WatSer$
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(6); USING "#########"; EleRead&; TAB(24); WatRead&
  FOR zz = 21 TO 43
    PRINT #RptHandle,
  NEXT
  PRINT #RptHandle, UBCustRec(1).CUSTNAME
  PRINT #RptHandle, UBCustRec(1).ADDR1
  PRINT #RptHandle, UBCustRec(1).SERVADDR
  PRINT #RptHandle, QPTrim$(UBCustRec(1).CITY); " "; UBCustRec(1).STATE; " "; UBCustRec(1).ZIPCODE
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(11); UBCustRec(1).BOOK; "-"; UBCustRec(1).SEQNUMB
  PRINT #RptHandle, "~"
  RETURN
  
  
LookForSecLights:
  SecFlag = False
  FOR FRCnt = 1 TO 4
    IF INSTR(UBCustRec(1).FlatRates(FRCnt).FRDESC, "SECUR") THEN
      SecFlag = True
      SecCnt = UBCustRec(1).FlatRates(FRCnt).NumMin
      EXIT FOR
    END IF
  NEXT
  RETURN
  
LookForTrash:
  TrashFlag = False
  IF LEN(QPTrim$(UBCustRec(1).Serv(9).RateCode)) > 0 THEN
    TrashFlag = True
    TCode$ = UBCustRec(1).Serv(9).RateCode
  END IF
  RETURN
  
PrintSkipHeader:
  FOR zz = 1 TO 9
    PRINT #RptHandle,
  NEXT
  RETURN
  
GetMeterFlags:
  WatFlag = False: WatMin = 0: WatSer$ = "": WatRead& = 0
  SewFlag = False: SewMin = 0:
  EleFlag = False: EleMin = 0: EleSer$ = "": EleRead& = 0
  
  FOR MtrCnt = 1 TO 7
    SELECT CASE UBCustRec(1).LocMeters(MtrCnt).MTRType
    CASE "C"
      WatFlag = True
      SewFlag = True
      WatMin = UBCustRec(1).LocMeters(MtrCnt).NumUser
      SewMin = UBCustRec(1).LocMeters(MtrCnt).NumUser
      WatRead& = UBCustRec(1).LocMeters(MtrCnt).CurRead
      WatSer$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MTRNUM)
      IF LEN(WatSer$) = 0 THEN
        WatSer$ = "?????"
      END IF
    CASE "W"
      WatFlag = True
      WatMin = UBCustRec(1).LocMeters(MtrCnt).NumUser
      WatSer$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MTRNUM)
      WatRead& = UBCustRec(1).LocMeters(MtrCnt).CurRead
      IF LEN(WatSer$) = 0 THEN
        WatSer$ = "?????"
      END IF
    CASE "S"
      SewFlag = True
      SewMin = UBCustRec(1).LocMeters(MtrCnt).NumUser
      WatSer$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MTRNUM)
      WatRead& = UBCustRec(1).LocMeters(MtrCnt).CurRead
      IF LEN(WatSer$) = 0 THEN
        WatSer$ = "?????"
      END IF
    CASE "E"
      EleFlag = True
      EleMin = UBCustRec(1).LocMeters(MtrCnt).NumUser
      EleSer$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MTRNUM)
      EleRead& = UBCustRec(1).LocMeters(MtrCnt).CurRead
      ECode$ = UBCustRec(1).Serv(3).RateCode
      IF LEN(EleSer$) = 0 THEN
        EleSer$ = "?????"
      END IF
    END SELECT
  NEXT
  IF WatRead& < 0 THEN
    WatRead& = 0
  END IF
  IF EleRead& < 0 THEN
    EleRead& = 0
  END IF
  RETURN
  
DoSheetMask:
  SaveScrn TempScrn()
  
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle     'Open Report File
  
  PRINT #RptHandle, "TOP"
  FOR zz = 1 TO 8
    PRINT #RptHandle, ""
  NEXT
  PRINT #RptHandle, " Electric  XXXXXXXXXX"
  PRINT #RptHandle, " Water  X  Min = X"
  PRINT #RptHandle, " Sewer  X  Min = X"
  PRINT #RptHandle, ""
  PRINT #RptHandle, ""
  PRINT #RptHandle, "               X"
  PRINT #RptHandle, ""
  PRINT #RptHandle, ""
  PRINT #RptHandle, "XXXXXXXXX              XXXXXXXXX"
  PRINT #RptHandle, ""
  PRINT #RptHandle, ""
  PRINT #RptHandle, "       XXXXXXX           XXXXXXX"
  FOR zz = 1 TO 23
    PRINT #RptHandle, ""
  NEXT
  PRINT #RptHandle, "XXXXXXX XXXXXXXXXXX"
  PRINT #RptHandle, "XX XXX XXX"
  PRINT #RptHandle, "XXXXXXXXXXX"
  PRINT #RptHandle, "XXXXXXXXX XX XXXXX"
  PRINT #RptHandle, ""
  PRINT #RptHandle, "          XX-XXXXXX"
  PRINT #RptHandle, "BOTTOM"
  
  CLOSE RptHandle
  
  CALL CursorOff
  
  PrintRptFile Header$, ReportFile$, 1, RetCode%, 4
  RestScrn TempScrn()
  Action = 1
  
  RETURN
  
CheckBookSequence:
  OkFlag = False
  IF In1 THEN
    Fld1 = 1
    Fld2 = 2
  ELSE
    Fld1 = 3
    Fld2 = 4
  END IF
  Form$(Fld1, 0) = FmtBook$(Form$(Fld1, 0))
  Form$(Fld2, 0) = FmtSeqN$(Form$(Fld2, 0))
  IF Chk4BookSeqNum(Form$(Fld1, 0), Form$(Fld2, 0)) = 0 THEN
    EditedFlag = False
    Frm(1).KeyCode = 0
    SaveScrn TempScrn()
    Ok = MsgBox%("UB", "BADACCTN")
    RestScrn TempScrn()
    IF In1 THEN
      Frm(1).FldNo = 1
    ELSE
      Frm(1).FldNo = 3
    END IF
  ELSE
    Frm(1).PageDisable = False
    OkFlag = True
  END IF
  Action = 2
  
  RETURN
  
ExitPrintSheets:
END SUB

