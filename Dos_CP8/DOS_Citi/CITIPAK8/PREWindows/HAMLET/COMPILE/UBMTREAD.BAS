DEFINT A-Z
DECLARE SUB InactiveConsReport ()
DECLARE SUB CursorOff ()
DECLARE SUB ProcessItron ()
DECLARE SUB ProcessLogicon ()
DECLARE SUB PrintHighLowReport ()
DECLARE SUB ProcessTelxon ()
DECLARE SUB DisplayHHScrn (ScrnName$)
DECLARE SUB ProcessHusky ()
DECLARE SUB PrintMeterNotes ()
DECLARE FUNCTION Chk4BookSeqNum (Book$, SeqNum$)
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
DECLARE FUNCTION FmtBook$ (Book$)
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

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubdraft.BI'
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'Newcust.bi'
  '$INCLUDE: 'ubgilsen.bi'
  '$INCLUDE: 'ubschlum.bi'

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

'  REDIM UBSenRdRec(1) AS UBSensusReadRecType
'  UBSenRdRecLen = LEN(UBSenRdRec(1))
'  REDIM UBSenGetRdRec(1) AS UBSensusGetReadRecType
'  UBSenGetRdRecLen = LEN(UBSenGetRdRec(1))
'  PRINT "send:"; UBSenRdRecLen
'  PRINT " get:"; UBSenGetRdRecLen
'  END
  
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
      'HandHeldProcessing             'unrem
    CASE 3
      EstMeterReading                'unrem
    CASE 4
      'PrintMeterList                 'unrem
    CASE 5
      PrintMeterReport               'unrem
    CASE 6
      'PrintHighLowReport             'unrem
    CASE 7
      'InactiveConsReport             'unrem
    CASE 8
      'PrintMeterNotes                'unrem
    CASE 9
      'PrintMeterSheets               'unrem
    CASE 10
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP
  IF INSTR(COMMAND$, "TEST") = 0 THEN
    RUN "ubbillin"
  ELSE
    HideCursor
    ClearScrn
  END IF

FUNCTION Chk4BookSeqNum (Book$, SeqNum$)
  
  Chk4BookSeqNum = False        'assume not found
  
  TBookSeq& = QPValL(Book$ + SeqNum$)
  
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

SUB DisplayH2Scrn (ScrnName$)
  LibFile2Scrn "UBHH", ScrnName$, MonoCode%, Attribute%, ErrorCode%
END SUB

SUB DisplayHHScrn (ScrnName$)
  LibFile2Scrn "UBSETUP", ScrnName$, MonoCode%, Attribute%, ErrorCode%
END SUB

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
            Prec& = IdxBuff(Cnt!).RecNum
            IF Prec& > 0 THEN
              GET UBFile, Prec&, UBCustRec(1)
              FOR BookCnt = 1 TO NumberofRoutes
                IF VAL(UBCustRec(1).Book) = Route(BookCnt) AND (UBCustRec(1).Status <> "F") THEN
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
            MeterID$ = UBCustRec(1).Book + UBCustRec(1).SeqNumb
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
  
  IF Book < 10 THEN
    Book$ = "0" + QPTrim$(STR$(Book))
  ELSE
    Book$ = QPTrim$(STR$(Book))
  END IF
  
  BlockClear
  ShowProcessingScrn "Estimating Book: " + Book$ + "  "
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  UBCustFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCustFile LEN = UBCustRecLen
  NumOfCust = LOF(UBCustFile) \ UBCustRecLen
  
  FOR LCnt& = 1 TO NumOfCust
    GET #UBCustFile, LCnt&, UBCustRec(1)
    DidEM = False
    IF UBCustRec(1).Status = "A" THEN
      IF UBCustRec(1).Book = Book$ THEN
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
  Book = VAL(Form$(2, 0))
  IF ESTDate < 0 OR Book < 1 THEN
    Ok2EST = False
  ELSE
    Ok2EST = True
  END IF
  
  
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
          Prec& = IdxBuff(Cnt!).RecNum
          IF NOT (Prec&) = 0 THEN
            GET UBFile, Prec&, UBCustRec(1)
            IF INSTR(UBCustRec(1).HHMSG1, "NOREAD") > 0 THEN
              GOTO HWriteSkip
            END IF
            DisplayAcct$ = UBCustRec(1).Book + "-" + UBCustRec(1).SeqNumb
            QPrintRC DisplayAcct$, 11, 45, -1
            FOR BookCnt! = 1 TO NumberofRoutes
              IF VAL(UBCustRec(1).Book) = Route(BookCnt!) AND (UBCustRec(1).Status <> "F") THEN
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
            MeterID$ = UBCustRec(1).Book + UBCustRec(1).SeqNumb
          END IF
          IF LEN(MeterID$) < 8 THEN
            MeterID$ = STRING$(8 - LEN(MeterID$), "0") + MeterID$
          END IF
          MeterID$ = LEFT$(MeterID$, 8)
          
          'Set Record Fields and Put On Disk
          UBPC3000RDRec(1).CUSTNAME = UBCustRec(1).CUSTNAME
          UBPC3000RDRec(1).ServAddress = LEFT$(UBCustRec(1).SERVADDR, 16)

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
          UBPC3000RDRec(1).MeterType$ = UBCustRec(1).LocMeters(MtrCnt!).MtrType
          UBPC3000RDRec(1).Book = VAL(UBCustRec(1).Book)
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
        Prec& = IdxBuff(Cnt!).RecNum
        LOCATE 11, 42: COLOR 15: PRINT USING "######"; Prec&
        IF NOT (Prec&) = 0 THEN
          GET UBFile, Prec&, UBCustRec(1)
          IF VAL(UBCustRec(1).Book) <> 0 THEN
            'Take Care of First Record Sent First
            IF FirstOnRoute = 0 THEN
              FirstOnRoute = VAL(UBCustRec(1).Book)
              UBItronFRec(1).RecordCode = "F"
              UBItronFRec(1).Route = UBCustRec(1).Book + "000000"
              UBItronFRec(1).Message = ""
              UBItronFRec(1).Filler = ""
              UBItronFRec(1).CrLf = CHR$(13) + CHR$(10)
              PUT #UBItronRdFile, 1, UBItronFRec(1)
            END IF
            
            'Compare to Check For End of Book Run
            IF FirstOnRoute < VAL(UBCustRec(1).Book) THEN
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
              FirstOnRoute = VAL(UBCustRec(1).Book)
              NumberofAccounts = 0
              UBItronFRec(1).RecordCode = "F"
              UBItronFRec(1).Route = UBCustRec(1).Book + "000000"
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
  CustomerName$ = QPTrim$(UBCustRec(1).CUSTNAME)
  'Set Customer Address
  CustomerAddress$ = QPTrim$(UBCustRec(1).SERVADDR)
  
  
  IF (ASC(UBCustRec(1).LocMeters(MtrCnt!).MtrType) > 32) THEN
    IF UBCustRec(1).LocMeters(MtrCnt!).MtrType = "C" OR UBCustRec(1).LocMeters(MtrCnt!).MtrType = "W" OR UBCustRec(1).LocMeters(MtrCnt!).MtrType = "S" THEN
      
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
      UBItronARec(1).Route = UBCustRec(1).Book + "000000"
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
      UBItronBRec(1).CUSTNAME = CustomerName$
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
  NumberofRoutes = 0
  
  
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
          UBLogRdRec(1).Book = ""
          UBLogRdRec(1).Future = ""
          UBLogRdRec(1).Recend = "X"
          UBLogRdRec(1).CrLf = CHR$(13) + CHR$(10)
          PUT UBLogRdFile, (LOF(UBLogRdFile) / UBLogRdRecLen) + 1, UBLogRdRec(1)
          
          'Open the Location Order for Reading
          IdxRecLen = 4         'we are using a integer
          IdxFileSize& = FileSize&("UBCUSTBK.IDX")
          IdxNumOfRecs = IdxFileSize& \ IdxRecLen
          REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
          FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs            'load it
          
          Cnt! = 1
          
          DO
            
            Prec& = IdxBuff(Cnt!).RecNum
            IF NOT (Prec&) = 0 THEN
              GET UBFile, Prec&, UBCustRec(1)
              FOR BookCnt! = 1 TO NumberofRoutes
                IF VAL(UBCustRec(1).Book) = Route(BookCnt!) AND (UBCustRec(1).Status <> "F") THEN
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
  Account$ = STR$(Prec&)
  Account$ = LEFT$(Account$, 6) + "-" + RIGHT$(STR$(MtrCnt!), 1)
  
  WHILE MtrCnt! < 8
    
    IF (ASC(UBCustRec(1).LocMeters(MtrCnt!).MtrType) > 32) THEN
      IF UBCustRec(1).LocMeters(MtrCnt!).MtrType = "C" OR UBCustRec(1).LocMeters(MtrCnt!).MtrType = "W" OR UBCustRec(1).LocMeters(MtrCnt!).MtrType = "T" OR UBCustRec(1).LocMeters(MtrCnt!).MtrType = "S" THEN
        
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
          MeterID$ = UBCustRec(1).Book + UBCustRec(1).SeqNumb
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
        UBLogRdRec(1).Book = UBCustRec(1).Book
        UBLogRdRec(1).Future = ""
        UBLogRdRec(1).Recend = "X"
        UBLogRdRec(1).CrLf = CHR$(13) + CHR$(10)
        PUT UBLogRdFile, (LOF(UBLogRdFile) / UBLogRdRecLen) + 1, UBLogRdRec(1)
      END IF
    END IF
    
SkipEmLC:
    MtrCnt! = MtrCnt! + 1
  WEND
  RETURN
  
FILENAM1LC:
  PathWay$ = QPTrim$(PathWay$)
  FileName$ = PathWay$ + "WBLOGNO" + LTRIM$(RouteID$) + ".DAT"
RETURN
  
GetLogicon:
  LibName$ = "UBSETUP"
  ScrnName$ = "UBLOG1G"
  Help$ = "Logicon Hand Held Processing"
  CursorOff
  
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
        RRDate = Date2Num(Form$(2, 0))
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
          Prec& = VAL(QPTrim$(UBLogGetRdRec(1).AcctNo))
          
          IF LEFT$(UBLogGetRdRec(1).CurRead, 1) <> "X" THEN
            IF NOT (Prec&) = 0 THEN
              GET UBFile, Prec&, UBCustRec(1)
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
  IF RRDate > 0 THEN
    MeterReadDate$ = Num2Date$(RRDate)
  ELSE
    MeterReadDate$ = MID$(UBLogGetRdRec(1).ReadDate, 3, 2) + "/" + MID$(UBLogGetRdRec(1).ReadDate, 5, 2) + "/" + RIGHT$(DATE$, 4)
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
  
FileNamLC:
  PathWay$ = QPTrim$(PathWay$)
  FileName$ = PathWay$ + "WBLOGNO" + LTRIM$(RouteID$) + ".DAT"
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
          
          Prec& = IdxBuff(Cnt!).RecNum
          IF NOT (Prec&) = 0 THEN
            GET UBFile, Prec&, UBCustRec(1)
            DisplayAcct$ = UBCustRec(1).Book + "-" + UBCustRec(1).SeqNumb
            QPrintRC DisplayAcct$, 11, 45, -1
            FOR BookCnt! = 1 TO NumberofRoutes
              IF VAL(UBCustRec(1).Book) = Route(BookCnt!) AND (UBCustRec(1).Status <> "F") THEN
                GOSUB TelxonPutRec
              END IF
            NEXT BookCnt!
          END IF
          Cnt! = Cnt! + 1
          
        LOOP UNTIL Cnt! > IdxNumOfRecs
        CLOSE
        
        Ok = MsgBox%("UBSETUP", "UBTELX1C")
        
        IF Ok = 1 THEN
          KillFile "TELXON.ERR"
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
    KillFile "TELXON.ERR"
    KillFile "UBCUSTTR.DAT"
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
    Prec& = VAL(QPTrim$(UBPC3000GetRDRec(1).Account))
    
    IF NOT (Prec&) = 0 THEN
      GET UBFile, Prec&, UBCustRec(1)
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
    UBCustRec(1).NewNotes = True
    UBCustRec(1).HHMSG1 = UBPC3000GetRDRec(1).Note1
    UBCustRec(1).HHMSG2 = UBPC3000GetRDRec(1).Note2
    UBCustRec(1).HHMSG3 = UBPC3000GetRDRec(1).Note3
  END IF
  '   ^^^
  PUT UBFile, Prec&, UBCustRec(1)
  RETURN
  
TelxonPutRec:
  'May Have Up to 7 Meters to Read
  FOR MtrCnt! = 1 TO 7
    MeterOK = False
    Account$ = STR$(Prec&)
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
            MeterID$ = UBCustRec(1).Book + UBCustRec(1).SeqNumb
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
          UBPC3000RDRec(1).Book = VAL(UBCustRec(1).Book)
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

