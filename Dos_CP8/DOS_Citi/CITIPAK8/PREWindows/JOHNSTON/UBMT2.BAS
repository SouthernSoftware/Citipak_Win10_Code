'need to add the Skip Inactive Flag Check.
DEFINT A-Z
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
DECLARE FUNCTION FindRateTbl% (RATECODE$, NumOfRates%, UBRateTbls() AS ANY)
DECLARE FUNCTION GetNumOfRevs% ()
DECLARE FUNCTION Round# (N#)
DECLARE FUNCTION GetNumOfCust% ()
DECLARE FUNCTION GetNumRateRecs% ()
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION FmtBook$ (Book$)
DECLARE FUNCTION FmtSeqN$ (SeqN$)
  '$INCLUDE: 'DefCnf.BI'
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)
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
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'ubdraft.BI'
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'PageInfo.bi'
  '$INCLUDE: 'Newcust.bi'
  '$INCLUDE: 'ubrate.BI'
  ''$INCLUDE: 'ubsetbal.BI'
  '$INCLUDE: 'ubgilsen.bi'
  
  CONST False = 0, True = NOT False
  
  'GOTO Skip2Here
  
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
  Col = ((80 - MaxLen) \ 2) - 2
  
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

SUB DisplayHHScrn (ScrnName$)
  LibFile2Scrn "UBSETUP", ScrnName$, MonoCode%, Attribute%, ErrorCode%
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
          IF LEN(QPTrim$(UBCustRec(1).LocMeters(zz).MTRType)) > 0 THEN
            IF UBCustRec(1).LocMeters(zz).UseCnt > 0 AND UBCustRec(1).LocMeters(zz).ReadFlag <> "Y" THEN
              DidEM = True
              UBCustRec(1).LocMeters(zz).PrevRead = UBCustRec(1).LocMeters(zz).CurRead
              UBCustRec(1).LocMeters(zz).PastDate = UBCustRec(1).LocMeters(zz).CurDate
              UBCustRec(1).LocMeters(zz).ReadFlag = "Y"
              IF BeechFlag THEN
                AvgUse& = 20
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
  
  'LOCATE 25, 1: PRINT "Did:"; DoneCnt;
  'WaitForAction
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

SUB HandHeldProcessing
  
  STATIC SName$, AcctNum&, MeterNum$, SAddr$, LNum$
  
  SHARED Choice$()
  REDIM TempScrn(0)
  REDIM Route(100)
  
  ' Check For Device Type and Run Appropriate Program
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen
  
  TownName$ = UBSetUpRec(1).UTILNAME
  
  IF INSTR(TownName$, "HERTFORD") THEN
    HertFlag = True
  END IF
  
  IF UBSetUpLen = 0 THEN
    PRINT CHR$(7);
    SaveScrn TempScrn()
    LOCATE , , 0
    DisplayHHScrn "UBNOHH"
    WaitForAction
    RestScrn TempScrn()
    LOCATE , , 1
    CLOSE
    EXIT SUB
  END IF
  
  SELECT CASE UBSetUpRec(1).HHDEVICE
  CASE "H", "S", "P", "C", "D", "M", "T", "L", "I"
    GoodHHDevice = True
  CASE ELSE
    PRINT CHR$(7);
    SaveScrn TempScrn()
    LOCATE , , 0
    DisplayUBScrn "UBNOHH"
    WaitForAction
    RestScrn TempScrn()
    LOCATE , , 1
    CLOSE
    EXIT SUB
  END SELECT
  
  ' Valid HandHeld Type Found
  
  HighVar = UBSetUpRec(1).HighRead
  LowVar = UBSetUpRec(1).LowRead
  
  SELECT CASE UBSetUpRec(1).HHDEVICE
  CASE "S"
    GOSUB Sensus
  CASE "P"
    IF UBSetUpRec(1).HHDEVICE = "P" THEN
      pc2000flag$ = "N"
      GOSUB PC3000
    END IF
  CASE "D"
    GOSUB DataGeneral
  CASE "C"
    GOSUB SysCom
  CASE "M"
    IF UBSetUpRec(1).HHDEVICE = "M" THEN
      pc2000flag$ = "Y"
      GOSUB PC3000
    END IF
  CASE "H"
    GOSUB Husky
  CASE "T"
    GOSUB Telxon
  CASE "L"
    GOSUB Logicon
  CASE "I"
    GOSUB Itron
  END SELECT
  
  EXIT SUB

Itron:
  ProcessItron
  RETURN
 
Logicon:
  ProcessLogicon
  RETURN
  
Husky:
  'ProcessHusky
  RETURN
  
Telxon:
  ProcessTelxon
  RETURN
  
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
  LOCATE 1, 1, 0
  
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
  NEXT F
  
  REDIM Choice$(3, 0)
  
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Send Info to Sensus"
  Choice$(2, 0) = "Get Info From Sensus"
  
  IF NumSensusRecs = 1 THEN
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
  Help$ = "Sensus Hand Held Processing"
  LOCATE 1, 1, 0
  
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
          INTERROGATOR = VAL(Form$(1, 0))
          COLOR 11
          LOCATE 12, 18
          PRINT "LOADING THESE ROUTES !!!                      "
          GOSUB FILENAM1
          ' Set Route Number Values
          FOR Cnt! = 1 TO NumberofRoutes
            Route(Cnt!) = VAL(Form$(Cnt! + 2, 0))
          NEXT Cnt!
          
          'Open Customer Data File
          REDIM UBCustRec(1) AS NewUBCustRecType
          UBCustRecLen = LEN(UBCustRec(1))
          UBFile = FREEFILE
          OPEN "UBCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBFile LEN = UBCustRecLen
          
          'if Gilbert sc then use different include
          '"HILLSVILLE" CHILHOWIE
          
          IF INSTR(TownName$, "GILBERT") OR INSTR(TownName$, "HILLSVILLE") OR INSTR(TownName$, "CHILHOWIE") THEN
            
            REDIM UBGilSenRdRec(1) AS UBGilSensusReadRecType
            UBSenRdRecLen = LEN(UBGilSenRdRec(1))
          ELSE
            REDIM UBSenRdRec(1) AS UBSensusReadRecType
            UBSenRdRecLen = LEN(UBSenRdRec(1))
          END IF
          
          UBSenRdFile = FREEFILE
          OPEN FileName$ FOR RANDOM SHARED AS UBSenRdFile LEN = UBSenRdRecLen
          CLOSE UBSenRdFile: KILL FileName$
          UBSenRdFile = FREEFILE
          OPEN FileName$ FOR RANDOM SHARED AS UBSenRdFile LEN = UBSenRdRecLen
          NumSenRdRecs = LOF(UBSenRdFile) / UBSenRdRecLen
          
        
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
              FOR BookCnt! = 1 TO NumberofRoutes
                IF VAL(UBCustRec(1).Book) = Route(BookCnt!) AND (UBCustRec(1).Status <> "F") THEN
                  IF HertFlag AND UBCustRec(1).Status = "I" THEN
                  ELSE
                    GOSUB WriteRecord
                  END IF
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
          PRINT "Interrogator Number "; INTERROGATOR
          LOCATE 9, 18
          COLOR 11
          PRINT "Information Sent to Sensus System Successfully"
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
            MeterID$ = UBCustRec(1).Book + UBCustRec(1).SEQNUMB
          END IF
          IF LEN(MeterID$) < 8 THEN
            MeterID$ = STRING$(8 - LEN(MeterID$), "0") + MeterID$
          END IF
          MeterID$ = LEFT$(MeterID$, 8)
          
          'Set Record Fields and Put On Disk
          IF INSTR(TownName$, "GILBERT") OR INSTR(TownName$, "HILLSVILLE") OR INSTR(TownName$, "CHILHOWIE") THEN
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
            UBGilSenRdRec(1).LocationNumber = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
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
  LOCATE 1, 1, 0
  
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
        PRINT "     Invalid Interrogator Number "
        LOCATE 14, 20: COLOR 11
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
        'Open Customer Data File
        REDIM UBCustRec(1) AS NewUBCustRecType
        UBCustRecLen = LEN(UBCustRec(1))
        UBFile = FREEFILE
        OPEN "UBCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBFile LEN = UBCustRecLen
        
        'Open and Initialize the Sensus Read Information File
        'if Gilbert sc then use different include
        IF INSTR(TownName$, "GILBERT") OR INSTR(TownName$, "HILLSVILLE") OR INSTR(TownName$, "CHILHOWIE") THEN
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
          IF INSTR(TownName$, "GILBERT") OR INSTR(TownName$, "HILLSVILLE") OR INSTR(TownName$, "CHILHOWIE") THEN
            GET UBSenGetRdFile, Cnt!, UBGilSenGetRdRec(1)
            Prec& = VAL(QPTrim$(UBGilSenGetRdRec(1).Account))
          ELSE
            GET UBSenGetRdFile, Cnt!, UBSenGetRdRec(1)
            Prec& = VAL(QPTrim$(UBSenGetRdRec(1).Account))
          END IF
          
          IF NOT (Prec&) = 0 THEN
            GET UBFile, Prec&, UBCustRec(1)
            GOSUB ExtractRecord
          END IF
          Cnt! = Cnt! + 1
          
        LOOP UNTIL Cnt! > NumSenGetRdRecs
        
        CLOSE
        Done = True
        FOR lc = 8 TO 10
          LOCATE lc, 16
          PRINT STRING$(40, 32)
        NEXT lc
        LOCATE 8, 18
        COLOR 11
        PRINT "Interrogator Number "; INTERROGATOR
        LOCATE 13, 18
        COLOR 11
        PRINT "Reading Updated From Sensus System Successfully"
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
  
  
ExtractRecord:
  UpdCnt! = UpdCnt! + 1
  LOCATE 13, 20: COLOR 11
  PRINT " Updated Count:  "; : COLOR 15: PRINT UpdCnt!
  COLOR 11
  
  
  'New Way To Activate
  IF INSTR(TownName$, "GILBERT") OR INSTR(TownName$, "HILLSVILLE") OR INSTR(TownName$, "CHILHOWIE") THEN
    MeterRecord = VAL(RIGHT$((QPTrim$(UBGilSenGetRdRec(1).Account)), 1))
  ELSE
    MeterRecord = VAL(RIGHT$((QPTrim$(UBSenGetRdRec(1).Account)), 1))
  END IF
  
  ' Check Meter Updated Flag
  
  
  ' Update Meter W/Reading
  IF INSTR(TownName$, "GILBERT") OR INSTR(TownName$, "HILLSVILLE") OR INSTR(TownName$, "CHILHOWIE") THEN
    CurReading# = VAL(UBGilSenGetRdRec(1).CurRead)
    MeterReadDate$ = LEFT$(DATE$, 6) + RIGHT$(DATE$, 2)
  ELSE
    CurReading# = VAL(UBSenGetRdRec(1).Reading)
    MeterReadDate$ = LEFT$(UBSenGetRdRec(1).DateRead, 2) + "/" + MID$(UBSenGetRdRec(1).DateRead, 3, 2) + "/" + RIGHT$(Form$(2, 0), 2)
  END IF
  IF UBCustRec(1).LocMeters(MeterRecord).MTRType = "T" THEN CurReading# = CurReading# * 1000
  
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
  
  IF INTERROGATOR = 1 THEN FileName$ = PathWay$ + "exssi001.dat": RETURN
  IF INTERROGATOR = 2 THEN FileName$ = PathWay$ + "exssi002.dat": RETURN
  IF INTERROGATOR = 3 THEN FileName$ = PathWay$ + "exssi003.dat": RETURN
  IF INTERROGATOR = 4 THEN FileName$ = PathWay$ + "exssi004.dat": RETURN
  IF INTERROGATOR = 5 THEN FileName$ = PathWay$ + "exssi005.dat": RETURN
  IF INTERROGATOR = 6 THEN FileName$ = PathWay$ + "exssi006.dat": RETURN
  IF INTERROGATOR = 7 THEN FileName$ = PathWay$ + "exssi007.dat": RETURN
  IF INTERROGATOR = 8 THEN FileName$ = PathWay$ + "exssi008.dat": RETURN
  IF INTERROGATOR = 9 THEN FileName$ = PathWay$ + "exssi009.dat": RETURN
  
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
  LOCATE 1, 1, 0
  
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
  NEXT F
  
  REDIM Choice$(5, 0)
  
  Choice$(0, 0) = "1"
  
  IF pc2000flag$ = "Y" THEN
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
        COLOR 12: LOCATE 11, 40: PRINT "Invalid Operation": BEEP:
        WaitForAction
        COLOR 15: LOCATE 11, 40: PRINT STRING$(25, 32)
        
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
  LOCATE 1, 1, 0
  
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
      WhatTypes$ = LEFT$(Form$(1, 0), 1)
      
      ' Next Check Routes
      IF NumberofRoutes > 0 THEN
        COLOR 11
        LOCATE 12, 18
        PRINT "LOADING THESE ROUTES !!!                      "
        GOSUB FILENAM1
        ' Set Route Number Values
        FOR Cnt! = 1 TO NumberofRoutes
          Route(Cnt!) = VAL(Form$(Cnt! + 2, 0))
        NEXT Cnt!
        PathWay$ = RTRIM$(PathWay$)
        
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
            DisplayAcct$ = "Processing Location #" + UBCustRec(1).Book + UBCustRec(1).SEQNUMB
            PrintHelp DisplayAcct$
            FOR BookCnt! = 1 TO NumberofRoutes
              IF VAL(UBCustRec(1).Book) = Route(BookCnt!) AND (UBCustRec(1).Status <> "F") THEN
                GOSUB PC3000WriteRecord
              END IF
            NEXT BookCnt!
          END IF
          Cnt! = Cnt! + 1
          
        LOOP UNTIL Cnt! > IdxNumOfRecs
        CLOSE
        CLS
        COLOR 11
        IF pc2000flag$ = "Y" THEN
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
        
        IF pc2000flag$ = "Y" THEN
          SHELL "REMDISK /com1"
          destination$ = "copy ubcusttr.dat  " + PathWay$
          SHELL destination$
          SHELL "REMDISK /u"
        ELSE
          SHELL "CONNECT"
          destination$ = "copy ubcusttr.dat  " + PathWay$
          SHELL destination$
          'END IF
        END IF
        PRINT
        PRINT "Press any key. . ."
        CALL WaitForAction
        
        Done = True
        DisplayUBScrn ScrnName$
        FOR lc = 8 TO 10
          LOCATE lc, 16
          PRINT STRING$(40, 32)
        NEXT lc
        LOCATE 9, 18
        COLOR 11
        IF pc2000flag$ = "Y" THEN
          PRINT "Information Sent to PC2000 Successfully"
        ELSE
          PRINT "Information Sent to PC3000 Successfully"
        END IF
        LOCATE 12, 18: PRINT "THESE ROUTES LOADED!!!                      "
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
            MeterID$ = UBCustRec(1).Book + UBCustRec(1).SEQNUMB
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
  
GetPC3000:
  
  LibName$ = "UB"
  ScrnName$ = "UBPCHH1G"
  '  help$ = "Sensus Hand Held Processing"
  LOCATE 1, 1, 0
  
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
  NEXT F
  
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
      IF pc2000flag$ = "Y" THEN
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
      IF pc2000flag$ = "Y" THEN
        SHELL "REMDISK /com1"
      ELSE
        SHELL "CONNECT"
      END IF
      
      DisplayUBScrn ScrnName$
      
      'Open Customer Data File
      REDIM UBCustRec(1) AS NewUBCustRecType
      UBCustRecLen = LEN(UBCustRec(1))
      UBFile = FREEFILE
      OPEN "UBCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBFile LEN = UBCustRecLen
      
      'Open and Initialize the PC3000 Read Information File
      REDIM UBPC3000GetRDRec(1) AS UBPC3000ReadRecType
      UBPC3000GetRdRecLen = LEN(UBPC3000GetRDRec(1))
      UBPC3000GetRdFile = FREEFILE
      FileNam$ = RTRIM$(PathWay$) + "UBCUSTTR.DAT"
      OPEN FileNam$ FOR RANDOM ACCESS READ WRITE SHARED AS UBPC3000GetRdFile LEN = UBPC3000GetRdRecLen
      NumPC3000GetRdRecs = LOF(UBPC3000GetRdFile) / UBPC3000GetRdRecLen
      
      IF NumPC3000GetRdRecs = 0 THEN
        FOR Cnt = 12 TO 16: LOCATE Cnt, 20: PRINT STRING$(40, 32): NEXT Cnt
          LOCATE 12, 20: COLOR 12
          PRINT "      [  E R R O R  ]"
          LOCATE 13, 20: COLOR 11
          PRINT " Nothing Found From MicroPalm "
          LOCATE 14, 20: COLOR 11
          PRINT " Please Check Unit Connection "
          BEEP
          WaitForAction
          CLOSE
          GOTO GetPC3000
        END IF
        
        Cnt! = 1                ' Initialize File Counter to 1
        DO
          GET UBPC3000GetRdFile, Cnt!, UBPC3000GetRDRec(1)
          
          ' New Way Undelete When Our System Sends and Retrieves
          Prec& = VAL(QPTrim$(UBPC3000GetRDRec(1).Account))
          
          IF NOT (Prec&) = 0 THEN
            GET UBFile, Prec&, UBCustRec(1)
            GOSUB PC3000ExtractRecord
          END IF
          Cnt! = Cnt! + 1
          
        LOOP UNTIL Cnt! > NumPC3000GetRdRecs
        
        CLOSE
        Done = True
        FOR lc = 12 TO 14
          LOCATE lc, 16
          PRINT STRING$(40, 32)
        NEXT lc
        LOCATE 12, 18
        COLOR 11
        PRINT "Readings Updated From PC3000 Successfully"
        PRINT CHR$(7);
        WaitForAction
        Done = True
        IF pc2000flag$ = "Y" THEN SHELL "REMDISK /u"
        
      CASE ESC
        Done = True
      CASE ELSE
        Done = False
      END SELECT
    LOOP UNTIL Done
    RETURN
    
    
PC3000ExtractRecord:
    UpdCnt! = UpdCnt! + 1
    LOCATE 14, 16: PRINT STRING$(30, 32);
    LOCATE 13, 16: COLOR 11
    PRINT "   Updated Count:  "; : COLOR 15: PRINT UpdCnt!
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
      UBCustRec(1).HHMSG1 = UBPC3000GetRDRec(1).Note1
      UBCustRec(1).HHMSG2 = UBPC3000GetRDRec(1).Note2
      UBCustRec(1).HHMSG3 = UBPC3000GetRDRec(1).Note3
    END IF
    '   ^^^
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
    Help$ = "Data Genie Hand Held Processing"
    LOCATE 1, 1, 0
    
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
    NEXT F
    
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
          COLOR 15: LOCATE 11, 40: PRINT STRING$(25, 32)
          
          Frm(1).FldNo = 1
          Action = 1
          Done = False
        ELSE
          
          PathWay$ = Form$(2, 0)
          UBDGRec(1).PathWay = PathWay$
          PUT UBDGFile, 1, UBDGRec(1)
          CLOSE UBDGFile
          IF LEFT$(Form$(1, 0), 1) = "S" THEN GOSUB SendDG
          IF LEFT$(Form$(1, 0), 1) = "G" THEN GOSUB GetDG
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
    Help$ = "Genie Hand Held Processing"
    LOCATE 1, 1, 0
    
    Help$ = "DownLoad HandHeld Operation"
    DisplayUBScrn ScrnName$
    COLOR 14
    LOCATE 7, 21: PRINT " DG 'Genie' File Maker "
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
        WhatTypes$ = LEFT$(Form$(1, 0), 1)
        
        ' Next Check Routes
        IF NumberofRoutes > 0 THEN
          COLOR 11
          LOCATE 12, 18
          PRINT "LOADING THESE ROUTES !!!                      "
          GOSUB FILENAM1
          ' Set Route Number Values
          FOR Cnt! = 1 TO NumberofRoutes
            Route(Cnt!) = VAL(Form$(Cnt! + 2, 0))
          NEXT Cnt!
          PathWay$ = RTRIM$(PathWay$)
          
          'Open Customer Data File
          REDIM UBCustRec(1) AS NewUBCustRecType
          UBCustRecLen = LEN(UBCustRec(1))
          UBFile = FREEFILE
          OPEN "UBCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBFile LEN = UBCustRecLen
          
          
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
              DisplayAcct$ = "Processing Location #" + UBCustRec(1).Book + UBCustRec(1).SEQNUMB
              PrintHelp DisplayAcct$
              FOR BookCnt! = 1 TO NumberofRoutes
                IF VAL(UBCustRec(1).Book) = Route(BookCnt!) AND (UBCustRec(1).Status <> "F") THEN
                  GOSUB DGWriteRecord
                END IF
              NEXT BookCnt!
            END IF
            Cnt! = Cnt! + 1
            
          LOOP UNTIL Cnt! > IdxNumOfRecs
          CLOSE
          FOR ll = 14 TO 16
            LOCATE ll, 17: PRINT STRING$(45, 32)
          NEXT ll
          LOCATE 15, 22: PRINT CHR$(7);
          PRINT "Data General 'Genie' File Now Ready"
          LOCATE 18, 17: PRINT STRING$(45, 32)
          LOCATE 18, 28: PRINT "Press any key to Continue"
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
          CASE "W"              'water/sewer
            IF MTRType$ = "W" OR MTRType$ = "S" OR MTRType$ = "C" THEN
              MeterType$ = "W"
              MeterOK = True
            END IF
          CASE "E"              'electric & demand elec.
            IF MTRType$ = "E" OR MTRType$ = "D" THEN
              MeterOK = True
            END IF
          CASE "G"              'gas
            IF MTRType$ = "G" THEN
              MeterType$ = "G"
              MeterOK = True
            END IF
          CASE "A", " "         'all meters
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
              MeterID$ = UBCustRec(1).Book + UBCustRec(1).SEQNUMB
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
    NEXT MtrCnt!
    RETURN
    
    
GetDG:
    DisplayUBScrn ScrnName$
    
    'Open Customer Data File
    REDIM UBCustRec(1) AS NewUBCustRecType
    UBCustRecLen = LEN(UBCustRec(1))
    UBFile = FREEFILE
    OPEN "UBCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBFile LEN = UBCustRecLen
    
    'Open and Initialize the Genie Information File
    REDIM UBDGGetRDRec(1) AS UBDGProcRecType
    UBDGGetRdRecLen = LEN(UBDGGetRDRec(1))
    UBDGGetRdFile = FREEFILE
    FileNam$ = RTRIM$(PathWay$) + "UBCUSTTR.DAT"
    OPEN FileNam$ FOR RANDOM ACCESS READ WRITE SHARED AS UBDGGetRdFile LEN = UBDGGetRdRecLen
    NumDGGetRdRecs = LOF(UBDGGetRdFile) / UBDGGetRdRecLen
    
    IF NumDGGetRdRecs = 0 THEN
      FOR Cnt = 12 TO 16: LOCATE Cnt, 20: PRINT STRING$(40, 32): NEXT Cnt
        LOCATE 12, 20: COLOR 12
        PRINT "      [  E R R O R  ]"
        LOCATE 13, 20: COLOR 11
        PRINT " Nothing Found From Genie "
        BEEP
        WaitForAction
        CLOSE
        GOTO DataGeneral
      END IF
      
      Cnt! = 1  ' Initialize File Counter to 1
      DO
        GET UBDGGetRdFile, Cnt!, UBDGGetRDRec(1)
        
        ' New Way Undelete When Our System Sends and Retrieves
        Prec& = VAL(QPTrim$(UBDGGetRDRec(1).Account))
        
        IF NOT (Prec&) = 0 THEN
          GET UBFile, Prec&, UBCustRec(1)
          GOSUB DGExtractRecord
        END IF
        Cnt! = Cnt! + 1
        
      LOOP UNTIL Cnt! > NumDGGetRdRecs
      
      CLOSE
      Done = True
      FOR lc = 12 TO 14
        LOCATE lc, 16
        PRINT STRING$(40, 32)
      NEXT lc
      LOCATE 12, 18
      COLOR 11
      PRINT "Readings Updated Successfully"
      PRINT CHR$(7);
      WaitForAction
      RETURN
      
      
DGExtractRecord:
      UpdCnt! = UpdCnt! + 1
      LOCATE 14, 16: PRINT STRING$(30, 32);
      LOCATE 13, 16: COLOR 11
      PRINT "   Updated Count:  "; : COLOR 15: PRINT UpdCnt!
      COLOR 11
      
      'New Way To Activate
      IF QPTrim$(UBDGGetRDRec(1).NewRdng) = "" THEN RETURN
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
      RETURN
      
      
      RETURN
      
SysCom:
      RETURN
      
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
        
        IdxRecLen = 4           'we are using a integer
        IdxFileSize& = FileSize&("UBCUSTBK.IDX")
        IdxNumOfRecs = IdxFileSize& \ IdxRecLen
        
        REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
        FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs              'load it
        
        Cnt = 1
        
        DO
          IF Cnt < 1 THEN Cnt = 1               ' Do Not Allow to Fall Below 1
          inputting = False     ' Set Edit Finish to No
          Prec& = IndexArray(Cnt).RecNum
          IF NOT (Prec&) = 0 THEN
            GOSUB GetCustRecNo
            IF process THEN
              GOSUB UpdateMeter
            END IF
          END IF
        LOOP UNTIL Cnt > NumOfRecs
        
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
      
      IdxRecLen = 4             'we are using a integer
      IdxFileSize& = FileSize&("UBCUSTBK.IDX")
      IdxNumOfRecs = IdxFileSize& \ IdxRecLen
      
      REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
      FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs                'load it
      
      Cnt = 1
      
      DO
        IF Cnt < 1 THEN
          Cnt = 1
          PageUp = False        ' Do Not Allow to Fall Below 1
        END IF
        inputting = False
        Prec& = IdxBuff(Cnt).RecNum
        QPrintRC STR$(Prec&), 25, 40, 112
        IF NOT (Prec&) = 0 THEN
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
      
      FOR MtrCnt = 1 TO 7       'find last active meter
        
        inputting = False
        TempRev$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MTRType)
        
        IF LEN(TempRev$) <> 0 THEN
          
          DisplayUBScrn ScrnName$
          PrintHelp Help$
          
          REM set information as known
          Form$(1, 0) = RIGHT$(STR$(BookNumber), 2)
          Form$(2, 0) = UBCustRec(1).CUSTNAME
          Form$(3, 0) = UBCustRec(1).Book
          Form$(4, 0) = UBCustRec(1).SEQNUMB
          Form$(5, 0) = STR$(Prec&)
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
            Form$(11, 0) = "0"  ' Demand Always = 0 for Previous Read
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
              
            CASE F5KEY          'Close Down and Exit
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
                UBCustRec(1).LocMeters(MtrCnt).PrevRead = 0     ' Demand Always = 0 for Previous Read
              END IF
              
              FOpenS "UBCUST.DAT", C2Handle     'open data file
              FPutRTA C2Handle, UBCustRec(1), Prec&, UBCustRecLen
              FClose C2Handle

              inputting = True
              
            CASE EscKey
              inputting = True
              
              'CASE IS > 0
              'QPrintRC CHR$(Frm(1).KeyCode), 1, 1, -1
              
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
      FOpenS "UBCUST.DAT", C2Handle             'open data file
      FGetRTA C2Handle, UBCustRec(1), Prec&, UBCustRecLen
      FClose C2Handle
      
      IF UBCustRec(1).Status <> "A" AND NoInactive THEN
        process = False
        IF PageUp THEN
          Cnt = Cnt - 1
        ELSE
          Cnt = Cnt + 1
        END IF
        GOTO NoInactRet
      END IF
      IF VAL(UBCustRec(1).Book) = BookNumber AND (UBCustRec(1).Status <> "F") THEN
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
      
      RecNo& = Prec&
      
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

SUB ProcessItron
  SHARED Choice$()
  DIM Route(10)
  REDIM UBItronRec(1) AS UBItronRecType
  REDIM UBItronFrec(1) AS ItronFType
  REDIM UBItronArec(1) AS ItronAType
  REDIM UBItronBrec(1) AS ItronBType
  REDIM UBItronHrec(1) AS ItronHType
  REDIM UBItronZrec(1) AS ItronZType

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
  LOCATE 1, 1, 0

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
  LOCATE 1, 1, 0

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

    CASE F10Key
    ' First Check For Valid Interrogator Number and Routes
          GOSUB FILENAM1Itron
         'Set Route Number Values
          FOR Cnt! = 1 TO NumberofRoutes
            Route(Cnt!) = VAL(Form$(Cnt! + 2, 0))
          NEXT Cnt!

         'Open Customer Data File
          REDIM UBCustRec(1) AS NewUBCustRecType
          UBCustRecLen = LEN(UBCustRec(1))
          UBFile = FREEFILE
          OPEN "UBCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBFile LEN = UBCustRecLen
         
         'Open Itron Date File
          UBItronRdRecLen = 80          '80 byte records
          UBItronRdFile = FREEFILE
          OPEN FileName$ FOR RANDOM SHARED AS UBItronRdFile LEN = 80
          CLOSE UBItronRdFile: KILL FileName$
          UBItronRdFile = FREEFILE
          OPEN FileName$ FOR RANDOM SHARED AS UBItronRdFile LEN = 80
          NumItronRdRecs = 0

         'Open the Location Order for Reading
          IdxRecLen = 4         'we are using a integer
          IdxFileSize& = FileSize&("UBCUSTBK.IDX")
          IdxNumOfRecs = IdxFileSize& \ IdxRecLen
          REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
          FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs            'load it
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
                     UBItronFrec(1).RecordCode = "F"
                     UBItronFrec(1).Route = UBCustRec(1).Book + "000000"
                     UBItronFrec(1).Message = ""
                     UBItronFrec(1).Filler = ""
                     UBItronFrec(1).CRLF = CHR$(13) + CHR$(10)
                     PUT #UBItronRdFile, 1, UBItronFrec(1)
                    END IF

                    'Compare to Check For End of Book Run
                    IF FirstOnRoute < VAL(UBCustRec(1).Book) THEN
                      'Create Z Record Trailor
                      Route$ = LTRIM$(STR$(FirstOnRoute))
                      IF LEN(Route$) = 1 THEN Route$ = "0" + Route$ + "000000"

                      NumberofAccounts$ = LTRIM$(STR$(NumberofAccounts))
                      IF LEN(NumberofAccounts$) < 4 THEN NumberofAccounts$ = STRING$(4 - LEN(NumberofAccounts$), "0") + NumberofAccounts$

                      UBItronZrec(1).RecordCode = "Z"
                      UBItronZrec(1).Route = Route$
                      UBItronZrec(1).NumberAccts = NumberofAccounts$
                      UBItronZrec(1).NumberMeters = NumberofAccounts$
                      UBItronZrec(1).Filler = ""
                      UBItronZrec(1).CRLF = CHR$(13) + CHR$(10)
                      PUT #UBItronRdFile, (LOF(UBItronRdFile) / 80) + 1, UBItronZrec(1)

                      'Create Next Route Header F
                      FirstOnRoute = VAL(UBCustRec(1).Book)
                      NumberofAccounts = 0
                      UBItronFrec(1).RecordCode = "F"
                      UBItronFrec(1).Route = UBCustRec(1).Book + "000000"
                      UBItronFrec(1).Message = ""
                      UBItronFrec(1).Filler = ""
                      UBItronFrec(1).CRLF = CHR$(13) + CHR$(10)
                      PUT #UBItronRdFile, (LOF(UBItronRdFile) / 80) + 1, UBItronFrec(1)
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

                      UBItronZrec(1).RecordCode = "Z"
                      UBItronZrec(1).Route = Route$
                      UBItronZrec(1).NumberAccts = NumberofAccounts$
                      UBItronZrec(1).NumberMeters = NumberofAccounts$
                      UBItronZrec(1).Filler = ""
                      UBItronZrec(1).CRLF = CHR$(13) + CHR$(10)
                      PUT #UBItronRdFile, (LOF(UBItronRdFile) / 80) + 1, UBItronZrec(1)


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
  CustomerType$ = LEFT$(UBCustRec(1).CUSTTYPE, 1)
  IF LEN(CustomerType$) = 0 THEN CustomerType$ = "R"

 'Set Customer Name
  CustomerName$ = QPTrim$(UBCustRec(1).CUSTNAME)
 'Set Customer Address
  CustomerAddress$ = QPTrim$(UBCustRec(1).SERVADDR)


 IF (ASC(UBCustRec(1).LocMeters(MtrCnt!).MTRType) > 32) THEN
    IF UBCustRec(1).LocMeters(MtrCnt!).MTRType = "C" OR UBCustRec(1).LocMeters(MtrCnt!).MTRType = "W" OR UBCustRec(1).LocMeters(MtrCnt!).MTRType = "S" THEN

        'Determine High and Low Reading
         Average# = UBCustRec(1).LocMeters(MtrCnt!).AvgUse
         ILowRead$ = LTRIM$(STR$((UBCustRec(1).LocMeters(MtrCnt!).CurRead)))
         IF LEN(ILowRead$) > 8 THEN ILowRead$ = RIGHT$(ILowRead$, 8)
         IF LEN(ILowRead$) < 8 THEN ILowRead$ = STRING$(8 - LEN(ILowRead$), "0") + ILowRead$

         
         
         HighRead# = FIX(Average# * (HighVar / 100)) + UBCustRec(1).LocMeters(MtrCnt!).CurRead
         IF Average# = 0 THEN
          HighRead# = VAL(ILowRead$) + 150            'Fudge Factor Until History Built
         END IF
         
         IHighRead$ = LTRIM$(STR$(HighRead#))
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
         UBItronArec(1).RecordCode = "A"
         UBItronArec(1).Route = UBCustRec(1).Book + "000000"
         UBItronArec(1).AcctNumb = Account$
         UBItronArec(1).Geo = ""
         UBItronArec(1).SEQNUMB = SeqNumber$
         UBItronArec(1).Message = "0"
         UBItronArec(1).AType = CustomerType$
         UBItronArec(1).Filler = ""
         UBItronArec(1).CRLF = CHR$(13) + CHR$(10)
         PUT #UBItronRdFile, (LOF(UBItronRdFile) / 80) + 1, UBItronArec(1)
        'Create B Record
         UBItronBrec(1).RecordCode = "B"
         UBItronBrec(1).CUSTNAME = CustomerName$
         UBItronBrec(1).CustAddr = CustomerAddress$
         UBItronBrec(1).Filler = ""
         UBItronBrec(1).CRLF = CHR$(13) + CHR$(10)
         PUT #UBItronRdFile, (LOF(UBItronRdFile) / 80) + 1, UBItronBrec(1)
        'Create H Record
         UBItronHrec(1).RecordCode = "H"
         UBItronHrec(1).MeterNumb = MeterID$
         UBItronHrec(1).Dials = "4"
         UBItronHrec(1).LowRead = ILowRead$
         UBItronHrec(1).HighRead = IHighRead$
         UBItronHrec(1).LastRead = ILowRead$
         UBItronHrec(1).MeterType = "W"
         UBItronHrec(1).Multiplier = "000000"
         UBItronHrec(1).NoMths = "0"
         UBItronHrec(1).MtrMake = ""
         UBItronHrec(1).DispCode = "  "
         UBItronHrec(1).NumbDec = "0"
         UBItronHrec(1).MustRead = "1"
         UBItronHrec(1).Status = "A"
         UBItronHrec(1).Filler = ""
         UBItronHrec(1).CRLF = CHR$(13) + CHR$(10)
         PUT #UBItronRdFile, (LOF(UBItronRdFile) / 80) + 1, UBItronHrec(1)

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
  LOCATE 1, 1, 0

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
        GOTO GetItron
        Frm(1).FldNo = 1
        Action = 1
        GOTO GetItron
      ELSE
        RouteID$ = Form$(1, 0)
        GOSUB FileNamItron           ' Get Itron File
        'Open Customer Data File
        REDIM UBCustRec(1) AS NewUBCustRecType
        UBCustRecLen = LEN(UBCustRec(1))
        UBFile = FREEFILE
        OPEN "UBCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBFile LEN = UBCustRecLen

        'REDIM UBLogGetRdRec(1) AS UBItronGetReadRecType
        'UBLogGetRdRecLen = LEN(UBLogGetRdRec(1))

        UBLogGetRdFile = FREEFILE
        OPEN FileName$ FOR RANDOM ACCESS READ WRITE SHARED AS UBLogGetRdFile LEN = UBLogGetRdRecLen
        NumLogGetRdRecs = LOF(UBLogGetRdFile) / UBLogGetRdRecLen

        IF NumLogGetRdRecs = 0 THEN
          LOCATE 12, 20: COLOR 12
          PRINT "            [  E R R O R  ]"
          LOCATE 13, 20: COLOR 11
          PRINT "     Nothing Found From Itron "
          LOCATE 14, 20: COLOR 11
          PRINT "   Please Reenter Route ID Number "
          BEEP
          WaitForAction
          CLOSE
          GOTO GetItron
        END IF

        Cnt! = 1                ' Initialize File Counter to 1
        DO
           ' GET UBLogGetRdFile, Cnt!, UBLogGetRdRec(1)
           ' Prec& = VAL(QPTrim$(UBLogGetRdRec(1).AcctNo))

 '         IF LEFT$(UBLogGetRdRec(1).CurRead, 1) <> "X" THEN
           IF NOT (Prec&) = 0 THEN
            GET UBFile, Prec&, UBCustRec(1)
            GOSUB ExtractRecordItron
           END IF
   '       END IF
          Cnt! = Cnt! + 1

        LOOP UNTIL Cnt! > NumLogGetRdRecs

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
      END IF
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


  'New Way To Activate
    MeterRecord = 1     'Only 1 meter per account for Itron
  ' Update Meter W/Reading
'    CurReading# = VAL(UBLogGetRdRec(1).CurRead)
'    MeterReadDate$ = MID$(UBLogGetRdRec(1).ReadDate, 3, 2) + "/" + MID$(UBLogGetRdRec(1).ReadDate, 5, 2) + "/" + RIGHT$(DATE$, 4)

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
  LOCATE 1, 1, 0

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
  LOCATE 1, 1, 0

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
            MeterID$ = UBCustRec(1).Book + UBCustRec(1).SEQNUMB
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
  LOCATE 1, 1, 0

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
        GOSUB FileNamLC           ' Get Logicon File
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
    MeterRecord = 1     'Only 1 meter per account for logicon
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
  PUT UBFile, Prec&, UBCustRec(1)
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
              IdxRecLen = 4     'we are using a integer
              MakeSequenceIndex "Sequence Number"
              IdxNumOfRecs = FileSize&("UBTEMP.IDX") \ 4
              REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
              FGetAH "UBTEMP.IDX", IdxBuff(1), 4, IdxNumOfRecs
            ELSE
              IdxRecLen = 4     'we are using a integer
              IdxFileSize& = FileSize&("UBCUSTBK.IDX")
              IdxNumOfRecs = IdxFileSize& \ IdxRecLen
              REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
              FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs        'load it
            END IF
            Cnt! = 1
            
            DO
              
              Prec& = IdxBuff(Cnt!).RecNum
              IF NOT (Prec&) = 0 THEN
                GET UBFile, Prec&, UBCustRec(1)
                DisplayAcct$ = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
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
      
      Cnt! = 1  ' Initialize File Counter to 1
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
        
        IF ASC(UBCustRec(1).LocMeters(MtrCnt!).MTRType) > 32 THEN
          MTRType$ = UBCustRec(1).LocMeters(MtrCnt!).MTRType
          IF MTRType$ = "W" OR MTRType$ = "S" OR MTRType$ = "C" OR MTRType$ = "E" OR MTRType$ = "D" OR MTRType$ = "G" THEN
            SELECT CASE WhatTypes$
            CASE "W"            'water/sewer
              IF MTRType$ = "W" OR MTRType$ = "S" OR MTRType$ = "C" THEN
                MeterOK = True
              END IF
            CASE "E"            'electric & demand elec.
              IF MTRType$ = "E" OR MTRType$ = "D" THEN
                MeterOK = True
              END IF
            CASE "G"            'gas
              IF MTRType$ = "G" THEN
                MeterOK = True
              END IF
            CASE "A", " "       'all meters
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
                MeterID$ = UBCustRec(1).Book + UBCustRec(1).SEQNUMB
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

SUB Search4LNumber (LocNum$, RecNo&, CLSFlag%)
      
      REDIM UBCustRec(1) AS NewUBCustRecType
      UBCustRecLen = LEN(UBCustRec(1))
      
      'REDIM UBLocatRec(1) AS UBLocationRecType
      'UBLocatRecLen = LEN(UBLocatRec(1))
      
      IdxRecLen = 4             'we are using a integer
      IdxFileSize& = FileSize&("UBCUSTBK.IDX")
      IdxNumOfRecs = IdxFileSize& \ IdxRecLen
      
      REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
      
      FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs                'load it
      
      SearchLen = LEN(LocNum$)
      
      Match = False
      FirstRec = 1
      LastRec = IdxNumOfRecs
      
      BotOffSet = 0
      TopOffSet = IdxNumOfRecs
      
      FOpenS "UBCUST.DAT", C1Handle             'open data file
      'FOpenS "UBLOCA.DAT", L1Handle               'open data file
      
      'Find matching record
      
      MidRec = (LastRec + FirstRec) \ 2
      
      DO
        IF LastSRec = MidRec THEN EXIT DO
        LastSRec = MidRec
        
        FGetRTA C1Handle, UBCustRec(1), CLNG(IdxBuff(MidRec).RecNum), UBCustRecLen
        
        UBSearchN$ = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
        
        IF LocNum$ = UBSearchN$ THEN
          
          IF MidRec - BotOffSet > 1 THEN
            MidRec = MidRec - 1
          ELSE
            FirstMatchRec = MidRec
          END IF
        ELSEIF LocNum$ < UBSearchN$ THEN        'lower
          TopOffSet = MidRec
          MidRec = TopOffSet - ((TopOffSet - BotOffSet) \ 2)
        ELSE    'higher
          BotOffSet = MidRec
          MidRec = BotOffSet + ((TopOffSet - BotOffSet) \ 2) + 1
          IF MidRec = IdxNumOfRecs + 1 THEN
            EXIT DO
          END IF
        END IF
        IF TopOffSet = BotOffSet THEN EXIT DO
      LOOP UNTIL FirstMatchRec
      
      FClose C1Handle
      FClose L1Handle
      
      IF FirstMatchRec = 0 THEN
        RecNo& = 0
      ELSE
        RecNo& = IdxBuff(FirstMatchRec).RecNum
      END IF
      
      
END SUB

