DECLARE SUB PrintHighLowReport ()
DECLARE SUB PrintMeterNotes ()
DEFINT A-Z
DECLARE SUB PrintMONList ()
DECLARE SUB EstMeterReading ()
DECLARE SUB PrintMeterReport ()
DECLARE SUB PrintMeterList ()
DECLARE SUB HandHeldProcessing ()
DECLARE SUB LoadUBSetUpFile (UBSetupRec() AS ANY, UBSetUpLen%)
DECLARE SUB PrintUtilBills ()
DECLARE SUB MakeSequenceIndex (IndexText$)
  'DECLARE SUB MakeBillFile (AbortFlag%)
DECLARE SUB PrintRptFile (RptTitle$, filename$, LPTPort%, RetCode%, EntryPoint%)
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
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE SUB ShowCursor ()
DECLARE SUB HideCursor ()
DECLARE SUB ClearScrn ()
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE FUNCTION FindRateTbl% (RATECODE$, NumOfRates%, UBRateTbls() AS ANY)
DECLARE FUNCTION GetNumOfRevs% ()
  'DECLARE FUNCTION GetRevCharge# (RateTbl AS ANY, TMeterConsp&)
DECLARE FUNCTION Round# (N#)
DECLARE FUNCTION GetNumOfCust% ()
DECLARE FUNCTION GetNumRateRecs% ()
DECLARE FUNCTION Exist% (filename$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION FileSize& (filename$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION FmtBook$ (BOOK$)
DECLARE FUNCTION FmtSeqN$ (SeqN$)
  
  '$INCLUDE: 'DefCnf.BI'
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)
DECLARE SUB FGetAH (filename$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FOpenS (filename$, Handle%)
DECLARE SUB FClose (Handle%)
DECLARE SUB FCreate (filename$)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutAH (filename$, SEG Element AS ANY, ElSize%, NumEls%)
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
  
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'ubdraft.BI'
  '$INCLUDE: 'ubsetup.bi'
  
  '$INCLUDE: 'PageInfo.bi'
  '$INCLUDE: 'Newcust.bi'
  ''$INCLUDE: 'ubcust.bi'
  ''$INCLUDE: 'ubloc.bi'
  '$INCLUDE: 'ubrate.BI'
  '$INCLUDE: 'ubsetbal.BI'
  
  CONST False = 0, True = NOT False

'GOTO Skip2Here

  REDIM MChoice$(1 TO 8)
  
  MChoice$(1) = "Manual Meter Reading Entry "
  MChoice$(2) = "Hand-Held Meter Reading Entry"
  MChoice$(3) = "Estimated Meter Reading Entry"
  MChoice$(4) = "Print Meter Reading List"
  MChoice$(5) = "Print MON Meter Reading List"
  MChoice$(6) = "Print Meter Reading Notes"
  MChoice$(7) = "Meter Reading High/Low Report"
  MChoice$(8) = "Exit to OS"
  
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
  Row = ((25 - (UBOUND(MChoice$))) \ 2)
  Col = ((80 - MaxLen) \ 2) - 2
  'Help$ = "Utility Billing Main Menu"
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    
    BlockClear
    TitleBox 3, Col, MaxLen + 3, "Meter Reading Menu ", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
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
      PrintMONList
      'PrintMeterReport
    CASE 6
      PrintMeterNotes
    CASE 7
      PrintHighLowReport

    CASE 8
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP

  
  'HideCursor
  'ClearScrn
  'END
  RUN "ubbillin"

'Skip2Here:
'    CLS
'    REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
'    UBCustRecLen = LEN(UBCustRec(1))
'
'    UBCustFile = FREEFILE
'    OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCustFile LEN = UBCustRecLen
'    NumOfCust = LOF(UBCustFile) \ UBCustRecLen
'
'    FOR LCnt& = 1 TO NumOfCust
'      GET #UBCustFile, LCnt&, UBCustRec(1)
'      PRINT LCnt&; "  ";
'      Book = QPValI(UBCustRec(1).Book)
'      FOR MtrCnt& = 1 TO 7  'find last active meter
'        IF UBCustRec(1).LocMeters(MtrCnt&).UseCnt > 0 THEN
'          IF MtrCnt& > 2 THEN STOP
'          PRINT UBCustRec(1).LocMeters(MtrCnt&).UseCnt; "  "; UBCustRec(1).LocMeters(MtrCnt&).AvgUse, ;
'          PRINT "Avg:"; ((UBCustRec(1).LocMeters(MtrCnt&).AvgUse) / UBCustRec(1).LocMeters(MtrCnt&).UseCnt)
'        END IF
'      NEXT
'      PRINT
'
'    zz$ = INPUT$(1)
'    IF zz$ = CHR$(27) THEN EXIT FOR
'    NEXT
'    CLOSE
'    END

SUB EstMeterReading

  REDIM TempScrn(0)
  LibName$ = "UB"
  ScrnName$ = "BOOK2EST"

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
        ok = MsgBox%("UB", "BADEINFO")
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
            IF UBCustRec(1).LocMeters(zz).UseCnt > 0 THEN
              DidEM = True
              'UBCustRec(1).LocMeters(zz).PrevRead = UBCustRec(1).LocMeters(zz).CurRead
              UBCustRec(1).LocMeters(zz).PastDate = UBCustRec(1).LocMeters(zz).CurDate
              UBCustRec(1).LocMeters(zz).ReadFlag = "E"
              'AvgUse& = ((UBCustRec(1).LocMeters(zz).AvgUse) / UBCustRec(1).LocMeters(zz).UseCnt)
              'UBCustRec(1).LocMeters(zz).CurRead = UBCustRec(1).LocMeters(zz).CurRead + AvgUse&
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
  BOOK = VAL(Form$(2, 0))
  IF ESTDate < 0 OR BOOK < 1 THEN
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
  REDIM UBSetupRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetupRec(), UBSetUpLen
  
  IF UBSetUpLen = 0 THEN
    PRINT CHR$(7);
    SaveScrn TempScrn()
    LOCATE , , 0
    DisplayUBScrn "UBNOHH"
    WaitForAction
    RestScrn TempScrn()
    LOCATE , , 1
    CLOSE
    EXIT SUB
  END IF
  
  IF UBSetupRec(1).HHDEVICE = "N" OR NOT (UBSetupRec(1).HHDEVICE = "S" OR UBSetupRec(1).HHDEVICE = "P" OR UBSetupRec(1).HHDEVICE = "C" OR UBSetupRec(1).HHDEVICE = "R") THEN
    PRINT CHR$(7);
    SaveScrn TempScrn()
    LOCATE , , 0
    DisplayUBScrn "UBNOHH"
    WaitForAction
    RestScrn TempScrn()
    LOCATE , , 1
    CLOSE
    EXIT SUB
  END IF
  
  ' Valid HandHeld Type Found
  
  HighVar = UBSetupRec(1).HighRead
  LowVar = UBSetupRec(1).LowRead
  
  IF UBSetupRec(1).HHDEVICE = "S" THEN
    GOSUB Sensus: EXIT SUB
  END IF
  
  IF UBSetupRec(1).HHDEVICE = "P" THEN
    GOSUB PC3000: EXIT SUB
  END IF
  
  IF UBSetupRec(1).HHDEVICE = "R" THEN
    GOSUB Radix: EXIT SUB
  END IF
  
  IF UBSetupRec(1).HHDEVICE = "C" THEN
    GOSUB SysCom: EXIT SUB
  END IF
  
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
  help$ = "Sensus Hand Held Processing"
  LOCATE 1, 1, 0
  
  help$ = "Choose HandHeld Operation"
  
  DisplayUBScrn ScrnName$
  PrintHelp help$
  
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
        COLOR 12: LOCATE 11, 40: PRINT "Invalid Operation": BEEP:
        WaitForAction
        COLOR 15: LOCATE 11, 40: PRINT STRING$(25, 32)
        
        Frm(1).FldNo = 1
        Action = 1
        DONE = False
      ELSE
        PathWay$ = Form$(2, 0)
        UBSensusRec(1).PathWay = PathWay$
        PUT UBSensusFile, 1, UBSensusRec(1)
        CLOSE UBSensusFile
        IF LEFT$(Form$(1, 0), 1) = "S" THEN GOSUB SendSensus
        IF LEFT$(Form$(1, 0), 1) = "G" THEN GOSUB GetSensus
        DONE = True
      END IF
    CASE ESC
      DONE = True
    CASE ELSE
      DONE = False
    END SELECT
  LOOP UNTIL DONE
  
  RETURN
  
SendSensus:
  LibName$ = "UB"
  ScrnName$ = "UBSSHH1S"
  help$ = "Sensus Hand Held Processing"
  LOCATE 1, 1, 0
  
  help$ = "DownLoad HandHeld Operation"
  DisplayUBScrn ScrnName$
  PrintHelp help$
  
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
        Frm(1).FldNo = 1: Action = 1: DONE = False
        
      ELSE
        
        IF NumberofRoutes > 0 THEN
          INTERROGATOR = VAL(Form$(1, 0))
          COLOR 11: LOCATE 12, 18: PRINT "LOADING THESE ROUTES !!!                      "
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
          
          'Open Location Data File
          'REDIM UBLocRec(1) AS UBLocationRecType
          'UBLocRecLen = LEN(UBLocRec(1))
          'UBLocFile = FREEFILE
          'OPEN "UBLOCA.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBLocFile LEN = UBLocRecLen
          'NumOfLocRecs = LOF(UBLocFile) / UBLocRecLen
          'Open and Initialize the Sensus Read Information File
          
          REDIM UBSenRdRec(1) AS UBSensusReadRecType
          UBSenRdRecLen = LEN(UBSenRdRec(1))
          UBSenRdFile = FREEFILE
          OPEN filename$ FOR RANDOM SHARED AS UBSenRdFile LEN = UBSenRdRecLen
          CLOSE UBSenRdFile: KILL filename$
          UBSenRdFile = FREEFILE
          OPEN filename$ FOR RANDOM SHARED AS UBSenRdFile LEN = UBSenRdRecLen
          NumSenRdRecs = LOF(UBSenRdFile) / UBSenRdRecLen
          'Open the Location Order for Reading
          IdxRecLen = 4         'we are using a integer
          IdxFileSize& = FileSize&("UBCUSTBK.IDX")
          IdxNumOfRecs = IdxFileSize& \ IdxRecLen
          REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
          FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs            'load it
          
          Cnt! = 1
          
          DO
            
            pREC& = IdxBuff(Cnt!).RecNum
            IF NOT (pREC&) = 0 THEN
              GET UBFile, pREC&, UBCustRec(1)
              FOR BookCnt! = 1 TO NumberofRoutes
                IF VAL(UBCustRec(1).BOOK) = Route(BookCnt!) AND (UBCustRec(1).Status <> "F") THEN
                  GOSUB WriteRecord
                END IF
              NEXT BookCnt!
            END IF
            Cnt! = Cnt! + 1
            
          LOOP UNTIL Cnt! > IdxNumOfRecs
          
          CLOSE
          DONE = True
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
        DONE = True
      END IF
      
    CASE ESC
      DONE = True
    CASE ELSE
      DONE = False
    END SELECT
  LOOP UNTIL DONE
  RETURN
  
  
WriteRecord:
  'May Have Up to 10 Meters to Read
  MtrCnt! = 1
  Account$ = STR$(pREC&)
  Account$ = LEFT$(Account$, 6) + "-" + RIGHT$(STR$(MtrCnt!), 1)
  
  WHILE MtrCnt! < 8
    IF (ASC(UBCustRec(1).LocMeters(MtrCnt!).MTRType) > 32) THEN
      IF UBCustRec(1).LocMeters(MtrCnt!).MTRType = "C" OR UBCustRec(1).LocMeters(MtrCnt!).MTRType = "W" OR UBCustRec(1).LocMeters(MtrCnt!).MTRType = "T" OR UBCustRec(1).LocMeters(MtrCnt!).MTRType = "S" THEN
        ' Determine Sensus Meter Type
        IF UBCustRec(1).LocMeters(MtrCnt!).MTRType = "T" THEN SensusType$ = "B" ELSE SensusType$ = "M"
        ' Determine High and Low Reading
        Average# = UBCustRec(1).LocMeters(MtrCnt!).AvgUse
        ilowread$ = RIGHT$(STR$((UBCustRec(1).LocMeters(MtrCnt!).CurRead)), 8)
        ilowread# = FIX(ilowread#)
        HighRead# = FIX(Average# * (HighVar / 100)) + UBCustRec(1).LocMeters(MtrCnt!).CurRead
        ihighread$ = STR$(HighRead#)
        ihighread$ = RIGHT$(ihighread$, 8)
        MeterID$ = LTRIM$(UBCustRec(1).LocMeters(MtrCnt!).MTRNUM)
        MeterID$ = RTRIM$(MeterID$)
        
        IF VAL(MeterID$) = 0 THEN
          MeterID$ = UBCustRec(1).BOOK + UBCustRec(1).SEQNUMB
        END IF
        IF LEN(MeterID$) < 8 THEN
          MeterID$ = STRING$(8 - LEN(MeterID$), "0") + MeterID$
        END IF
        MeterID$ = RIGHT$(MeterID$, 8)
        
        'Set Record Fields and Put On Disk
        UBSenRdRec(1).ServAddress = UBCustRec(1).SERVADDR
        UBSenRdRec(1).MeterID = MeterID$
        UBSenRdRec(1).LowRead = ilowread$
        UBSenRdRec(1).HighRead = ihighread$
        UBSenRdRec(1).Account = Account$
        UBSenRdRec(1).SensusType = SensusType$
        PUT UBSenRdFile, (LOF(UBSenRdFile) / UBSenRdRecLen) + 1, UBSenRdRec(1)
      END IF
    END IF
    
    MtrCnt! = MtrCnt! + 1
  WEND
  RETURN
  
FILENAM1:
  
  PathWay$ = QPTrim$(PathWay$)
  IF INTERROGATOR = 1 THEN filename$ = PathWay$ + "ssi001.rte": RETURN
  IF INTERROGATOR = 2 THEN filename$ = PathWay$ + "ssi002.rte": RETURN
  IF INTERROGATOR = 3 THEN filename$ = PathWay$ + "ssi003.rte": RETURN
  IF INTERROGATOR = 4 THEN filename$ = PathWay$ + "ssi004.rte": RETURN
  IF INTERROGATOR = 5 THEN filename$ = PathWay$ + "ssi005.rte": RETURN
  IF INTERROGATOR = 6 THEN filename$ = PathWay$ + "ssi006.rte": RETURN
  IF INTERROGATOR = 7 THEN filename$ = PathWay$ + "ssi007.rte": RETURN
  IF INTERROGATOR = 8 THEN filename$ = PathWay$ + "ssi008.rte": RETURN
  IF INTERROGATOR = 9 THEN filename$ = PathWay$ + "ssi009.rte": RETURN
  RETURN
  
GetSensus:
  
  LibName$ = "UB"
  ScrnName$ = "UBSSHH1G"
  help$ = "Sensus Hand Held Processing"
  LOCATE 1, 1, 0
  
  help$ = "Upload HandHeld Operation"
  DisplayUBScrn ScrnName$
  PrintHelp help$
  
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
        
        'Open Location Data File
        'REDIM UBCustRec(1) AS UBLocationRecType
        'UBLocRecLen = LEN(UBLocRec(1))
        'UBLocFile = FREEFILE
        'OPEN "UBLOCA.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBLocFile LEN = UBLocRecLen
        'NumOfLocRecs = LOF(UBLocFile) / UBLocRecLen
        'Open and Initialize the Sensus Read Information File
        
        REDIM UBSenGetRdRec(1) AS UBSensusGetReadRecType
        UBSenGetRdRecLen = LEN(UBSenGetRdRec(1))
        UBSenGetRdFile = FREEFILE
        OPEN filename$ FOR RANDOM ACCESS READ WRITE SHARED AS UBSenGetRdFile LEN = UBSenGetRdRecLen
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
          GET UBSenGetRdFile, Cnt!, UBSenGetRdRec(1)
          REM old way
          LocationNumber$ = UBSenGetRdRec(1).Account
          kk = INSTR(LocationNumber$, "/")
          IF kk > 0 THEN
            LocationNumber$ = LEFT$(LocationNumber$, kk - 1)
          END IF
          
          GOSUB FindAcct
          REM end old way
          ' New Way Undelete When Our System Sends and Retrieves
          pREC& = VAL(QPTrim$(UBSenGetRdRec(1).Account))
          IF NOT (pREC&) = 0 THEN
            GET UBFile, pREC&, UBCustRec(1)
            GOSUB ExtractRecord
          END IF
          Cnt! = Cnt! + 1
          
        LOOP UNTIL Cnt! > NumSenGetRdRecs
        
        CLOSE
        DONE = True
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
      DONE = True
      
    CASE ESC
      DONE = True
    CASE ELSE
      DONE = False
    END SELECT
  LOOP UNTIL DONE
  RETURN
  
  
ExtractRecord:
  UpdCnt! = UpdCnt! + 1
  LOCATE 13, 20: COLOR 11
  PRINT " Updated Count:  "; : COLOR 15: PRINT UpdCnt!
  COLOR 11
  
  
  'New Way To Activate
  MeterRecord = VAL(RIGHT$((QPTrim$(UBSenGetRdRec(1).Account)), 1))
  
  ' Check Meter Updated Flag
  
  
  ' Update Meter W/Reading
  Curreading# = VAL(UBSenGetRdRec(1).Reading)
  MeterReadDate$ = Form$(2, 0)
  IF UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y" THEN
    UBCustRec(1).LocMeters(MeterRecord).CurRead = Curreading#
    UBCustRec(1).LocMeters(MeterRecord).CurDate = Date2Num(MeterReadDate$)
  ELSE
    UBCustRec(1).LocMeters(MeterRecord).PrevRead = UBCustRec(1).LocMeters(MeterRecord).CurRead
    UBCustRec(1).LocMeters(MeterRecord).PastDate = UBCustRec(1).LocMeters(MeterRecord).CurDate
    UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y"
    UBCustRec(1).LocMeters(MeterRecord).CurDate = Date2Num(MeterReadDate$)
    UBCustRec(1).LocMeters(MeterRecord).CurRead = Curreading#
  END IF
  PUT UBFile, pREC&, UBCustRec(1)
  RETURN
  
  
FileNam:
  PathWay$ = QPTrim$(PathWay$)
  IF INTERROGATOR = 1 THEN filename$ = PathWay$ + "exssi001.dat": RETURN
  IF INTERROGATOR = 2 THEN filename$ = PathWay$ + "exssi002.dat": RETURN
  IF INTERROGATOR = 3 THEN filename$ = PathWay$ + "exssi003.dat": RETURN
  IF INTERROGATOR = 4 THEN filename$ = PathWay$ + "exssi004.dat": RETURN
  IF INTERROGATOR = 5 THEN filename$ = PathWay$ + "exssi005.dat": RETURN
  IF INTERROGATOR = 6 THEN filename$ = PathWay$ + "exssi006.dat": RETURN
  IF INTERROGATOR = 7 THEN filename$ = PathWay$ + "exssi007.dat": RETURN
  IF INTERROGATOR = 8 THEN filename$ = PathWay$ + "exssi008.dat": RETURN
  IF INTERROGATOR = 9 THEN filename$ = PathWay$ + "exssi009.dat": RETURN
  RETURN
  
FindAcct:
  OK2Search = False
  LNum$ = QPTrim$(LocationNumber$)
  DashPos = INSTR(LNum$, "-")
  
  IF DashPos > 1 THEN
    BOOK$ = FmtBook$(LEFT$(LNum$, DashPos - 1))
    SeqN$ = FmtSeqN$(MID$(LNum$, DashPos + 1))
    LNum$ = BOOK$ + "-" + SeqN$
  ELSE
    BOOK$ = FmtBook$(LEFT$(LNum$, 2))
    SeqN$ = FmtSeqN$(MID$(LNum$, 3))
    LNum$ = BOOK$ + "-" + SeqN$
  END IF
  
  Search4LNumber LNum$, RecNo&, CLSFlag
  
  IF RecNo& > 0 THEN
    GOSUB GetCusRecNo
    pREC& = RecNo&
    RETURN
  END IF
  pREC& = 0
  RETURN
  
  
GetCusRecNo:
  '  REDIM UBLocatRec(1) AS UBLocationRecType
  '  UBLocatRecLen = LEN(UBLocatRec(1))
  '  FOpenS "UBLOCA.DAT", L1Handle 'open data file
  '  FGetRTA L1Handle, UBLocatRec(1), RecNo&, UBLocatRecLen
  '  FClose L1Handle
  '  RecNo& = UBLocatRec(1).CustRec
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
  help$ = "PC3000 Hand Held Processing"
  LOCATE 1, 1, 0
  
  help$ = "Choose HandHeld Operation"
  
  DisplayUBScrn ScrnName$
  PrintHelp help$
  
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
  Choice$(1, 0) = "Send Info to PC3000"
  Choice$(2, 0) = "Get Info From PC3000"
  
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
        DONE = False
      ELSE
        
        PathWay$ = Form$(2, 0)
        UBPC3000Rec(1).PathWay = PathWay$
        PUT UBPC3000File, 1, UBPC3000Rec(1)
        CLOSE UBPC3000File
        IF LEFT$(Form$(1, 0), 1) = "S" THEN GOSUB SendPC3000
        IF LEFT$(Form$(1, 0), 1) = "G" THEN GOSUB GetPC3000
        DONE = True
      END IF
    CASE ESC
      DONE = True
    CASE ELSE
      DONE = False
    END SELECT
  LOOP UNTIL DONE
  
  RETURN
  
SendPC3000:
  LibName$ = "UB"
  ScrnName$ = "UBPCHH1S"
  help$ = "PC3000 Hand Held Processing"
  LOCATE 1, 1, 0
  
  help$ = "DownLoad HandHeld Operation"
  DisplayUBScrn ScrnName$
  PrintHelp help$
  
  
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
        
        'Open Location Data File
        'REDIM UBLocRec(1) AS UBLocationRecType
        'UBLocRecLen = LEN(UBLocRec(1))
        'UBLocFile = FREEFILE
        'OPEN "UBLOCA.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBLocFile LEN = UBLocRecLen
        
        'NumOfLocRecs = LOF(UBLocFile) / UBLocRecLen
        
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
        
        'Open the Location Order for Reading
        IdxRecLen = 4           'we are using a integer
        IdxFileSize& = FileSize&("UBCUSTBK.IDX")
        IdxNumOfRecs = IdxFileSize& \ IdxRecLen
        REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
        FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs              'load it
        
        Cnt! = 1
        
        DO
          
          pREC& = IdxBuff(Cnt!).RecNum
          IF NOT (pREC&) = 0 THEN
            GET UBFile, pREC&, UBCustRec(1)
            DisplayAcct$ = "Processing Location #" + UBCustRec(1).BOOK + UBCustRec(1).SEQNUMB
            PrintHelp DisplayAcct$
            FOR BookCnt! = 1 TO NumberofRoutes
              IF VAL(UBCustRec(1).BOOK) = Route(BookCnt!) AND (UBCustRec(1).Status <> "F") THEN
                GOSUB PC3000WriteRecord
              END IF
            NEXT BookCnt!
          END IF
          Cnt! = Cnt! + 1
          
        LOOP UNTIL Cnt! > IdxNumOfRecs
        CLOSE
        CLS
        COLOR 11
        PRINT "PC 3000 Connect"
        COLOR 15
        PRINT "Make Sure the PC 3000 is in the Cradle or Connected by Data Cable"
        PRINT "The Screen Should be at Serial Communications in Progress"
        PRINT "You get to the Communications Screen by Pressing the 'L' key when"
        PRINT "the directory is being displayed."
        PRINT
        COLOR 11
        
        SHELL "CONNECT"
        destination$ = "copy ubcusttr.dat  " + PathWay$
        SHELL destination$
        PRINT
        PRINT "Press any key. . ."
        CALL WaitForAction
        
        DONE = True
        DisplayUBScrn ScrnName$
        FOR lc = 8 TO 10
          LOCATE lc, 16
          PRINT STRING$(40, 32)
        NEXT lc
        LOCATE 9, 18
        COLOR 11
        PRINT "Information Sent to PC3000 Successfully"
        LOCATE 12, 18: PRINT "THESE ROUTES LOADED!!!                      "
        PRINT CHR$(7);
        WaitForAction
      END IF
      DONE = True
      
      
    CASE ESC
      DONE = True
    CASE ELSE
      DONE = False
    END SELECT
  LOOP UNTIL DONE
  RETURN
  
  
PC3000WriteRecord:
  'Get Customer File for Name Information
  '  IF UBCustRec(1).CustRec < 1 OR UBCustRec(1).CustRec > LOF(UBFile) / LEN(UBCustRec(1)) THEN
  '    'do nothing
  '  ELSE
  '    GET UBFile, UBLocRec(1).CustRec, UBCustRec(1)
  '  END IF
  
  'May Have Up to 7 Meters to Read
  FOR MtrCnt! = 1 TO 7
    MeterOK = False
    Account$ = STR$(pREC&)
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
          
          ilowread$ = RIGHT$(STR$((UBCustRec(1).LocMeters(MtrCnt!).CurRead)), 8)
          LowRead# = FIX(ilowread#)
          HighRead# = FIX(Average# * (HighVar / 100)) + UBCustRec(1).LocMeters(MtrCnt!).CurRead
          
          MeterID$ = LTRIM$(UBCustRec(1).LocMeters(MtrCnt!).MTRNUM)
          MeterID$ = RTRIM$(MeterID$)
          MeterID$ = RIGHT$(MeterID$, 8)
          
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
          UBPC3000RDRec(1).Note1 = ""
          UBPC3000RDRec(1).Note2 = ""
          UBPC3000RDRec(1).Note3 = ""
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
  help$ = "Sensus Hand Held Processing"
  LOCATE 1, 1, 0
  
  help$ = "Upload HandHeld Operation"
  DisplayUBScrn ScrnName$
  PrintHelp help$
  
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
      PRINT "Make Sure the PC 3000 is in the Cradle or Connected by Data Cable"
      PRINT "The Screen Should be at Serial Communications in Progress"
      PRINT "You get to the Communications Screen by Pressing the 'L' key when"
      PRINT "the directory is being displayed."
      PRINT
      COLOR 11
      
      SHELL "CONNECT"
      
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
          PRINT " Nothing Found From PC3000 "
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
          pREC& = VAL(QPTrim$(UBPC3000GetRDRec(1).Account))
          
          IF NOT (pREC&) = 0 THEN
            GET UBFile, pREC&, UBCustRec(1)
            GOSUB PC3000ExtractRecord
          END IF
          Cnt! = Cnt! + 1
          
        LOOP UNTIL Cnt! > NumPC3000GetRdRecs
        
        CLOSE
        DONE = True
        FOR lc = 12 TO 14
          LOCATE lc, 16
          PRINT STRING$(40, 32)
        NEXT lc
        LOCATE 12, 18
        COLOR 11
        PRINT "Readings Updated From PC3000 Successfully"
        PRINT CHR$(7);
        WaitForAction
        DONE = True
        
      CASE ESC
        DONE = True
      CASE ELSE
        DONE = False
      END SELECT
    LOOP UNTIL DONE
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
    Curreading# = UBPC3000GetRDRec(1).CurRead
    
    IF UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y" THEN
      UBCustRec(1).LocMeters(MeterRecord).CurRead = Curreading#
      UBCustRec(1).LocMeters(MeterRecord).CurDate = UBPC3000GetRDRec(1).ReadDate
    ELSE
      UBCustRec(1).LocMeters(MeterRecord).PrevRead = UBCustRec(1).LocMeters(MeterRecord).CurRead
      UBCustRec(1).LocMeters(MeterRecord).PastDate = UBCustRec(1).LocMeters(MeterRecord).CurDate
      UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y"
      UBCustRec(1).LocMeters(MeterRecord).CurDate = UBPC3000GetRDRec(1).ReadDate
      UBCustRec(1).LocMeters(MeterRecord).CurRead = Curreading#
    END IF
    IF UBPC3000GetRDRec(1).NoteStatus = "P" THEN
      UBCustRec(1).NewNotes = True
      UBCustRec(1).HHMSG1 = UBPC3000GetRDRec(1).Note1
      UBCustRec(1).HHMSG2 = UBPC3000GetRDRec(1).Note2
      UBCustRec(1).HHMSG3 = UBPC3000GetRDRec(1).Note3
    END IF
    
    PUT UBFile, pREC&, UBCustRec(1)
    RETURN
    
    
Radix:
    RETURN
    
SysCom:
    RETURN
    
    
    
END SUB

SUB ManualMeterReading
    
    REDIM TempScrn(0)
    
    REDIM UBCustRec(1) AS NewUBCustRecType
    UBCustRecLen = LEN(UBCustRec(1))
    
    help$ = "Meter Reading Input"
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
    
    REDIM UBSetupRec(1) AS UBSetupRecType
    LoadUBSetUpFile UBSetupRec(), UBSetUpLen
    
    HighVar = UBSetupRec(1).HighRead
    LowVar = UBSetupRec(1).LowRead
    
    IF UBSetupRec(1).UseSeq = "Y" THEN
      MakeSequenceIndex "Sequence Number"
      NumOfRecs = FileSize&("UBTEMP.IDX") \ 4
      REDIM IndexArray(1 TO NumOfRecs) AS UBCustIndexRecType
      FGetAH "UBTEMP.IDX", IndexArray(1), 4, NumOfRecs
      
      DO
        IF Cnt < 1 THEN Cnt = 1 ' Do Not Allow to Fall Below 1
        inputting = False       ' Set Edit Finish to No
        pREC& = IndexArray(Cnt).RecNum
        IF NOT (pREC&) = 0 THEN
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
    
    IdxRecLen = 4               'we are using a integer
    IdxFileSize& = FileSize&("UBCUSTBK.IDX")
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    
    REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
    FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs  'load it
    
    Cnt = 1
    
    DO
      IF Cnt < 1 THEN
        Cnt = 1
        PageUp = False          ' Do Not Allow to Fall Below 1
      END IF
      inputting = False
      pREC& = IdxBuff(Cnt).RecNum
      
      IF NOT (pREC&) = 0 THEN
        GOSUB GetCustRecNo
        IF process THEN
          GOSUB UpdateMeter
        END IF
      END IF
    LOOP UNTIL Cnt > IdxNumOfRecs
    
    EXIT SUB
    'End of Location Order Processing *******************************************
    
    ' Update the Meter Reading Here *********************************************
UpdateMeter:
    
    help$ = "Record #" + STR$(Cnt) + " of " + STR$(IdxNumOfRecs)
    
    FOR MtrCnt = 1 TO 7         'find last active meter
      
      inputting = False
      TempRev$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MTRType)
      
      IF LEN(TempRev$) <> 0 THEN
        
        DisplayUBScrn ScrnName$
        PrintHelp help$
        
        REM set information as known
        Form$(1, 0) = RIGHT$(STR$(BookNumber), 2)
        Form$(2, 0) = UBCustRec(1).CUSTNAME
        Form$(3, 0) = UBCustRec(1).BOOK
        Form$(4, 0) = UBCustRec(1).SEQNUMB
        Form$(5, 0) = STR$(pREC&)
        Form$(6, 0) = UBCustRec(1).SERVADDR
        
        GOSUB GetMeterType
        
        Form$(7, 0) = MeterType$
        Form$(8, 0) = UBCustRec(1).LocMeters(MtrCnt).MTRNUM
        Form$(9, 0) = MtrReadDate$
        
        IF UBCustRec(1).LocMeters(MtrCnt).ReadFlag = "Y" THEN
          Form$(10, 0) = STR$(UBCustRec(1).LocMeters(MtrCnt).CurRead)
          Form$(11, 0) = STR$(UBCustRec(1).LocMeters(MtrCnt).PrevRead)
        ELSE
          Form$(11, 0) = STR$(UBCustRec(1).LocMeters(MtrCnt).CurRead)
        END IF
        
        Form$(13, 0) = "N"
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
        
        
        Frm(1).FldNo = 10
        Action = 1
        
        DO
          EditForm Form$(), Fld(), Frm(1), Cnf, Action
          IF Frm(1).PrevFld = 10 AND Frm(1).FldNo > 10 THEN
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
            IF AvgCnt! > 0 AND Average# > 0 THEN
              IF Calcusage# > (Average# * (HighVar / 100)) THEN
                HighRead = True
                LOCATE 17, 48: COLOR 12: PRINT "<-- HIGH READ": BEEP
              END IF
              IF Calcusage# < (Average# * (LowVar / 100)) THEN
                LOCATE 17, 48: COLOR 12: PRINT "<--LOW READ": BEEP
                LowRead = True
              END IF
            END IF
            Action = 1
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
            
            
          CASE F5KEY            'Close Down and Exit
            EXIT SUB
            
          CASE F7KEY
            PRINT CHR$(7);
            SaveScrn TempScrn()
            LOCATE , , 0
            
            DisplayUBScrn "UBNOEST"
            WaitForAction
            RestScrn TempScrn()
            LOCATE , , 1
            
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
            END IF
            
            FOpenS "UBCUST.DAT", C2Handle       'open data file
            FPutRTA C2Handle, UBCustRec(1), pREC&, UBCustRecLen
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
    FOpenS "UBCUST.DAT", C2Handle               'open data file
    FGetRTA C2Handle, UBCustRec(1), pREC&, UBCustRecLen
    FClose C2Handle
    
    IF VAL(UBCustRec(1).BOOK) = BookNumber AND (UBCustRec(1).Status <> "F") THEN
      process = True
    ELSE
      process = False
      IF PageUp THEN
        Cnt = Cnt - 1
      ELSE
        Cnt = Cnt + 1
      END IF
      RETURN
    END IF
    
    RecNo& = pREC&
    
    IF RecNo& <= 0 THEN
      process = False
      Cnt = Cnt + 1
      RETURN
    END IF
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
    
END SUB

SUB PrintMeterList
    
    REDIM TempScrn(0)
    
    REDIM UBCustRec(1) AS NewUBCustRecType
    UBCustRecLen = LEN(UBCustRec(1))

    SHARED Choice$()
    DevFld = 3
    
ListMainBody:
    
    help$ = "Print Meter Reading List"
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
    
    Choice$(0, 0) = STR$(DevFld)                'Choices for Output Device
    Choice$(1, 0) = "Screen"
    Choice$(2, 0) = "Prn#1"
    Choice$(3, 0) = "Prn#2"
    LineCnt = 0
    
    Action = 1
    DisplayUBScrn ScrnName$
    
    DO
      
      EditForm Form$(), Fld(), Frm(1), Cnf, Action
      
      '--Check for Key presses
      
      SELECT CASE Frm(1).KeyCode
        
      CASE F10Key
        ' Check for Proper Order of Route Questions
        IF VAL(Form$(1, 0)) = 0 THEN
          BegRoute = 1: EndRoute = 99
          inputting = True
        ELSE
          BegRoute = VAL(Form$(1, 0)): EndRoute = VAL(Form$(2, 0))
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
    COLOR 7
    LOCATE 25, 42: PRINT STRING$(39, 219);
    
    IF LEN(LTRIM$(RTRIM$(Form$(DevFld, 0)))) = 0 THEN
      dev$ = "S"
      LPTNo = 1
    ELSE
      dev$ = LEFT$(Form$(DevFld, 0), 1)
      LPTNo = VAL(RIGHT$(RTRIM$(Form$(DevFld, 0)), 1))
    END IF
    
    'Free Up Some Memory
    ERASE Form$, Fld, Choice$, Frm
    
    'Open Report File
    ReportFile$ = "WBMTRLST.RPT"

    RptHandle = FREEFILE
    OPEN ReportFile$ FOR OUTPUT AS #RptHandle
    
    'Open the Utility Setup File to Grab Meter List Order (Seq or Loc)
    REDIM UBSetupRec(1) AS UBSetupRecType
    LoadUBSetUpFile UBSetupRec(), UBSetUpLen
    
    
    IF UBSetupRec(1).UseSeq = "Y" THEN
      MakeSequenceIndex "Sequence Number"
      NumOfRecs = FileSize&("UBTEMP.IDX") \ 4
      REDIM IndexArray(1 TO NumOfRecs) AS UBCustIndexRecType
      FGetAH "UBTEMP.IDX", IndexArray(1), 4, NumOfRecs
      
      DO
        
        IF Cnt < 1 THEN Cnt = 1 ' Do Not Allow to Fall Below 1
        inputting = False       ' Set Edit Finish to No
        pREC& = IndexArray(Cnt).RecNum
        IF NOT pREC& = 0 THEN
          GOSUB GetCustRecNoPrint
          IF process THEN
            GOSUB PrintLine
          END IF
        END IF
        Cnt = Cnt + 1
      LOOP UNTIL Cnt > NumOfRecs
      
      CLOSE
      EXIT SUB
    END IF
    ' End of Seq Order **********************************************************
    ' Location Order ??? ********************************************************
    
    IdxRecLen = 4               'we are using a integer
    IdxFileSize& = FileSize&("UBCUSTBK.IDX")
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    
    REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
    IdxFile = FREEFILE
    OPEN "UBCUSTBK.IDX" FOR RANDOM SHARED AS #IdxFile LEN = IdxRecLen
    FOR Cnt = 1 TO IdxNumOfRecs
      GET #IdxFile, Cnt, IdxBuff(Cnt)
    NEXT
    CLOSE IdxFile
    Cnt = 1
    
    DO
      inputting = False
      pREC& = IdxBuff(Cnt).RecNum
      
      IF NOT (pREC&) = 0 THEN
        GOSUB GetCustRecNoPrint
        IF process THEN
          GOSUB PrintLine
        END IF
      END IF
      
      Cnt = Cnt + 1
      IF EndProcess THEN EXIT DO
    LOOP UNTIL Cnt > IdxNumOfRecs
    CLOSE

    'GOSUB ProcessReport
    'EXIT SUB
    
    'End of Location Order Processing *******************************************
    
ProcessReport:
    SELECT CASE dev$
    CASE "S"
      EntryPoint = 2
    CASE "P"
      EntryPoint = 5
    END SELECT
    Header$ = "Customer Meter Listing Report"
    
    
    PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
'    CALL KILLFILE(ReportFile$)

EXIT SUB
    
    
    'Print the Meter Reading Here *********************************************
    
PrintLine:
    
    IF LineCnt = 0 THEN GOSUB PrintHeading
    help$ = "Process Location Record #" + STR$(Cnt) + " of " + STR$(IdxNumOfRecs)
    PrintHelp help$
    
    MtrCnt = 0
    ValidCustomer = False
    DO
      MtrCnt = MtrCnt + 1       'Check For Meter This Customer
      TempRev$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MTRType)
      IF LEN(TempRev$) <> 0 THEN ValidCustomer = True: EXIT DO
    LOOP UNTIL MtrCnt = 7
    IF ValidCustomer = False THEN RETURN
    
    IF LineCnt >= 53 THEN
      PRINT #RptHandle, CHR$(12); : GOSUB PrintHeading
    END IF
    
    GOSUB GetMeterStatusPrint
    
    PRINT #RptHandle, UBCustRec(1).CUSTNAME; TAB(40); LEFT$(UBCustRec(1).SERVADDR, 28);
    PRINT #RptHandle, TAB(70); MeterStatus$
    LineCnt = LineCnt + 1
    
    FOR MtrCnt = 1 TO 7         'find last active meter
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
    PRINT #RptHandle, TAB(27); "Meter Reading Listing Report"
    PRINT #RptHandle, "Beginning Route: "; BegRoute
    PRINT #RptHandle, "   Ending Route: "; EndRoute; TAB(65); "Page #"; page
    PRINT #RptHandle, ""
    PRINT #RptHandle, "Customer Name"; TAB(40); "Service Address"; TAB(70); "Status"
    PRINT #RptHandle, "Location"; TAB(15); "Meter Number"; TAB(35); "Mtr Type"; TAB(55); "Cur Read"; TAB(70); "New Read"
    PRINT #RptHandle, STRING$(80, "=")
    LineCnt = 5
    
    RETURN
    
    
GetCustRecNoPrint:
    
    RecNo& = pREC&
    
    FOpenS "UBCUST.DAT", L2Handle               'open data file
    FGetRTA L2Handle, UBCustRec(1), RecNo&, UBCustRecLen
    FClose L2Handle
    IF FirstCust THEN
      FirstCust = False
      WhatBook = VAL(UBCustRec(1).BOOK)
    END IF

    IF VAL(UBCustRec(1).BOOK) > EndRoute THEN EndProcess = True
    
    IF VAL(UBCustRec(1).BOOK) >= BegRoute AND VAL(UBCustRec(1).BOOK) <= EndRoute THEN
      IF VAL(UBCustRec(1).BOOK) <> WhatBook THEN
         PRINT #RptHandle, CHR$(12);
         DoHeaderFlag = True
         LineCnt = 0
         'GOSUB PrintHeading
         WhatBook = VAL(UBCustRec(1).BOOK)
      END IF
      process = True
    ELSE
      process = False
      RETURN
    END IF
    
    
    IF RecNo& <= 0 THEN
      process = False
      RETURN
    END IF
    
    RETURN
    
    
GetMeterTypePrint:
    IF UBCustRec(1).LocMeters(MtrCnt).MTRType = "C" THEN MeterType$ = "Water/Sewer": RETURN
    IF UBCustRec(1).LocMeters(MtrCnt).MTRType = "W" THEN MeterType$ = "Water Only": RETURN
    IF UBCustRec(1).LocMeters(MtrCnt).MTRType = "S" THEN MeterType$ = "Sewer Only": RETURN
    IF UBCustRec(1).LocMeters(MtrCnt).MTRType = "T" THEN MeterType$ = "Touch Read": RETURN
    IF UBCustRec(1).LocMeters(MtrCnt).MTRType = "E" THEN MeterType$ = "Electric": RETURN
    IF UBCustRec(1).LocMeters(MtrCnt).MTRType = "G" THEN MeterType$ = "Gas": RETURN
    MeterType$ = "Undefined"
    RETURN
    
GetMeterStatusPrint:
    IF UBCustRec(1).Status = "A" THEN MeterStatus$ = "Active": RETURN
    IF UBCustRec(1).Status = "F" THEN MeterStatus$ = "Final": RETURN
    IF UBCustRec(1).Status = "I" THEN MeterStatus$ = "Vacant": RETURN
    MeterStatus$ = "Undef."
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

SUB PrintMONList

    REDIM TempScrn(0)

    REDIM UBCustRec(1) AS NewUBCustRecType
    UBCustRecLen = LEN(UBCustRec(1))

    SHARED Choice$()
    DevFld = 3

    help$ = "Print MON Meter Reading List"
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

    Choice$(0, 0) = STR$(DevFld)                'Choices for Output Device
    Choice$(1, 0) = "Screen"
    Choice$(2, 0) = "Prn#1"
    Choice$(3, 0) = "Prn#2"
    LineCnt = 0

    Action = 1
    DisplayUBScrn ScrnName$

    DO

      EditForm Form$(), Fld(), Frm(1), Cnf, Action

      '--Check for Key presses

      SELECT CASE Frm(1).KeyCode

      CASE F10Key
        ' Check for Proper Order of Route Questions
        IF VAL(Form$(1, 0)) = 0 THEN
          BegRoute = 1: EndRoute = 99
          inputting = True
        ELSE
          BegRoute = VAL(Form$(1, 0)): EndRoute = VAL(Form$(2, 0))
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
    COLOR 7
    LOCATE 25, 42: PRINT STRING$(39, 219);

    IF LEN(LTRIM$(RTRIM$(Form$(DevFld, 0)))) = 0 THEN
      dev$ = "S"
      LPTNo = 1
    ELSE
      dev$ = LEFT$(Form$(DevFld, 0), 1)
      LPTNo = VAL(RIGHT$(RTRIM$(Form$(DevFld, 0)), 1))
    END IF

    'Free Up Some Memory
    ERASE Form$, Fld, Choice$, Frm

    'Open Report File
    ReportFile$ = "WBMTRLST.RPT"

    RptHandle = FREEFILE
    OPEN ReportFile$ FOR OUTPUT AS #RptHandle

    'Open the Utility Setup File to Grab Meter List Order (Seq or Loc)
    REDIM UBSetupRec(1) AS UBSetupRecType
    LoadUBSetUpFile UBSetupRec(), UBSetUpLen


    IF UBSetupRec(1).UseSeq = "Y" THEN
      MakeSequenceIndex "Sequence Number"
      NumOfRecs = FileSize&("UBTEMP.IDX") \ 4
      REDIM IndexArray(1 TO NumOfRecs) AS UBCustIndexRecType
      FGetAH "UBTEMP.IDX", IndexArray(1), 4, NumOfRecs

      DO

        IF Cnt < 1 THEN Cnt = 1 ' Do Not Allow to Fall Below 1
        inputting = False       ' Set Edit Finish to No
        pREC& = IndexArray(Cnt).RecNum
        IF NOT pREC& = 0 THEN
          GOSUB GetCustMONRecNoPrint
          IF process THEN
            GOSUB PrintMONLine
          END IF
        END IF
        Cnt = Cnt + 1
      LOOP UNTIL Cnt > NumOfRecs

      CLOSE
      EXIT SUB
    END IF
    ' End of Seq Order **********************************************************
    ' Location Order ??? ********************************************************

    IdxRecLen = 4               'we are using a integer
    IdxFileSize& = FileSize&("UBCUSTBK.IDX")
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen

    REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
    IdxFile = FREEFILE
    OPEN "UBCUSTBK.IDX" FOR RANDOM SHARED AS #IdxFile LEN = IdxRecLen
    FOR Cnt = 1 TO IdxNumOfRecs
      GET #IdxFile, Cnt, IdxBuff(Cnt)
    NEXT
    CLOSE IdxFile
    Cnt = 1

    DO
      inputting = False
      pREC& = IdxBuff(Cnt).RecNum

      IF NOT (pREC&) = 0 THEN
        GOSUB GetCustMONRecNoPrint
        IF process THEN
          GOSUB PrintMONLine
        END IF
      END IF

      Cnt = Cnt + 1
      IF EndProcess THEN EXIT DO
    LOOP UNTIL Cnt > IdxNumOfRecs
    CLOSE

    'GOSUB ProcMONReport
    'EXIT SUB

    'End of Location Order Processing *******************************************

ProcMONReport:
    SELECT CASE dev$
    CASE "S"
      EntryPoint = 2
    CASE "P"
      EntryPoint = 5
    END SELECT
    Header$ = "Customer Meter Listing Report"


    PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
'    CALL KILLFILE(ReportFile$)

EXIT SUB


    'Print the Meter Reading Here *********************************************

PrintMONLine:

    IF LineCnt = 0 THEN GOSUB PrintMONHeading
    help$ = "Process Location Record #" + STR$(Cnt) + " of " + STR$(IdxNumOfRecs)
    PrintHelp help$

    MtrCnt = 0
    ValidCustomer = False
    DO
      MtrCnt = MtrCnt + 1       'Check For Meter This Customer
      TempRev$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MTRType)
      IF LEN(TempRev$) <> 0 THEN ValidCustomer = True: EXIT DO
    LOOP UNTIL MtrCnt = 7
    IF ValidCustomer = False THEN RETURN

    IF LineCnt >= 53 THEN
      PRINT #RptHandle, CHR$(12); : GOSUB PrintMONHeading
    END IF

    GOSUB MeterStatusPrintMON

    PRINT #RptHandle, UBCustRec(1).CUSTNAME; TAB(40); LEFT$(UBCustRec(1).SERVADDR, 28);
    PRINT #RptHandle, TAB(70); MeterStatus$
    LineCnt = LineCnt + 1

    FOR MtrCnt = 1 TO 7         'find last active meter
      TempRev$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MTRType)
      IF LEN(TempRev$) <> 0 THEN
        GOSUB MONGetMeterTypePrint
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


PrintMONHeading:
    page = page + 1
    PRINT #RptHandle, TAB(27); "Meter Reading Listing Report"
    PRINT #RptHandle, "Beginning Route: "; BegRoute
    PRINT #RptHandle, "   Ending Route: "; EndRoute; TAB(65); "Page #"; page
    PRINT #RptHandle, ""
    PRINT #RptHandle, "Customer Name"; TAB(40); "Service Address"; TAB(70); "Status"
    PRINT #RptHandle, "Location"; TAB(15); "Meter Number"; TAB(35); "Mtr Type"; TAB(55); "Cur Read"; TAB(70); "New Read"
    PRINT #RptHandle, STRING$(80, "=")
    LineCnt = 5

    RETURN


GetCustMONRecNoPrint:

    RecNo& = pREC&

    FOpenS "UBCUST.DAT", L2Handle               'open data file
    FGetRTA L2Handle, UBCustRec(1), RecNo&, UBCustRecLen
    FClose L2Handle
    IF FirstCust THEN
      FirstCust = False
      WhatBook = VAL(UBCustRec(1).BOOK)
    END IF

    IF VAL(UBCustRec(1).BOOK) > EndRoute THEN EndProcess = True

    IF VAL(UBCustRec(1).BOOK) >= BegRoute AND VAL(UBCustRec(1).BOOK) <= EndRoute THEN
      IF UBCustRec(1).CustType = "MON" THEN
        IF VAL(UBCustRec(1).BOOK) <> WhatBook THEN
           PRINT #RptHandle, CHR$(12);
           DoHeaderFlag = True
           LineCnt = 0
           'GOSUB PrintMONHeading
           WhatBook = VAL(UBCustRec(1).BOOK)
        END IF
        process = True
      ELSE
        process = False
        RETURN
      END IF
    ELSE
      process = False
      RETURN
    END IF


    IF RecNo& <= 0 THEN
      process = False
      RETURN
    END IF

    RETURN


MONGetMeterTypePrint:
    IF UBCustRec(1).LocMeters(MtrCnt).MTRType = "C" THEN MeterType$ = "Water/Sewer": RETURN
    IF UBCustRec(1).LocMeters(MtrCnt).MTRType = "W" THEN MeterType$ = "Water Only": RETURN
    IF UBCustRec(1).LocMeters(MtrCnt).MTRType = "S" THEN MeterType$ = "Sewer Only": RETURN
    IF UBCustRec(1).LocMeters(MtrCnt).MTRType = "T" THEN MeterType$ = "Touch Read": RETURN
    IF UBCustRec(1).LocMeters(MtrCnt).MTRType = "E" THEN MeterType$ = "Electric": RETURN
    IF UBCustRec(1).LocMeters(MtrCnt).MTRType = "G" THEN MeterType$ = "Gas": RETURN
    MeterType$ = "Undefined"
    RETURN

MeterStatusPrintMON:
    IF UBCustRec(1).Status = "A" THEN MeterStatus$ = "Active": RETURN
    IF UBCustRec(1).Status = "F" THEN MeterStatus$ = "Final": RETURN
    IF UBCustRec(1).Status = "I" THEN MeterStatus$ = "Vacant": RETURN
    MeterStatus$ = "Undef."
    RETURN

END SUB

SUB Search4LNumber (LocNum$, RecNo&, CLSFlag%)
    
    REDIM UBCustRec(1) AS NewUBCustRecType
    UBCustRecLen = LEN(UBCustRec(1))
    
    'REDIM UBLocatRec(1) AS UBLocationRecType
    'UBLocatRecLen = LEN(UBLocatRec(1))
    
    IdxRecLen = 4               'we are using a integer
    IdxFileSize& = FileSize&("UBCUSTBK.IDX")
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    
    REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
    
    FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs  'load it
    
    SearchLen = LEN(LocNum$)
    
    Match = False
    FirstRec = 1
    LastRec = IdxNumOfRecs
    
    BotOffSet = 0
    TopOffSet = IdxNumOfRecs
    
    FOpenS "UBCUST.DAT", C1Handle               'open data file
    
    'Find matching record
    
    MidRec = (LastRec + FirstRec) \ 2
    
    DO
      IF LastSRec = MidRec THEN EXIT DO
      LastSRec = MidRec
      
      FGetRTA C1Handle, UBCustRec(1), CLNG(IdxBuff(MidRec).RecNum), UBCustRecLen
      
      UBSearchN$ = UBCustRec(1).BOOK + "-" + UBCustRec(1).SEQNUMB
      
      IF LocNum$ = UBSearchN$ THEN
        
        IF MidRec - BotOffSet > 1 THEN
          MidRec = MidRec - 1
        ELSE
          FirstMatchRec = MidRec
        END IF
      ELSEIF LocNum$ < UBSearchN$ THEN          'lower
        TopOffSet = MidRec
        MidRec = TopOffSet - ((TopOffSet - BotOffSet) \ 2)
      ELSE      'higher
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

