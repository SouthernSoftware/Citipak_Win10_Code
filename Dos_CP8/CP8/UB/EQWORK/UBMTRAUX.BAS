DECLARE SUB DisplayH2Scrn (ScrnName$)
DECLARE SUB ProcessIntermec ()
DEFINT A-Z
DECLARE SUB KILLFile (FileName$)
DECLARE SUB DisplayHHScrn (ScrnName$)
DECLARE SUB ProcessItron ()
DECLARE SUB ProcessLogicon ()
DECLARE SUB ProcessTelxon ()
DECLARE SUB CursorOff ()
DECLARE SUB DisplayHuskyScrn (ScrnName$)
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
DECLARE SUB MakeZipCodeIndex (IndexText$)
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
DECLARE FUNCTION FmtBook$ (Book$)
DECLARE FUNCTION FmtSeqN$ (SeqN$)
  '$INCLUDE: 'DefCnf.BI'
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FOpenS (FileName$, Handle%)
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
  ''$INCLUDE: 'ubtrans.BI'
  ''$INCLUDE: 'ubdraft.BI'
  '$INCLUDE: 'ubsetup.bi'
  ''$INCLUDE: 'PageInfo.bi'
  '$INCLUDE: 'Newcust.bi'
  ''$INCLUDE: 'ubrate.BI'
  ''$INCLUDE: 'ubsetbal.BI'
  '$INCLUDE: 'ubgilsen.bi'
 '$INCLUDE: 'ubInter.bi'

  
  CONST False = 0, True = NOT False

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
  
  IF UBSetUpLen = 0 THEN
    BEEP
    CursorOff
    DisplayHHScrn "UBNOHH"
    WaitForAction
    CLOSE
    EXIT SUB
  END IF
  
  SELECT CASE UBSetUpRec(1).HHDEVICE
  CASE "H", "S", "C", "D", "T", "L", "I", "W"
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
'  CASE "S"
'    GOSUB Sensus
'  CASE "D"
'    GOSUB DataGeneral
'  CASE "C"
'    GOSUB SysCom
'  CASE "H"
    'ProcessHusky
'  CASE "T"
'    ProcessTelxon
'  CASE "L"
'    ProcessLogicon
'  CASE "I"
'    ProcessItron

  CASE "W"
    ProcessIntermec
  END SELECT

EXIT SUB

'Sensus:
  
  
FILENAM1:
  PathWay$ = QPTrim$(PathWay$)
  FileName$ = PathWay$ + "ssi00" + LTRIM$(STR$(INTERROGATOR)) + ".rte"
RETURN

  
FileNam:
  PathWay$ = QPTrim$(PathWay$)
  FileName$ = PathWay$ + "exssi00" + QPTrim$(STR$(INTERROGATOR)) + ".dat"
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

SUB PrintMeterList
  
  REDIM TempScrn(0)
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  SHARED Choice$()
  DevFld = 3
  
ListMainBody:
  
  Help$ = "Print Meter Reading List"
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
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen
  
  IF UBSetUpRec(1).UseSeq = "Y" THEN
    MakeSequenceIndex "Sequence Number"
    NumOfRecs = FileSize&("UBTEMP.IDX") \ 4
    IdxNumOfRecs = NumOfRecs
    REDIM IndexArray(1 TO NumOfRecs) AS UBCustIndexRecType
    FGetAH "UBTEMP.IDX", IndexArray(1), 4, NumOfRecs
    DO
      IF Cnt < 1 THEN Cnt = 1   ' Do Not Allow to Fall Below 1
      inputting = False         ' Set Edit Finish to No
      Prec& = IndexArray(Cnt).RecNum
      IF NOT Prec& = 0 THEN
        GOSUB GetCustRecNoPrint
        IF process THEN
          GOSUB PrintLine
        END IF
      END IF
      Cnt = Cnt + 1
    LOOP UNTIL Cnt > NumOfRecs
    CLOSE
    GOTO ProcessReport
  END IF
  ' End of Seq Order **********************************************************
  ' Location Order ??? ********************************************************
  
  IdxRecLen = 4 'we are using a integer
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
    Prec& = IdxBuff(Cnt).RecNum
    
    IF NOT (Prec&) = 0 THEN
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
  Help$ = "Process Location Record #" + STR$(Cnt) + " of " + STR$(IdxNumOfRecs)
  PrintHelp Help$
  
  MtrCnt = 0
  ValidCustomer = False
  DO
    MtrCnt = MtrCnt + 1         'Check For Meter This Customer
    TempRev$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MTRType)
    IF LEN(TempRev$) <> 0 THEN ValidCustomer = True: EXIT DO
  LOOP UNTIL MtrCnt = 7
  IF ValidCustomer = False THEN RETURN
  
  IF LineCnt >= 53 THEN
    PRINT #RptHandle, CHR$(12); : GOSUB PrintHeading
  END IF
  
  GOSUB GetMeterStatusPrint
  
  PRINT #RptHandle, LEFT$(UBCustRec(1).CustName, 30);
  IF UBSetUpRec(1).UseSeq = "Y" THEN
    IF UBCustRec(1).Seq < 0 THEN UBCustRec(1).Seq = 0
    PRINT #RptHandle, TAB(32); USING "######"; UBCustRec(1).Seq;
  END IF
  PRINT #RptHandle, TAB(40); LEFT$(UBCustRec(1).SERVADDR, 28);
  PRINT #RptHandle, TAB(70); MeterStatus$
  LineCnt = LineCnt + 1
  
  FOR MtrCnt = 1 TO 7           'find last active meter
    TempRev$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MTRType)
    IF LEN(TempRev$) <> 0 THEN
      GOSUB GetMeterTypePrint
      PRINT #RptHandle, UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB;
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
  IF UBSetUpRec(1).UseSeq = "Y" THEN
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
    WhatBook = VAL(UBCustRec(1).Book)
  END IF
  
  IF (UBCustRec(1).DelFlag <> 0) OR INSTR(UBCustRec(1).HHMSG1, "NOREAD") > 0 THEN
    process = False
    RETURN
  END IF
  
  IF VAL(UBCustRec(1).Book) > EndRoute THEN EndProcess = True
  
  IF VAL(UBCustRec(1).Book) >= BegRoute AND VAL(UBCustRec(1).Book) <= EndRoute THEN
    IF VAL(UBCustRec(1).Book) <> WhatBook THEN
      PRINT #RptHandle, CHR$(12);
      DoHeaderFlag = True
      LineCnt = 0
      WhatBook = VAL(UBCustRec(1).Book)
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
  
  'LineCnt& = 0
  
  Action = 1
  FirstTime = True
  BlockClear
  DisplayUBScrn ScrnName$
  QPrintRC "Notes  ", 8, 38, -1
  
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
      
    CASE F10Key
      ' Check for Proper Order of Route Questions
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
  
  'Open the Utility Setup File to Grab Meter List Order (Seq or Loc)
  'REDIM UBSetupRec(1) AS UBSetupRecType
  'LoadUBSetUpFile UBSetupRec(), UBSetUpLen
  
  CustName$ = SPACE$(30)
  
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
    Book = QPValI(UBCustRec(1).Book)
    IF Book >= BegRoute AND Book <= EndRoute THEN
      LSET CustName$ = UBCustRec(1).CustName
      IF UBCustRec(1).DelFlag = 0 THEN
        IF LEN(QPTrim$(UBCustRec(1).HHMSG1)) > 0 OR LEN(QPTrim$(UBCustRec(1).HHMSG2)) > 0 OR LEN(QPTrim$(UBCustRec(1).HHMSG3)) > 0 THEN
          PRINT #RptHandle, UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB; "  "; CustName$; LEFT$(UBCustRec(1).SERVADDR, 30)
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
  
GetNotesMeterType:
  PrintMrtFlag = False
  SELECT CASE UBCustRec(1).LocMeters(MtrCnt&).MTRType
  CASE "C"
    MeterType$ = "Water/Sewer"
    PrintMrtFlag = True
  CASE "W"
    MeterType$ = "Water Only"
    PrintMrtFlag = True
  CASE "S"
    MeterType$ = "Sewer Only"
    PrintMrtFlag = True
  CASE "T"
    MeterType$ = "Touch Read"
    PrintMrtFlag = True
  CASE "E", "D"
    MeterType$ = "Electric"
    PrintMrtFlag = True
  CASE "G"
    MeterType$ = "Gas"
    PrintMrtFlag = True
  CASE ELSE
    MeterType$ = "Undefined"
    PrintMrtFlag = True
  END SELECT
  RETURN
  
  'FOR MtrCnt& = 1 TO 7                'find last active meter
  '  TempRev$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt&).MTRType)
  '  IF LEN(TempRev$) > 0 THEN
  '    GOSUB GetNotesMeterType
  '    IF PrintMrtFlag THEN
  '      DidOne = True
  '      PRINT #RptHandle, TAB(14); MeterType$;
  '      PRINT #RptHandle, TAB(31); USING "##########"; UBCustRec(1).LocMeters(MtrCnt&).CurRead;
  '      PRINT #RptHandle, TAB(42); USING "##########"; UBCustRec(1).LocMeters(MtrCnt&).PrevRead;
  '      MeterConsp& = UBCustRec(1).LocMeters(MtrCnt&).CurRead - UBCustRec(1).LocMeters(MtrCnt&).PrevRead
  '      IF MeterConsp& < 0 THEN
  '        MaxMeterAmt& = 10& ^ (LEN(STR$(UBCustRec(1).LocMeters(MtrCnt&).PrevRead)) - 1)
  '        MeterConsp& = (MaxMeterAmt& - UBCustRec(1).LocMeters(MtrCnt&).PrevRead) + UBCustRec(1).LocMeters(MtrCnt&).CurRead
  '      END IF
  '      PRINT #RptHandle, TAB(53); USING "##########"; MeterConsp&;
  '      PRINT #RptHandle, TAB(67); Num2Date$(UBCustRec(1).LocMeters(MtrCnt&).CurDate)
  '      'PRINT #RptHandle, " "; UBCustRec(1).LocMeters(MtrCnt&).ReadFlag
  '      LineCnt = LineCnt + 1
  '    END IF
  
  
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
      
    CASE F10Key
      ' Check for Proper Order of Route Questions
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
  'REDIM UBSetupRec(1) AS UBSetupRecType
  'LoadUBSetUpFile UBSetupRec(), UBSetUpLen
  
  CustName$ = SPACE$(30)
  
  ' Location Order ********************************************************
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
    IF NOT UBCustRec(1).DelFlag THEN
      Book = QPValI(UBCustRec(1).Book)
      IF Book >= BegRoute AND Book <= EndRoute THEN
        LSET CustName$ = UBCustRec(1).CustName
        IF UBCustRec(1).Status > "" THEN
          PRINT #RptHandle, UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB; "  "; CustName$; LEFT$(UBCustRec(1).SERVADDR, 30)
          LineCnt = LineCnt + 1
          FOR MtrCnt& = 1 TO 7  'find last active meter
            TempRev$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt&).MTRType)
            IF LEN(TempRev$) > 0 THEN
              GOSUB GetReadMeterTypePrint
              IF PrintMrtFlag THEN
                DidOne = True
                IF UBCustRec(1).LocMeters(MtrCnt&).ReadFlag <> "Y" THEN
                  PRINT #RptHandle, "UNREAD";
                END IF
                PRINT #RptHandle, TAB(14); MeterType$;
                PRINT #RptHandle, TAB(31); USING "##########"; UBCustRec(1).LocMeters(MtrCnt&).CurRead;
                PRINT #RptHandle, TAB(42); USING "##########"; UBCustRec(1).LocMeters(MtrCnt&).PrevRead;
                IF UBCustRec(1).LocMeters(MtrCnt&).CurRead < 0 OR UBCustRec(1).LocMeters(MtrCnt&).PrevRead < 0 THEN
                  MeterConsp& = 0
                ELSE
                  MeterConsp& = UBCustRec(1).LocMeters(MtrCnt&).CurRead - UBCustRec(1).LocMeters(MtrCnt&).PrevRead
                END IF
                IF MeterConsp& < 0 THEN
                  MaxMeterAmt& = 10& ^ (LEN(STR$(UBCustRec(1).LocMeters(MtrCnt&).PrevRead)) - 1)
                  MeterConsp& = (MaxMeterAmt& - UBCustRec(1).LocMeters(MtrCnt&).PrevRead) + UBCustRec(1).LocMeters(MtrCnt&).CurRead
                END IF
                PRINT #RptHandle, TAB(53); USING "##########"; MeterConsp&;
                PRINT #RptHandle, TAB(67); Num2Date$(UBCustRec(1).LocMeters(MtrCnt&).CurDate)
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
      IF LineCnt >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintReadHeading
      END IF
    END IF
    ShowPctComp LCnt&, IdxNumOfRecs
  NEXT
  
  PRINT #RptHandle, FF$
  CLOSE
  
  Header$ = "Customer Meter Reading Report"
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  EXIT SUB
  
PrintReadHeading:
  page = page + 1
  PRINT #RptHandle, TAB(30); "Meter Reading Report"; TAB(65); "Date: "; DATE$
  PRINT #RptHandle, "Beginning Route: "; BegRoute
  PRINT #RptHandle, "   Ending Route: "; EndRoute; TAB(70); "Page #"; page
  PRINT #RptHandle, ""
  PRINT #RptHandle, "Location   Customer Name"; TAB(41); "Service Address"
  PRINT #RptHandle, "              Mtr Type"; TAB(31); "   Current   Previous    Consump     Read Date"
  PRINT #RptHandle, STRING$(80, "=")
  LineCnt = 7
  RETURN
  
GetReadMeterTypePrint:
  PrintMrtFlag = False
  SELECT CASE UBCustRec(1).LocMeters(MtrCnt&).MTRType
  CASE "C"
    MeterType$ = "Water/Sewer"
    PrintMrtFlag = True
  CASE "W"
    MeterType$ = "Water Only"
    PrintMrtFlag = True
  CASE "S"
    MeterType$ = "Sewer Only"
    PrintMrtFlag = True
  CASE "T"
    MeterType$ = "Touch Read"
    PrintMrtFlag = True
  CASE "E", "D"
    MeterType$ = "Electric"
    PrintMrtFlag = True
  CASE "G"
    MeterType$ = "Gas"
    PrintMrtFlag = True
  CASE ELSE
    MeterType$ = "Undefined"
    PrintMrtFlag = True
  END SELECT
  
  RETURN
  
  
END SUB

SUB PrintMeterSheets
  
  ReportFile$ = "UBMTRSHT.RPT"
  
  REDIM TempScrn(0)
  
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
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
  
  ' Location Order ********************************************************
  IdxRecLen = 4 'we are using a long integer
  IdxFileSize& = FileSize&("UBCUSTBK.IDX")
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  
  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it
  
  UBCustFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCustFile LEN = UBCustRecLen
  
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle     'Open Report File
  
  ShowProcessingScrn "Reading Meter Information"
  FOR LCnt& = 1 TO IdxNumOfRecs
    GET #UBCustFile, IdxBuff(LCnt&).RecNum, UBCustRec(1)
    Book = QPValL(UBCustRec(1).Book)
    Sequ& = QPValL(UBCustRec(1).SEQNUMB)
    IF Book >= Book1& AND Book <= Book2& THEN
      IF Sequ& >= Sequ1& AND Sequ& <= Sequ2& THEN
        IF UBCustRec(1).DelFlag = 0 THEN
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
          PRINT #RptHandle, UBCustRec(1).CustName
          PRINT #RptHandle, UBCustRec(1).ADDR1
          PRINT #RptHandle, UBCustRec(1).SERVADDR
          PRINT #RptHandle, QPTrim$(UBCustRec(1).CITY); " "; UBCustRec(1).STATE; " "; UBCustRec(1).ZIPCODE
          PRINT #RptHandle,
          PRINT #RptHandle, TAB(11); UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB
          PRINT #RptHandle, "~"
        END IF
      END IF
    END IF
    
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

SUB ProcessIntermec

  CEReadFile$ = "CEMTREAD.DAT"

  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen

  q$ = CHR$(34)
  cb$ = SPACE$(45)
  SHARED Choice$()
  REDIM Route(100)

  LibName$ = "UBHH"
  ScrnName$ = "UBINTRHH"
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
  Choice$(1, 0) = "Export file TO Handheld"
  Choice$(2, 0) = "Import file FROM Handheld"

  Action = 1

  DisplayH2Scrn ScrnName$

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode

    CASE F10Key
      WhatAction$ = LEFT$(Form$(1, 0), 1)
      SELECT CASE WhatAction$
      CASE "E"
        Done = True
      CASE "I"
        Done = True
      CASE ELSE
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
      END SELECT
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done

  SELECT CASE LEFT$(Form$(1, 0), 1)
  CASE "E"
    GOSUB ExportIntermec
  CASE "I"
    GOSUB ImportIntermec
  CASE ELSE
  END SELECT

EXIT SUB

ExportIntermec:
  LibName$ = "UBHH"
  ScrnName$ = "UBINTR1S"

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
      IF NumberofRoutes > 0 THEN             ' Next Check Routes
        QPrintRC "      Scanning Account:        ", 11, 21, -1
        FOR Cnt! = 1 TO NumberofRoutes       ' Set Route Number Values
          Route(Cnt!) = VAL(Form$(Cnt! + 2, 0))
        NEXT Cnt!
        
        REDIM UBCustRec(1) AS NewUBCustRecType
        UBCustRecLen = LEN(UBCustRec(1))
        UBFile = FREEFILE               'Open Customer Data File
        OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

        REDIM UBInterRDRec(1)   AS UBIntermecHHRecType

        UBInterRDRec(1).CEVariant = CHR$(8) + CHR$(0)
        UBInterRDRec(1).CEStrLen = CHR$(165) + CHR$(0)

        UBInterRDRecLen = LEN(UBInterRDRec(1))

        HHFile = FREEFILE
        OPEN CEReadFile$ FOR RANDOM SHARED AS HHFile LEN = UBInterRDRecLen
        CLOSE HHFile

        CALL KILLFile(CEReadFile$)

        HHFile = FREEFILE
        OPEN CEReadFile$ FOR RANDOM SHARED AS HHFile LEN = UBInterRDRecLen
        
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
              GOTO IWriteSkip
            END IF
            DisplayAcct$ = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
            QPrintRC DisplayAcct$, 11, 45, -1
            FOR BookCnt! = 1 TO NumberofRoutes
              IF VAL(UBCustRec(1).Book) = Route(BookCnt!) AND (UBCustRec(1).Status <> "F") THEN
                GOSUB IntermecWriteRec
              END IF
            NEXT BookCnt!
          END IF

IWriteSkip:
          Cnt! = Cnt! + 1
        LOOP UNTIL Cnt! > IdxNumOfRecs

        CLOSE

        Ok = MsgBox%("UBSETUP", "UBCEMT03")
        Done = True
      END IF
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done

  RETURN

ImportIntermec:

  Ok = MsgBox%("UBSETUP", "UBCEMT01")

  DisplayHHScrn "UBHUSK3C"
  QPrintRC "Intermec File Operation ", 7, 21, -1

  IF Ok = 2 THEN
    QPrintRC "        Import Aborted.    ", 10, 25, -1
    QPrintRC "  Press any key to continue.  ", 14, 25, -1
    BEEP
    CALL WaitForAction
    Done = True
    GOTO IntermecGetExit:
  END IF

  '    'Open Customer Data File

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBFile LEN = UBCustRecLen
  
  REDIM UBInterRDRec(1)   AS UBIntermecHHRecType
  UBInterRDRecLen = LEN(UBInterRDRec(1))

  HHFile = FREEFILE
  OPEN CEReadFile$ FOR RANDOM SHARED AS HHFile LEN = UBInterRDRecLen

  NumIntermecRdRecs = LOF(HHFile) / UBInterRDRecLen

  IF NumIntermecRdRecs = 0 THEN
    CLOSE
    QPrintRC " ERROR:  READING FILE NOT FOUND!", 9, 25, -1
    QPrintRC " The file 'CEMTREAD.DAT' MUST BE", 11, 25, -1
    QPrintRC "     IN THE CITIPAK FOLDER.     ", 13, 25, -1
    QPrintRC "                                ", 14, 24, 15
    QPrintRC "   Press any key to continue.   ", 15, 25, 14
    BEEP: BEEP: BEEP
    WaitForAction
    Done = True
    GOTO IntermecGetExit:
  END IF

  QPrintRC "  Processing Meter Readings.  ", 9, 25, 15
  QPrintRC "                              ", 10, 25, 11
  QPrintRC "    Updated Count:            ", 11, 25, 11
  QPrintRC "                              ", 14, 25, 11

  Cnt! = 1      ' Initialize File Counter to 1
  DO
    GET HHFile, Cnt!, UBInterRDRec(1)
    ' New Way Undelete When Our System Sends and Retrieves
    Prec& = VAL(QPTrim$(UBInterRDRec(1).Account))
    IF NOT (Prec&) = 0 THEN
      GET UBFile, Prec&, UBCustRec(1)
      GOSUB IntermecExtractRec
    END IF
    Cnt! = Cnt! + 1
  LOOP UNTIL Cnt! > NumIntermecRdRecs

  CLOSE
  Done = True

  QPrintRC " Readings Updated Successfully. ", 13, 24, 15
  QPrintRC "  Press any key to continue.  ", 15, 25, 14
  BEEP
  WaitForAction
  Done = True

IntermecGetExit:
  RETURN

IntermecExtractRec:
  UpdCnt! = UpdCnt! + 1
  QPrintRC STR$(UpdCnt!), 11, 44, 15
  'New Way To Activate

  IF UBInterRDRec(1).ReadFlag = "Y" THEN
    MeterRecord = VAL(RIGHT$((QPTrim$(UBInterRDRec(1).Account)), 1))
    ' Check Meter Updated Flag
    ' Update Meter W/Reading
'NOTE: New current reading stored in the pastreading field from CEMTREAD.DAT
    CurReading# = VAL(UBInterRDRec(1).PastRead)
    'CurReading# = VAL(UBInterRDRec(1).CurRead)

    ThisDate$ = LEFT$(UBInterRDRec(1).ReadDate, 2) + "/" + MID$(UBInterRDRec(1).ReadDate, 3, 2) + "/" + RIGHT$(UBInterRDRec(1).ReadDate, 4)
    ReadingDate = Date2Num(ThisDate$)

    IF UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y" THEN
      UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
      UBCustRec(1).LocMeters(MeterRecord).CurDate = ReadingDate
    ELSE
      UBCustRec(1).LocMeters(MeterRecord).PrevRead = UBCustRec(1).LocMeters(MeterRecord).CurRead
      UBCustRec(1).LocMeters(MeterRecord).PastDate = UBCustRec(1).LocMeters(MeterRecord).CurDate
      UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y"
      UBCustRec(1).LocMeters(MeterRecord).CurDate = ReadingDate
      UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
    END IF
    'Modifed 04-28-97
    'IF UBPC3000GetRDRec(1).NoteStatus = "P" THEN

    MMsg1$ = QPTrim$(UBInterRDRec(1).Note1)
    CMsg1$ = QPTrim$(UBCustRec(1).HHMSG1)
    MMsg2$ = QPTrim$(UBInterRDRec(1).Note2)
    CMsg2$ = QPTrim$(UBCustRec(1).HHMSG2)
    MMsg3$ = QPTrim$(UBInterRDRec(1).Note3)
    CMsg3$ = QPTrim$(UBCustRec(1).HHMSG3)

    IF MMsg1$ <> CMsg1$ THEN
      GOSUB UpDateNoteInfo
      GOTO DoneINNotes
    END IF
    IF MMsg2$ <> CMsg2$ THEN
      GOSUB UpDateNoteInfo
      GOTO DoneINNotes
    END IF
    IF MMsg3$ <> CMsg3$ THEN
      GOSUB UpDateNoteInfo
    END IF

DoneINNotes:
    PUT UBFile, Prec&, UBCustRec(1)
  END IF
RETURN

UpDateNoteInfo:
  'UBCustRec(1).NewNotes = True
  UBCustRec(1).HHMSG1 = MMsg1$
  UBCustRec(1).HHMSG2 = MMsg2$
  UBCustRec(1).HHMSG3 = MMsg3$
RETURN

IntermecWriteRec:
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

          UBInterRDRec(1).CustName = UBCustRec(1).CustName
          UBInterRDRec(1).ServAddress = LEFT$(UBCustRec(1).SERVADDR, 16)
          UBInterRDRec(1).ReadDate = ""

          C$ = QPTrim$(UBCustRec(1).USERCODE1)
          IF LEN(C$) > 0 THEN
            MID$(UBInterRDRec(1).ServAddress, 19, 1) = LEFT$(C$, 1)
          END IF
          C$ = QPTrim$(UBCustRec(1).USERCODE2)
          IF LEN(C$) > 0 THEN
            MID$(UBInterRDRec(1).ServAddress, 20, 1) = LEFT$(C$, 1)
          END IF
          
          UBInterRDRec(1).MeterID = MeterID$
          UBInterRDRec(1).LowRead = QPTrim$(STR$(LowRead#))
          UBInterRDRec(1).HighRead = QPTrim$(STR$(HighRead#))
          UBInterRDRec(1).Account = Account$
          UBInterRDRec(1).MeterType$ = UBCustRec(1).LocMeters(MtrCnt!).MTRType
          UBInterRDRec(1).Book = UBCustRec(1).Book
          UBInterRDRec(1).CurRead = QPTrim$(STR$(UBCustRec(1).LocMeters(MtrCnt!).CurRead))
          UBInterRDRec(1).PastRead = "0"
          UBInterRDRec(1).ReadFlag = "N"
          'Modifed 04-28-97
          UBInterRDRec(1).Note1 = QPTrim$(UBCustRec(1).HHMSG1)
          UBInterRDRec(1).Note2 = QPTrim$(UBCustRec(1).HHMSG2)
          UBInterRDRec(1).Note3 = QPTrim$(UBCustRec(1).HHMSG3)
          '         ^^^
          UBInterRDRec(1).NoteStatus = ""
          PUT HHFile, (LOF(HHFile) / UBInterRDRecLen) + 1, UBInterRDRec(1)
        END IF
      END IF
    END IF
  NEXT MtrCnt!
  RETURN

END SUB

SUB Search4LNumber (LocNum$, RecNo&, CLSFlag%)
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  'REDIM UBLocatRec(1) AS UBLocationRecType
  'UBLocatRecLen = LEN(UBLocatRec(1))
  
  IdxRecLen = 4 'we are using a integer
  IdxFileSize& = FileSize&("UBCUSTBK.IDX")
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  
  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  
  FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it
  
  SearchLen = LEN(LocNum$)
  
  Match = False
  FirstRec = 1
  LastRec = IdxNumOfRecs
  
  BotOffSet = 0
  TopOffSet = IdxNumOfRecs
  
  FOpenS "UBCUST.DAT", C1Handle 'open data file
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
    ELSEIF LocNum$ < UBSearchN$ THEN            'lower
      TopOffSet = MidRec
      MidRec = TopOffSet - ((TopOffSet - BotOffSet) \ 2)
    ELSE        'higher
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

