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

  ShowCursor
  PrintMeterList
  HideCursor
  RUN "ubmtread"

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
    'SHELL
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
  
  PRINT #RptHandle, LEFT$(UBCustRec(1).CustName, 30);
  IF UBSetupRec(1).UseSeq = "Y" THEN
    IF UBCustRec(1).SEQ < 0 THEN UBCustRec(1).SEQ = 0
    PRINT #RptHandle, TAB(32); USING "######"; UBCustRec(1).SEQ;
  END IF
  PRINT #RptHandle, TAB(40); LEFT$(UBCustRec(1).ServAddr, 28);
  PRINT #RptHandle, TAB(70); MeterStatus$
  LineCnt = LineCnt + 1
  
  FOR MtrCnt = 1 TO 7           'find last active meter
    TempRev$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MTRType)
    IF LEN(TempRev$) <> 0 THEN
      GOSUB GetMeterTypePrint
      PRINT #RptHandle, UBCustRec(1).BOOK; "-"; UBCustRec(1).SeqNumb;
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

