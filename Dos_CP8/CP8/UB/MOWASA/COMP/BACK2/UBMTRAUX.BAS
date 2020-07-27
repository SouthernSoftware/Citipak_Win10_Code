DECLARE SUB ProcessBadger ()
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

SUB EstMeterReading
  
  REDIM TempScrn(0)
  LibName$ = "UB"
  ScrnName$ = "BOOK2EST"
  
  REDIM UBSetupRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetupRec(), UBSetUpLen
  
  TownName$ = UBSetupRec(1).UTILNAME
  
  'IF INSTR(TownName$, "BEECH MOUNTAIN") THEN
  '  BeechFlag = True
  'END IF
  
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
              'IF BeechFlag THEN
              '  AvgUse& = 20
              'ELSE
                AvgUse& = UBCustRec(1).LocMeters(zz).AvgUse
              'END IF
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
  
  
  IF UBSetUpLen = 0 THEN
    BEEP
    CursorOff
    DisplayHHScrn "UBNOHH"
    WaitForAction
    CLOSE
    EXIT SUB
  END IF
  
  SELECT CASE UBSetupRec(1).HHDEVICE
  CASE "H", "S", "E", "P", "C", "D", "M", "I", "B"
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
  CASE "H"
    ProcessHusky
  CASE "I"
    ProcessItron
  CASE "B"
    ProcessBadger

  END SELECT
  
  EXIT SUB
  
  
SysCom:
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

