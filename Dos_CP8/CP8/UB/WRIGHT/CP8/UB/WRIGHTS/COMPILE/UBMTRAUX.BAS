  DEFINT A-Z
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE FUNCTION Chk4BookSeqNum (Book$, SeqNum$)
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FLof& (FileHand%)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION FmtBook$ (Book$)
DECLARE FUNCTION FmtSeqN$ (SeqN$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValI% (Numb$)
DECLARE FUNCTION QPValL& (Numb$)
DECLARE FUNCTION Round# (N#)
DECLARE SUB BlockClear ()
DECLARE SUB ClearScrn ()
DECLARE SUB CursorOff ()
DECLARE SUB DisplayHHScrn (ScrnName$)
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB EnterMeterReadingsMenu ()
DECLARE SUB EstMeterReading ()
DECLARE SUB FClose (Handle%)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB HideCursor ()
DECLARE SUB KillFile (FileName$)
DECLARE SUB LoadUBSetUpFile (UBSetupRec() AS ANY, UBSetUpLen%)
DECLARE SUB MakeSequenceIndex (IndexText$)
DECLARE SUB ManualMeterReading ()
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintMeterList ()
DECLARE SUB PrintMeterNotes ()
DECLARE SUB PrintMeterReport ()
DECLARE SUB PrintMeterSheets ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB ProcessHusky ()
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB RestScrn (TempScrn())
DECLARE SUB SaveScrn (TempScrn())
DECLARE SUB ShowCursor ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB WaitForAction ()
  
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'Newcust.bi'
  ''$INCLUDE: 'ubgilsen.bi'
  
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

SUB ConditionReport
  REDIM TempScrn(0)
  
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  MaxLines = 58
  FF$ = CHR$(12)
  Dash131$ = STRING$(129, "=")
  Dash2$ = STRING$(129, "-")

  LibName$ = "UB"
  ScrnName$ = "UBCONDIT"
  
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
      FirstTime = False
      LSET Form$(1, 0) = "1"
      LSET Form$(2, 0) = "99"
      LSET Form$(3, 0) = "Y"
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
        ClrNew = Form$(3, 0) = "Y"
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
    
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 19   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 37 TO 46           '--F5 button
          PressButton F10Key, 19, 37, 46
        CASE 49 TO 60           '--Cancel button
          PressButton 27, 19, 49, 60
        END SELECT
      END SELECT                'row
    END IF
  LOOP WHILE inputting = False
  'Free Up Some Memory
  ERASE Form$, Fld, Frm
  
  'Open the Utility Setup File to Grab Meter List Order (Seq or Loc)
  REDIM UBSetupRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetupRec(), UBSetUpLen
  
  HighVar# = UBSetupRec(1).HighRead
  LowVar# = UBSetupRec(1).LowRead
  
  'Open Report File
  ReportFile$ = "WBMTRLST.RPT"
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  Custname$ = SPACE$(25)
  
  ' Location Order ********************************************************
  IdxRecLen = 4 'we are using a long integer
  IdxFileSize& = FileSize&("UBCUSTBK.IDX")
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  
  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it
  
  UBCustFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCustFile LEN = UBCustRecLen
  Cnt& = 1
  GOSUB ConditHeading
  ShowProcessingScrn "Reading Meter Information"
  FOR LCnt& = 1 TO IdxNumOfRecs
    GET #UBCustFile, IdxBuff(LCnt&).RecNum, UBCustRec(1)
    IF UBCustRec(1).DelFlag = 0 THEN
      IF INSTR(UBCustRec(1).HHMsg1, "NOREAD") = 0 THEN
        Book = QPValI(UBCustRec(1).Book)
        IF (Book >= BegRoute AND Book <= EndRoute) AND UBCustRec(1).Status = "A" THEN
          
          IF UBCustRec(1).NewNotes <> True THEN
            GOTO SkipAcct
          ELSE
            IF ClrNew THEN
              UBCustRec(1).NewNotes = False
              PUT #UBCustFile, IdxBuff(LCnt&).RecNum, UBCustRec(1)
            END IF
          END IF
          DidName = False       'Flag used to print a customers name once
          FOR MtrCnt& = 1 TO 7  'find last active meter
            HighFlag = False
            LowFlag = False
            StillFlag = False
            IF LEN(QPTrim$(UBCustRec(1).LocMeters(MtrCnt&).MtrType)) > 0 THEN
              IF UBCustRec(1).LocMeters(MtrCnt&).CurRead < 0 THEN
                UBCustRec(1).LocMeters(MtrCnt&).CurRead = 0
              END IF
              IF UBCustRec(1).LocMeters(MtrCnt&).PrevRead < 0 THEN
                UBCustRec(1).LocMeters(MtrCnt&).PrevRead = 0
              END IF
              '060199 Added
              HiUsed# = UBCustRec(1).MeterAux(MtrCnt&).MaxUsed
              IF HiUsed# = 0 THEN
                HiUsed# = 1
              END IF
              LowUsed# = UBCustRec(1).MeterAux(MtrCnt&).MinUsed
              IF LowUsed# = 0 THEN
                LowUsed# = 1
              END IF
              LowUsed# = LowUsed# + 1
              AvgUse# = UBCustRec(1).LocMeters(MtrCnt&).AvgUse
              UseCnt = UBCustRec(1).LocMeters(MtrCnt&).UseCnt
              IF AvgUse# < 0 THEN
                AvgUse# = 1
              END IF
              IF UseCnt < 0 THEN
                UseCnt = 1
              END IF
              MeterConsp& = UBCustRec(1).LocMeters(MtrCnt&).CurRead - UBCustRec(1).LocMeters(MtrCnt&).PrevRead
              IF MeterConsp& < 0 THEN
                MaxMeterAmt& = 10& ^ (LEN(STR$(UBCustRec(1).LocMeters(MtrCnt&).PrevRead)) - 1)
                MeterConsp& = (MaxMeterAmt& - UBCustRec(1).LocMeters(MtrCnt&).PrevRead) + UBCustRec(1).LocMeters(MtrCnt&).CurRead
              END IF
              MaxRead& = (HiUsed# * (HighVar# / 100))
              IF MeterConsp& > MaxRead& THEN
                HighFlag = True
              END IF
              MinRead& = (LowUsed# * (LowVar# / 100))
              IF MeterConsp& < MinRead& THEN
                LowFlag = True
              END IF
              IF UBCustRec(1).LocMeters(MtrCnt&).CurRead = UBCustRec(1).LocMeters(MtrCnt&).PrevRead THEN
                StillFlag = True
              END IF
              IF LineCnt >= MaxLines THEN
                PRINT #RptHandle, FF$
                GOSUB ConditHeading
              END IF
              IF NOT DidName THEN
                DidName = True
                LSET Custname$ = UBCustRec(1).Custname
                PRINT #RptHandle, USING "#####  "; IdxBuff(LCnt&).RecNum;
                PRINT #RptHandle, UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb; "  "; Custname$; LEFT$(UBCustRec(1).ServAddr, 25);
              END IF
              GOSUB ConditMtrType
              IF PrintMrtFlag THEN
                DidOne = True
                PRINT #RptHandle, TAB(69); QPTrim$(UBCustRec(1).LocMeters(MtrCnt&).MTRNUM); TAB(82); QPTrim$(UBCustRec(1).UserCode2);
                PRINT #RptHandle, TAB(86); USING "###########"; UBCustRec(1).LocMeters(MtrCnt&).CurRead; UBCustRec(1).LocMeters(MtrCnt&).PrevRead; MeterConsp&; AvgUse#  '; MinRead&; MaxRead&; UseCnt
                PRINT #RptHandle, "N1: "; UBCustRec(1).HHMsg1; TAB(27); "N2: "; UBCustRec(1).HHMsg2; TAB(52); "N3: "; UBCustRec(1).HHMsg3
                LineCnt = LineCnt + 2
                IF HighFlag OR LowFlag OR StillFlag THEN
                  PRINT #RptHandle, "*** WARNING ***      Consumption Failed boundary test. Codes: ";
                  IF HighFlag THEN
                    PRINT #RptHandle, "H";
                  END IF
                  IF LowFlag THEN
                    PRINT #RptHandle, "L";
                  END IF
                  IF StillFlag THEN
                    PRINT #RptHandle, "S";
                  END IF
                  PRINT #RptHandle,
                  LineCnt = LineCnt + 1
                END IF
              END IF
            END IF
          NEXT
          DidOne = False
          IF DidName THEN
            PRINT #RptHandle, Dash2$
            LineCnt = LineCnt + 1
          END IF
          '        END IF
        END IF
      END IF
    END IF
SkipAcct:
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
    ShowPctComp LCnt&, IdxNumOfRecs
  NEXT
  
  PRINT #RptHandle, FF$
  CLOSE
  
  ERASE IdxBuff
  
  IF NOT AbortFlag THEN
    Header$ = "Meter High/Low Reading Report"
    PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  END IF
  EXIT SUB
  
ConditHeading:
  page = page + 1
  PRINT #RptHandle, TAB(50); "Master Meter Reading Report"; TAB(110); USING "Page: ###"; page
  PRINT #RptHandle, "Start Route: "; BegRoute; "  End Route:"; EndRoute; TAB(110); "Date: "; DATE$
  PRINT #RptHandle, ""
  PRINT #RptHandle, "Acct.  Location   Customer Name            Service Address          Mtr No.    Size      Current   Previous      Usage    Average"
  PRINT #RptHandle, Dash131$
  LineCnt = 5
RETURN
  
ConditMtrType:
  PrintMrtFlag = False
  SELECT CASE UBCustRec(1).LocMeters(MtrCnt&).MtrType
  CASE "C", "W", "S", "T", "E", "D", "G"
    PrintMrtFlag = True
    'CASE ELSE
    '  PrintMrtFlag = True
  END SELECT
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
  
  Custname$ = SPACE$(25)
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
      IF INSTR(UBCustRec(1).HHMsg1, "NOREAD") = 0 THEN
        LSET Custname$ = UBCustRec(1).Custname
        LSET SAddr$ = LEFT$(UBCustRec(1).ServAddr, 30)
        IF UBCustRec(1).Status = "I" THEN       '<> "A") AND (UBCustRec(1).Status <> "F") THEN
          DidOne = False
          FOR MtrCnt& = 1 TO 7  'find last active meter
            TempRev$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt&).MtrType)
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
                    PRINT #RptHandle, " "; UBCustRec(1).Status; "   "; UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb; USING "  #####"; LCnt&;
                    PRINT #RptHandle, TAB(25); Custname$; TAB(50); SAddr$
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
  SELECT CASE UBCustRec(1).LocMeters(MtrCnt&).MtrType
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

SUB MasterReadReport
  
  REDIM TempScrn(0)
  
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  MaxLines = 58
  FF$ = CHR$(12)
  Dash131$ = STRING$(129, "=")
  Dash2$ = STRING$(129, "-")

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
  QPrintRC "Master Reading Report ", 6, 24, -1
  QPrintRC "]ÄÄÄÄÄÄ", 6, 46, 10
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      FirstTime = False
      LSET Form$(1, 0) = "1"
      LSET Form$(2, 0) = "99"
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
    
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 19   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 37 TO 46           '--F5 button
          PressButton F10Key, 19, 37, 46
        CASE 49 TO 60           '--Cancel button
          PressButton 27, 19, 49, 60
        END SELECT
      END SELECT                'row
    END IF
  LOOP WHILE inputting = False
  'Free Up Some Memory
  ERASE Form$, Fld, Frm
  
  'Open the Utility Setup File to Grab Meter List Order (Seq or Loc)
  REDIM UBSetupRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetupRec(), UBSetUpLen
  
  HighVar# = UBSetupRec(1).HighRead
  LowVar# = UBSetupRec(1).LowRead
  
  'Open Report File
  ReportFile$ = "WBMTRLST.RPT"
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  Custname$ = SPACE$(25)
  
  ' Location Order ********************************************************
  IdxRecLen = 4 'we are using a long integer
  IdxFileSize& = FileSize&("UBCUSTBK.IDX")
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  
  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it
  
  UBCustFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCustFile LEN = UBCustRecLen
  
  Cnt& = 1
  GOSUB HighLowHeading
  ShowProcessingScrn "Reading Meter Information"
  FOR LCnt& = 1 TO IdxNumOfRecs
    GET #UBCustFile, IdxBuff(LCnt&).RecNum, UBCustRec(1)
    IF UBCustRec(1).DelFlag = 0 THEN
      IF INSTR(UBCustRec(1).HHMsg1, "NOREAD") = 0 THEN
        Book = QPValI(UBCustRec(1).Book)
        IF (Book >= BegRoute AND Book <= EndRoute) AND UBCustRec(1).Status = "A" THEN
          DidName = False       'Flag used to print a customers name once
          FOR MtrCnt& = 1 TO 7  'find last active meter
            HighFlag = False
            LowFlag = False
            StillFlag = False
            IF LEN(QPTrim$(UBCustRec(1).LocMeters(MtrCnt&).MtrType)) > 0 THEN
              IF UBCustRec(1).LocMeters(MtrCnt&).CurRead < 0 THEN
                UBCustRec(1).LocMeters(MtrCnt&).CurRead = 0
              END IF
              IF UBCustRec(1).LocMeters(MtrCnt&).PrevRead < 0 THEN
                UBCustRec(1).LocMeters(MtrCnt&).PrevRead = 0
              END IF
              '060199 Added
              HiUsed# = UBCustRec(1).MeterAux(MtrCnt&).MaxUsed
              IF HiUsed# = 0 THEN
                HiUsed# = 1
              END IF
              LowUsed# = UBCustRec(1).MeterAux(MtrCnt&).MinUsed
              IF LowUsed# = 0 THEN
                LowUsed# = 1
              END IF
              LowUsed# = LowUsed# + 1
              AvgUse# = UBCustRec(1).LocMeters(MtrCnt&).AvgUse
              UseCnt = UBCustRec(1).LocMeters(MtrCnt&).UseCnt
              IF AvgUse# < 0 THEN
                AvgUse# = 1
              END IF
              IF UseCnt < 0 THEN
                UseCnt = 1
              END IF
              MeterConsp& = UBCustRec(1).LocMeters(MtrCnt&).CurRead - UBCustRec(1).LocMeters(MtrCnt&).PrevRead
              IF MeterConsp& < 0 THEN
                MaxMeterAmt& = 10& ^ (LEN(STR$(UBCustRec(1).LocMeters(MtrCnt&).PrevRead)) - 1)
                MeterConsp& = (MaxMeterAmt& - UBCustRec(1).LocMeters(MtrCnt&).PrevRead) + UBCustRec(1).LocMeters(MtrCnt&).CurRead
              END IF
              MaxRead& = (HiUsed# * (HighVar# / 100))
              IF MeterConsp& > MaxRead& THEN
                HighFlag = True
              END IF
              MinRead& = (LowUsed# * (LowVar# / 100))
              IF MeterConsp& < MinRead& THEN
                LowFlag = True
              END IF
              IF UBCustRec(1).LocMeters(MtrCnt&).CurRead = UBCustRec(1).LocMeters(MtrCnt&).PrevRead THEN
                StillFlag = True
              END IF
              IF NOT DidName THEN
                DidName = True
                PRINT #RptHandle, USING "#####  "; IdxBuff(LCnt&).RecNum;
                LSET Custname$ = UBCustRec(1).Custname
                PRINT #RptHandle, UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb; "  "; Custname$; LEFT$(UBCustRec(1).ServAddr, 25);
              END IF
              GOSUB HighLowMtrType
              IF PrintMrtFlag THEN
                DidOne = True
                PRINT #RptHandle, TAB(69); QPTrim$(UBCustRec(1).LocMeters(MtrCnt&).MTRNUM); TAB(82); QPTrim$(UBCustRec(1).UserCode2);
                PRINT #RptHandle, TAB(86); USING "###########"; UBCustRec(1).LocMeters(MtrCnt&).CurRead; UBCustRec(1).LocMeters(MtrCnt&).PrevRead; MeterConsp&; AvgUse#  '; MinRead&; MaxRead&; UseCnt
                PRINT #RptHandle, TAB(10); "N1: "; UBCustRec(1).HHMsg1; TAB(40); "N2: "; UBCustRec(1).HHMsg2; TAB(65); "N3: "; UBCustRec(1).HHMsg3
                LineCnt = LineCnt + 2
                IF HighFlag OR LowFlag OR StillFlag THEN
                  PRINT #RptHandle, "*** WARNING ***      Consumption Failed boundary test. Codes: ";
                  IF HighFlag THEN
                    PRINT #RptHandle, "H";
                  END IF
                  IF LowFlag THEN
                    PRINT #RptHandle, "L";
                  END IF
                  IF StillFlag THEN
                    PRINT #RptHandle, "S";
                  END IF
                  PRINT #RptHandle,
                  LineCnt = LineCnt + 1
                END IF
              END IF
            END IF
          NEXT
          DidOne = False
          IF DidName THEN
            PRINT #RptHandle, Dash2$
            LineCnt = LineCnt + 1
          END IF
        END IF
      END IF
    END IF
    IF LineCnt >= MaxLines THEN
      PRINT #RptHandle, FF$
      GOSUB HighLowHeading
    END IF
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
    ShowPctComp LCnt&, IdxNumOfRecs
  NEXT
  
  PRINT #RptHandle, FF$
  CLOSE
  
  ERASE IdxBuff
  
  IF NOT AbortFlag THEN
    Header$ = "Meter High/Low Reading Report"
    PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  END IF
  EXIT SUB
  
HighLowHeading:
  page = page + 1
  PRINT #RptHandle, TAB(50); "Master Meter Reading Report"; TAB(110); USING "Page: ###"; page
  PRINT #RptHandle, "Start Route: "; BegRoute; "  End Route:"; EndRoute; TAB(110); "Date: "; DATE$
  PRINT #RptHandle, ""
  PRINT #RptHandle, "Acct.  Location   Customer Name            Service Address          Mtr No.    Size      Current   Previous      Usage    Average"
  PRINT #RptHandle, Dash131$
  LineCnt = 5
RETURN
  
HighLowMtrType:
  PrintMrtFlag = False
  SELECT CASE UBCustRec(1).LocMeters(MtrCnt&).MtrType
  CASE "C", "W", "S", "T", "E", "D", "G"
    PrintMrtFlag = True
    'CASE ELSE
    '  PrintMrtFlag = True
  END SELECT
  RETURN
  
  
END SUB

SUB PrintMeterNotes
  REDIM TempScrn(0)
  
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  MaxLines = 50
  FF$ = CHR$(12)
  
  LibName$ = "UB"
  ScrnName$ = "UBNOTES"
  
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
  
  Custname$ = SPACE$(30)
  
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
      IF UBCustRec(1).DelFlag = 0 THEN
        IF NewFlag THEN
          IF UBCustRec(1).NewNotes <> True THEN
            GOTO SkipNotes
          ELSE
            UBCustRec(1).NewNotes = False
            'PUT #UBCustFile, IdxBuff(LCnt&).RecNum, UBCustRec(1)
          END IF
        END IF
        LSET Custname$ = UBCustRec(1).Custname
        IF LEN(QPTrim$(UBCustRec(1).HHMsg1)) > 0 OR LEN(QPTrim$(UBCustRec(1).HHMsg2)) > 0 OR LEN(QPTrim$(UBCustRec(1).HHMsg3)) > 0 THEN
          PRINT #RptHandle, UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb; "  "; Custname$; LEFT$(UBCustRec(1).ServAddr, 30)
          PRINT #RptHandle, TAB(3); UBCustRec(1).HHMsg1; TAB(26); UBCustRec(1).HHMsg2; TAB(49); UBCustRec(1).HHMsg3
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
  
  Custname$ = SPACE$(30)
  
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
      IF INSTR(UBCustRec(1).HHMsg1, "NOREAD") = 0 THEN
        Book = QPValI(UBCustRec(1).Book)
        IF Book >= BegRoute AND Book <= EndRoute THEN
          LSET Custname$ = UBCustRec(1).Custname
          IF UBCustRec(1).Status > "" THEN
            PRINT #RptHandle, UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb; "  "; Custname$; LEFT$(UBCustRec(1).ServAddr, 30)
            'IF LEN(QPTrim$(UBCustRec(1).EstFlag)) > 0 THEN STOP
            LineCnt = LineCnt + 1
            FOR MtrCnt& = 1 TO 7                'find last active meter
              TempRev$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt&).MtrType)
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
                  'IF MtrCnt& > 1 THEN STOP
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
  SELECT CASE UBCustRec(1).LocMeters(MtrCnt&).MtrType
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

