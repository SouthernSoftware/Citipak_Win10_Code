DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValI% (Numb$)
DECLARE FUNCTION QPValL& (Numb$)
DECLARE SUB BlockClear ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetUpLen%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB WaitForAction ()
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE FUNCTION Round# (DblNum#)
DECLARE SUB MakeSequenceIndex (Text$)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FClose (Handle%)
DECLARE SUB QPrintRC (Text$, Row%, Col%, Attrib%)
DECLARE SUB SaveScrn (TempScrn())
DECLARE SUB RestScrn (TempScrn())
DECLARE FUNCTION FmtBook$ (Text$)
DECLARE FUNCTION FmtSeqN$ (Text$)
DECLARE FUNCTION Chk4BookSeqNum% (Text1$, Text2$)
DECLARE FUNCTION MsgBox% (Lib$, Scrn$)


  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'Newcust.bi'
  
  CONST False = 0, True = NOT False

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
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen
  
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
            TempRev$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt&).MtrType)
            IF LEN(TempRev$) > 0 THEN
              GOSUB GetInactMtrTypePrint
              IF PrintMrtFlag THEN
                Multi# = UBCustRec(1).LocMeters(MtrCnt&).MTRMulti
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

SUB PrintHighLowReport
  
  REDIM TempScrn(0)
  
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  MaxLines = 40
  FF$ = CHR$(12)
  
  LibName$ = "UB"
  ScrnName$ = "UBHIGHLO"
  
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
      FOR zz = 3 TO 5
        LSET Form$(zz, 0) = "Y"
      NEXT
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
      IF Form$(3, 0) = "Y" THEN
        UseHigh = True
      END IF
      IF Form$(4, 0) = "Y" THEN
        UseLow = True
      END IF
      IF Form$(5, 0) = "Y" THEN
        UseUnRead = True
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
      CASE 18   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 37 TO 46           '--F5 button
          PressButton F10Key, 18, 37, 46
        CASE 49 TO 60           '--Cancel button
          PressButton 27, 18, 49, 60
        END SELECT
      END SELECT                'row
    END IF
  

  LOOP WHILE inputting = False
  'Free Up Some Memory
  ERASE Form$, Fld, Frm
  
  'Open the Utility Setup File to Grab Meter List Order (Seq or Loc)
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen
  
  HighVar# = UBSetUpRec(1).HighRead
  LowVar# = UBSetUpRec(1).LowRead
  
  'Open Report File
  ReportFile$ = "WBMTRLST.RPT"
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
  GOSUB HighLowHeading
  ShowProcessingScrn "Reading Meter Information"
  FOR LCnt& = 1 TO IdxNumOfRecs
    GET #UBCustFile, IdxBuff(LCnt&).RecNum, UBCustRec(1)

    'IF IdxBuff(LCnt&).RecNum = 879 THEN STOP

    IF UBCustRec(1).DelFlag = 0 THEN
      IF INSTR(UBCustRec(1).HHMSG1, "NOREAD") = 0 THEN
        BOOK = QPValI(UBCustRec(1).BOOK)
        IF (BOOK >= BegRoute AND BOOK <= EndRoute) AND UBCustRec(1).Status = "A" THEN
          'IF UBCustRec(1).Status = "A" OR INSTR(UBCustRec(1).CustName, "VACA") > 0 THEN
          DidName = False           'Flag used to print a customers name once
          FOR MtrCnt& = 1 TO 7      'find last active meter
            UNReadFlag = False
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
          
              IF UseHigh THEN
                HiAvg# = (AvgUse# * (HighVar# / 100))
                IF MeterConsp& > HiAvg# THEN
                  HighFlag = True
                END IF
              END IF
              IF UseLow THEN
                LoAvg# = (AvgUse# * (LowVar# / 100))
                IF MeterConsp& < LoAvg# THEN
                  LowFlag = True
                END IF
              END IF

              IF UseUnRead THEN
                IF LEN(QPTrim$(UBCustRec(1).LocMeters(MtrCnt&).ReadFlag)) = 0 THEN
                  UNReadFlag = True
                END IF
              END IF

              IF UBCustRec(1).LocMeters(MtrCnt&).CurRead = UBCustRec(1).LocMeters(MtrCnt&).PrevRead THEN
                StillFlag = True
              END IF
            
              IF HighFlag OR LowFlag OR UNReadFlag OR StillFlag THEN
                IF LineCnt >= MaxLines THEN
                  PRINT #RptHandle, FF$
                  GOSUB HighLowHeading
                END IF

                IF NOT DidName THEN
                  DidName = True
                  LSET CUSTNAME$ = UBCustRec(1).CUSTNAME
                  PRINT #RptHandle, UBCustRec(1).BOOK; "-"; UBCustRec(1).SEQNUMB; "  "; CUSTNAME$; LEFT$(UBCustRec(1).SERVADDR, 30)
                END IF
                GOSUB HighLowMtrType
                IF PrintMrtFlag THEN
                  DidOne = True
                  IF UBCustRec(1).LocMeters(MtrCnt&).ReadFlag <> "Y" THEN
                    UNReadFlag = True
                  END IF
                  PRINT #RptHandle, MeterType$; TAB(12); QPTrim$(UBCustRec(1).LocMeters(MtrCnt&).MTRNUM);
                  PRINT #RptHandle, TAB(24); USING "##########"; UBCustRec(1).LocMeters(MtrCnt&).CurRead; UBCustRec(1).LocMeters(MtrCnt&).PrevRead; MeterConsp&;
                  PRINT #RptHandle, " ";
                  IF HighFlag THEN
                    PRINT #RptHandle, "H";
                  END IF
                  IF LowFlag THEN
                    PRINT #RptHandle, "L";
                  END IF
                  IF StillFlag THEN
                    PRINT #RptHandle, "S";
                  END IF
                  IF UNReadFlag THEN
                    PRINT #RptHandle, "U";
                    UNReadFlag = False
                  END IF

                  IF LEN(QPTrim$(UBCustRec(1).EstFlag)) > 0 THEN
                    PRINT #RptHandle, UBCustRec(1).EstFlag;
                  END IF
                  PRINT #RptHandle, TAB(61); Num2Date$(UBCustRec(1).LocMeters(MtrCnt&).CurDate);
                  PRINT #RptHandle, USING "########"; AvgUse#
                  LineCnt = LineCnt + 1
                END IF
              END IF
            END IF
          NEXT
          DidOne = False
          IF DidName THEN
            PRINT #RptHandle, STRING$(79, "-")
            LineCnt = LineCnt + 1
          END IF
'        END IF
        END IF
      END IF
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
  PRINT #RptHandle, TAB(27); "Meter Reading High/Low Report"; TAB(65); "Date: "; DATE$
  PRINT #RptHandle, "Start Route: "; BegRoute; "  End Route:"; EndRoute; TAB(70); "Page #"; page
  PRINT #RptHandle, ""
  PRINT #RptHandle, "Location   Customer Name               Service Address                   Average"
  PRINT #RptHandle, "Mtr Type   Mtr No.        Current    Previous     Usage      Read Date    Usage "
  PRINT #RptHandle, STRING$(80, "=")
  LineCnt = 5
RETURN
  
HighLowMtrType:
  PrintMrtFlag = False
  SELECT CASE UBCustRec(1).LocMeters(MtrCnt&).MtrType
  CASE "C"
    MeterType$ = "Wat-Sew"
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
    MeterType$ = "?????"
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
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen
  
  IF UBSetUpRec(1).UseSeq = "Y" THEN
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
    PRec& = IndexArray(Cnt).RecNum
    IF NOT PRec& = 0 THEN
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
    TempRev$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrType)
    IF LEN(TempRev$) <> 0 THEN ValidCustomer = True: EXIT DO
  LOOP UNTIL MtrCnt = 7
  IF ValidCustomer = False THEN RETURN
  
  IF LineCnt >= 53 THEN
    PRINT #RptHandle, CHR$(12);
    GOSUB PrintHeading
  END IF
  
  GOSUB GetMeterStatusPrint
  
  PRINT #RptHandle, LEFT$(UBCustRec(1).CUSTNAME, 30);
  IF UBSetUpRec(1).UseSeq = "Y" THEN
    IF UBCustRec(1).SEQ < 0 THEN UBCustRec(1).SEQ = 0
    PRINT #RptHandle, TAB(32); USING "######"; UBCustRec(1).SEQ;
  END IF
  PRINT #RptHandle, TAB(40); LEFT$(UBCustRec(1).SERVADDR, 28);
  PRINT #RptHandle, TAB(70); MeterStatus$
  LineCnt = LineCnt + 1
  
  FOR MtrCnt = 1 TO 7           'find last active meter
    TempRev$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrType)
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
  IF UBSetUpRec(1).UseSeq = "Y" THEN
    PRINT #RptHandle, TAB(32); "Seq #";
  END IF
  PRINT #RptHandle, TAB(40); "Service Address"; TAB(70); "Status"
  PRINT #RptHandle, "Location"; TAB(15); "Meter Number"; TAB(35); "Mtr Type"; TAB(55); "Cur Read"; TAB(70); "New Read"
  PRINT #RptHandle, STRING$(80, "=")
  LineCnt = 5
  RETURN
  
GetCustRecNoPrint:
  RecNo& = PRec&
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
  SELECT CASE UBCustRec(1).LocMeters(MtrCnt).MtrType
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
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen
  
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
                  Multi# = UBCustRec(1).LocMeters(MtrCnt&).MTRMulti
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

SUB PrintMeterSheets
  BlockClear
  
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen

  IF INSTR(UBSetUpRec(1).UTILNAME, "WHITE LAKE") THEN
     WhiteLakeFlag = 1
  END IF
  
  IF UBSetUpRec(1).UseSeq = "Y" THEN
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
    AcctNumber& = IdxBuff(LCnt&).RecNum
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
  IF WhiteLakeFlag <> 1 THEN
   GOSUB PrintSkipHeader
  END IF

  ZONE$ = QPTrim$(UBCustRec(1).ZONE)
  SELECT CASE LEFT$(QPTrim$(UBCustRec(1).CUSTTYPE), 1)
  CASE "B"
    CustT$ = "Commerical"
  CASE "R"
    CustT$ = "Residential"
  CASE ELSE
    CustT$ = "??????????"
  END SELECT


  IF WhiteLakeFlag = 1 THEN
   PRINT #RptHandle,
   PRINT #RptHandle,
   PRINT #RptHandle, TAB(15); "Acct #"; AcctNumber&
   PRINT #RptHandle,
   PRINT #RptHandle, TAB(15); TAB(15); UBCustRec(1).CUSTNAME
   PRINT #RptHandle,
   PRINT #RptHandle,
   PRINT #RptHandle, TAB(15); UBCustRec(1).ADDR1
   PRINT #RptHandle, TAB(15); UBCustRec(1).ADDR2
   PRINT #RptHandle, TAB(15); QPTrim$(UBCustRec(1).CITY); " "; UBCustRec(1).STATE; " "; UBCustRec(1).ZIPCODE
   PRINT #RptHandle,
   PRINT #RptHandle,
   PRINT #RptHandle, TAB(15); "SERVICE AT "; UBCustRec(1).SERVADDR
   PRINT #RptHandle,
   PRINT #RptHandle, TAB(14); RIGHT$(DATE$, 2); TAB(43); UBCustRec(1).BOOK; "-"; UBCustRec(1).SEQNUMB
   FOR zz = 16 TO 43
   PRINT #RptHandle,
   NEXT zz
   PRINT #RptHandle, TAB(15); WatRead&
   PRINT #RptHandle,
   PRINT #RptHandle,
   PRINT #RptHandle,
   PRINT #RptHandle,
   PRINT #RptHandle,
   PRINT #RptHandle, TAB(15); WatSer$
   FOR zz = 51 TO 59
    PRINT #RptHandle,
   NEXT zz
    PRINT #RptHandle,
    RETURN
  END IF
  



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
    SELECT CASE UBCustRec(1).LocMeters(MtrCnt).MtrType
    CASE "C"
      WatFlag = True
      SewFlag = True
      WatMin = UBCustRec(1).LocMeters(MtrCnt).NumUser
      SewMin = UBCustRec(1).LocMeters(MtrCnt).NumUser
      WatRead& = UBCustRec(1).LocMeters(MtrCnt).PrevRead
      WatSer$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MTRNUM)
      IF LEN(WatSer$) = 0 THEN
        WatSer$ = "?????"
      END IF
    CASE "W"
      WatFlag = True
      WatMin = UBCustRec(1).LocMeters(MtrCnt).NumUser
      WatSer$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MTRNUM)
      WatRead& = UBCustRec(1).LocMeters(MtrCnt).PrevRead
      IF LEN(WatSer$) = 0 THEN
        WatSer$ = "?????"
      END IF
    CASE "S"
      SewFlag = True
      SewMin = UBCustRec(1).LocMeters(MtrCnt).NumUser
      WatSer$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MTRNUM)
      WatRead& = UBCustRec(1).LocMeters(MtrCnt).PrevRead
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
  IF WhiteLakeFlag = 1 THEN
   PRINT #RptHandle, "TOP"
   PRINT #RptHandle,
   PRINT #RptHandle, TAB(15); "Acct # XXXXXX"
   PRINT #RptHandle,
   PRINT #RptHandle, TAB(15); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
   PRINT #RptHandle,
   PRINT #RptHandle,
   PRINT #RptHandle, TAB(15); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
   PRINT #RptHandle, TAB(15); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
   PRINT #RptHandle, TAB(15); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
   PRINT #RptHandle, TAB(15); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
   PRINT #RptHandle,
   PRINT #RptHandle,
   PRINT #RptHandle,
   PRINT #RptHandle, TAB(14); RIGHT$(DATE$, 2); TAB(43); "XX-XXXXXXX"
   FOR zz = 16 TO 43
   PRINT #RptHandle,
   NEXT zz
   PRINT #RptHandle, TAB(15); "XXXXXXXXXX"
   PRINT #RptHandle,
   PRINT #RptHandle,
   PRINT #RptHandle,
   PRINT #RptHandle,
   PRINT #RptHandle,
   PRINT #RptHandle, TAB(15); "XXXXXXXXXXX"
   FOR zz = 51 TO 59
   PRINT #RptHandle,
   NEXT zz
   PRINT #RptHandle, "BOTTOM"
   ELSE
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
  END IF
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

