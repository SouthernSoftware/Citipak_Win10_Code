DEFINT A-Z
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
DECLARE SUB PrintUtilBills ()
DECLARE SUB MakeSequenceIndex (IndexText$)
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
DECLARE FUNCTION FindRateTbl% (RateCode$, NumOfRates%, UBRateTbls() AS ANY)
DECLARE FUNCTION GetNumOfRevs% ()
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
DECLARE FUNCTION FLof& (FileHand%)
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
  '$INCLUDE: 'ubsetbal.BI'
  '$INCLUDE: 'ubgilsen.bi'

  CONST False = 0, True = NOT False

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
      
      Choice$(0, 0) = STR$(DevFld)              'Choices for Output Device
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
          
          IF Cnt < 1 THEN Cnt = 1               ' Do Not Allow to Fall Below 1
          inputting = False     ' Set Edit Finish to No
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
      
      IdxRecLen = 4             'we are using a integer
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
      help$ = "Process Location Record #" + STR$(Cnt) + " of " + STR$(IdxNumOfRecs)
      PrintHelp help$
      
      MtrCnt = 0
      ValidCustomer = False
      DO
        MtrCnt = MtrCnt + 1     'Check For Meter This Customer
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
        IF UBCustRec(1).SEQ < 0 THEN UBCustRec(1).SEQ = 0
        PRINT #RptHandle, TAB(32); USING "######"; UBCustRec(1).SEQ;
      END IF
      PRINT #RptHandle, TAB(40); LEFT$(UBCustRec(1).ServAddr, 28);
      PRINT #RptHandle, TAB(70); MeterStatus$
      LineCnt = LineCnt + 1
      
      FOR MtrCnt = 1 TO 7       'find last active meter
        TempRev$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MTRType)
        IF LEN(TempRev$) <> 0 THEN
          GOSUB GetMeterTypePrint
          PRINT #RptHandle, UBCustRec(1).BOOK; "-"; UBCustRec(1).SeqNumb;
          PRINT #RptHandle, TAB(15); QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum);
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
      
      FOpenS "UBCUST.DAT", L2Handle             'open data file
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
      ReportFile$ = "UBMTRNOT.RPT"
      RptHandle = FREEFILE
      OPEN ReportFile$ FOR OUTPUT AS #RptHandle
      
      'Open the Utility Setup File to Grab Meter List Order (Seq or Loc)
      'REDIM UBSetupRec(1) AS UBSetupRecType
      'LoadUBSetUpFile UBSetupRec(), UBSetUpLen
      
      CustName$ = SPACE$(30)
      
      ' Location Order ********************************************************
      IdxRecLen = 4             'we are using a long integer
      IdxFileSize& = FileSize&("UBCUSTBK.IDX")
      IdxNumOfRecs = IdxFileSize& \ IdxRecLen
      
      REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
      FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs                'load it
      
      UBCustFile = FREEFILE
      OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCustFile LEN = UBCustRecLen
      
      Cnt& = 1
      GOSUB PrintNotesHeading
      ShowProcessingScrn "Reading Meter Information"
      FOR LCnt& = 1 TO IdxNumOfRecs
        GET #UBCustFile, IdxBuff(LCnt&).RecNum, UBCustRec(1)
        BOOK = QPValI(UBCustRec(1).BOOK)
        IF BOOK >= BegRoute AND BOOK <= EndRoute THEN
          LSET CustName$ = UBCustRec(1).CustName
          IF UBCustRec(1).DelFlag = 0 THEN
            IF LEN(QPTrim$(UBCustRec(1).HHMSG1)) > 0 OR LEN(QPTrim$(UBCustRec(1).HHMSG2)) > 0 OR LEN(QPTrim$(UBCustRec(1).HHMSG3)) > 0 THEN
              PRINT #RptHandle, UBCustRec(1).BOOK; "-"; UBCustRec(1).SeqNumb; "  "; CustName$; LEFT$(UBCustRec(1).ServAddr, 30)
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
      PRINT #RptHandle, TAB(30); "Meter Reading Notes"
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
      IdxRecLen = 4             'we are using a long integer
      IdxFileSize& = FileSize&("UBCUSTBK.IDX")
      IdxNumOfRecs = IdxFileSize& \ IdxRecLen
      
      REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
      FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs                'load it
      
      UBCustFile = FREEFILE
      OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCustFile LEN = UBCustRecLen
      
      Cnt& = 1
      GOSUB PrintReadHeading
      ShowProcessingScrn "Reading Meter Information"
      FOR LCnt& = 1 TO IdxNumOfRecs
        GET #UBCustFile, IdxBuff(LCnt&).RecNum, UBCustRec(1)
        BOOK = QPValI(UBCustRec(1).BOOK)
        IF BOOK >= BegRoute AND BOOK <= EndRoute THEN
          LSET CustName$ = UBCustRec(1).CustName
          IF UBCustRec(1).Status > "" THEN
            PRINT #RptHandle, UBCustRec(1).BOOK; "-"; UBCustRec(1).SeqNumb; "  "; CustName$; LEFT$(UBCustRec(1).ServAddr, 30)
            FOR MtrCnt& = 1 TO 7                'find last active meter
              TempRev$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt&).MTRType)
              IF LEN(TempRev$) > 0 THEN
                GOSUB GetReadMeterTypePrint
                IF PrintMrtFlag THEN
                  DidOne = True
                  PRINT #RptHandle, TAB(14); MeterType$;
                  PRINT #RptHandle, TAB(31); USING "##########"; UBCustRec(1).LocMeters(MtrCnt&).CurRead;
                  PRINT #RptHandle, TAB(42); USING "##########"; UBCustRec(1).LocMeters(MtrCnt&).PrevRead;
                  MeterConsp& = UBCustRec(1).LocMeters(MtrCnt&).CurRead - UBCustRec(1).LocMeters(MtrCnt&).PrevRead
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
        ShowPctComp LCnt&, IdxNumOfRecs
      NEXT
      
      PRINT #RptHandle, FF$
      CLOSE
      
      Header$ = "Customer Meter Reading Report"
      PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
      
      EXIT SUB
      
PrintReadHeading:
      page = page + 1
      PRINT #RptHandle, TAB(30); "Meter Reading Report"
      PRINT #RptHandle, "Beginning Route: "; BegRoute
      PRINT #RptHandle, "   Ending Route: "; EndRoute; TAB(70); "Page #"; page
      PRINT #RptHandle, ""
      PRINT #RptHandle, "Location   Customer Name"; TAB(41); "Service Address"
      PRINT #RptHandle, "              Mtr Type"; TAB(31); "   Current   Previous    Consump     Read Date"
      PRINT #RptHandle, STRING$(80, "=")
      LineCnt = 5
      
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
        
        IF Frm(1).FldNo = 3 AND Frm(1).PrevFld < 3 THEN         '>2
          In1 = True
          GOSUB CheckBookSequence
        ELSEIF Frm(1).FldNo > 4 AND Frm(1).PrevFld = 4 THEN     '>2
          In1 = False
          GOSUB CheckBookSequence
        END IF
        
        SELECT CASE Frm(1).KeyCode              '--Check for Key presses
          
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
          
        CASE F5Key
          GOSUB DoSheetMask
          
        CASE EscKey
          GOTO ExitPrintSheets
          
        END SELECT
        
      LOOP UNTIL Done
      
      'Free Up Some Memory
      ERASE Form$, Fld, Frm
      
      ' Location Order ********************************************************
      IdxRecLen = 4             'we are using a long integer
      IdxFileSize& = FileSize&("UBCUSTBK.IDX")
      IdxNumOfRecs = IdxFileSize& \ IdxRecLen
      
      REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
      FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs                'load it
      
      UBCustFile = FREEFILE
      OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCustFile LEN = UBCustRecLen
      
      RptHandle = FREEFILE
      OPEN ReportFile$ FOR OUTPUT AS #RptHandle 'Open Report File
      
      ShowProcessingScrn "Reading Meter Information"
      FOR LCnt& = 1 TO IdxNumOfRecs
        GET #UBCustFile, IdxBuff(LCnt&).RecNum, UBCustRec(1)
        BOOK = QPValL(UBCustRec(1).BOOK)
        Sequ& = QPValL(UBCustRec(1).SeqNumb)
        IF BOOK >= Book1& AND BOOK <= Book2& THEN
          IF Sequ& >= Sequ1& AND Sequ& <= Sequ2& THEN
            IF UBCustRec(1).DelFlag = 0 THEN
              GOSUB GetMeterFlags
              GOSUB LookForSecLights
              GOSUB LookForTrash
              GOSUB PrintSkipHeader
              Zone$ = QPTrim$(UBCustRec(1).Zone)
              SELECT CASE LEFT$(QPTrim$(UBCustRec(1).CustType), 1)
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
                PRINT #RptHandle, " Water"; "  "; Zone$; "  Min ="; WatMin
              ELSE
                PRINT #RptHandle, ""
              END IF
              IF SewFlag THEN
                PRINT #RptHandle, " Sewer"; "  "; Zone$; "  Min ="; SewMin
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
              PRINT #RptHandle, UBCustRec(1).Addr1
              PRINT #RptHandle, UBCustRec(1).ServAddr
              PRINT #RptHandle, QPTrim$(UBCustRec(1).City); " "; UBCustRec(1).State; " "; UBCustRec(1).ZipCode
              PRINT #RptHandle,
              PRINT #RptHandle, TAB(11); UBCustRec(1).BOOK; "-"; UBCustRec(1).SeqNumb
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
          WatSer$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)
          IF LEN(WatSer$) = 0 THEN
            WatSer$ = "?????"
          END IF
        CASE "W"
          WatFlag = True
          WatMin = UBCustRec(1).LocMeters(MtrCnt).NumUser
          WatSer$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)
          WatRead& = UBCustRec(1).LocMeters(MtrCnt).CurRead
          IF LEN(WatSer$) = 0 THEN
            WatSer$ = "?????"
          END IF
        CASE "S"
          SewFlag = True
          SewMin = UBCustRec(1).LocMeters(MtrCnt).NumUser
          WatSer$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)
          WatRead& = UBCustRec(1).LocMeters(MtrCnt).CurRead
          IF LEN(WatSer$) = 0 THEN
            WatSer$ = "?????"
          END IF
        CASE "E"
          EleFlag = True
          EleMin = UBCustRec(1).LocMeters(MtrCnt).NumUser
          EleSer$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)
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
      OPEN ReportFile$ FOR OUTPUT AS #RptHandle 'Open Report File
      
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

