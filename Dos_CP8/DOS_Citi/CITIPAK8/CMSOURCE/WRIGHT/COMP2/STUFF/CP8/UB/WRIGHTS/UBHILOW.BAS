  DEFINT A-Z
DECLARE FUNCTION FileSize& (filename$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValI% (Numb$)
DECLARE FUNCTION QPValL& (Numb$)
DECLARE SUB BlockClear ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB FGetAH (filename$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB LoadUBSetUpFile (UBSetupRec() AS ANY, UBSetUpLen%)
DECLARE SUB PrintRptFile (RptTitle$, filename$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB WaitForAction ()
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
  
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'Newcust.bi'
  
  CONST False = 0, True = NOT False

SUB PrintHighLowReport

  REDIM TempScrn(0)
  
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  MaxLines = 58
  FF$ = CHR$(12)
  Dash131$ = STRING$(129, "=")
  Dash2$ = STRING$(129, "-")

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
        UseStl = True
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
              IF UseHigh THEN
                MaxRead& = (HiUsed# * (HighVar# / 100))
                IF MeterConsp& > MaxRead& THEN
                  HighFlag = True
                END IF
              END IF
              IF UseLow THEN
                MinRead& = (LowUsed# * (LowVar# / 100))
                IF MeterConsp& < MinRead& THEN
                  LowFlag = True
                END IF
              END IF
              IF UseStl THEN
                IF UBCustRec(1).LocMeters(MtrCnt&).CurRead = UBCustRec(1).LocMeters(MtrCnt&).PrevRead THEN
                  StillFlag = True
                END IF
              END IF
              IF HighFlag OR LowFlag OR StillFlag THEN
                IF LineCnt >= MaxLines THEN
                  PRINT #RptHandle, FF$
                  GOSUB HighLowHeading
                END IF
                IF NOT DidName THEN
                  DidName = True
                  LSET Custname$ = UBCustRec(1).Custname
                  PRINT #RptHandle, USING "#####  "; IdxBuff(LCnt&).RecNum;
                  PRINT #RptHandle, UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb; "  "; Custname$; LEFT$(UBCustRec(1).ServAddr, 25);
                END IF
                GOSUB HighLowMtrType
                IF PrintMrtFlag THEN
                  DidOne = True
                  PRINT #RptHandle, TAB(69); QPTrim$(UBCustRec(1).LocMeters(MtrCnt&).MTRNUM); TAB(82); QPTrim$(UBCustRec(1).UserCode2);
                  PRINT #RptHandle, TAB(86); USING "###########"; UBCustRec(1).LocMeters(MtrCnt&).CurRead; UBCustRec(1).LocMeters(MtrCnt&).PrevRead; MeterConsp&; AvgUse#  '; MinRead&; MaxRead&; UseCnt
                  LineCnt = LineCnt + 1
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
  Page = Page + 1
  PRINT #RptHandle, TAB(50); "Meter Reading High/Low Report"; TAB(110); USING "Page: ###"; Page
  PRINT #RptHandle, "Start Route: "; BegRoute; "  End Route:"; EndRoute; TAB(110); "Date: "; DATE$
  PRINT #RptHandle, ""
  PRINT #RptHandle, "Acct.  Location   Customer Name            Service Address          Mtr No.    Size      Current   Previous      Usage    Average"
  PRINT #RptHandle, Dash131$
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

