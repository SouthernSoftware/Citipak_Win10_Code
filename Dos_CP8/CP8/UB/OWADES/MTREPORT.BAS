DEFINT A-Z

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
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize&("UBCUSTBK.IDX")
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    
    REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
    FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs  'load it
    
    UBCustFile = FREEFILE
    OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCustFile LEN = UBCustRecLen
    
    Cnt& = 1
    GOSUB PrintReadHeading
    ShowProcessingScrn "Reading Meter Information"
    FOR LCnt& = 1 TO IdxNumOfRecs
      GET #UBCustFile, IdxBuff(LCnt&).RecNum, UBCustRec(1)
      Book = QPValI(UBCustRec(1).Book)
      IF Book >= BegRoute AND Book <= EndRoute THEN
        LSET CustName$ = UBCustRec(1).CustName
        IF UBCustRec(1).Status > "" THEN
          PRINT #RptHandle, UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb; "  "; CustName$; LEFT$(UBCustRec(1).SERVADDR, 30)
          FOR MtrCnt& = 1 TO 7  'find last active meter
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

