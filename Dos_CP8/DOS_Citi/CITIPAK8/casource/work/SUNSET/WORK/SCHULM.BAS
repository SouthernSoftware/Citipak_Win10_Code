DECLARE SUB DisplayH2Scrn (ScrnName$)
DECLARE SUB DisplayHHScrn (ScrnName$)
DEFINT A-Z

DEFSNG A-Z
DEFINT A-Z
SUB ProcessSchlum
  
  ExportFileName$ = "F:\ezroute\HOST2PC.IMP"
  ImportFileName$ = "F:\ezroute\PC2HOST.EXP"
  
  FF$ = CHR$(12)
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen
  
  HighVar = UBSetUpRec(1).HighRead
  LowVar = UBSetUpRec(1).LowRead
  
  CrLf$ = CHR$(13) + CHR$(10)
  q$ = CHR$(34)
  cb$ = SPACE$(45)
  SHARED Choice$()
  REDIM Route(100)
  
  LibName$ = "UBHH"
  ScrnName$ = "UBSCHUMI"
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
  Choice$(1, 0) = "Export Meter Readings"
  Choice$(2, 0) = "Import Meter Readings"
  
  Action = 1
  
  DisplayH2Scrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      WhatOp$ = LEFT$(Form$(1, 0), 1)
      IF WhatOp$ <> "E" AND WhatOp$ <> "I" THEN
        QPrintRC "Invalid Operation", 12, 40, 12
        BEEP
        WaitForAction
        QPrintRC STRING$(25, 32), 12, 40, 15
        Frm(1).FldNo = 1
        Action = 1
        Done = False
      ELSE
        Done = True
      END IF
    CASE ESC
      EXIT SUB
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  
  SELECT CASE WhatOp$
  CASE "E"
    GOSUB ExportSchlum
  CASE "I"
    GOSUB ImportSchlum
  CASE ELSE
  END SELECT
  
  EXIT SUB
  
ExportSchlum:
  
  LibName$ = "UBHH"
  ScrnName$ = "UBSCHL1S"
  
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
      IF NumberofRoutes > 0 THEN                ' Next Check Routes
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

        'Open and Initialize the Reading File
        
        KillFile ExportFileName$
        
        REDIM UBSchlumHHRec(1) AS SchlumHHType
        UBSchlumHHRecLen = LEN(UBSchlumHHRec(1))
        UBSchlFile = FREEFILE
        OPEN ExportFileName$ FOR RANDOM SHARED AS UBSchlFile LEN = UBSchlumHHRecLen
        
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
        WriteCnt& = 0
        PageCnt& = 0
        
        DO
          PRec& = IdxBuff(Cnt!).RecNum
          IF NOT PRec& = 0 THEN
            GET UBFile, PRec&, UBCustRec(1)
            IF INSTR(UBCustRec(1).HHMSG1, "NOREAD") > 0 THEN
              GOTO HWriteSkip
            END IF
            DisplayAcct$ = UBCustRec(1).Book + "-" + UBCustRec(1).SeqNumb
            QPrintRC DisplayAcct$, 11, 45, -1
            FOR BookCnt! = 1 TO NumberofRoutes
              IF VAL(UBCustRec(1).Book) = Route(BookCnt!) AND (UBCustRec(1).Status <> "F") THEN
                GOSUB SchlumWriteRec
              END IF
            NEXT BookCnt!
          END IF
HWriteSkip:
          Cnt! = Cnt! + 1
        LOOP UNTIL Cnt! > IdxNumOfRecs
        CLOSE
        DisplayH2Scrn "EXPORTOK"
        WaitForAction
        Done = True
      END IF
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  
  RETURN
  
SchlumWriteRec:
  'May Have Up to 7 Meters to Read
  FOR MtrCnt! = 1 TO 7
    MeterOK = False
    Account$ = STR$(PRec&)
    Account$ = LEFT$(Account$, 6) + "-" + RIGHT$(STR$(MtrCnt!), 1)
    MtrType$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt!).MtrType)
    IF LEN(MtrType$) > 0 THEN
      IF MtrType$ = "W" OR MtrType$ = "S" OR MtrType$ = "C" OR MtrType$ = "E" OR MtrType$ = "D" OR MtrType$ = "G" THEN
        SELECT CASE WhatTypes$
        CASE "W"                'water/sewer
          IF MtrType$ = "W" OR MtrType$ = "S" OR MtrType$ = "C" THEN
            MeterOK = True
          END IF
        CASE "E"                'electric & demand elec.
          IF MtrType$ = "E" OR MtrType$ = "D" THEN
            MeterOK = True
          END IF
        CASE "G"                'gas
          IF MtrType$ = "G" THEN
            MeterOK = True
          END IF
        CASE "A", " "           'all meters
          MeterOK = True
        END SELECT
        
        IF MeterOK = True THEN  ' Determine High and Low Reading
          HasZ = 0
          SRouteID$ = UBCustRec(1).Book + UBCustRec(1).SeqNumb + "0" + QPTrim$(STR$(MtrCnt!))
          'LSET UBSchlumHHRec(1).Route = SRouteID$
          LSET UBSchlumHHRec(1).Route = QPTrim$(UBCustRec(1).Book)
          PageCnt& = PageCnt& + 10
          WriteCnt& = WriteCnt& + 1
          
          WalkSeq$ = "0000" + QPTrim$(STR$(WriteCnt&))
          PageNum$ = "0000" + QPTrim$(STR$(PageCnt&))
          
          'LSET UBSchlumHHRec(1).WalkSeq$ = RIGHT$(WalkSeq$, 4)
          LSET UBSchlumHHRec(1).WalkSeq$ = ""
          
          UBSchlumHHRec(1).PageNum = RIGHT$(PageNum$, 4)
          UBSchlumHHRec(1).ReadSeq = "01"       'UBSchlumHHRec(1).WalkSeq$
          UBSchlumHHRec(1).HHID = ""            'HH Number 'no information
          UBSchlumHHRec(1).ReadDir = "L"
          UBSchlumHHRec(1).NumDial = ""
          
          MeterID$ = UCASE$(QPTrim$(UBCustRec(1).LocMeters(MtrCnt!).MTRNUM))

          IF RIGHT$(MeterID$, 1) = "Z" THEN
            HasZ = -1
            MeterID$ = MID$(MeterID$, 1, (LEN(MeterID$) - 1))
          END IF

          LSET UBSchlumHHRec(1).IDExpected = MeterID$
          LSET UBSchlumHHRec(1).IDCaptured = ""
          LSET UBSchlumHHRec(1).IDOverride = ""
          LSET UBSchlumHHRec(1).Decimals = ""
          LSET UBSchlumHHRec(1).MtrRead = ""    'current from HH
          LSET UBSchlumHHRec(1).ReadOVRide = ""
          
          Average# = UBCustRec(1).LocMeters(MtrCnt!).AvgUse
          ILowRead$ = QPTrim$(STR$((UBCustRec(1).LocMeters(MtrCnt!).CurRead + 1)))
          PrevRead$ = QPTrim$(STR$(UBCustRec(1).LocMeters(MtrCnt!).PrevRead))

          PrevDate$ = Num2Date(UBCustRec(1).LocMeters(MtrCnt!).PastDate)
          PrevDate$ = LEFT$(PrevDate$, 2) + MID$(PrevDate$, 4, 2) + RIGHT$(PrevDate$, 2)
          
          LowRead# = FIX(ILowRead#)
          
          HighRead# = FIX((Average# * (HighVar / 100) - Average#) + UBCustRec(1).LocMeters(MtrCnt!).CurRead)
          HighRead$ = QPTrim$(STR$(HighRead#))
          LSET UBSchlumHHRec(1).HighLimit = HighRead$
          LSET UBSchlumHHRec(1).LowLimit = ILowRead$
          LSET UBSchlumHHRec(1).Date2Read = ""
          LSET UBSchlumHHRec(1).Date2Exp = ""
          LSET UBSchlumHHRec(1).NoteCodes = ""
          LSET UBSchlumHHRec(1).LocatCode = ""
          LSET UBSchlumHHRec(1).MtrRCode = ""
          LSET UBSchlumHHRec(1).RecType = "EU"
          IF HasZ THEN
            LSET UBSchlumHHRec(1).RecStatus = "Z"
          ELSE
            LSET UBSchlumHHRec(1).RecStatus = ""
          END IF
          LSET UBSchlumHHRec(1).ReadDate = ""
          LSET UBSchlumHHRec(1).ReadTime = ""
          LSET UBSchlumHHRec(1).ReadType = ""
          LSET UBSchlumHHRec(1).NetNumb = ""
          LSET UBSchlumHHRec(1).ReadAtmpt = ""
          LSET UBSchlumHHRec(1).UserChar = ""
          LSET UBSchlumHHRec(1).HHManufac = ""
          LSET UBSchlumHHRec(1).ActStatus = UBCustRec(1).Status
          LSET UBSchlumHHRec(1).MtrType = ""
          LSET UBSchlumHHRec(1).ReadFailCode = ""
          LSET UBSchlumHHRec(1).PrevRead = PrevRead$
          LSET UBSchlumHHRec(1).PrevDate = PrevDate$

          'LSET UBSchlumHHRec(1).HHDisp1 = QPTrim$(UBCustRec(1).SERVADDR)
          'LSET UBSchlumHHRec(1).HHDisp2 = QPTrim$(UBCustRec(1).ADDR1)
          'LSET UBSchlumHHRec(1).HHDisp3 = QPTrim$(UBCustRec(1).ADDR2)
          'LSET UBSchlumHHRec(1).HHDisp4 = QPTrim$(UBCustRec(1).CUSTNAME)

          LSET UBSchlumHHRec(1).HHDisp1 = QPTrim$(UBCustRec(1).SERVADDR)
          LSET UBSchlumHHRec(1).HHDisp2 = QPTrim$(UBCustRec(1).CUSTNAME)
          LSET UBSchlumHHRec(1).HHDisp3 = QPTrim$(UBCustRec(1).HHMSG1)
          LSET UBSchlumHHRec(1).HHDisp4 = MeterID$

          LSET UBSchlumHHRec(1).Notes1 = QPTrim$(UBCustRec(1).HHMSG1)
          LSET UBSchlumHHRec(1).Notes2 = QPTrim$(UBCustRec(1).HHMSG2)
          LSET UBSchlumHHRec(1).Notes3 = QPTrim$(UBCustRec(1).HHMSG3)
          LSET UBSchlumHHRec(1).Notes4 = ""
          LSET UBSchlumHHRec(1).Notes5 = ""
          LSET UBSchlumHHRec(1).Notes6 = ""
          LSET UBSchlumHHRec(1).Notes7 = ""
          LSET UBSchlumHHRec(1).Notes8 = SRouteID$
          LSET UBSchlumHHRec(1).OpCode = ""
          LSET UBSchlumHHRec(1).UBAcctNo = QPTrim$(STR$(PRec&))
          LSET UBSchlumHHRec(1).MtrSlot = QPTrim$(STR$(MtrCnt!))
          LSET UBSchlumHHRec(1).UtilFld = ""
          LSET UBSchlumHHRec(1).CrLf = CrLf$
          PUT UBSchlFile, (LOF(UBSchlFile) / UBSchlumHHRecLen) + 1, UBSchlumHHRec(1)
        END IF
      END IF
    END IF
  NEXT MtrCnt!
  
RETURN
  
  
ImportSchlum:
  
  Ok = MsgBox%("UBHH", "UBSCHL1C")
  
  DisplayHHScrn "UBHUSK3C"
  QPrintRC " Import Meter Readings ", 7, 21, -1
  'overwrite the box header text.
  
  IF Ok = 1 THEN
    Done = True
  ELSE
    QPrintRC "   Import Readings Aborted.   ", 10, 25, -1
    QPrintRC "  Press any key to continue.  ", 14, 25, -1
    CursorOff
    CALL WaitForAction
    Done = True
    GOTO HuskyGetExit:
  END IF
  
  ReportFile$ = "IMPREAD.RPT"
  MaxLines = 55
  
  REDIM UBSchlumHHRec(1) AS SchlumHHType
  UBSchlumHHRecLen = LEN(UBSchlumHHRec(1))
  
  UBSchlFile = FREEFILE
  OPEN ImportFileName$ FOR RANDOM SHARED AS UBSchlFile LEN = UBSchlumHHRecLen
  
  NumTRGetRecs = LOF(UBSchlFile) / UBSchlumHHRecLen
  
  IF NumTRGetRecs = 0 THEN
    CLOSE
    QPrintRC "                                ", 10, 24, 15
    QPrintRC "                                ", 14, 24, 15
    QPrintRC " ERROR: IMPORT FILE NOT FOUND ", 9, 25, -1
    QPrintRC " Make sure that: 'HOST2PC.IMP'", 11, 25, -1
    QPrintRC " is in the CITIPAK directory!!", 13, 25, -1
    QPrintRC "  Press any key to continue.  ", 15, 25, 14
    BEEP
    WaitForAction
    EXIT SUB
  END IF
  
  QPrintRC "  Loading Customer Locations  ", 9, 25, 15
  QPrintRC "                              ", 10, 25, 11
  QPrintRC "            Count:            ", 11, 25, 11
  QPrintRC "                              ", 14, 25, 11
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  IdxRecLen = 4                           'we are using a integer
  IdxFileSize& = FileSize&("UBCUSTBK.IDX")
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it
  
  REDIM SeqBuff(1 TO IdxNumOfRecs)  AS UBLocaReIndexRecType
  
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBFile LEN = UBCustRecLen
  
  FOR IdxCnt = 1 TO IdxNumOfRecs
    GET UBFile, IdxBuff(IdxCnt).RecNum, UBCustRec(1)
    'IF IdxCnt >= 800 THEN STOP
    SeqBuff(IdxCnt).Book = UBCustRec(1).Book
    SeqBuff(IdxCnt).SeqNumb = UBCustRec(1).SeqNumb
    SeqBuff(IdxCnt).RecNum = IdxBuff(IdxCnt).RecNum
    QPrintRC STR$(IdxCnt), 11, 44, 15
  NEXT
  
  ERASE IdxBuff
  
  QPrintRC "  Processing Meter Readings.  ", 9, 25, 15
  QPrintRC "                              ", 10, 25, 11
  QPrintRC "    Updated Count:            ", 11, 25, 11
  QPrintRC "                              ", 14, 25, 11
  
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS RptHandle
  
  GOSUB ImpRptHeader
  
  FOR Cnt& = 1 TO NumTRGetRecs
    GET UBSchlFile, Cnt&, UBSchlumHHRec(1)
    PRec& = 0
    SchlSeq$ = QPTrim$(LEFT$(UBSchlumHHRec(1).Notes8, 8))
    IF LEN(SchlSeq$) = 0 THEN
      GOTO BadSkip
    END IF
    FOR IdxCnt = 1 TO IdxNumOfRecs
      CustSeq$ = (SeqBuff(IdxCnt).Book + SeqBuff(IdxCnt).SeqNumb)
      IF SchlSeq$ = CustSeq$ THEN
        PRec& = SeqBuff(IdxCnt).RecNum
        'IF PRec& = 435 THEN STOP
        EXIT FOR
      END IF
    NEXT
    
    IF PRec& > 0 THEN
      GET UBFile, PRec&, UBCustRec(1)
      GOSUB SchlumExtractRecord
      IF LineCnt > MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB ImpRptHeader
      END IF
    ELSE
      
    END IF
BadSkip:
  NEXT
  
  GOSUB ImpRptTotal
  PRINT #RptHandle, FF$
  
  CLOSE
  Done = True
  
  QPrintRC " Readings Updated Successfully. ", 13, 24, 15
  QPrintRC "  Press any key to continue.  ", 15, 25, 14
  '  BEEP
  WaitForAction
  
  Header$ = "Imported Meter Readings"
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  Done = True
'EXIT SUB

HuskyGetExit:
  RETURN
  
SchlumExtractRecord:
  BadDate = False
  'UpdCnt = UpdCnt + 1
  QPrintRC STR$(Cnt&), 11, 44, 15
  'New Way To Activate
  'MeterSlot = VAL(QPTrim$(UBSchlumHHRec(1).MtrSlot))
  'MeterSlot = VAL(QPTrim$(RIGHT$(UBSchlumHHRec(1).Route, 2)))

  MeterSlot = VAL(QPTrim$(MID$(UBSchlumHHRec(1).Notes8, 9, 2)))

  CurReading# = VAL(QPTrim$(UBSchlumHHRec(1).MtrRead))
  RYear$ = QPTrim$(RIGHT$(UBSchlumHHRec(1).ReadDate, 2))
  
  IF LEN(RYear$) < 2 THEN
    BadDate = True
  END IF
  
  WhatYear = VAL(RYear$)
  IF WhatYear < 65 THEN
    ReadYear$ = "-20" + RYear$
  ELSE
    ReadYear$ = "-19" + RYear$
  END IF
  DateRead$ = LEFT$(UBSchlumHHRec(1).ReadDate, 2) + "-" + MID$(UBSchlumHHRec(1).ReadDate, 3, 2) + ReadYear$
  ReadDate = Date2Num(DateRead$)

  IF CurReading# >= 9999999999# THEN
    CurReading# = 999999999
  END IF

  IF UBCustRec(1).LocMeters(MeterSlot).ReadFlag = "Y" THEN
    UBCustRec(1).LocMeters(MeterSlot).CurRead = CurReading#
    UBCustRec(1).LocMeters(MeterSlot).CurDate = ReadDate
  ELSE
    UBCustRec(1).LocMeters(MeterSlot).PrevRead = UBCustRec(1).LocMeters(MeterSlot).CurRead
    UBCustRec(1).LocMeters(MeterSlot).PastDate = UBCustRec(1).LocMeters(MeterSlot).CurDate
    UBCustRec(1).LocMeters(MeterSlot).ReadFlag = "Y"
    UBCustRec(1).LocMeters(MeterSlot).CurDate = ReadDate
    UBCustRec(1).LocMeters(MeterSlot).CurRead = CurReading#
  END IF
  
  c1$ = QPTrim$(UBCustRec(1).HHMSG1)
  s1$ = QPTrim$(UBSchlumHHRec(1).Notes1)
  IF LEN(s1$) > 0 THEN
    IF s1$ <> c1$ THEN
      UBCustRec(1).NewNotes = True
      UBCustRec(1).HHMSG1 = s1$
    END IF
  END IF
  c1$ = QPTrim$(UBCustRec(1).HHMSG2)
  s1$ = QPTrim$(UBSchlumHHRec(1).Notes2)
  IF LEN(s1$) > 0 THEN
    IF s1$ <> c1$ THEN
      UBCustRec(1).NewNotes = True
      UBCustRec(1).HHMSG2 = s1$
    END IF
  END IF
  c1$ = QPTrim$(UBCustRec(1).HHMSG3)
  s1$ = QPTrim$(UBSchlumHHRec(1).Notes3)
  IF LEN(s1$) > 0 THEN
    IF s1$ <> c1$ THEN
      UBCustRec(1).NewNotes = True
      UBCustRec(1).HHMSG3 = s1$
    END IF
  END IF

  BookSeq$ = LEFT$(UBSchlumHHRec(1).Notes8, 2) + "-" + MID$(UBSchlumHHRec(1).Notes8, 3, 6)
  PRINT #RptHandle, BookSeq$; TAB(10); USING "######"; PRec&; TAB(19);
  PRINT #RptHandle, UBSchlumHHRec(1).HHDisp4; TAB(45); USING "#########"; CurReading#;
  PRINT #RptHandle, TAB(60); DateRead$; TAB(74); UBSchlumHHRec(1).ReadType
  LineCnt = LineCnt + 1
  
  PUT UBFile, PRec&, UBCustRec(1)
  
  RETURN
  
ImpRptHeader:
  Page = Page + 1
  PRINT #RptHandle, "Date: "; DATE$; TAB(27); "Imported Meter Reading Report"; TAB(70); "Page:"; Page
  PRINT #RptHandle, "Location Account  Customer Name               Reading       ReadDate     Status"
  PRINT #RptHandle, STRING$(80, "-")
  LineCnt = 3
  RETURN
  
ImpRptTotal:
  PRINT #RptHandle, STRING$(80, "-")
  PRINT #RptHandle, USING "Imported Count:######"; Cnt& - 1
RETURN
  
END SUB

