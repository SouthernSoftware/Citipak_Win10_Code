DECLARE SUB DisplayH2Scrn (ScrnName$)
DECLARE SUB DisplayHHScrn (ScrnName$)
DEFINT A-Z

DEFSNG A-Z
DEFINT A-Z
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
  Choice$(1, 0) = "Export reading file for Intermec"
  Choice$(2, 0) = "Import reading file from Intermec"

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
            DisplayAcct$ = UBCustRec(1).BOOK + "-" + UBCustRec(1).SEQNUMB
            QPrintRC DisplayAcct$, 11, 45, -1
            FOR BookCnt! = 1 TO NumberofRoutes
              IF VAL(UBCustRec(1).BOOK) = Route(BookCnt!) AND (UBCustRec(1).Status <> "F") THEN
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
  UBCustRec(1).NewNotes = True
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
            MeterID$ = UBCustRec(1).BOOK + UBCustRec(1).SEQNUMB
          END IF
          IF LEN(MeterID$) < 8 THEN
            MeterID$ = STRING$(8 - LEN(MeterID$), "0") + MeterID$
          END IF
          MeterID$ = LEFT$(MeterID$, 8)

          'Set Record Fields and Put On Disk

          UBInterRDRec(1).CUSTNAME = UBCustRec(1).CUSTNAME
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
          UBInterRDRec(1).BOOK = UBCustRec(1).BOOK
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

