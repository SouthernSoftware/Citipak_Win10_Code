DECLARE SUB DisplayHuskyScrn (ScrnName$)
DEFINT A-Z

DEFSNG A-Z
DEFINT A-Z
SUB ProcessTelxon

      q$ = CHR$(34)
      cb$ = SPACE$(45)
      SHARED Choice$()
      REDIM Route(100)

      LibName$ = "UBSETUP"
      ScrnName$ = "UBTELXHH"
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

      REDIM Choice$(5, 1)
      Choice$(0, 0) = "1"
      Choice$(1, 0) = "Create Readings file"
      Choice$(2, 0) = "Load Readings file"

      Choice$(0, 1) = "2"
      Choice$(1, 1) = "Water/Sewer"
      Choice$(2, 1) = "Electric"
      Choice$(3, 1) = "Gas Meters"
      Choice$(4, 1) = "ALL METERS"

      Action = 1
      NumberofRoutes = 0
      DisplayHuskyScrn ScrnName$
      DO

        EditForm Form$(), Fld(), Frm(1), Cnf, Action
        IF Frm(1).PrevFld = 1 AND Frm(1).FldNo > 1 THEN
          SELECT CASE QPTrim$(LEFT$(Form$(1, 0), 1))
          CASE "C"
            QPrintRC "Selected Meter Type:", 10, 20, 11
            QPrintRC "Enter Route to Process:", 12, 17, 11
          CASE "L"
            QPrintRC "                    ", 10, 20, 11
            QPrintRC "                       ", 12, 17, 11
            Fld(2).Protected = True
            Fld(3).Protected = True
            Action = 2
          CASE ELSE
            Fld(2).Protected = False
            Fld(3).Protected = False
            Frm(1).FldNo = 1
            Action = 2
          END SELECT
        END IF

        IF Frm(1).PrevFld = 14 AND Frm(1).FldNo > 3 THEN
          IF VAL(Form$(3, 0)) <> 0 THEN
            NumberofRoutes = NumberofRoutes + 1
            IF NumberofRoutes <= 10 THEN
              Form$(NumberofRoutes + 3, 0) = Form$(3, 0)
              Frm(1).FldNo = 3
              Form$(3, 0) = ""
              Action = 1
            END IF
          END IF
        END IF

        SELECT CASE Frm(1).KeyCode

        CASE F10Key

           WhatTypes$ = QPTrim$(LEFT$(Form$(2, 0), 1))
           WhatAction$ = QPTrim$(LEFT$(Form$(1, 0), 1))
           IF LEN(WhatTypes$) = 0 AND LEN(WhatAction$) = 0 THEN
             Done = False
           ELSEIF LEN(WhatAction$) > 0 THEN

             SELECT CASE WhatAction$
             CASE "C"
               ReadFileOK = False
               GOSUB CreateReadFile
               IF ReadFileOK THEN
                 Done = True
               END IF
             CASE "L"
               ReadFileOK = False
               GOSUB LoadReadFile
               IF ReadFileOK THEN
                 Done = True
               END IF

             END SELECT
           END IF
        CASE ESC
          Done = True
        CASE ELSE
          Done = False
        END SELECT
      LOOP UNTIL Done

EXIT SUB
      'RETURN

CreateReadFile:
  WhatTypes$ = LEFT$(Form$(2, 0), 1)
  ' Next Check Routes
  IF NumberofRoutes > 0 THEN
    DisplayHuskyScrn "UBHUSK3C"
    QPrintRC "Create Telxon Read File ]", 7, 21, -1
    QPrintRC "      Scanning Account:        ", 11, 21, -1

    ' Set Route Number Values
    FOR Cnt! = 1 TO NumberofRoutes
      Route(Cnt!) = VAL(Form$(Cnt! + 3, 0))
    NEXT Cnt!

    'Open Customer Data File
    REDIM UBCustRec(1) AS NewUBCustRecType
    UBCustRecLen = LEN(UBCustRec(1))
    UBFile = FREEFILE
    OPEN "UBCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBFile LEN = UBCustRecLen

    'Open and Initialize the PC3000 Read Information File
    REDIM UBPC3000RDRec(1) AS UBPC3000ReadRecType
    UBPC3000RdRecLen = LEN(UBPC3000RDRec(1))
    UBPC3000RdFile = FREEFILE
    OPEN "UBCUSTTR.DAT" FOR RANDOM SHARED AS UBPC3000RdFile LEN = UBPC3000RdRecLen
    CLOSE UBPC3000RdFile
    KILL "UBCUSTTR.DAT"
    UBPC3000RdFile = FREEFILE
    OPEN "UBCUSTTR.DAT" FOR RANDOM SHARED AS UBPC3000RdFile LEN = UBPC3000RdRecLen
    NumPC3000RdRecs = LOF(UBPC3000RdFile) / UBPC3000RdRecLen

    REDIM UBSetUpRec(1) AS UBSetupRecType
    LoadUBSetUpFile UBSetUpRec(), UBSetUpLen


    'Open the Correct Order for Reading
    IF UBSetUpRec(1).UseSeq = "Y" THEN
      IdxRecLen = 4     'we are using a integer
      MakeSequenceIndex "Sequence Number"
      IdxNumOfRecs = FileSize&("UBTEMP.IDX") \ 4
      REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
      FGetAH "UBTEMP.IDX", IdxBuff(1), 4, IdxNumOfRecs
    ELSE
      IdxRecLen = 4     'we are using a integer
      IdxFileSize& = FileSize&("UBCUSTBK.IDX")
      IdxNumOfRecs = IdxFileSize& \ IdxRecLen
      REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
      FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs        'load it
    END IF
    Cnt! = 1

    DO

      Prec& = IdxBuff(Cnt!).RecNum
      IF NOT (Prec&) = 0 THEN
        GET UBFile, Prec&, UBCustRec(1)
        DisplayAcct$ = UBCustRec(1).BOOK + "-" + UBCustRec(1).SeqNumb
        QPrintRC DisplayAcct$, 11, 45, -1

        FOR BookCnt! = 1 TO NumberofRoutes
          IF VAL(UBCustRec(1).BOOK) = Route(BookCnt!) AND (UBCustRec(1).Status <> "F") THEN
            GOSUB TelxonPutRec
          END IF
        NEXT BookCnt!
      END IF
      Cnt! = Cnt! + 1

    LOOP UNTIL Cnt! > IdxNumOfRecs
    CLOSE

    DisplayHuskyScrn "UBTELXCD"
    'Ok = MsgBox%("UBSETUP", "UBHUSK1C")

    WaitForAction
    ReadFileOK = True
  ELSE
    Frm(1).FldNo = 3
    Action = 2
  END IF


RETURN


LoadReadFile:

      DisplayHuskyScrn "UBHUSK3C"
      QPrintRC "Telxon Load Reading File ]", 7, 21, -1

      REDIM UBCustRec(1) AS NewUBCustRecType
      UBCustRecLen = LEN(UBCustRec(1))
      UBFile = FREEFILE
      OPEN "UBCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBFile LEN = UBCustRecLen

      'Open and Initialize the PC3000 Read Information File
      REDIM UBPC3000GetRDRec(1) AS UBPC3000ReadRecType
      UBPC3000GetRdRecLen = LEN(UBPC3000GetRDRec(1))
      UBPC3000GetRdFile = FREEFILE
      FileNam$ = "UBCUSTTR.DAT"
      OPEN FileNam$ FOR RANDOM ACCESS READ WRITE SHARED AS UBPC3000GetRdFile LEN = UBPC3000GetRdRecLen
      NumPC3000GetRdRecs = LOF(UBPC3000GetRdFile) / UBPC3000GetRdRecLen

      IF NumPC3000GetRdRecs = 0 THEN
        CLOSE
        QPrintRC "                                ", 10, 24, 15
        QPrintRC "                                ", 14, 24, 15
        QPrintRC "   ERROR: NO READINGS FOUND   ", 9, 25, -1
        QPrintRC " Check the handheld connection", 11, 25, -1
        QPrintRC " and try the transfer again!!!", 13, 25, -1
        QPrintRC "   Press any key to continue. ", 15, 25, 14
        BEEP: BEEP: BEEP
        WaitForAction
        Done = True
        RETURN
      END IF

      QPrintRC "  Processing Meter Readings.  ", 9, 25, 15
      QPrintRC "                              ", 10, 25, 11
      QPrintRC "    Updated Count:            ", 11, 25, 11
      QPrintRC "                              ", 14, 25, 11

      Cnt! = 1  ' Initialize File Counter to 1
      DO
        GET UBPC3000GetRdFile, Cnt!, UBPC3000GetRDRec(1)
        'IF UBPC3000GetRDRec(1).ReadFlag = "Y" THEN
          DidOne = True
          ' New Way Undelete When Our System Sends and Retrieves
          Prec& = VAL(QPTrim$(UBPC3000GetRDRec(1).Account))
          IF NOT (Prec&) = 0 THEN
            GET UBFile, Prec&, UBCustRec(1)
            GOSUB TelxonExtRec
          END IF
        'END IF
        Cnt! = Cnt! + 1
      LOOP UNTIL Cnt! > NumPC3000GetRdRecs
      CLOSE

      Done = True
      IF DidOne THEN
        QPrintRC " Readings Updated Successfully. ", 13, 24, 15
        QPrintRC "  Press any key to continue.  ", 15, 25, 14
        BEEP
      ELSE
        QPrintRC "      NO READINGS UPDATED       ", 13, 24, 15
        QPrintRC "  Press any key to continue.  ", 15, 25, 14
      END IF
      WaitForAction
      Done = True

TelxonGetExit:
      RETURN

TelxonExtRec:
      UpdCnt! = UpdCnt! + 1
      QPrintRC STR$(UpdCnt!), 11, 44, 15
      'New Way To Activate
      MeterRecord = VAL(RIGHT$((QPTrim$(UBPC3000GetRDRec(1).Account)), 1))
      ' Check Meter Updated Flag
      ' Update Meter W/Reading
      CurReading# = UBPC3000GetRDRec(1).CurRead

      'IF UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y" THEN
      'ELSE
        UBCustRec(1).LocMeters(MeterRecord).PrevRead = UBPC3000GetRDRec(1).PastRead
        UBCustRec(1).LocMeters(MeterRecord).PastDate = UBCustRec(1).LocMeters(MeterRecord).CurDate
        UBCustRec(1).LocMeters(MeterRecord).ReadFlag = UBPC3000GetRDRec(1).ReadFlag
        UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
        UBCustRec(1).LocMeters(MeterRecord).CurDate = UBPC3000GetRDRec(1).ReadDate
      'Modifed 04-28-97
      IF UBPC3000GetRDRec(1).NoteStatus = "P" THEN
        UBCustRec(1).HHMSG1 = UBPC3000GetRDRec(1).Note1
        UBCustRec(1).HHMSG2 = UBPC3000GetRDRec(1).Note2
        UBCustRec(1).HHMSG3 = UBPC3000GetRDRec(1).Note3
      END IF
      '   ^^^
      PUT UBFile, Prec&, UBCustRec(1)
      RETURN


TelxonPutRec:
      'May Have Up to 7 Meters to Read
      FOR MtrCnt! = 1 TO 7
        MeterOK = False
        Account$ = STR$(Prec&)
        Account$ = LEFT$(Account$, 6) + "-" + RIGHT$(STR$(MtrCnt!), 1)

        IF ASC(UBCustRec(1).LocMeters(MtrCnt!).MTRType) > 32 THEN
          MTRType$ = UBCustRec(1).LocMeters(MtrCnt!).MTRType
          IF MTRType$ = "W" OR MTRType$ = "S" OR MTRType$ = "C" OR MTRType$ = "E" OR MTRType$ = "D" OR MTRType$ = "G" THEN
            SELECT CASE WhatTypes$
            CASE "W"            'water/sewer
              IF MTRType$ = "W" OR MTRType$ = "S" OR MTRType$ = "C" THEN
                MeterOK = True
              END IF
            CASE "E"            'electric & demand elec.
              IF MTRType$ = "E" OR MTRType$ = "D" THEN
                MeterOK = True
              END IF
            CASE "G"            'gas
              IF MTRType$ = "G" THEN
                MeterOK = True
              END IF
            CASE "A", " "       'all meters
              MeterOK = True
            END SELECT

            IF MeterOK = True THEN

              ' Determine High and Low Reading
              Average# = UBCustRec(1).LocMeters(MtrCnt!).AvgUse

              ILowRead$ = RIGHT$(STR$((UBCustRec(1).LocMeters(MtrCnt!).CurRead)), 8)
              LowRead# = FIX(ILowRead#)
              HighRead# = FIX(Average# * (HighVar / 100)) + UBCustRec(1).LocMeters(MtrCnt!).CurRead
              
              MeterID$ = LTRIM$(UBCustRec(1).LocMeters(MtrCnt!).MtrNum)
              MeterID$ = RTRIM$(MeterID$)

              IF VAL(MeterID$) = 0 THEN
                MeterID$ = UBCustRec(1).BOOK + UBCustRec(1).SeqNumb
              END IF
              IF LEN(MeterID$) < 8 THEN
                MeterID$ = STRING$(8 - LEN(MeterID$), "0") + MeterID$
              END IF
              MeterID$ = LEFT$(MeterID$, 8)

              'Set Record Fields and Put On Disk
              UBPC3000RDRec(1).CustName = UBCustRec(1).CustName
              UBPC3000RDRec(1).ServAddress = UBCustRec(1).ServAddr
              UBPC3000RDRec(1).MeterID = MeterID$
              UBPC3000RDRec(1).LowRead = LowRead#
              UBPC3000RDRec(1).HighRead = HighRead#
              UBPC3000RDRec(1).Account = Account$
              UBPC3000RDRec(1).MeterType$ = UBCustRec(1).LocMeters(MtrCnt!).MTRType
              UBPC3000RDRec(1).BOOK = VAL(UBCustRec(1).BOOK)
              UBPC3000RDRec(1).CurRead = UBCustRec(1).LocMeters(MtrCnt!).CurRead
              UBPC3000RDRec(1).PastRead = 0
              UBPC3000RDRec(1).ReadFlag = "N"
              'Modifed 04-28-97
              UBPC3000RDRec(1).Note1 = UBCustRec(1).HHMSG1
              UBPC3000RDRec(1).Note2 = UBCustRec(1).HHMSG2
              UBPC3000RDRec(1).Note3 = UBCustRec(1).HHMSG3
              '         ^^^
              UBPC3000RDRec(1).NoteStatus = ""
              PUT UBPC3000RdFile, (LOF(UBPC3000RdFile) / UBPC3000RdRecLen) + 1, UBPC3000RDRec(1)
            END IF
          END IF
        END IF
      NEXT MtrCnt!
      RETURN

END SUB

