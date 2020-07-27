DECLARE SUB DisplayHHScrn (ScrnName$)

  '    'Open Customer Data File
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
    FOR zz = 8 TO 16
      QPrintRC SPACE$(49), zz, 16, 15
    NEXT

    DisplayHHScrn "UBTELX3C"
    QPrintRC "ERROR: NO READINGS FOUND", 9, 27, -1
    QPrintRC " Check the handheld connection", 10, 24, -1
    QPrintRC " and try the transfer again!!!", 11, 24, -1
    QPrintRC "   Press any key to continue. ", 13, 24, 14
    BEEP
    WaitForAction
    GOTO GetTelxon
  END IF
  DisplayHHScrn "UBTELX3C"

  QPrintRC "  Processing Meter Readings.  ", 9, 25, 15
  QPrintRC "    Updated Count:            ", 11, 25, 11

  Cnt! = 1      ' Initialize File Counter to 1
  DO
    GET UBPC3000GetRdFile, Cnt!, UBPC3000GetRDRec(1)
    ' New Way Undelete When Our System Sends and Retrieves
    Prec& = VAL(QPTrim$(UBPC3000GetRDRec(1).Account))

    IF NOT (Prec&) = 0 THEN
      GET UBFile, Prec&, UBCustRec(1)
      GOSUB TelxonExtRec
    END IF
    Cnt! = Cnt! + 1

  LOOP UNTIL Cnt! > NumPC3000GetRdRecs

  CLOSE
  Done = True

TelxonExtRec:
  UpdCnt! = UpdCnt! + 1
  QPrintRC STR$(UpdCnt!), 11, 44, 15
  'New Way To Activate
  MeterRecord = VAL(RIGHT$((QPTrim$(UBPC3000GetRDRec(1).Account)), 1))
  ' Check Meter Updated Flag
  ' Update Meter W/Reading
  CurReading# = UBPC3000GetRDRec(1).CurRead

  IF UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y" THEN
    UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
    UBCustRec(1).LocMeters(MeterRecord).CurDate = UBPC3000GetRDRec(1).ReadDate
  ELSE
    UBCustRec(1).LocMeters(MeterRecord).PrevRead = UBCustRec(1).LocMeters(MeterRecord).CurRead
    UBCustRec(1).LocMeters(MeterRecord).PastDate = UBCustRec(1).LocMeters(MeterRecord).CurDate
    UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y"
    UBCustRec(1).LocMeters(MeterRecord).CurDate = UBPC3000GetRDRec(1).ReadDate
    UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
  END IF
  'Modifed 04-28-97
  IF UBPC3000GetRDRec(1).NoteStatus = "P" THEN
    UBCustRec(1).NewNotes = True
    UBCustRec(1).HHMSG1 = UBPC3000GetRDRec(1).Note1
    UBCustRec(1).HHMSG2 = UBPC3000GetRDRec(1).Note2
    UBCustRec(1).HHMSG3 = UBPC3000GetRDRec(1).Note3
  END IF
  '   ^^^
  PUT UBFile, Prec&, UBCustRec(1)
  RETURN

