DECLARE SUB DisplayHHScrn (ScrnName$)
DECLARE SUB PrintHelp (Help$)
DEFINT A-Z

  SHARED Choice$()
  DIM Route(10)
  REDIM UBLogiconRec(1) AS UBLogiconRecType
  UBLogiconRecLen = LEN(UBLogiconRec(1))
  UBLogiconFile = FREEFILE
  OPEN "UBLogcon.DAT" FOR RANDOM SHARED AS UBLogiconFile LEN = UBLogiconRecLen
  NumLogiconRecs = LOF(UBLogiconFile) / UBLogiconRecLen
  IF NumLogiconRecs = 1 THEN
    GET UBLogiconFile, 1, UBLogiconRec(1)
    PathWay$ = UBLogiconRec(1).PathWay
  END IF

  LibName$ = "UBSETUP"
  ScrnName$ = "UBLOGHH"
  Help$ = "Logicon Hand Held Processing"
  CursorOff
  'LOCATE 1, 1, 0

  Help$ = "Choose HandHeld Operation"
  DisplayHHScrn ScrnName$
  PrintHelp Help$

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

  REDIM Choice$(3, 0)

  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Send Info to Logicon"
  Choice$(2, 0) = "Get Info From Logicon"

  IF NumLogiconRecs = 1 THEN
    Form$(2, 0) = PathWay$
    'Fld(2).Protected = True
  END IF

  Action = 1

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode

    CASE F10Key
      IF NOT (LEFT$(Form$(1, 0), 1) = "G" OR LEFT$(Form$(1, 0), 1) = "S") THEN
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

      ELSE
        PathWay$ = Form$(2, 0)
        UBLogiconRec(1).PathWay = PathWay$
        PUT UBLogiconFile, 1, UBLogiconRec(1)
        CLOSE UBLogiconFile
        IF LEFT$(Form$(1, 0), 1) = "S" THEN GOSUB SendLogicon
        IF LEFT$(Form$(1, 0), 1) = "G" THEN GOSUB GetLogicon
        Done = True
      END IF
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done

  EXIT DEF

SendLogicon:
  LibName$ = "UBSETUP"
  ScrnName$ = "UBLOG1S"
  Help$ = "Logicon Hand Held Processing"
  CursorOff

  Help$ = "DownLoad HandHeld Operation"
  DisplayHHScrn ScrnName$
  PrintHelp Help$

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

  Action = 1
  NumberofRoutes = 0


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
      ' First Check For Valid Interrogator Number and Routes
      IF VAL(Form$(1, 0)) < 1 THEN
        Frm(1).FldNo = 1: Action = 1: Done = False

      ELSE

        IF NumberofRoutes > 0 THEN
          RouteID$ = Form$(1, 0)
          COLOR 11
          LOCATE 12, 18
          PRINT "LOADING THESE ROUTES !!!                      "
          GOSUB FILENAM1LC
          'Set Route Number Values
          FOR Cnt! = 1 TO NumberofRoutes
            Route(Cnt!) = VAL(Form$(Cnt! + 2, 0))
          NEXT Cnt!

          'Open Customer Data File
          REDIM UBCustRec(1) AS NewUBCustRecType
          UBCustRecLen = LEN(UBCustRec(1))
          UBFile = FREEFILE
          OPEN "UBCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBFile LEN = UBCustRecLen

          'Open Logicon Date File
          REDIM UBLogRdRec(1) AS UBLogiconReadRecType
          UBLogRdRecLen = LEN(UBLogRdRec(1))
          UBLogRdFile = FREEFILE
          OPEN FileName$ FOR RANDOM SHARED AS UBLogRdFile LEN = UBLogRdRecLen
          CLOSE UBLogRdFile: KILL FileName$
          UBLogRdFile = FREEFILE
          OPEN FileName$ FOR RANDOM SHARED AS UBLogRdFile LEN = UBLogRdRecLen
          NumLogRdRecs = LOF(UBLogRdFile) / UBLogRdRecLen

          'Write First Record With Route Information
          UBLogRdRec(1).RecType = "H"
          UBLogRdRec(1).RouteNo = RouteID$
          UBLogRdRec(1).AcctNo = ""
          UBLogRdRec(1).RecName = ""
          UBLogRdRec(1).ServAddress = ""
          UBLogRdRec(1).ReadDate = ""
          UBLogRdRec(1).ReadTime = ""
          UBLogRdRec(1).Consumption = ""
          UBLogRdRec(1).PrevRead = ""
          UBLogRdRec(1).CurRead = ""
          UBLogRdRec(1).LowRead = ""
          UBLogRdRec(1).HighRead = ""
          UBLogRdRec(1).MtrNumb = ""
          UBLogRdRec(1).CountChg = ""
          UBLogRdRec(1).ForceFlag = ""
          UBLogRdRec(1).ReportCode = ""
          UBLogRdRec(1).Remark = ""
          UBLogRdRec(1).Label = ""
          UBLogRdRec(1).PrintFlag = ""
          UBLogRdRec(1).MessageOut = ""
          UBLogRdRec(1).Book = ""
          UBLogRdRec(1).Future = ""
          UBLogRdRec(1).Recend = "X"
          UBLogRdRec(1).CrLf = CHR$(13) + CHR$(10)
          PUT UBLogRdFile, (LOF(UBLogRdFile) / UBLogRdRecLen) + 1, UBLogRdRec(1)

          'Open the Location Order for Reading
          IdxRecLen = 4         'we are using a integer
          IdxFileSize& = FileSize&("UBCUSTBK.IDX")
          IdxNumOfRecs = IdxFileSize& \ IdxRecLen
          REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
          FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs            'load it

          Cnt! = 1

          DO

            Prec& = IdxBuff(Cnt!).RecNum
            IF NOT (Prec&) = 0 THEN
              GET UBFile, Prec&, UBCustRec(1)
              FOR BookCnt! = 1 TO NumberofRoutes
                IF VAL(UBCustRec(1).Book) = Route(BookCnt!) AND (UBCustRec(1).Status <> "F") THEN
                  GOSUB WriteRecordLC
                END IF
              NEXT BookCnt!
            END IF
            Cnt! = Cnt! + 1

          LOOP UNTIL Cnt! > IdxNumOfRecs

          CLOSE
          Done = True
          FOR lc = 8 TO 10
            LOCATE lc, 16
            PRINT STRING$(40, 32)
          NEXT lc
          LOCATE 8, 18
          COLOR 11
          PRINT "Route Number "; RouteID$
          LOCATE 9, 18
          COLOR 11
          PRINT "Information Sent to Logicon System Successfully"
          LOCATE 12, 18: PRINT "THESE ROUTES LOADED!!!                      "
          PRINT CHR$(7);
          WaitForAction
        END IF
        Done = True
      END IF

    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  RETURN


WriteRecordLC:
  'May Have Up to 10 Meters to Read
  MtrCnt! = 1

  Account$ = SPACE$(6)
  LSET Account$ = QPTrim$(STR$(Prec&))

  'Account$ = LEFT$(Account$, 6) + "-" + RIGHT$(STR$(MtrCnt!), 1)

  WHILE MtrCnt! < 8

    IF (ASC(UBCustRec(1).LocMeters(MtrCnt!).MtrType) > 32) THEN
      IF UBCustRec(1).LocMeters(MtrCnt!).MtrType = "C" OR UBCustRec(1).LocMeters(MtrCnt!).MtrType = "W" OR UBCustRec(1).LocMeters(MtrCnt!).MtrType = "T" OR UBCustRec(1).LocMeters(MtrCnt!).MtrType = "S" THEN
        MID$(Account$, 6, 1) = QPTrim$(STR$(MtrCnt!))
        ' Determine High and Low Reading
        Average# = UBCustRec(1).LocMeters(MtrCnt!).AvgUse
        ILowRead$ = RIGHT$(STR$((UBCustRec(1).LocMeters(MtrCnt!).CurRead)), 8)
        ILowRead# = VAL(ILowRead$)
        HighRead# = FIX(Average# * (HighVar / 100)) + UBCustRec(1).LocMeters(MtrCnt!).CurRead
        IF FIX(HighRead#) = ILowRead# THEN HighRead# = HighRead# + 12000
        IHighRead$ = STR$(HighRead#)
        IHighRead$ = RIGHT$(IHighRead$, 8)
        MeterID$ = LTRIM$(UBCustRec(1).LocMeters(MtrCnt!).MTRNUM)

        MeterID$ = RTRIM$(MeterID$)
        IF VAL(MeterID$) = 0 THEN
          MeterID$ = UBCustRec(1).Book + UBCustRec(1).SEQNUMB
        END IF
        IF LEN(MeterID$) < 8 THEN
          MeterID$ = STRING$(8 - LEN(MeterID$), "0") + MeterID$
        END IF
        MeterID$ = LEFT$(MeterID$, 8)

        UBLogRdRec(1).RecType = "A"
        UBLogRdRec(1).RouteNo = RouteID$
        UBLogRdRec(1).AcctNo = Account$
        UBLogRdRec(1).RecName = UBCustRec(1).CUSTNAME
        UBLogRdRec(1).ServAddress = UBCustRec(1).SERVADDR
        UBLogRdRec(1).ReadDate = "      "
        UBLogRdRec(1).ReadTime = "      "
        UBLogRdRec(1).Consumption = "        "
        UBLogRdRec(1).PrevRead = ILowRead$
        UBLogRdRec(1).CurRead = "XXXXXXXX"
        UBLogRdRec(1).LowRead = ILowRead$
        UBLogRdRec(1).HighRead = IHighRead$
        UBLogRdRec(1).MtrNumb = MeterID$
        UBLogRdRec(1).CountChg = "0"
        UBLogRdRec(1).ForceFlag = " "
        UBLogRdRec(1).ReportCode = "--"
        UBLogRdRec(1).Remark = ""
        UBLogRdRec(1).Label = ""
        UBLogRdRec(1).PrintFlag = ""
        UBLogRdRec(1).MessageOut = "OutPut to Meter Reader"
        UBLogRdRec(1).Book = UBCustRec(1).Book
        UBLogRdRec(1).Future = ""
        UBLogRdRec(1).Recend = "X"
        UBLogRdRec(1).CrLf = CHR$(13) + CHR$(10)
        PUT UBLogRdFile, (LOF(UBLogRdFile) / UBLogRdRecLen) + 1, UBLogRdRec(1)
      END IF
    END IF

SkipEmLC:
    MtrCnt! = MtrCnt! + 1
  WEND
  RETURN

FILENAM1LC:
  PathWay$ = QPTrim$(PathWay$)
  FileName$ = PathWay$ + "WBLOGNO" + LTRIM$(RouteID$) + ".DAT"
  RETURN

GetLogicon:
  LibName$ = "UBSETUP"
  ScrnName$ = "UBLOG1G"
  Help$ = "Logicon Hand Held Processing"
  CursorOff

  Help$ = "Upload HandHeld Operation"
  DisplayHHScrn ScrnName$
  PrintHelp Help$

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

  Action = 1
  NumberofRoutes = 0


  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode

    CASE F10Key
      IF VAL(Form$(1, 0)) < 1 OR VAL(Form$(2, 0)) < 1 THEN
        LOCATE 12, 20: COLOR 12
        PRINT "            [  E R R O R  ]"
        LOCATE 13, 20: COLOR 11
        PRINT "     Invalid Route ID Number "
        LOCATE 14, 20: COLOR 11
        PRINT "      or Invalid Reading Date "
        BEEP
        WaitForAction
        CLOSE
        GOTO GetLogicon
        Frm(1).FldNo = 1
        Action = 1
        GOTO GetLogicon
      ELSE
        RouteID$ = Form$(1, 0)
        RRDate = Date2Num(Form$(2, 0))
        GOSUB FileNamLC         ' Get Logicon File
        'Open Customer Data File
        REDIM UBCustRec(1) AS NewUBCustRecType
        UBCustRecLen = LEN(UBCustRec(1))
        UBFile = FREEFILE
        OPEN "UBCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBFile LEN = UBCustRecLen

        REDIM UBLogGetRdRec(1) AS UBLogiconGetReadRecType
        UBLogGetRdRecLen = LEN(UBLogGetRdRec(1))

        UBLogGetRdFile = FREEFILE
        OPEN FileName$ FOR RANDOM ACCESS READ WRITE SHARED AS UBLogGetRdFile LEN = UBLogGetRdRecLen
        NumLogGetRdRecs = LOF(UBLogGetRdFile) / UBLogGetRdRecLen

        IF NumLogGetRdRecs = 0 THEN
          LOCATE 12, 20: COLOR 12
          PRINT "            [  E R R O R  ]"
          LOCATE 13, 20: COLOR 11
          PRINT "     Nothing Found From Logicon "
          LOCATE 14, 20: COLOR 11
          PRINT "   Please Reenter Route ID Number "
          BEEP
          WaitForAction
          CLOSE
          GOTO GetLogicon
        END IF

        Cnt! = 1                ' Initialize File Counter to 1
        DO
          GET UBLogGetRdFile, Cnt!, UBLogGetRdRec(1)

          TAcct$ = LEFT$(UBLogGetRdRec(1).AcctNo, 5)
          Prec& = VAL(TAcct$)

          IF LEFT$(UBLogGetRdRec(1).CurRead, 1) <> "X" THEN
            IF Prec& > 0 THEN
              GET UBFile, Prec&, UBCustRec(1)
              GOSUB ExtractRecordLC
            END IF
          END IF
          Cnt! = Cnt! + 1

        LOOP UNTIL Cnt! > NumLogGetRdRecs

        CLOSE
        Done = True
        FOR lc = 8 TO 10
          LOCATE lc, 16
          PRINT STRING$(40, 32)
        NEXT lc
        LOCATE 8, 18
        COLOR 11
        PRINT "Logicon Unit Number "; RouteID$
        LOCATE 13, 18
        COLOR 11
        PRINT "Reading Updated From Logicon System Successfully"
        PRINT CHR$(7);
        WaitForAction
      END IF
      Done = True

    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  RETURN


ExtractRecordLC:
  UpdCnt! = UpdCnt! + 1
  LOCATE 13, 20: COLOR 11
  PRINT " Updated Count:  "; : COLOR 15: PRINT UpdCnt!
  COLOR 11


  'New Way To Activate

  MeterRecord = VAL(RIGHT$(UBLogGetRdRec(1).AcctNo, 1))
  'Only 1 meter per account for logicon 'no; don't think so Tim.

  ' Update Meter W/Reading
  CurReading# = VAL(UBLogGetRdRec(1).CurRead)
  IF RRDate > 0 THEN
    MeterReadDate$ = Num2Date$(RRDate)
  ELSE
    MeterReadDate$ = MID$(UBLogGetRdRec(1).ReadDate, 3, 2) + "/" + MID$(UBLogGetRdRec(1).ReadDate, 5, 2) + "/" + RIGHT$(DATE$, 4)
  END IF

  IF UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y" THEN
    UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
    UBCustRec(1).LocMeters(MeterRecord).CurDate = Date2Num(MeterReadDate$)
  ELSE
    UBCustRec(1).LocMeters(MeterRecord).PrevRead = UBCustRec(1).LocMeters(MeterRecord).CurRead
    UBCustRec(1).LocMeters(MeterRecord).PastDate = UBCustRec(1).LocMeters(MeterRecord).CurDate
    UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y"
    UBCustRec(1).LocMeters(MeterRecord).CurDate = Date2Num(MeterReadDate$)
    UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
  END IF
  PUT UBFile, Prec&, UBCustRec(1)
  RETURN

FileNamLC:
  PathWay$ = QPTrim$(PathWay$)
  FileName$ = PathWay$ + "WBLOGNO" + LTRIM$(RouteID$) + ".DAT"
RETURN


