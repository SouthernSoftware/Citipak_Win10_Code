DEFINT A-Z    'New hand held program for Version 8.2w
DECLARE SUB BackLightToggle ()
DECLARE SUB SetFont ()
DECLARE SUB WInput (Edit$, GoodKey$, Row%, Col%, ExitCode%)
DECLARE SUB CLocate (BYVAL Row%, BYVAL Col%)
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE SUB QPSound (Frequency%, Duration%)
DECLARE FUNCTION GetPower ()
DECLARE FUNCTION BiosKey% ()
  
  '$INCLUDE: 'ubsetup.bi'
  
  CONST False = 0, True = NOT False
  
  CONST F1Key = 59
  CONST F2Key = 60
  CONST F3Key = 61
  CONST F4Key = 62
  CONST EscKey = -27
  
  DIM SHARED GoodKey2$, GoodKey3$
  DIM SHARED NoBackLight
  
  GoodKey2$ = " ABCDEFGHIJKLMNOPQRSTUVWXYZ"
  GoodKey3$ = "1234567890"
  
  StatBar$ = " Press ESC for Help."             'CHR$(25) + "=Next " + CHR$(24) + "=Prev"
  
  CALL SetFont
  Ast$ = STRING$(20, "*")
  Blank$ = STRING$(20, " ")
  Fmt2$ = "        "
  REDIM ReadRec(1) AS UBPC3000ReadRecType
  ReadRecLen = LEN(ReadRec(1))
  
Top:
  
  CLS
  PRINT "ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿"        '1
  PRINT "³  Foxfire Meter Read v1.00  ³"        '3
  PRINT "³ (C)97-07 Southern Software ³"        '4
  PRINT "³                            ³"        '5
  PRINT "³     Press ESC to Exit.     ³"        '6
  PRINT "³                            ³"        '5
  PRINT "³ Any other key to continue. ³"        '7
  PRINT "ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ";       '8
  
GetRoute:
  'Route$ = "   "
  ExitCode = True
  WInput Route$, GoodKey3$, 6, 15, ExitCode
  
  IF ExitCode = -27 THEN
    CLS
    CLocate 2, 1
    PRINT "Program Exited."
    PRINT
    END
  END IF
  
  'IF LEN(QPTrim$(Route$)) = False THEN
  '  GOTO GetRoute
  'END IF
  
  CurBook = VAL(Route$)
  CurBook = 1

  CLS
  PRINT
  PRINT "Searching for Route: "; CurBook
  CLocate 4, 4
  PRINT "Scanning [ ]"
  GOSUB Scanning
  GOSUB OpenFile
  LastRecord = LOF(1) / ReadRecLen
  GOSUB CloseFile
  CurRecord = 1
  ScanCode = 1
  GOTO 1060
  
1059
  IF ScanCode = 1 OR BackFlag = True THEN
    CLS
    CLocate 4, 4
    PRINT "Scanning [ ]"
  END IF
  
1060
  IF CurRecord < 1 THEN
    CurRecord = 1
    BackFlag = False
  END IF
  
  IF CurRecord > LastRecord THEN
    BackFlag = False
    GOTO 1080
  ELSE
    X = CurRecord
    GOSUB OpenFile
    GET 1, X, ReadRec(1)
    GOSUB Scanning
    GOSUB CloseFile
    IF ScanCode = 1 THEN
      IF TrReadCode$ = "Y" THEN
        CurRecord = CurRecord + 1
        GOTO 1059
      END IF
    END IF
  END IF
  'IF ReadRec(1).Book <> CurBook THEN 1072
  ScanCode = False
  BackFlag = False
  GOSUB 1100
  
  IF (CF = EscKey) AND (CurRecord <> LastRecord) THEN
    GOTO 1060
  END IF
  
1072
  IF BackFlag = True THEN
    CurRecord = CurRecord - 1
  ELSE
    CurRecord = CurRecord + 1
  END IF
  GOTO 1060
  
1080
  IF BeenHere THEN
    CLS
    QPSound 1000, 1
    PRINT "NO MORE METERS!!!"
    QPSound 1000, 3
    PRINT
    PRINT
    PRINT
    PRINT "Press any Key...";
    ZZ = BiosKey
    GOTO Top
  ELSE
    CLS
    'QPSound 1000, 1
    PRINT "ROUTE: "; CurBook
    PRINT
    PRINT "NOT FOUND"
    PRINT
    PRINT
    PRINT "Press any Key. . .";
    ZZ = BiosKey
    GOTO Top
  END IF
  
1100
  CLS
  BeenHere = True
  IF ReadRec(1).ReadFlag = "Y" THEN
    PRINT "*";
  ELSE
    PRINT " ";
  END IF
  PRINT ReadRec(1).CustName
  PRINT " "; ReadRec(1).ServAddress
  PRINT " Mtr# "; ReadRec(1).MeterID;
  SELECT CASE ReadRec(1).MeterType
  CASE "E"
    PRINT " *Electric"
  CASE "D"
    PRINT " *Demand"
  CASE "W"
    PRINT " *Water"
  CASE "C"
    PRINT " *Wtr/Swr"
  CASE "S"
    PRINT " *Sewer"
  CASE "G"
    PRINT " *Gas"
  END SELECT
  PRINT " Previous:";
  IF ReadRec(1).ReadFlag = "Y" THEN
    RSET Fmt2$ = QPTrim$(STR$(ReadRec(1).PastRead))
    PRINT Fmt2$;
  ELSE
    RSET Fmt2$ = QPTrim$(STR$(ReadRec(1).CurRead))
    PRINT Fmt2$;
  END IF
  IF ReadRec(1).ReadFlag = "Y" THEN
    PreSetCur = True
  END IF
  IF ASC(LEFT$(ReadRec(1).Note1, 1)) > 32 OR ASC(LEFT$(ReadRec(1).Note2, 1)) > 32 OR ASC(LEFT$(ReadRec(1).Note3, 1)) > 32 THEN
    CLocate 6, 2
    PRINT "Note!";
    QPSound 1000, 1
  END IF
  GOSUB ShowStatBar
  CLocate 5, 1
  PRINT " Current.."; Fmt2$;
  
GetInput:
  
  CF = False
  fl = 9
  
  IF PreSetCur THEN
    PreSetCur = False
    In$ = QPTrim$(STR$(ReadRec(1).CurRead)) + SPACE$(fl)
    In$ = LEFT$(In$, fl)
  ELSE
    ExitCode = False
    In$ = SPACE$(fl)
  END IF
  
  OIn$ = In$
  WInput In$, GoodKey3$, 5, 11, ExitCode
  
  CF = ExitCode
  
  IF OIn$ <> In$ THEN
    SELECT CASE ExitCode
    CASE 59, 62, 72, 80
      ExitCode = False
    END SELECT
  END IF
  
  SELECT CASE ExitCode
  CASE EscKey
    GOSUB ShowHelp
    GOTO 1100
  CASE F1Key    'f1
    GOSUB F1Funcs
    IF ScanCode = 1 THEN
      CurRecord = 1
      GOTO 1059
    END IF
    GOTO 1100
  CASE F2Key       'f2key
    GOSUB DisplayNotes
    GOTO 1100
    'F3 is handled inside WInput
  CASE F4Key       'F4
    GOTO Top
  CASE 72       'UpArrow
    CurRecord = CurRecord - 1
    IF CurRecord < 1 THEN CurRecord = 1
    IF BackFlag = False THEN
      BackFlag = True
      GOTO 1059
    ELSE
      GOTO 1060
    END IF
    
  CASE 80       'DownArrow
    BackFlag = False
    CurRecord = CurRecord + 1
    ScanCode = 1
    IF CurRecord > LastRecord THEN
      CurRecord = LastRecord
      GOTO 1080
    ELSE
      GOTO 1059
    END IF
  END SELECT
  
  IF CF = EscKey THEN RETURN
  
  SELECT CASE CF
  CASE 11, 12
    GOTO GetInput
  END SELECT
  
  CLocate 5, 11
  CurRead# = VAL(In$)
  RSET Fmt2$ = QPTrim$(STR$(CurRead#))
  PRINT Fmt2$;
  
  CLocate 6, 10
  PRINT "---------"
  IF ReadRec(1).ReadFlag = "Y" THEN
    IF CurRead# < ReadRec(1).PastRead THEN
      GOSUB ChkMeterTrip
    ELSE
      CalcUsage# = CurRead# - ReadRec(1).PastRead
    END IF
  ELSE
    IF CurRead# < ReadRec(1).CurRead THEN
      GOSUB 1755
    ELSE
      CalcUsage# = CurRead# - ReadRec(1).CurRead
    END IF
  END IF
  
  RSET Fmt2$ = QPTrim$(STR$(CalcUsage#))
  PRINT " Usage:   "; Fmt2$;
  IF CalcUsage# < ReadRec(1).LowRead THEN
    CLocate 7, 20
    PRINT "L"
    QPSound 1000, 2
  END IF
  
  IF CalcUsage# > ReadRec(1).HighRead THEN
    CLocate 7, 20
    PRINT "H"
    QPSound 1000, 2
  END IF
  
  
1152
  CLocate 8, 1
  PRINT "   ESC=NO  Enter=OK ";
  ExitCode = True
  kk$ = " "
  WInput kk$, "", 8, 21, ExitCode
  
  CF = ExitCode
  SELECT CASE CF
  CASE EscKey
    GOTO 1100
  CASE True
    ExitCode = False
    GOTO UpdateReading
  CASE ELSE
    GOTO 1152
  END SELECT
  
UpdateReading:
  IF ReadRec(1).ReadFlag <> "Y" THEN
    ReadRec(1).PastRead = ReadRec(1).CurRead
  END IF
  ReadRec(1).CurRead = CurRead#
  ReadRec(1).ReadTime = LEFT$(TIME$, 5)
  ReadRec(1).ReadDate = Date2Num%(DATE$)
  ReadRec(1).ReadFlag = "Y"
  GOSUB OpenFile
  PUT 1, CurRecord, ReadRec(1)
  GOSUB CloseFile
  BackFlag = False
RETURN
  
ChkMeterTrip:
  'check for meter triping
  J = LEN(STR$(ReadRec(1).PastRead)) - 1
  Max# = 10 ^ J
  CalcUsage# = Max# - ReadRec(1).PastRead + CurRead#
RETURN
  
1755          REM check for meter triping
  J = LEN(STR$(ReadRec(1).CurRead)) - 1
  Max# = 10 ^ J
  CalcUsage# = Max# - ReadRec(1).CurRead + CurRead#
RETURN
  
  
F1Funcs:
  CLS
  PRINT "Battery Power: "; GetPower + 1; "%"
  PRINT "Select Search:"
  PRINT "F1 = First un-read"
  PRINT "F2 = Last un-read"
  PRINT "F3 = Meter Number"
  PRINT "F4 = Service Addr."
  PRINT "ESC = Cancel"
  PRINT "==>";
  
  CF = False
  fl = 1
  MNum$ = SPACE$(fl)
  
  DO

    NoBackLight = True  'This temporarily disables f3 backlight toggle
    WInput MNum$, "", 8, 6, ExitCode
    
    CF = ExitCode
    MNum$ = UCASE$(MNum$)
    
    SELECT CASE ExitCode
    CASE EscKey
      EXIT DO
    CASE F1Key
      CurRecord = 1
      GOSUB OpenFile
      TNumOfRecs& = LOF(1) \ ReadRecLen
      FOR CurRecord = 1 TO TNumOfRecs&
        GET 1, CurRecord, ReadRec(1)
        IF ReadRec(1).Book = CurBook THEN
          IF ReadRec(1).ReadFlag <> "Y" THEN
            EXIT FOR
          END IF
        END IF
      NEXT
      GOSUB CloseFile
      EXIT DO
    CASE F2Key
      Last = CurRecord
      GOSUB OpenFile
      TNumOfRecs& = LOF(1) \ ReadRecLen
      FOR CurRecord = TNumOfRecs& TO 1 STEP -1
        GET 1, CurRecord, ReadRec(1)
        IF ReadRec(1).Book = CurBook THEN
          IF ReadRec(1).ReadFlag <> "Y" THEN
            EXIT FOR
          END IF
        END IF
      NEXT
      GET 1, CurRecord, ReadRec(1)
      GOSUB CloseFile
      EXIT DO
    CASE F3Key
      GOSUB FindMeter
      EXIT DO
    CASE F4Key
      NoBackLight = False
      GOSUB DoServAddr
      EXIT DO
    CASE True
      EXIT DO
    CASE ELSE
      ScanCode = False
      IF CF = EscKey THEN
        EXIT DO
      END IF
    END SELECT
    IF CF >= 60 AND CF <= 70 THEN
      GOTO F1Funcs
    END IF
  LOOP
  
  NoBackLight = False
RETURN
  
FindMeter:
  GotMtr = False

FindMeterTop:
  MNum$ = SPACE$(20)
  LSET MNum$ = WhatMeter$
  CLS
  PRINT "Enter meter number."
  PRINT "=>";
  WInput MNum$, GoodKey2$ + GoodKey3$, 2, 3, ExitCode

  IF ExitCode = EscKey THEN
    GOTO FindMtrRet
  END IF

  WhatMeter$ = QPTrim$(MNum$)
  IF LEN(WhatMeter$) = False THEN
    GOTO F1Funcs
  END IF

  GOSUB OpenFile
  FOR X = 1 TO LastRecord
    GET 1, X, ReadRec(1)
    GOSUB Scanning
    trmeter$ = QPTrim$(ReadRec(1).MeterID)
    IF WhatMeter$ = trmeter$ THEN
      IF ReadRec(1).ReadFlag <> "Y" THEN
        GotMtr = True
        CurRecord = X
        GOSUB CloseFile
        GOTO FindMtrRet
      END IF
    END IF
  NEXT X
  GET 1, CurRecord, ReadRec(1)
  GOSUB CloseFile
  GOSUB ClearArea
  CLocate 4, 1
  PRINT " METER: " + WhatMeter$
  PRINT " NOT Located"
  PRINT " Press <ENTER>"
  ZZ = BiosKey
  IF GotMtr = False THEN
    GOTO FindMeterTop
  END IF

FindMtrRet:
RETURN
  
DisplayNotes:
  IF CF = F2Key THEN
    CLS
    CLocate 1, 1
    PRINT Ast$
    CLocate 1, 8
    PRINT "Notes"
    PRINT Ast$
    PRINT ReadRec(1).Note1
    PRINT ReadRec(1).Note2
    PRINT ReadRec(1).Note3
    PRINT Ast$
    
dp1:
    CLocate 8, 1
    PRINT "New Notes (Y/N) [ ]" ';
    YN$ = " "
    
    WInput YN$, "YN", 8, 18, ExitCode
    CF = ExitCode
    
    IF CF = 76 THEN RETURN
    
    'IF NOT (YN$ = "Y" OR YN$ = "N") THEN GOTO dp1
    IF YN$ = "N" OR CF = True OR CF = EscKey THEN RETURN
    
dp2:
    CLocate 8, 1
    PRINT Blank$
    fl = 20
    Note1$ = LEFT$(QPTrim$(ReadRec(1).Note1) + SPACE$(fl), 20)
    Note2$ = LEFT$(QPTrim$(ReadRec(1).Note2) + SPACE$(fl), 20)
    Note3$ = LEFT$(QPTrim$(ReadRec(1).Note3) + SPACE$(fl), 20)
    
Line1:
    
    WInput Note1$, GoodKey2$ + GoodKey3$, 3, 1, ExitCode
    CF = ExitCode
    IF CF = EscKey THEN RETURN
    LSET ReadRec(1).Note1 = QPTrim$(Note1$)
    IF CF = 72 THEN
      GOTO Line1
    END IF
    IF CF = 80 THEN
      GOTO Line2
    END IF
    
Line2:
    
    WInput Note2$, GoodKey2$ + GoodKey3$, 4, 1, ExitCode
    CF = ExitCode
    IF CF = EscKey THEN RETURN
    LSET ReadRec(1).Note2 = QPTrim$(Note2$)
    IF CF = 72 THEN
      GOTO Line1
    END IF
    IF CF = 80 THEN
      GOTO Line3
    END IF
    
    
Line3:
    
    WInput Note3$, GoodKey2$ + GoodKey3$, 5, 1, ExitCode
    CF = ExitCode
    IF CF = EscKey THEN RETURN
    LSET ReadRec(1).Note3 = QPTrim$(Note3$)
    IF CF = True THEN
      GOTO dp3
    END IF
    IF CF = 72 THEN
      GOTO Line2
    END IF
    IF CF = 13 THEN
      GOTO dp3
    ELSE
      GOTO Line3
    END IF
    
dp3:
    CLocate 8, 1
    PRINT "Save Notes (Y/N) [ ]"
    YN$ = " "
    WInput YN$, "YN", 8, 19, ExitCode
    CF = ExitCode
    IF CF = EscKey THEN noteflag = False: RETURN
    IF YN$ = "Y" THEN
dp4:
      CLocate 8, 1
      PRINT "Perm/Temp (P/T)  [ ] "             ';
      CLocate 8, 18
      
      PT$ = " "
      WInput PT$, "PT", 8, 19, ExitCode
      CF = ExitCode
      IF CF = EscKey THEN
        noteflag = False
        RETURN
      END IF
      PT$ = UCASE$(PT$)
      IF (PT$ <> "P") AND (PT$ <> "T") THEN GOTO dp4
      noteflag = 1
      ReadRec(1).NoteStatus = PT$
      GOSUB OpenFile
      PUT 1, CurRecord, ReadRec(1)
      GOSUB CloseFile
      RETURN
    ELSEIF YN$ = "N" THEN
      RETURN
    END IF
    IF PT$ = "N" THEN GOTO dp2
    GOTO dp3
  END IF
  
RETURN
  'must go back to first page
  
Scanning:
  ScnCnt = ScnCnt + 1
  IF ScnCnt > 5 THEN
    RSET Fmt2$ = QPTrim$(STR$(CurRecord))
    ScnCnt = False
    Char = Char + 1
    IF Char > 4 THEN Char = 1
    LOCATE 4, 14, 0
    PRINT MID$("\|/-", Char, 1)
    CLocate 6, 4
    PRINT "Rec:"; Fmt2$;
  END IF
  RETURN
  
OpenFile:
  OPEN "UBCUSTTR.DAT" FOR RANDOM SHARED AS #1 LEN = ReadRecLen
RETURN
  
CloseFile:
  CLOSE #1
RETURN
  
ShowStatBar:
  CLocate 8, 1
  PRINT StatBar$;
RETURN
  
DoServAddr:
  ExitCode = True
  GotSAdd = False
  SAdd$ = SPACE$(20)
  CLS
  PRINT "Enter partial service addr."
  PRINT "=>";
  WInput SAdd$, "." + GoodKey2$ + GoodKey3$, 2, 3, ExitCode
  
  SAdd$ = QPTrim$(SAdd$)
  IF LEN(SAdd$) = False THEN
    GOTO ServAddrRet
  END IF
  
  GOSUB OpenFile
  FOR X = 1 TO LastRecord
    GET 1, X, ReadRec(1)
    GOSUB Scanning
    TSAdd$ = UCASE$(QPTrim$(ReadRec(1).ServAddress))
    IF INSTR(TSAdd$, SAdd$) > 0 THEN
      GOSUB AskSAddOK
      IF GotSAdd THEN
        CurRecord = X
        EXIT FOR
      END IF
    END IF
  NEXT X
  GOSUB CloseFile
  
  '    GET 1, CurRecord, ReadRec(1)
  '    GOSUB CloseFile
  '    CLocate 7, 1
  '    PRINT " METER: " + WhatMeter$
  '    PRINT " NOT Located"
  '    PRINT " Press <ENTER>"
  '    zz = BiosKey
  '    GOTO F1Funcs
  '
ServAddrRet:
RETURN
  
AskSAddOK:
  GOSUB CloseFile
  OK$ = SPACE$(1)
  GOSUB ClearArea
  CLocate 4, 1
  PRINT ReadRec(1).CustName
  PRINT TSAdd$
  CLocate 6, 1
  PRINT "OK? (Y/N) "
  WInput OK$, "YN", 6, 11, ExitCode
  IF OK$ = "Y" THEN
    GotSAdd = True
  END IF
  GOSUB OpenFile
  GOSUB ClearArea
RETURN
  
ClearArea:
  FOR ZZ = 3 TO 8
    CLocate ZZ, 3
    PRINT Blank$;
  NEXT
RETURN

ShowHelp:
  CLS
  PRINT "ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿"        '1
  PRINT "³ (C)97-99 Southern Software ³"        '4
  PRINT "³    "; CHR$(24); " =Next    "; CHR$(25); " =Prev      ³"        '3Read           ³"   '3
  PRINT "³   F1 =Scan   F2 =Notes     ³"        '3
  PRINT "³   F3 =BackLight Toggle     ³"        '6
  PRINT "³   F4 =New Route / Exit     ³"        '7
  PRINT "³ Press any key to continue. ³"        '3
  PRINT "ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ";       '8
  CLocate 7, 29
  HKey = BiosKey
RETURN

SUB CLocate (BYVAL Row%, BYVAL Col%)
  LOCATE Row, Col
END SUB

SUB WInput (Edit$, GoodKey$, Row, Col, ExitCode)
  
  'Return Codes:
  ' Enter Key Exit
  '     0
  ' Escape Key Exit
  '   -27 Escape Key was Pressed to exit sub
  '       Edit string reset to original string
  'SHARED NoBackLight
  
  STATIC InsertMode
  
  GOSUB ShowCursor
  
  Original$ = Edit$
  ExitCode = 0
  
  MaxLength = LEN(Edit$)        'max length of input string
  Ptr = LEN(RTRIM$(Edit$))
  IF Ptr = MaxLength THEN Ptr = MaxLength - 1
  'added code to just look for function keys with no scrn output
  IF Ptr <= 0 THEN Ptr = 0
  
  DO            'main processing loop
    
    LOCATE Row, Col, 1          'update cursor position.
    PRINT Edit$;                'and display the line
    LOCATE Row, Col + Ptr       'update cursor position.
    
    PressedKey = BiosKey        'wait for a key press
    
    SELECT CASE PressedKey      'process the key
      
      '****  Exit keys
    CASE 13     'ENTER - Accept line and, exit
      Edit$ = QPTrim$(Edit$)
      IF LEN(Edit$) = 0 THEN
        ExitCode = True
      ELSE
        ExitCode = 13
      END IF
      
      EXIT DO
      
    CASE 27     'ESCAPE - Abort operation exit
      ExitCode = -27
      Edit$ = Original$
      EXIT DO
      
    CASE -61    'F3
      IF NoBackLight THEN
        ExitCode = ABS(PressedKey)
        EXIT DO
      ELSE
        BackLightToggle
      END IF
    CASE -59, -60, -62, -72, -80                'F1 F2 F4 DnArrow UpArrow
      'UpArrow DnArrow F1 F2 F4
      ExitCode = ABS(PressedKey)
      EXIT DO
      
      'add alt exits here
      ' i.e. downarrow
      'End of Exit keys
      
      '****  Cursor keys
    CASE -82    'INSERT - Toggle insert or, overtype mode
      InsertMode = NOT InsertMode
      GOSUB ShowCursor
    CASE -75    'LEFTARROW - Moves cursor left
      IF Ptr THEN Ptr = Ptr - 1
      IF Col + Ptr <= 0 THEN
        Ptr = Ptr + 1
      END IF
    CASE -77    'RIGHTARROW - Moves cursor right
      IF Ptr < MaxLength - 1 THEN Ptr = Ptr + 1
    CASE -71    'HOME - Moves cursor to first position
      Ptr = 0
    CASE -79    'END - Moves cursor to last position
      Ptr = LEN(RTRIM$(Edit$))
      IF Ptr = MaxLength THEN Ptr = MaxLength - 1
      'End of Cursor keys
      
      '****   Editing Keys & Normal Input
      '    CASE -117   'CONTROLEND - Erase from cursor to end of Edit$
      '      IF Ptr THEN
      '        LSET Edit$ = LEFT$(Edit$, Ptr)
      '      ELSE
      '        Edit$ = SPACE$(LEN(Edit$))
      '      END IF
      '    CASE -22    'ALT-U     - Restore to orginial string
      '      Edit$ = Original$
      '      Ptr = 0
      
    CASE 8      'BACKSPACE - Moves cursor left and erase characters
      IF Ptr THEN
        Edit$ = LEFT$(Edit$, Ptr - 1) + MID$(Edit$, Ptr + 1) + " "
        Ptr = Ptr - 1
      END IF
      
    CASE -83    'DELETE - Erases characters at cursor
      Edit$ = LEFT$(Edit$, Ptr) + MID$(Edit$, Ptr + 2) + " "
      
    CASE ELSE   'ALL OTHER NORMAL KEYS
      IF PressedKey > 1 AND PressedKey < 256 THEN
        K$ = CHR$(PressedKey)
        K$ = UCASE$(K$)
        IF INSTR(GoodKey$, K$) THEN             'if key is in the keylist
          IF InsertMode THEN    'Insert or Overtype the key
            Edit$ = LEFT$(Edit$, Ptr) + K$ + MID$(Edit$, Ptr + 1)
            Edit$ = LEFT$(Edit$, MaxLength)
          ELSE
            IF Ptr < MaxLength THEN MID$(Edit$, Ptr + 1, 1) = K$
          END IF
          IF Ptr < MaxLength - 1 THEN Ptr = Ptr + 1             'Are we at max length
        END IF
      END IF
      'End of Editing & Input keys
    END SELECT
  LOOP
  EXIT SUB
  
ShowCursor:
  IF InsertMode THEN
    LOCATE , , , 0, 6
  ELSE          'show appropriate cursor type
    LOCATE , , , 5, 6
  END IF
  RETURN
  
END SUB

