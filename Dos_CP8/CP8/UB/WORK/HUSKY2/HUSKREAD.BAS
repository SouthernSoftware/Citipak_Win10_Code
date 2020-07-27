DEFINT A-Z    'New hand held program for Version 8.2w
DECLARE SUB HVertMenu (Item$(), HChoice%, Ky)
DECLARE SUB BackLightToggle ()
DECLARE SUB SetFont ()
DECLARE SUB WInput (Edit$, GoodKey$, Row%, Col%, ExitCode%)
DECLARE SUB CLocate (BYVAL Row%, BYVAL Col%)
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE SUB QPSound (Frequency%, Duration%)
DECLARE FUNCTION GetPower ()
DECLARE FUNCTION BiosKey% ()
DECLARE SUB APrint0 (BYVAL Address, NumEls, MidFirst, NumChars, Colr)
DECLARE SUB PaintBox0 (ULRow, ULCol, LRRow, LRCol, Colr)
DECLARE SUB QPrintRC (x$, Row, Col, Colr)
DECLARE FUNCTION QPValL& (Numb$)
  
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'ubhusky.bi'
  
  CONST False = 0, True = NOT False
  CONST F1Key = 59
  CONST F2Key = 60
  CONST F3Key = 61
  CONST F4Key = 62
  CONST EscKey = -27
  
  DIM SHARED GoodKey2$, GoodKey3$
  DIM SHARED F3Disable, NotesRecs

  DIM Note$(1 TO 3)
  DIM NoteRec AS MtrNotesType

  BoxBot = 6
  ScrlClr = 112
  BoxSClr = 7
  MaxLen = 20

  '---- determine how many HChoices to display, and where to put the menu
  DspTop = 2                     'put the menu at the current
  LeftMargin = 5                 'cursor location

  NoteRecLen = 20
  OPEN "UBHHCODE.DAT" FOR RANDOM AS #1 LEN = NoteRecLen
  NotesRecs = LOF(1) / NoteRecLen

  IF NotesRecs > 0 THEN
    REDIM Item$(1 TO NotesRecs)
    FOR Cnt = 1 TO NotesRecs
      GET #1, , NoteRec
      Item$(Cnt) = NoteRec.Note
    NEXT
  ELSE
    NotesRecs = 1
    REDIM Item$(NotesRecs)
    Item$(NotesRecs) = SPACE$(NoteRecLen)
    LSET Item$(NotesRecs) = "NO NOTES DEFINED!"
  END IF
  CLOSE
  
  GoodKey2$ = " "
  FOR Cnt = 65 TO 90
    GoodKey2$ = GoodKey2$ + CHR$(Cnt)
  NEXT
  GoodKey3$ = "1234567890"

  SetFont

  Ast$ = "********************"
  Blank$ = SPACE$(26)
  Fmt2$ = SPACE$(9)

  REDIM ReadRec(1 TO 2) AS UBHuskyHHRecType
  ReadRecLen = LEN(ReadRec(1))
  
Top:
  
  CLS
  PRINT "ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿"                          '1
  PRINT "³ Meter Reading System v4.01 ³"        '3
  PRINT "³(C)97-2000 Southern Software³"        '4
  PRINT "³                            ³"        '5
  PRINT "³      Enter Route:          ³"        '6
  PRINT "³                            ³"        '5
  PRINT "³     Press ESC to Exit.     ³"        '7
  PRINT "ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ";     '8
  
GetRoute:
  Route$ = "   "
  WInput Route$, GoodKey3$, 5, 21, ExitCode
  
  IF ExitCode = EscKey THEN
    CLS
    CLocate 2, 1
    PRINT "Program Exited."
    END
  END IF
  
  IF LEN(QPTrim$(Route$)) = False THEN
    GOTO GetRoute
  END IF
  
  CurBook = QPValL&(Route$)
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
    x = CurRecord
    GOSUB OpenFile
    GET 1, x, ReadRec(1)
    GOSUB Scanning
    GOSUB CloseFile
    IF ScanCode = 1 THEN
      IF TrReadCode$ = "Y" THEN
        CurRecord = CurRecord + 1
        GOTO 1059
      END IF
    END IF
  END IF
  IF ReadRec(1).Book <> CurBook THEN 1072
  ScanCode = False
  BackFlag = False
  GOSUB 1100
  
  IF (ExitCode = EscKey) AND (CurRecord <> LastRecord) THEN
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
  CLS
  IF BeenHere THEN
    PRINT "NO MORE METERS!!!"
  ELSE
    PRINT "ROUTE: "; CurBook
    PRINT
    PRINT "NOT FOUND"
  END IF
  CLocate 7, 1
  PRINT "Press any Key...";
  GOSUB MakeSound
  zz = BiosKey
  GOTO Top
  
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
  PRINT " Mt# "; ReadRec(1).MeterID; " ";
  SELECT CASE ReadRec(1).MeterType
  CASE "E"
    PRINT "Ele";
  CASE "D"
    PRINT "Dem";
  CASE "W"
    PRINT "Wat";
  CASE "C"
    PRINT "W/S";
  CASE "S"
    PRINT "Sew";
  CASE "G"
    PRINT "Gas";
  END SELECT
  PRINT " C1:" + ReadRec(1).UCode1 + " C2:" + ReadRec(1).UCode2
  IF ReadRec(1).ReadFlag = "Y" THEN
    PreSetCur = True
    RSET Fmt2$ = QPTrim$(STR$(ReadRec(1).PastRead))
  ELSE
    RSET Fmt2$ = QPTrim$(STR$(ReadRec(1).CurRead))
  END IF
  PRINT " Previous: "; Fmt2$;
  
  FOR NCnt = 1 TO 3
    IF LEN(QPTrim$(ReadRec(1).Notes(NCnt).Note)) > 0 THEN
      CLocate 4, 23
      PRINT "Note!";
      GOSUB MakeSound
      EXIT FOR
    END IF
  NEXT
  GOSUB ShowStatBar
  CLocate 5, 1
  PRINT "  Current: "; Fmt2$;
  
GetInput:
  
  fl = 9
  IF PreSetCur THEN
    PreSetCur = False
    In$ = QPTrim$(STR$(ReadRec(1).CurRead)) + SPACE$(fl)
    In$ = LEFT$(In$, fl)
  ELSE
    In$ = SPACE$(fl)
  END IF
  
  OIn$ = In$
  WInput In$, GoodKey3$, 5, 12, ExitCode

  IF ExitCode = True AND LEN(In$) = 0 THEN
    GOTO 1100
  END IF

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
    
  CASE 84       'Shift F1
    GOSUB ShowRouteStat
    GOTO 1100
    
  CASE F1Key    'f1
    GOSUB F1Funcs
    IF ScanCode = 1 THEN
      CurRecord = 1
      GOTO 1059
    END IF
    GOTO 1100
  CASE F2Key    'f2key
    GOSUB DisplayNotes
    GOTO 1100

  'F3 is handled inside WInput

  CASE F4Key    'F4
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
  
  IF ExitCode = EscKey THEN RETURN
  
  SELECT CASE ExitCode
  CASE 11, 12
    GOTO GetInput
  END SELECT
  
  CLocate 5, 12
  CurRead# = QPValL&(In$)
  RSET Fmt2$ = QPTrim$(STR$(CurRead#))
  PRINT Fmt2$;
  
  CLocate 6, 12
  PRINT "---------"
  IF ReadRec(1).ReadFlag = "Y" THEN
    IF CurRead# < ReadRec(1).PastRead THEN
      J = LEN(STR$(ReadRec(1).PastRead)) - 1
      Max# = 10 ^ J
      CalcUsage# = Max# - ReadRec(1).PastRead + CurRead#
    ELSE
      CalcUsage# = CurRead# - ReadRec(1).PastRead
    END IF
  ELSE
    IF CurRead# < ReadRec(1).CurRead THEN
      J = LEN(STR$(ReadRec(1).CurRead)) - 1
      Max# = 10 ^ J
      CalcUsage# = Max# - ReadRec(1).CurRead + CurRead#
    ELSE
      CalcUsage# = CurRead# - ReadRec(1).CurRead
    END IF
  END IF
  
  RSET Fmt2$ = QPTrim$(STR$(CalcUsage#))
  PRINT "    Usage: "; Fmt2$;
  IF CalcUsage# < ReadRec(1).LowRead THEN
    PRINT "L"
    GOSUB MakeSound
  END IF
  
  IF CalcUsage# > ReadRec(1).HighRead THEN
    PRINT "H"
    GOSUB MakeSound
  END IF
  
  
1152
  CLocate 8, 1
  PRINT "   ESC=NO  Enter=OK ";
  WInput "", "", 8, 21, ExitCode
  
  SELECT CASE ExitCode
  CASE EscKey
    GOTO 1100
  CASE True
    GOTO UpdateReading
  CASE ELSE
    GOTO 1152
  END SELECT
  
UpdateReading:
  IF ReadRec(1).ReadFlag <> "Y" THEN
    ReadRec(1).PastRead = ReadRec(1).CurRead
  END IF
  ReadRec(1).CurRead = CurRead#
  ReadRec(1).ReadDate = Date2Num%(DATE$)
  ReadRec(1).ReadFlag = "Y"
  GOSUB OpenFile
  PUT 1, CurRecord, ReadRec(1)
  GOSUB CloseFile
  BackFlag = False
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
  
  DO
    
    F3Disable = True       'Temporarily disables f3 backlight toggle
    WInput "", "", 8, 6, ExitCode
    
    SELECT CASE ExitCode
    CASE EscKey
      EXIT DO
    CASE F1Key
      CurRecord = 1
      GOSUB OpenFile
      'LastRecord = LOF(1) \ ReadRecLen
      FOR CurRecord = 1 TO LastRecord
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
      'LastRecord = LOF(1) \ ReadRecLen
      FOR CurRecord = LastRecord TO 1 STEP -1
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
      F3Disable = False
      GOSUB DoServAddr
      EXIT DO
    CASE True
      EXIT DO
    CASE ELSE
      ScanCode = False
      IF ExitCode = EscKey THEN
        EXIT DO
      END IF
    END SELECT
    IF ExitCode >= 60 AND ExitCode <= 70 THEN
      GOTO F1Funcs
    END IF
  LOOP
  F3Disable = False
RETURN

DisplayNotes:
  
  GOSUB NotesDisplay
  
  CLocate 8, 1
  PRINT "New Notes (Y/N) [ ]"   ';
  YN$ = "N"
  
  WInput YN$, "YN", 8, 18, ExitCode
  HChoice = 1

  IF YN$ = "N" OR ExitCode = True OR ExitCode = EscKey THEN
    GOTO NotesReturn
  END IF
  
  CLocate 8, 1
  PRINT Blank$
  
  NoteCnt = 1
  
  DO
    CLocate 8, 1
    'PRINT "Press F2 for Picklist.";
    PRINT "Press F2 to Select.";
    Note$(NoteCnt) = LEFT$(QPTrim$(ReadRec(1).Notes(NoteCnt).Note) + SPACE$(NoteRecLen), NoteRecLen)
    WInput Note$(NoteCnt), GoodKey2$ + GoodKey3$, NoteCnt + 2, 1, ExitCode
    ExitCode = ExitCode

    IF ExitCode = EscKey THEN
      GOTO NotesReturn
    END IF

    IF ExitCode = 60 THEN   'abs(F2KEY)
      GOSUB HVertMenu  'Item$(), HChoice, Ky
      IF Ky = 13 THEN
        Note$(NoteCnt) = LEFT$(QPTrim$(Item$(HChoice)) + SPACE$(NoteRecLen), NoteRecLen)
      END IF
      GOSUB NotesDisplay
    END IF
    
    LSET ReadRec(1).Notes(NoteCnt).Note = QPTrim$(Note$(NoteCnt))
    SELECT CASE ExitCode
    CASE 72             'ABS(UP ARROW)
      IF NoteCnt > 1 THEN
        NoteCnt = NoteCnt - 1
      END IF
    CASE 13, True
      IF NoteCnt = 3 THEN
        EXIT DO
      ELSE
        NoteCnt = NoteCnt + 1
      END IF
    CASE 80             'ABS(DOWN ARROW)
      IF NoteCnt < 3 THEN
        NoteCnt = NoteCnt + 1
      END IF
    END SELECT
  LOOP
  
  CLocate 8, 1
  PRINT "Save Notes (Y/N) [ ]"
        
  YN$ = " "
  WInput YN$, "YN", 8, 19, ExitCode

  IF ExitCode = EscKey THEN
    GOTO NotesReturn
  END IF
  
  IF YN$ = "Y" THEN
dp4:
    YN$ = "P"
    CLocate 8, 1
    PRINT "Perm/Temp (P/T)  [ ] "
    CLocate 8, 18
    WInput YN$, "PT", 8, 19, ExitCode
    SELECT CASE ExitCode
    CASE EscKey
      GOTO NotesReturn
    CASE 13
      IF (YN$ <> "P") AND (YN$ <> "T") THEN
        GOTO dp4
      END IF
      ReadRec(1).NoteStatus = YN$
      GOSUB OpenFile
      PUT 1, CurRecord, ReadRec(1)
      GOSUB CloseFile
    CASE ELSE
      GOTO dp4
    END SELECT
  END IF
  
NotesReturn:
RETURN
  
NotesDisplay:
  CLS
  CLocate 1, 8
  PRINT "Notes"
  PRINT Ast$
  PRINT ReadRec(1).Notes(1).Note
  PRINT ReadRec(1).Notes(2).Note
  PRINT ReadRec(1).Notes(3).Note
  PRINT Ast$
RETURN
  'must go back to first page
  
Scanning:
  ScnCnt = ScnCnt + 1
  IF ScnCnt > 5 THEN
    RSET Fmt2$ = QPTrim$(STR$(CurRecord))
    ScnCnt = False
    Char = Char + 1
    IF Char > 4 THEN Char = 1
    CLocate 4, 14
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
  PRINT " Press ESC for Help.";
RETURN
'***********************************
FindMeter:
  GotMtr = False

FindMeterTop:
  MNum$ = SPACE$(20)
'  LSET MNum$ = WhatMeter$
  CLS
  PRINT "Enter partial meter number."
  PRINT "=>";
  WInput MNum$, GoodKey2$ + GoodKey3$, 2, 3, ExitCode
  IF ExitCode = EscKey THEN
    GOTO FindMtrRet
  END IF

  WhatMeter$ = QPTrim$(MNum$)
  IF LEN(WhatMeter$) = False THEN
    GOTO FindMtrRet
  END IF

  GOSUB OpenFile
  FOR x = 1 TO LastRecord
    GET 1, x, ReadRec(1)
    GOSUB Scanning
    IF ReadRec(1).Book = CurBook THEN
      trmeter$ = QPTrim$(ReadRec(1).MeterID)
      IF INSTR(trmeter$, WhatMeter$) THEN
        GOSUB AskMtrOK
        IF GotMtr THEN
          CurRecord = x
          EXIT FOR
        ELSEIF ExitCode = EscKey THEN
          EXIT FOR
        END IF
      END IF
    END IF
  NEXT x
  IF NOT GotMtr THEN
    GET 1, CurRecord, ReadRec(1)
  END IF
  GOSUB CloseFile

FindMtrRet:
RETURN

AskMtrOK:
  GOSUB CloseFile
  YN$ = " "
  GOSUB ClearArea
  CLocate 4, 1
  PRINT ReadRec(1).CustName
  PRINT trmeter$
  CLocate 6, 1
  PRINT "OK? (Y/N) "
  WInput YN$, "YN", 6, 11, ExitCode
  IF YN$ = "Y" THEN
    GotMtr = True
  END IF
  GOSUB OpenFile
  GOSUB ClearArea
RETURN
'***********************************

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
  FOR x = 1 TO LastRecord
    GET 1, x, ReadRec(1)
    GOSUB Scanning
    TSAdd$ = UCASE$(QPTrim$(ReadRec(1).ServAddress))
    IF INSTR(TSAdd$, SAdd$) > 0 THEN
      GOSUB AskSAddOK
      IF GotSAdd THEN
        CurRecord = x
        EXIT FOR
      ELSEIF ExitCode = EscKey THEN
        EXIT FOR
      END IF
    END IF
  NEXT x
  IF NOT GotSAdd THEN
    GET 1, CurRecord, ReadRec(1)
  END IF
  GOSUB CloseFile
  
ServAddrRet:
RETURN
  
AskSAddOK:
  GOSUB CloseFile
  YN$ = " "
  GOSUB ClearArea
  CLocate 4, 1
  PRINT ReadRec(1).CustName
  PRINT TSAdd$
  CLocate 6, 1
  PRINT "OK? (Y/N) "
  WInput YN$, "YN", 6, 11, ExitCode
  IF YN$ = "Y" THEN
    GotSAdd = True
  END IF
  GOSUB OpenFile
  GOSUB ClearArea
RETURN

ClearArea:
  FOR zz = 3 TO 8
    CLocate zz, 3
    PRINT Blank$;
  NEXT
RETURN
  
ShowHelp:
  CLS

  PRINT "ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿"                          '1
  PRINT "³(C)97-2000 Southern Software³"        '2
  PRINT "³   "; CHR$(24); "=Next      "; CHR$(25); "=Prev       ³"              '3
  PRINT "³  F1=Search   F3=BkLight    ³"        '4
  PRINT "³  F2=Notes    F4=Route/Exit ³"        '5
  PRINT "³   SHIFT+F1=Route Status    ³"        '6
  PRINT "³ Press any key to continue. ³"        '7
  PRINT "ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ";     '8
  CLocate 7, 29
  HKey = BiosKey
RETURN
  
ShowRouteStat:
  CLS
  PRINT "Route Stats for:"; ReadRec(1).Book
  RouteTot = False
  ReadCnt = False
  ReadSkip = False
  
  GOSUB OpenFile
  FOR RCnt = 1 TO LastRecord
    WasRead = False
    CLocate 3, 1
    PRINT "Scanning:"; RCnt; " of"; LastRecord;
    GET #1, RCnt, ReadRec(2)
    IF ReadRec(2).Book = ReadRec(1).Book THEN
      RouteTot = RouteTot + 1
      IF ReadRec(2).ReadFlag = "Y" THEN
        WasRead = True
      END IF
      IF RCnt < CurRecord THEN
        IF WasRead THEN
          ReadCnt = ReadCnt + 1
        ELSE
          ReadSkip = ReadSkip + 1
        END IF
      ELSE
        IF WasRead THEN
          ReadCnt = ReadCnt + 1
        END IF
        IF RCnt = CurRecord THEN
          RouteCur = RouteTot   'RouteTot is Current record at this point
        END IF
      END IF
    END IF
  NEXT
  CLOSE
  
  CLocate 3, 1
  PRINT Blank$
  CLocate 3, 1
  PRINT "Route Total:"; RouteTot
  PRINT "    Current:"; RouteCur
  PRINT "    Skipped:"; ReadSkip
  PRINT "  Remaining:"; RouteTot - RouteCur
  CLocate 8, 1
  PRINT "Press any key to continue.";
  HKey = BiosKey
RETURN

MakeSound:
  QPSound 1000, 1
RETURN

HVertMenu:
  DspBtm = DspTop + NotesRecs - 1
  IF BoxBot > DspBtm THEN
    BoxBot = DspBtm + 1
  END IF

  LastRow = False

  IF DspBtm >= BoxBot THEN      '<---/-- BoxBot sets lowest row to display
    DspBtm = BoxBot - 1         '<-/
  END IF
  WinSize = DspBtm - DspTop + 1

  RightMargin = 28
  RCol = 27

  Offset = False
  IF HChoice = False OR HChoice > NotesRecs THEN
    HChoice = 1
  END IF

  IF HChoice > Offset + WinSize THEN
    Offset = HChoice - WinSize
  END IF

'----- draw border and display items on the screen
  QPrintRC "ÉÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ»", DspTop - 1, LeftMargin, BoxSClr        ' Cnf.MenBox
  FOR x = DspTop TO DspBtm                          'Draw sides, display items
    QPrintRC "º                      º", x, LeftMargin, BoxSClr
  NEXT
  QPrintRC "ÈÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼", BoxBot, LeftMargin, BoxSClr
                                                'Draw the DspBtm of box
  GOSUB VertDisplay

  DO

    PressedKey = BiosKey        'wait for a key press

    SELECT CASE PressedKey      'ASC(Ky$)
    CASE 13, 27                 'Enter or Escape key
      Ky = PressedKey
      EXIT DO
    CASE -72                    'up arrow
      IF HChoice > 1 THEN
        HChoice = HChoice - 1
        IF Row - 1 < DspTop THEN Offset = Offset - 1
      ELSE
        HChoice = NotesRecs
      END IF
    CASE -80                    'down arrow
      IF HChoice < NotesRecs THEN
        HChoice = HChoice + 1
        IF Row + 1 > DspBtm THEN Offset = Offset + 1
      ELSE
        HChoice = 1
      END IF
    CASE -71    'home key
      HChoice = 1
      Offset = False
    CASE -79    'end key
      HChoice = NotesRecs
      Offset = NotesRecs - (DspBtm - DspTop + 1)
    CASE ELSE

    END SELECT
    GOSUB VertDisplay

  LOOP          'loop until they're done

HVertReturn:
RETURN

  '----- Print the HChoices in the window
VertDisplay:
  IF HChoice <= Offset THEN Offset = HChoice - 1
  IF HChoice > Offset + WinSize THEN
    Offset = HChoice - WinSize
  END IF

  CLocate DspTop, LeftMargin + 2
  APrint0 VARPTR(Item$(Offset + 1)), DspBtm - DspTop + 1, 1, MaxLen, -1

  Row = DspTop + (HChoice - Offset) - 1

  IF LastRow > 0 AND Row <> LastRow THEN
    PaintBox0 LastRow, LeftMargin + 1, LastRow, RCol, BoxSClr   '
  END IF
  PaintBox0 Row, LeftMargin + 1, Row, RCol, ScrlClr      'set inverse color
  LastRow = Row

  CLocate Row, LeftMargin + 2

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
'

'SHARED F3Disable
  
  STATIC InsertMode
  
  GOSUB ShowCursor
  
  Original$ = Edit$
  ExitCode = False
  
  MaxLength = LEN(Edit$)        'max length of input string
  Ptr = LEN(RTRIM$(Edit$))
  IF Ptr = MaxLength THEN Ptr = MaxLength - 1
  'added code to just look for function keys with no scrn output
  IF Ptr < False THEN Ptr = False
  
  DO            'main processing loop
    
    LOCATE Row, Col, 1          'update cursor position.
    PRINT Edit$;                'and display the line
    CLocate Row, Col + Ptr      'update cursor position.
    
    PressedKey = BiosKey        'wait for a key press
    
    SELECT CASE PressedKey      'process the key
      '****  Exit keys
    CASE 13     'ENTER - Accept line and, exit
      Edit$ = QPTrim$(Edit$)
      IF LEN(Edit$) = False THEN
        ExitCode = True
      ELSE
        ExitCode = 13
      END IF
      
      EXIT DO
      
    CASE 27     'ESCAPE - Abort operation exit
      ExitCode = EscKey
      Edit$ = Original$
      EXIT DO
      
    CASE -61    'F3
      IF F3Disable THEN
        ExitCode = ABS(PressedKey)
        EXIT DO
      ELSE
        BackLightToggle
      END IF
    CASE -59, -60, -62, -72, -80, -84           'F1 F2 F4 DnArrow UpArrow
      'UpArrow DnArrow F1 F2 F4
      ExitCode = ABS(PressedKey)
      EXIT DO
      
      '****  Cursor keys
    CASE -82    'INSERT - Toggle insert or, overtype mode
      InsertMode = NOT InsertMode
      GOSUB ShowCursor
    CASE -75    'LEFTARROW - Moves cursor left
      IF Ptr THEN Ptr = Ptr - 1
      IF Col + Ptr <= False THEN
        Ptr = Ptr + 1
      END IF
    CASE -77    'RIGHTARROW - Moves cursor right
      IF Ptr < MaxLength - 1 THEN Ptr = Ptr + 1
    CASE -71    'HOME - Moves cursor to first position
      Ptr = 0
    CASE -79    'END - Moves cursor to last position
      Ptr = LEN(RTRIM$(Edit$))
      IF Ptr = MaxLength THEN Ptr = MaxLength - 1
      
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

