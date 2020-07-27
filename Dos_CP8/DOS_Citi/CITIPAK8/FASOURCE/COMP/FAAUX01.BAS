DEFINT A-Z
DECLARE FUNCTION FAAssetLookUp% (TempCode$)
DECLARE SUB AcctPickList (Items() AS ANY, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB FASearch4SerNo (SerialNum$, RecNo&, CLSFlag%, ActiveOnly%)
DECLARE SUB FASearch4Desc (TText$, RecNo&, CLSFlag%, ActiveOnly%)
DECLARE SUB FAShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB FAShowProcessingScrn (RptTitle$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE SUB FASearch4Tag (TNumb$, RecNo&, CLSFlag%, ActiveOnly%)
DECLARE SUB DisplayFAScrn (ScrnName$)
DECLARE SUB ClearScrn ()
DECLARE SUB PrintHelp (H$)
DECLARE SUB BlockClear ()
DECLARE SUB PrintTitle (Title$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION WEnvTest% ()
DECLARE FUNCTION Round# (DblNum#)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB HideCursor ()
DECLARE SUB CursorOff ()
DECLARE SUB TextCursor (MouseFg%, MouseBg%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE SUB SaveScrn (Array())
DECLARE SUB RestScrn (Array())
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE FUNCTION QPValL& (Number$)
DECLARE FUNCTION QPStrI$ (Num%)
DECLARE FUNCTION QPStrL$ (Num&)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
DECLARE SUB FOpenS (FileName$, Handle)
DECLARE SUB FClose (Handle%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FGetA (Handle%, SEG Dest AS ANY, NumBytes%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB WaitForAction ()
DECLARE SUB MPaintBox (UlRow%, UlCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice, MaxLen%, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION Num2Date$ (TheDate%)


CONST False = 0, True = NOT False
CONST BlankInt% = -32767      'Specifies blank integer

  TYPE FLen
    V AS STRING * 42
  END TYPE

  TYPE FLen2
    V AS STRING * 64
  END TYPE

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'newfa.BI'
  '$INCLUDE: 'Qscr.BI'
  '$INCLUDE: 'SetCnf.BI'


  DIM SHARED PctC(1) AS STRING * 3

  STACK 5000

SUB CodeLookUp (RecNo, CLSFlag%, ActiveOnly%)

  REDIM CodeRec(1) AS FAAssetCodeRecType
  CodeRecLen = LEN(CodeRec(1))

  REDIM ScrnArray(0)
  SaveScrn ScrnArray()

  WPos = 1
  
  CodeFileSize& = FileSize(FACodeFile)
  NumCodeRec = CodeFileSize& \ CodeRecLen

  CodeFile = FREEFILE
  OPEN FACodeFile FOR RANDOM AS CodeFile LEN = CodeRecLen

  REDIM MChoice(1 TO 1) AS FLen2

'load up the list of asset codes.
  MatchCnt = 0
  FOR Cnt = 1 TO NumCodeRec
    REDIM PRESERVE MChoice(1 TO Cnt) AS FLen2
    GET CodeFile, Cnt, CodeRec(1)
    LSET MChoice(Cnt).V = QPTrim$(CodeRec(1).AssetCode)
    MID$(MChoice(Cnt).V, 18, 20) = CodeRec(1).AssetDesc
    MID$(MChoice(Cnt).V, 50, 10) = QPTrim$(CodeRec(1).AssetStatus)
    MID$(MChoice(Cnt).V, 61) = MKI$(Cnt)
    FAShowPctComp Cnt, NumCodeRec
  NEXT
  
  'QPrintRC "Sorting. . .  ", 11, 34, -1

  MaxLen = 59 'Set menu width to zero
  Action = 0  '0 means stay in the menu until they select something
  IF Choice = 0 THEN
    Choice = 1                'Pre-load choice to highlight
  ELSEIF Choice > MatchCnt THEN
    Choice = 1                'Pre-load choice to highlight
  END IF
  Title$ = SPACE$(MaxLen + 4)
  LSET Title$ = " Asset Code            Description             Asset Status"
  '--Find max menu width
  '--Center Menu within Screen

  Row = 4
  Col = ((80 - 60) \ 2) - 1

  IF CLSFlag THEN
    Row = 4
    BoxBot = 17               'limit the box length
    BlockClear
  ELSE
    Row = 6
    BoxBot = 14               'limit the box length to go no lower than line 20
    RestScrn ScrnArray()
  END IF

  LOCATE Row, Col, 0

  DO
    TitleBox BoxBot + 3, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
'    QPrintRC "Matched:" + STR$(MatchCnt), BoxBot + 4, Col + 2, 15
    QPrintRC Title$, Row - 1, Col, 112
    MPaintBox Row, Col + MaxLen + 4, Row, Col + MaxLen + 5, 8
    VertMenuT2 MChoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    IF Ky$ = CHR$(27) THEN
      RecNo = -1
      EXIT DO 'choice = 0
    END IF
    RecNo = CVI(MID$(MChoice(Choice).V, 61, 2))
  LOOP UNTIL RecNo > 0
  'ELSE
  '  RecNo& = CVL(MID$(MChoice(1).V, 61, 4))
  'END IF

CodeExitDescSearch:

  RestScrn ScrnArray()
  ERASE ScrnArray, MChoice
  EXIT SUB


END SUB

SUB DisplayFAScrn (ScrnName$)
  LibFile2Scrn "FA", ScrnName$, MonoCode, Attribute%, ErrCode
END SUB

FUNCTION FAAssetLookUp (TempCode$)

  REDIM CodeRec(1) AS FAAssetCodeRecType
  CodeRecLen = LEN(CodeRec(1))
  
  REDIM ScrBuf(0)
  SaveScrn ScrBuf()

  MatchAccts = 0
  Lookup$ = UCASE$(QPTrim$(TempCode$))

  Lookuplen = LEN(Lookup$)

  CodeFileSize& = FileSize(FACodeFile)
  NumCodeRec = CodeFileSize& \ CodeRecLen
  IF NumCodeRec < 1 THEN
    GOTO ExitAssetLookUp
  END IF
  
  CodeFile = FREEFILE
  OPEN FACodeFile FOR RANDOM AS CodeFile LEN = CodeRecLen
  
  REDIM CodeList(1 TO NumCodeRec) AS FLen       'Dim array to hold trans list
  FOR Cnt = 1 TO NumCodeRec     'Read the file
    GET CodeFile, Cnt, CodeRec(1)
    IF LEFT$(CodeRec(1).AssetCode, Lookuplen) = Lookup$ THEN
      MatchAccts = MatchAccts + 1
      LSET CodeList(MatchAccts).V = ""
      CodeList(MatchAccts).V = CodeRec(1).AssetCode
      MID$(CodeList(MatchAccts).V, 8) = CodeRec(1).AssetDesc
      MID$(CodeList(MatchAccts).V, 41) = MKI$(Cnt)
    END IF
  NEXT

  CLOSE

  SELECT CASE MatchAccts
  CASE 0
    Ok = MsgBox%("FA", "BADASSET")
    Action = 1
  CASE 1
    CodeRecNo = CVI(RIGHT$(CodeList(1).V, 2))
    Choice = 1
  CASE ELSE
    Row = 5
    Col = 15
    LOCATE Row, Col
    Action = 0
    MaxLen = 39
    BoxBot = 18
    ListDescTxt$ = SPACE$(MaxLen + 4)
    LSET ListDescTxt$ = " Asset Code     Description"
    QPrintRC ListDescTxt$, Row - 1, Col, 112       'Cnf.ActivCh
    DO          'call vertmenu to display matching recs to choose from
      AcctPickList CodeList(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        EXIT DO
      END IF
      IF Choice > 0 THEN
        CodeRecNo = CVI(RIGHT$(CodeList(Choice).V, 2))
        EXIT DO
      END IF
    LOOP
  END SELECT

  IF CodeRecNo > 0 THEN
    TempCode$ = LEFT$(QPTrim$(CodeList(Choice).V), 4)
    FAAssetLookUp% = CodeRecNo
  ELSE
    FAAssetLookUp% = 0
  END IF

  RestScrn ScrBuf()
  ERASE CodeList

ExitAssetLookUp:
  ERASE ScrBuf

END FUNCTION

SUB FALookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, ActiveOnly%)
  
  'STATIC TNumb$, AcctNum&, SerialNum$, SAddr$, LNum$
  
  REDIM FAItemRec(1) AS FAItemRecType
  FAItemRecLen = LEN(FAItemRec(1))
  
  TNumb$ = ""
  AcctNum& = 0
  SerialNum$ = ""

  IF INSTR(COMMAND$, "DEBUG") THEN
    DebugFlag = True
  END IF
  
  LScrn = 1
  
  CursorOff
  
  REDIM ScrnArray(0)
  REDIM ScrnArray2(0)
  
  SaveScrn ScrnArray()
  
  REDIM LText$(3)
  
  MScrn = 3
  
  LText$(1) = "     Tag Number:"
  LText$(2) = "    Description:"
  LText$(3) = "  Serial Number:"
  
  LibName$ = "FA"
  ScrnName$ = "LUPACCT"
  
  '--Initialize the form name array
  '--Get the total number of fields from all pages
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  '--Clear all fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT
  Text$ = " Look-Up"
  TextLen = LEN(Text$)
  TCol = ((80 - TextLen) \ 2)
  DisplayFAScrn ScrnName$
  
  QPrintRC Text$, 8, TCol, -1
  
  GOSUB DisplayLookupText
  
  ShowCursor
  
  Action = 1
  FirstTime = True
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE -68, 13                'F10Key    Proceed with look up
      CursorOff
      SELECT CASE LScrn
      CASE 1    'Item Number
        TNumb$ = LEFT$(QPTrim$(Form$(0, 0)), 10)
        IF LEN(TNumb$) = 0 THEN
          TNumb$ = SPACE$(10)
        END IF
        SaveScrn ScrnArray2()
        RestScrn ScrnArray()
        FASearch4Tag TNumb$, RecNo&, CLSFlag, ActiveOnly
        IF RecNo& > 0 THEN
          OKFlag = True
        ELSEIF RecNo& = 0 THEN
          Ok = MsgBox%("FA.QSL", "NOMATCH")
        END IF
        RestScrn ScrnArray2()
        Action = 1
      
      CASE 2    'Description
        TText$ = LEFT$(QPTrim$(Form$(0, 0)), 10)
        IF LEN(TText$) = 0 THEN
          TText$ = SPACE$(10)
        END IF
        SaveScrn ScrnArray2()
        RestScrn ScrnArray()
        FASearch4Desc TText$, RecNo&, CLSFlag, ActiveOnly
        IF RecNo& > 0 THEN
          OKFlag = True
        ELSEIF RecNo& = 0 THEN
          Ok = MsgBox%("FA.QSL", "NOMATCH")
        END IF
        RestScrn ScrnArray2()
        Action = 1
      CASE 3    'Serial Number
        SerialNum$ = QPTrim$(Form$(0, 0))
        IF LEN(SerialNum$) = 0 THEN
          Ok = MsgBox%("FA.QSL", "BADSEARH")
          Action = 1
          Frm(1).FldNo = 1
        ELSE
          SaveScrn ScrnArray2()
          RestScrn ScrnArray()
          FASearch4SerNo SerialNum$, RecNo&, CLSFlag, ActiveOnly
          IF RecNo& > 0 THEN
            OKFlag = True
          ELSEIF RecNo& = 0 THEN
            Ok = MsgBox%("FA.QSL", "NOMATCH")
          END IF
          RestScrn ScrnArray2()
          Action = 1
        END IF
      END SELECT
    CASE -65    'F7Key
      IF LScrn < MScrn THEN
        LScrn = LScrn + 1
      ELSE
        LScrn = 1
      END IF
      LSET Form$(1, 0) = ""
      Action = 1
      FirstTime = True
      SaveField 0, Form$(), Fld(), BadField
      GOSUB DisplayLookupText
    CASE 27
      RecNo& = 0
      ExitFlag = True
    END SELECT
    
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 16
        SELECT CASE Frm(1).MCol
        CASE 22 TO 33           'ESC Cancel button
          PressButton 27, 16, 22, 33
        CASE 35 TO 45           'F7 Toggle Choice
          PressButton -65, 16, 35, 45
        CASE 47 TO 59           'F10 Save Button
          PressButton -68, 16, 47, 59
        END SELECT
      END SELECT
    END IF
    
  LOOP UNTIL ExitFlag OR OKFlag
  RestScrn ScrnArray()

  ERASE Frm, Form$, Fld
  ERASE ScrnArray, ScrnArray2
  ERASE LText$
  
  EXIT SUB
  
DisplayLookupText:
  QPrintRC LText$(LScrn), 12, 15, -1
RETURN
  
  
END SUB

SUB FASearch4Desc (TText$, RecNo&, CLSFlag%, ActiveOnly%)
  
  REDIM ScrnArray(0)
  SaveScrn ScrnArray()
  
  WPos = 1
  
  FAShowProcessingScrn "Searching Item Descriptions."
  
  REDIM FAItemRec(1) AS FAItemRecType
  FAItemRecLen = LEN(FAItemRec(1))

  ItemFileSize& = FileSize(FAItemFile)
  NumItemRecs = ItemFileSize& \ FAItemRecLen

  REDIM RecBuff(1 TO 1) AS LONG
  
  FOpenS FAItemFile, C1Handle 'open data file
  'Find matching record

  MatchCnt = 0
  FOR Cnt = 1 TO NumItemRecs
    FGetRTA C1Handle, FAItemRec(1), CLNG(Cnt), FAItemRecLen
    GOSUB CheckEM2
    FAShowPctComp Cnt, NumItemRecs
  NEXT
  
  IF MatchCnt = 0 THEN
    RecNo& = 0
    FClose C1Handle
    GOTO ExitDescSearch
  END IF
  IF NOT DebugFlag THEN
    FreeMem& = FRE(-1)
    IF FreeMem& >= 65536 THEN
      FreeMem& = 65536
    END IF
    MemNeeded& = MatchCnt * 64&
    IF MemNeeded& > FreeMem& THEN
      FClose C1Handle
      QPrintRC "Matched:>" + QPStrI(MatchCnt), 24, 1, 15
      QPrintRC " Needed: " + QPStrL(MemNeeded&) + "  Free: " + QPStrL(FreeMem&), 25, 1, 15
      RecNo& = -1
      WaitForAction
      GOTO ExitDescSearch
    END IF
  END IF
  REDIM MChoice(1 TO MatchCnt) AS FLen2
  
  FOR Cnt = 1 TO MatchCnt
    FGetRTA C1Handle, FAItemRec(1), RecBuff(Cnt), FAItemRecLen
    TItemTag$ = QPTrim$(FAItemRec(1).ITEMTAG)
    LSET MChoice(Cnt).V = TItemTag$
    MID$(MChoice(Cnt).V, 15, 15) = " " + FAItemRec(1).SERIALNO
    MID$(MChoice(Cnt).V, 31, 25) = " " + FAItemRec(1).IDESC1
    MID$(MChoice(Cnt).V, 61) = MKL$(RecBuff(Cnt))
  NEXT
  
  FClose C1Handle
  
  IF MatchCnt = 0 THEN
    RecNo& = 0
  ELSEIF MatchCnt > 1 THEN
    QPrintRC "Sorting. . .  ", 11, 34, -1
    
    MaxLen = 59 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something
    IF Choice = 0 THEN
      Choice = 1                'Pre-load choice to highlight
    ELSEIF Choice > MatchCnt THEN
      Choice = 1                'Pre-load choice to highlight
    END IF
    Title$ = SPACE$(MaxLen + 4)
    LSET Title$ = " Tag#            Serial#           Description Line 1"
    '--Find max menu width
    
    '--Center Menu within Screen
    
    Row = 4
    Col = ((80 - 60) \ 2) - 1
    
    IF CLSFlag THEN
      Row = 4
      BoxBot = 17               'limit the box length
      BlockClear
    ELSE
      Row = 6
      BoxBot = 14               'limit the box length to go no lower than line 20
      RestScrn ScrnArray()
    END IF
    
    LOCATE Row, Col, 0
    
    DO
      TitleBox BoxBot + 3, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
      QPrintRC "Matched:" + STR$(MatchCnt), BoxBot + 4, Col + 2, 15
      QPrintRC Title$, Row - 1, Col, 112
      MPaintBox Row, Col + MaxLen + 4, Row, Col + MaxLen + 5, 8
      VertMenuT2 MChoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        RecNo& = -1
        EXIT DO 'choice = 0
      END IF
      RecNo& = CVL(MID$(MChoice(Choice).V, 61, 4))
    LOOP UNTIL RecNo& > 0
  ELSE
    RecNo& = CVL(MID$(MChoice(1).V, 61, 4))
  END IF
  
ExitDescSearch:
  
  RestScrn ScrnArray()
  ERASE ScrnArray, RecBuff, MChoice
  EXIT SUB
  
CheckEM2:
    IF INSTR(FAItemRec(1).IDESC1, TText$) > 0 THEN
      MatchCnt = MatchCnt + 1
      REDIM PRESERVE RecBuff(1 TO MatchCnt) AS LONG
      RecBuff(MatchCnt) = Cnt
      GOTO DoneCheckEM2
    END IF
    IF INSTR(FAItemRec(1).IDESC2, TText$) > 0 THEN
      MatchCnt = MatchCnt + 1
      REDIM PRESERVE RecBuff(1 TO MatchCnt) AS LONG
      RecBuff(MatchCnt) = Cnt
      GOTO DoneCheckEM2
    END IF
    'IF INSTR(FAItemRec(1).IDESC3, TText$) > 0 THEN
    '  MatchCnt = MatchCnt + 1
    '  REDIM PRESERVE RecBuff(1 TO MatchCnt) AS LONG
    '  RecBuff(MatchCnt) = Cnt
    'END IF
DoneCheckEM2:

RETURN
  
END SUB

SUB FASearch4SerNo (SerialNum$, RecNo&, CLSFlag%, ActiveOnly%)

  REDIM ScrnArray(0)
  SaveScrn ScrnArray()

  WPos = 1

  FAShowProcessingScrn "Searching Serial Numbers."

  REDIM FAItemRec(1) AS FAItemRecType
  FAItemRecLen = LEN(FAItemRec(1))

  ItemFileSize& = FileSize(FAItemFile)
  NumItemRecs = ItemFileSize& \ FAItemRecLen

  REDIM RecBuff(1 TO 1) AS LONG

  FOpenS FAItemFile, C1Handle 'open data file
  'Find matching record

  MatchCnt = 0
  FOR Cnt = 1 TO NumItemRecs
    FGetRTA C1Handle, FAItemRec(1), CLNG(Cnt), FAItemRecLen
    IF INSTR(FAItemRec(1).SERIALNO, SerialNum$) > 0 THEN
      MatchCnt = MatchCnt + 1
      REDIM PRESERVE RecBuff(1 TO MatchCnt) AS LONG
      RecBuff(MatchCnt) = Cnt
    END IF
    FAShowPctComp Cnt, NumItemRecs
  NEXT

  IF MatchCnt = 0 THEN
    RecNo& = 0
    FClose C1Handle
    GOTO ExitSerialSearch
  END IF
  IF NOT DebugFlag THEN
    FreeMem& = FRE(-1)
    IF FreeMem& >= 65536 THEN
      FreeMem& = 65536
    END IF
    MemNeeded& = MatchCnt * 64&
    IF MemNeeded& > FreeMem& THEN
      FClose C1Handle
      QPrintRC "Matched:>" + QPStrI(MatchCnt), 24, 1, 15
      QPrintRC " Needed: " + QPStrL(MemNeeded&) + "  Free: " + QPStrL(FreeMem&), 25, 1, 15
      RecNo& = -1
      WaitForAction
      GOTO ExitSerialSearch
    END IF
  END IF

  REDIM MChoice(1 TO MatchCnt) AS FLen2

  FOR Cnt = 1 TO MatchCnt
    'ShowSearchWheel 12, 44
    FGetRTA C1Handle, FAItemRec(1), RecBuff(Cnt), FAItemRecLen
    TItemTag$ = QPTrim$(FAItemRec(1).ITEMTAG)
    LSET MChoice(Cnt).V = TItemTag$
    MID$(MChoice(Cnt).V, 15, 15) = " " + FAItemRec(1).SERIALNO
    MID$(MChoice(Cnt).V, 31, 25) = " " + FAItemRec(1).IDESC1
    MID$(MChoice(Cnt).V, 61) = MKL$(RecBuff(Cnt))
  NEXT

  FClose C1Handle



  FClose C1Handle

  IF MatchCnt = 0 THEN
    RecNo& = 0
  ELSEIF MatchCnt > 1 THEN
    QPrintRC "Sorting. . .  ", 11, 34, -1
    MaxLen = 59 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something
    IF Choice = 0 THEN
      Choice = 1                'Pre-load choice to highlight
    ELSEIF Choice > MatchCnt THEN
      Choice = 1                'Pre-load choice to highlight
    END IF
    Title$ = SPACE$(MaxLen + 4)
    LSET Title$ = " Tag#            Serial#           Description Line 1"
    '--Find max menu width

    '--Center Menu within Screen

    Row = 4
    Col = ((80 - 60) \ 2) - 1

    IF CLSFlag THEN
      Row = 4
      BoxBot = 17               'limit the box length
      BlockClear
    ELSE
      Row = 6
      BoxBot = 14               'limit the box length to go no lower than line 20
      RestScrn ScrnArray()
    END IF

    LOCATE Row, Col, 0

    DO
      TitleBox BoxBot + 3, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
      QPrintRC "Matched:" + STR$(MatchCnt), BoxBot + 4, Col + 2, 15
      QPrintRC Title$, Row - 1, Col, 112
      MPaintBox Row, Col + MaxLen + 4, Row, Col + MaxLen + 5, 8
      VertMenuT2 MChoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        RecNo& = -1
        EXIT DO 'choice = 0
      END IF
      RecNo& = CVL(MID$(MChoice(Choice).V, 61, 4))
    LOOP UNTIL RecNo& > 0
  ELSE
    RecNo& = CVL(MID$(MChoice(1).V, 61, 4))
  END IF

ExitSerialSearch:

  RestScrn ScrnArray()
  ERASE ScrnArray, RecBuff, MChoice
  EXIT SUB


END SUB

SUB FASearch4Tag (TNumb$, RecNo&, CLSFlag%, ActiveOnly%)
  
  REDIM ScrnArray(0)
  SaveScrn ScrnArray()
  
  FAShowProcessingScrn "Searching Tag Numbers."
  
  REDIM FAItemRec(1) AS FAItemRecType
  FAItemRecLen = LEN(FAItemRec(1))
  
  ItemFileSize& = FileSize(FAItemFile)
  NumItemRecs = ItemFileSize& \ FAItemRecLen
  
  REDIM RecBuff(1 TO 1) AS LONG
  
  FOpenS FAItemFile, C1Handle 'open data file
  'Find matching record
  
  MatchCnt = 0
  FOR Cnt = 1 TO NumItemRecs
    FGetRTA C1Handle, FAItemRec(1), CLNG(Cnt), FAItemRecLen
    IF INSTR(FAItemRec(1).ITEMTAG, TNumb$) > 0 THEN
      MatchCnt = MatchCnt + 1
      REDIM PRESERVE RecBuff(1 TO MatchCnt) AS LONG
      RecBuff(MatchCnt) = Cnt
    END IF
    FAShowPctComp Cnt, NumItemRecs
  NEXT
  
  IF MatchCnt = 0 THEN
    RecNo& = 0
    FClose C1Handle
    GOTO ExitMeterSearch
  END IF
  IF NOT DebugFlag THEN
    FreeMem& = FRE(-1)
    IF FreeMem& >= 65536 THEN
      FreeMem& = 65536
    END IF
    MemNeeded& = MatchCnt * 64&
    IF MemNeeded& > FreeMem& THEN
      FClose C1Handle
      QPrintRC "Matched:>" + QPStrI(MatchCnt), 24, 1, 15
      QPrintRC " Needed: " + QPStrL(MemNeeded&) + "  Free: " + QPStrL(FreeMem&), 25, 1, 15
      RecNo& = -1
      WaitForAction
      GOTO ExitMeterSearch
    END IF
  END IF

  REDIM MChoice(1 TO MatchCnt) AS FLen2

  FOR Cnt = 1 TO MatchCnt
    'ShowSearchWheel 12, 44
    FGetRTA C1Handle, FAItemRec(1), RecBuff(Cnt), FAItemRecLen
    TItemTag$ = QPTrim$(FAItemRec(1).ITEMTAG)
    LSET MChoice(Cnt).V = TItemTag$
    'Book$ = FAItemRec(1).Book + "-" + FAItemRec(1).SEQNUMB
    MID$(MChoice(Cnt).V, 15, 15) = " " + FAItemRec(1).SERIALNO
    MID$(MChoice(Cnt).V, 31, 25) = " " + FAItemRec(1).IDESC1
    MID$(MChoice(Cnt).V, 61) = MKL$(RecBuff(Cnt))
  NEXT
  
  FClose C1Handle
  
  IF MatchCnt = 0 THEN
    RecNo& = 0
  ELSEIF MatchCnt > 1 THEN
    QPrintRC "Sorting. . .  ", 11, 34, -1
    MaxLen = 59 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something
    IF Choice = 0 THEN
      Choice = 1                'Pre-load choice to highlight
    ELSEIF Choice > MatchCnt THEN
      Choice = 1                'Pre-load choice to highlight
    END IF
    Title$ = SPACE$(MaxLen + 4)
    LSET Title$ = " Tag#            Serial#           Description Line 1"
    '--Find max menu width
    Row = 4
    Col = ((80 - 60) \ 2) - 1
    
    IF CLSFlag THEN
      Row = 4
      BoxBot = 17               'limit the box length
      BlockClear
    ELSE
      Row = 6
      BoxBot = 14               'limit the box length to go no lower than line 20
      RestScrn ScrnArray()
    END IF
    
    LOCATE Row, Col, 0
    
    DO
      TitleBox BoxBot + 3, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
      QPrintRC "Matched:" + STR$(MatchCnt), BoxBot + 4, Col + 2, 15
      QPrintRC Title$, Row - 1, Col, 112
      MPaintBox Row, Col + MaxLen + 4, Row, Col + MaxLen + 5, 8
      VertMenuT2 MChoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        RecNo& = -1
        EXIT DO 'choice = 0
      END IF
      RecNo& = CVL(MID$(MChoice(Choice).V, 61, 4))
    LOOP UNTIL RecNo& > 0
  ELSE
    RecNo& = CVL(MID$(MChoice(1).V, 61, 4))
  END IF
  
ExitMeterSearch:
  
  RestScrn ScrnArray()
  ERASE ScrnArray, RecBuff, MChoice
  EXIT SUB

END SUB

SUB FAShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%) STATIC
  RSET PctC(1) = QPStrI$(INT((RecNo / NumOfRecs) * 100))
  'HideCursor
  QPrintRC PctC(1), 13, 40, Cnf.HiLite
  'ShowCursor
  '  QPrintRC STR$(FRE("")), 25, 1, Cnf.HiLite
END SUB

SUB FAShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&) STATIC
  RSET PctC(1) = QPStrL$(INT((RecNo& / NumOfRecs&) * 100))
  'HideCursor
  QPrintRC PctC(1), 13, 40, Cnf.HiLite
  'ShowCursor
  '  QPrintRC STR$(FRE("")), 25, 1, Cnf.HiLite
END SUB

SUB FAShowProcessingScrn (RptTitle$)
  TitleRow = 9
  TitleCol = 40 - (LEN(RptTitle$) \ 2) + 1
  CursorOff
  BlockClear
  DisplayFAScrn "PRORPT"
  HideCursor
  QPrintRC RptTitle$, TitleRow, TitleCol, 126
  QPrintRC "Processing:    % Completed.", 13, 28, Cnf.HiLite
  ShowCursor
END SUB

