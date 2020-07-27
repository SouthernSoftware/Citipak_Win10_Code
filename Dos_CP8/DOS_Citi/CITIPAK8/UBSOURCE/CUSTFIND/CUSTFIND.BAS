DECLARE SUB nCursorOff ()
DEFINT A-Z
DECLARE SUB nBlockClear ()
DECLARE SUB DisplayCFScrn (ScrnName$)
DECLARE SUB ShowSearchThing (BYVAL Row%, BYVAL Col%)
DECLARE SUB WaitForAction ()
DECLARE SUB BlockClear ()
DECLARE SUB LookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, ActiveOnly%)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPStrI$ (Num%)
DECLARE FUNCTION QPStrL$ (Num&)
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB HideCursor ()
DECLARE SUB ShowCursor ()
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE SUB WaitForAction ()
DECLARE SUB SaveScrn (Array())
DECLARE SUB RestScrn (Array())
DECLARE SUB CursorOff ()
DECLARE FUNCTION GetValidateDataPath% (DataPath$)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB MPaintBox (UlRow%, UlCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice, MaxLen%, BoxBot, Ky$, Action, Cnf AS ANY)

  TYPE FLen2
    V AS STRING * 64
  END TYPE

  '$INCLUDE: 'defcnf.bi'
  '$INCLUDE: 'formedit.bi'
  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'setcnf.bi'
  '$INCLUDE: 'custfind.bi'

  CONST False = 0, True = NOT False
  CONST Wheel$ = "|/Ä\"

  DIM SHARED DebugFlag AS INTEGER
  REDIM TempScrn(0)
  REDIM MainScrn(0)

  LOCATE 1, 1, 0
  nBlockClear

  IsOkFlag = GetValidateDataPath%(DataPath$)
  IF NOT IsOkFlag THEN
    GOTO ExitSearch
  END IF

  REDIM CustRec(1) AS PineBluffCustType
  CustRecLen = LEN(CustRec(1))

  LibName$ = "CUSTFIND"
  ScrnName$ = "LOOKUP"

  '--Get the total number of fields from all pages
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1

  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  '--Increment StartEl to next

  DisplayCFScrn "LOOKUP"
  ShowCursor

Top:
  ExitFlag = False
  OKFlag = False
  Action = 1
  MatchCnt = 0

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key   ', 13           'F10
      GOSUB CheckParseSearchText
      IF OKFlag THEN
        ExitFlag = True
      ELSE
        Action = 1
      END IF
    CASE EscKey
      ExitFlag = True
    END SELECT

    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 16
        SELECT CASE Frm(1).MCol
        CASE 23 TO 31
          PressButton F1Key, 16, 23, 31
        CASE 33 TO 44
          PressButton EscKey, 16, 33, 44
        CASE 46 TO 57          '--F10 Save Button
          PressButton F10Key, 16, 46, 57
        END SELECT
      END SELECT                'row
    END IF

    '--Check screen page
  LOOP UNTIL ExitFlag

  IF NOT OKFlag THEN
    HideCursor
    CLS
    END
  END IF
  SaveScrn MainScrn()
  HideCursor
  nBlockClear
  DisplayCFScrn "SHOWSCRH"

  'END

  REDIM MChoice(1 TO 1) AS FLen2

  CustFile = FREEFILE
  OPEN DataPath$ FOR RANDOM ACCESS READ SHARED AS CustFile LEN = CustRecLen
  NumOfRec = LOF(CustFile) \ CustRecLen

  FOR Cnt = 1 TO NumOfRec
    GET #CustFile, Cnt, CustRec(1)
    IF UseF THEN
      IF INSTR(CustRec(1).FName, FSearch$) = 0 THEN
        GOTO SkipEM
      END IF
    END IF
    IF UseL THEN
      IF INSTR(CustRec(1).LName, LSearch$) = 0 THEN
        GOTO SkipEM
      END IF
    END IF
    MatchCnt = MatchCnt + 1
    GOSUB CustLoadEM

SkipEM:
    ShowSearchThing 12, 44
  NEXT
  CLOSE

  IF MatchCnt > 0 THEN
    GOSUB ShowPickList
  END IF
  RestScrn MainScrn()
  GOTO Top

ExitSearch:

END

CustLoadEM:
  REDIM PRESERVE MChoice(1 TO MatchCnt) AS FLen2
  CName$ = QPTrim$(LEFT$(CustRec(1).LName, 20)) + ", " + QPTrim$(LEFT$(CustRec(1).FName, 20))
  LSET MChoice(MatchCnt).V = CName$
  MID$(MChoice(MatchCnt).V, 32) = CustRec(1).Addr1
  MID$(MChoice(MatchCnt).V, 50) = CustRec(1).Phone
  MID$(MChoice(MatchCnt).V, 63) = MKI$(Cnt)

RETURN

CheckParseSearchText:
 UseF = False
 UseL = False

 FSearch$ = QPTrim$(Form$(1, 0))
 LSearch$ = QPTrim$(Form$(2, 0))

 FLen = LEN(FSearch$)
 LLen = LEN(LSearch$)

 IF FLen = 0 AND LLen = 0 THEN
   nCursorOff
   SaveScrn TempScrn()
   DisplayCFScrn "BADSERCH"
   WaitForAction
   RestScrn TempScrn()
 ELSE
   IF FLen THEN
     UseF = True
   END IF
   IF LLen THEN
     UseL = True
   END IF
   OKFlag = True
 END IF

RETURN

ShowPickList:

    SortT MChoice(1), MatchCnt, 0, 64, 0, 64
    MaxLen = 59   'Set menu width to zero
    VAction = 0  '0 means stay in the menu until they select something
    nBlockClear
    IF Choice < 1 THEN
      Choice = 1                'Pre-load choice to highlight
    END IF

    Title$ = SPACE$(MaxLen + 4)
    LSET Title$ = "   Name                          Address                       "
    Col = 9
    Row = 6
    BoxBot = 14               'limit the box length to go no lower than line 20

    LOCATE Row, Col, 0

    DO
      TitleBox BoxBot + 3, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
      QPrintRC "Found:" + STR$(MatchCnt), BoxBot + 4, Col + 2, 15
      QPrintRC Title$, Row - 1, Col, 112
      MPaintBox Row, Col + MaxLen + 4, Row, Col + MaxLen + 5, 8
      VertMenuT2 MChoice(), Choice, MaxLen, BoxBot, Ky$, VAction, Cnf
      SELECT CASE ASC(Ky$)
      CASE 27
        RecNo = 0
        EXIT DO 'choice = 0
      CASE 13
        RecNo = CVI(MID$(MChoice(Choice).V, 63, 2))
      END SELECT

    LOOP UNTIL RecNo > 0
    IF RecNo > 0 THEN
      GOSUB DisplayDetail
    END IF

RETURN

DisplayDetail:
  nCursorOff
  CustFile = FREEFILE
  OPEN DataPath$ FOR RANDOM ACCESS READ SHARED AS CustFile LEN = CustRecLen
  GET #CustFile, RecNo, CustRec(1)
  CLOSE
  nBlockClear
  DisplayCFScrn "CFDETAIL"
  CName$ = QPTrim$(CustRec(1).FName) + " " + QPTrim$(CustRec(1).LName)
  QPrintRC QPTrim$(CName$), 8, 26, -1
  QPrintRC QPTrim$(CustRec(1).Addr1), 10, 26, -1
  QPrintRC QPTrim$(CustRec(1).Addr2), 11, 26, -1
  QPrintRC QPTrim$(CustRec(1).City), 12, 26, -1
  QPrintRC QPTrim$(CustRec(1).State), 12, 52, -1
  QPrintRC QPTrim$(CustRec(1).Zip), 12, 61, -1
  QPrintRC QPTrim$(CustRec(1).SAddr), 14, 26, -1
  WaitForAction
RETURN

SUB DisplayCFScrn (ScrnName$)
  LibFile2Scrn "CUSTFIND", ScrnName$, MonoCode, -1, ErrorCode
END SUB

SUB nBlockClear
  LibFile2Scrn "CUSTFIND", "BAKCLEAR", MonoCode, -1, ErrorCode
END SUB

SUB nCursorOff
LOCATE , , 0
END SUB

SUB ShowSearchThing (BYVAL Row%, BYVAL Col%) STATIC
  WheelPos = WheelPos + 1
  IF WheelPos > 4 THEN WheelPos = 1
  HideCursor
  QPrintRC MID$(Wheel$, WheelPos, 1), Row, Col, -1
  ShowCursor
  'smallpause
END SUB

