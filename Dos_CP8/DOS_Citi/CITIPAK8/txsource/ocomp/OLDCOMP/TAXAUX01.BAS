DEFINT A-Z
DECLARE FUNCTION GetTaxCustCnt& ()
DECLARE FUNCTION IsCustDeleted% (AcctNum&)
DECLARE SUB ClearBack ()
DECLARE SUB Search4Cust (SEARCH$, RecNo&, CLSFlag%, ActiveOnly%)
DECLARE SUB ShowSearchPCT (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE SUB DisplayTaxScrn (ScrnName$)
DECLARE FUNCTION PromptSaveData% ()
DECLARE FUNCTION Monitor% ()
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)

'$INCLUDE: 'DefCnf.BI'
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE SUB ShowCursor ()
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)
DECLARE SUB HideCursor ()
DECLARE SUB CursorOff ()
DECLARE FUNCTION PromptSaveData% ()
DECLARE FUNCTION QPStrL$ (Num&)
DECLARE FUNCTION QPStrI$ (Num%)
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE FUNCTION QPValL& (Number$)
DECLARE SUB FOpenS (FileName$, Handle)
DECLARE SUB FClose (Handle%)
DECLARE SUB FGetA (Handle%, SEG Element AS ANY, NumBytes AS ANY)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE FUNCTION FLof& (Handle%)
DECLARE SUB MPaintBox (UlRow%, UlCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice, MaxLen%, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB RestScrn (Array%())
DECLARE SUB SaveScrn (Array%())

  CONST False = 0, True = NOT False

  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'

  '$INCLUDE: 'TaxCust.BI'

  TYPE FLen2
    V AS STRING * 64
  END TYPE

  DIM SHARED PctC(1) AS STRING * 4

SUB ClearBack
  LibFile2Scrn "TAX", "BAKCLEAR", MonoCode%, Attribute%, ErrorCode%
END SUB

SUB ClearScrn STATIC
  WazzWind 1, 1, 25, 80, 7, 0, 0
END SUB

SUB CursorOff STATIC
  LOCATE , , 0
END SUB

SUB DisplayTaxScrn (ScrnName$)
  LibFile2Scrn "TAX", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  IF ErrorCode% <> 0 THEN
    PRINT "Screen Error: "; ScrnName$
    END
  END IF
END SUB

FUNCTION GetTaxCustCnt&

  REDIM TAXCust(1)  AS TaxCustType
  TaxCustLen = LEN(TAXCust(1))              'Length of Cust Record Structure

  TAXFile = FREEFILE
  OPEN "TAXCUST.DAT" FOR RANDOM SHARED AS TAXFile LEN = TaxCustLen
  NumOfRecs& = LOF(TAXFile) \ TaxCustLen
  CLOSE TAXFile

  ERASE TAXCust

  GetTaxCustCnt& = NumOfRecs&

END FUNCTION

FUNCTION IsCustDeleted (AcctNum&)
  IsCustDeleted = False

  REDIM TAXCust(1)  AS TaxCustType
  TaxCustLen = LEN(TAXCust(1))              'Length of Cust Record Structure
  TAXFile = FREEFILE
  OPEN "TAXCUST.DAT" FOR RANDOM SHARED AS TAXFile LEN = TaxCustLen
  GET TAXFile, AcctNum&, TAXCust(1)
  CLOSE TAXFile

  IF TAXCust(1).Deleted <> 0 THEN
    IsCustDeleted = True
  END IF

  ERASE TAXCust

END FUNCTION

SUB LookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, ActiveOnly%)
  
'  STATIC SNAME$, AcctNum&, MeterNum$, SAddr$, LNum$

  REDIM TAXCust(1) AS TaxCustType
  TaxCustLen = LEN(TAXCust(1))
  
  SNAME$ = ""
  AcctNum& = 0
  
'  SELECT CASE QPValI(UBSetUpRec(1).DefLook)
'  CASE 1
'    LScrn = 1
'  CASE 2
'    LScrn = 2
'  CASE 3
'    LScrn = 3
'  CASE 4
'    LScrn = 4
'  CASE 5
'    LScrn = 5
'  CASE ELSE
'    LScrn = DefaultLook
'  END SELECT

  LScrn = 2

  CursorOff
  
  REDIM ScrnArray(0)
  REDIM ScrnArray2(0)
  
  SaveScrn ScrnArray()
  
  REDIM LText(1 TO 2) AS STRING * 17
  
  MScrn = 2
  
  LText(1) = " Account Number:"
  LText(2) = "    Search Name:"
  
  LibName$ = "TAX"
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
  Text$ = Text$ + " Look-Up"
  TextLen = LEN(Text$)
  TCol = ((80 - TextLen) \ 2)
  DisplayTaxScrn ScrnName$
  
  QPrintRC Text$, 8, TCol, -1
  
  GOSUB DisplayLookupText
  
  ShowCursor
  
  Action = 1
  FirstTime = True
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = NOT FirstTime
      SELECT CASE LScrn
      CASE 1
        IF AcctNum& > 0 THEN
          Form$(1, 0) = QPStrL$(AcctNum&)
        END IF
      CASE 2
        Form$(1, 0) = SNAME$
      END SELECT
      Action = 1
    END IF
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE -68, 13                'F10Key    Proceed with look up
      CursorOff
      SELECT CASE LScrn
      CASE 1    'account lookup

        AcctNum& = QPValL(Form$(1, 0))

        IF AcctNum& < 1 OR AcctNum& > GetTaxCustCnt& THEN
          Ok = MsgBox%("TAX.QSL", "BADACCTN")

        ELSEIF IsCustDeleted(AcctNum&) THEN
          Ok = MsgBox%("TAX.QSL", "DELACCTN")

        ELSE
          RecNo& = AcctNum&
          OKFlag = True
        END IF
        Action = 1

      CASE 2    'Name lookup
        SNAME$ = LEFT$(QPTrim$(Form$(0, 0)), 10)
        IF LEN(SNAME$) = 0 THEN
          SNAME$ = SPACE$(10)
        END IF
        SaveScrn ScrnArray2()
        RestScrn ScrnArray()
        Search4Cust SNAME$, RecNo&, CLSFlag, ActiveOnly
        IF RecNo& > 0 THEN
          OKFlag = True
        ELSEIF RecNo& = 0 THEN
          Ok = MsgBox%("TAX.QSL", "NOMATCH")
        END IF
        RestScrn ScrnArray2()
        Action = 1
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

  ERASE TAXCust, ScrnArray, ScrnArray2
  ERASE Frm, Form$, Fld, LText

EXIT SUB
  
DisplayLookupText:
  QPrintRC LText(LScrn), 12, 15, -1
RETURN
  
END SUB

FUNCTION PromptSaveData%
  
  REDIM TempScrn(0)
  SaveScrn TempScrn()
  
  LibName$ = "TAX"
  SaveFlag = 2
  
  FormName$ = "SAVE1ST"
  NumFlds = LibNumberOfFields(LibName$, FormName$)
  
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)       'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo               'DIM the field information array
  StartEl = 0   'Load first form at array start
  LibGetFldDef LibName$, FormName$, StartEl, Fld(), Form$(), ErrCode
  
  
  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1
  
  '----- Setup TYPE for setting and reading form editing information.
  Frm(1).FldNo = 1              'Start editing on field #1
  Frm(1).InsStat = False        'Set insert state (True = Insert on)
  Frm(1).StartEl = 0            'Set form starting element to 0 and
  
  DisplayTaxScrn FormName$
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    SELECT CASE Frm(1).KeyCode
    CASE F0Key
      SaveFlag = True
    CASE EscKey
      SaveFlag = 1
    CASE 88, 120                'X Key
      SaveFlag = False
    END SELECT
    
  LOOP WHILE SaveFlag = 2       'proper key not set
  
  PromptSaveData = SaveFlag
  CursorOff
  
  RestScrn TempScrn()
  
  ERASE TempScrn, Form$, Fld, Frm
  
END FUNCTION

SUB Search4Cust (SEARCH$, RecNo&, CLSFlag%, ActiveOnly%)

  'LocationFlag%
  STATIC Choice, LastSEARCH$
  Acct$ = SPACE$(5)

  IF LastSEARCH$ <> SEARCH$ THEN
    LastSEARCH$ = SEARCH$
    Choice = 1
  END IF

  REDIM ScrnArray(0)
  SaveScrn ScrnArray()

  WazzWind 10, 22, 14, 58, 10, 2, True
  QPrintRC "Searching:    % Completed.", 12, 28, 14
  
  CustBlock = 64

  REDIM MChoice(1 TO 1) AS FLen2
  REDIM TAXCust(1 TO CustBlock) AS TaxCustType

  TaxCustLen = LEN(TAXCust(1))
 
  SearchLen = LEN(SEARCH$)
  Match = False

  FOpenS "TAXCUST.DAT", TAXFile     'open data file

  NumOfCust& = FLof&(TAXFile) / TaxCustLen
  NumChunks& = NumOfCust& / CustBlock

  OddRecs& = NumOfCust& MOD CustBlock

  BlockSize& = (0& + TaxCustLen) * CustBlock
  '            ^^^^^ stops an overflow error

  'Find matching record
  FOR CCnt& = 1 TO NumChunks&
    FGetA TAXFile, TAXCust(1), BlockSize&
    FOR RecCnt = 1 TO CustBlock
      WhatRec& = ((CCnt& - 1) * CustBlock) + RecCnt
      UBSearchN$ = LEFT$(TAXCust(RecCnt).SNAME, SearchLen)
      IF (SEARCH$ = UBSearchN$) THEN
        GOSUB CustLoadEM2
      END IF

DelSkip2:
      ShowSearchPCT WhatRec&, NumOfCust&
    NEXT
  NEXT

  IF OddRecs& > 0 THEN
    NextRec& = (NumChunks& * CustBlock) + 1
    RecCnt = 1
    FOR CCnt& = NextRec& TO NumOfCust&
      FGetRTA TAXFile, TAXCust(1), CCnt&, TaxCustLen
      UBSearchN$ = LEFT$(TAXCust(RecCnt).SNAME, SearchLen)
      WhatRec& = CCnt&
      IF (SEARCH$ = UBSearchN$) THEN
        GOSUB CustLoadEM2
      END IF
DelSkip3:
      ShowSearchPCT WhatRec&, NumOfCust&
    NEXT
  END IF
  
    FClose TAXFile

    IF DCnt = 0 THEN
      RecNo& = 0
      GOTO ExitSearch2
    ELSE

    SortT MChoice(1), DCnt, Direction%, LEN(MChoice(1).V), 0, 18

    MaxLen = 59 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something

    IF Choice < 1 THEN
      Choice = 1                'Pre-load choice to highlight
    END IF

    Title$ = SPACE$(MaxLen + 4)
    LSET Title$ = "  Last Name          First               City         Acct No"
    '--Find max menu width

    '--Center Menu within Screen

    Row = 4
    Col = ((80 - 60) \ 2) - 1

    IF CLSFlag THEN
      Row = 4
      BoxBot = 17               'limit the box length
      ClearBack
    ELSE
      Row = 6
      BoxBot = 14               'limit the box length to go no lower than line 20
      RestScrn ScrnArray()
    END IF

    LOCATE Row, Col, 0

    DO
      TitleBox BoxBot + 3, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
      QPrintRC "Matched:" + STR$(DCnt), BoxBot + 4, Col + 2, 15
      QPrintRC Title$, Row - 1, Col, 112
      MPaintBox Row, Col + MaxLen + 4, Row, Col + MaxLen + 5, 8
      VertMenuT2 MChoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        RecNo& = -1
        EXIT DO 'choice = 0
      END IF
      RecNo& = CVL(MID$(MChoice(Choice).V, 61, 4))
    LOOP UNTIL RecNo& > 0
  END IF

ExitSearch2:
  RestScrn ScrnArray()

  ERASE ScrnArray, MChoice, TAXCust

EXIT SUB

CustLoadEM2:
  
  DCnt = DCnt + 1
  REDIM PRESERVE MChoice(1 TO DCnt) AS FLen2
  RSET Acct$ = QPTrim$(STR$(WhatRec&))
  'LSET MChoice(DCnt).V = Acct$
  LSET MChoice(DCnt).V = LEFT$(QPTrim$(TAXCust(RecCnt).LName), 18)
  MID$(MChoice(DCnt).V, 20, 18) = LEFT$(TAXCust(RecCnt).FNAME, 18)
  MID$(MChoice(DCnt).V, 40, 11) = TAXCust(RecCnt).City
  MID$(MChoice(DCnt).V, 54, 5) = Acct$
  'MID$(MChoice(DCnt).V, 59, 1) = TAXCust(RecCnt).STATUS
  MID$(MChoice(DCnt).V, 61) = MKL$(WhatRec&)
RETURN


END SUB

SUB ShowSearchPCT (BYVAL RecNo&, BYVAL NumOfRecs&) STATIC
  RSET PctC(1) = QPStrI$(INT((RecNo& / NumOfRecs&) * 100))
  HideCursor
  QPrintRC PctC(1), 12, 38, 15
  ShowCursor
'  ZZ = 0
'  FOR ZZZ = 1 TO 2000
'    ZZ = ZZZ
'  NEXT
END SUB

