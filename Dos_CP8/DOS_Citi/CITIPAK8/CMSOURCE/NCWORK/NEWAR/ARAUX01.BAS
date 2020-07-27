DEFINT A-Z
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE FUNCTION GetARCustCnt& ()
DECLARE FUNCTION Round# (N#)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION Unique$ (Path$)
DECLARE FUNCTION OK2UPDateCust% ()
DECLARE FUNCTION GetARCustCnt& ()
DECLARE FUNCTION IsCustDeleted% (AcctNum&)
DECLARE SUB ClearBack ()
DECLARE SUB Search4Cust (Search$, RecNo&, ChkBalFlag%, CLSFlag%, SSNFlag%)
DECLARE SUB ShowSearchPCT (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE SUB DisplayARScrn (ScrnName$)
DECLARE FUNCTION PromptSaveData% ()
DECLARE FUNCTION Monitor% ()
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB DelPropAbstract (PropRecs() AS LONG, WhatProp%, CustRec&)
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION FUsing$ (Number$, Image$)

DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE SUB ShowCursor ()
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB HideCursor ()
DECLARE SUB CursorOff ()
DECLARE FUNCTION PromptSaveData% ()
DECLARE FUNCTION QPStrL$ (Num&)
DECLARE FUNCTION QPStrI$ (Num%)
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE FUNCTION QPValL& (Number$)
DECLARE SUB FOpenS (FileName$, Handle)
DECLARE SUB FClose (Handle%)
DECLARE SUB FGETA (Handle%, SEG Element AS ANY, NumBytes AS ANY)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE FUNCTION FLof& (Handle%)
DECLARE SUB MPaintBox (UlRow%, UlCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice, MaxLen%, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB RestScrn (Array%())
DECLARE SUB SaveScrn (Array%())
DECLARE FUNCTION DoesCustOwe% (ARCustRec AS ANY)
DECLARE FUNCTION GetCustName$ (CustRec&)
DECLARE FUNCTION ParseBillNum$ (Text$)


  CONST False = 0, True = NOT False

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ARV95.BI'
  
  TYPE FLen2
    V AS STRING * 64
  END TYPE
  
  TYPE SortStruct
    who AS STRING * 14
    RecNum AS INTEGER
  END TYPE
  
  DIM SHARED PctC(1) AS STRING * 4

SUB ClearBack
  LibFile2Scrn "AR", "MENUBAK", MonoCode%, Attribute%, ErrorCode%
END SUB

SUB ClearScrn STATIC
  WazzWind 1, 1, 25, 80, 7, 0, 0
END SUB

SUB CursorOff
LOCATE , , 0
END SUB

FUNCTION GetARCustCnt&
  
  REDIM ARCust(1)  AS ARCustRecType
  ARCustLen = LEN(ARCust(1))  'Length of Cust Record Structure
  
  ARFile = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM SHARED AS ARFile LEN = ARCustLen
  NumOfRecs& = LOF(ARFile) \ ARCustLen
  CLOSE ARFile
  
  ERASE ARCust
  
  GetARCustCnt& = NumOfRecs&
  
END FUNCTION

SUB LookUp (RecNo&, Text$, ChkBalFlag%, CLSFlag%, CNumFlag%)

LookUpTop:

  REDIM Hlp$(1 TO 2)
  Hlp$(1) = "Enter an account number to look-up here."
  Hlp$(2) = "Enter all or part of the Customer Search Name here."
  
  REDIM ARCust(1) AS ARCustRecType
  ARCustLen = LEN(ARCust(1))
  
  SName$ = ""
  AcctNum& = 0
  LScrn = 2
  
  CursorOff
  
  REDIM ScrnArray(0)
  REDIM ScrnArray2(0)
  
  SaveScrn ScrnArray()
  
  REDIM LText(1 TO 2) AS STRING * 17
  
  MScrn = 2
  
  LText(1) = "Customer Number:"
  LText(2) = "    Search Name:"
  
  LibName$ = "AR"
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

  ClearBack

  DisplayARScrn ScrnName$

  QPrintRC Text$, 8, TCol, -1
  GOSUB DisplayLookupText
  
  ShowCursor
  
  Action = 1
  FirstTime = True
  Frm(1).StayOnField = True
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      FirstTime = NOT FirstTime
      SELECT CASE LScrn
      CASE 1, 2
        LSET Form$(1, 0) = ""
        Fld(1).FType = 2
      END SELECT
      Form$(1, 1) = Hlp$(LScrn)
      Action = 1
    END IF
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE -68, 13                'F10Key    Proceed with look up
      CursorOff
      SELECT CASE LScrn

      CASE 1    'account lookup
        SName$ = LEFT$(QPTrim$(Form$(0, 0)), 10)
        IF LEN(SName$) = 0 THEN
          SName$ = SPACE$(10)
        END IF
        SaveScrn ScrnArray2()
        RestScrn ScrnArray()
        Search4Cust SName$, RecNo&, ChkBalFlag, CLSFlag, True
        IF RecNo& > 0 THEN
          OKFlag = True
        ELSEIF RecNo& = 0 THEN
          Ok = MsgBox%("AR.QSL", "NOMATCH")
        END IF
        RestScrn ScrnArray2()
        Action = 1
'*******************
'        AcctNum& = QPValL(Form$(1, 0))
'        IF AcctNum& < 1 OR AcctNum& > GetARCustCnt& THEN
'          Ok = MsgBox%("AR.QSL", "BADACCTN")
'        ELSE
'          REDIM ARCust(1) AS ARCustRecType
'          ARCustLen = LEN(ARCust(1))
'          FOpenS "ARCUST.DAT", ARFile   'open data file
'          FGetRTA ARFile, ARCust(1), AcctNum&, ARCustLen
'          FClose ARFile
'            RecNo& = AcctNum&
'            OKFlag = True
'        END IF
'        Action = 1
'**********************
      CASE 2    'Name lookup
        SName$ = LEFT$(QPTrim$(Form$(0, 0)), 10)
        IF LEN(SName$) = 0 THEN
          SName$ = SPACE$(10)
        END IF
        SaveScrn ScrnArray2()
        'RestScrn ScrnArray()
        Search4Cust SName$, RecNo&, ChkBalFlag, CLSFlag, False
        IF RecNo& > 0 THEN
          OKFlag = True
        ELSEIF RecNo& = 0 THEN
          Ok = MsgBox%("AR.QSL", "NOMATCH")
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
  
  ERASE ARCust, ScrnArray, ScrnArray2
  ERASE Frm, Form$, Fld, LText, Hlp$
  
  EXIT SUB
  
DisplayLookupText:
  QPrintRC LText(LScrn), 12, 15, -1
RETURN

END SUB

SUB MakeCustIndex (IdxType)
  
  ShowProcessingScrn "Creating Customer Name Index"
  REDIM ARCust(1) AS ARCustRecType
  ARRecLen = LEN(ARCust(1))
  
  CustFile = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM AS CustFile LEN = ARRecLen
  NumOfCRecs = LOF(CustFile) / ARRecLen
  
  QPrintRC "Reading Customer Information", 11, 27, -1
  
  REDIM Array(1 TO NumOfCRecs) AS SortStruct
  FOR Cnt = 1 TO NumOfCRecs
    GET CustFile, Cnt, ARCust(1)
    SELECT CASE IdxType
    CASE 1
      Array(Cnt).who = LEFT$(QPTrim$(ARCust(1).BILLNAME), 14)
    CASE 2
      Array(Cnt).who = QPTrim$(ARCust(1).SORTNAME)
    END SELECT
    Array(Cnt).RecNum = Cnt
    ShowPctComp Cnt, NumOfCRecs
  NEXT
  CLOSE
  
  QPrintRC "Sorting Customer Information", 11, 27, -1

  'Sort Them Here
  SortT Array(1), NumOfCRecs, 0, LEN(Array(1)), 0, 14
  
  
  QPrintRC "   Writing Customer Index   ", 11, 27, -1
  
  Idxfile = FREEFILE
  OPEN "ARTEMP.IDX" FOR RANDOM AS Idxfile LEN = 2
  FOR Cnt = 1 TO NumOfCRecs
    PUT #Idxfile, Cnt, Array(Cnt).RecNum
    ShowPctComp Cnt, NumOfCRecs
  NEXT
  CLOSE
  
END SUB

FUNCTION PromptSaveData%
  
  REDIM TempScrn(0)
  SaveScrn TempScrn()
  
  LibName$ = "AR"
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
  
  DisplayARScrn FormName$
  
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

'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
'07-01-98
'corrected a bug which could cause certain numbers to round incorrectly
FUNCTION Round# (N#)
  Round# = INT(N# * 100 + .5000001#) / 100
END FUNCTION

SUB Search4Cust (Search$, RecNo&, ChkBalFlag%, CLSFlag%, CNumFlag%)
  
  STATIC Choice, LastSEARCH$
  Acct$ = SPACE$(10)

  IF LastSEARCH$ <> Search$ THEN   'if this search is different than the
    LastSEARCH$ = Search$          'last search, then set choice to first
    Choice = 1                     'match
  END IF
  
  REDIM TScrnArray(0)
  REDIM ScrnArray(0)
  SaveScrn ScrnArray()
  
  WazzWind 10, 22, 14, 58, 10, 2, True
  QPrintRC "Searching:    % Completed.", 12, 28, 14
  
  REDIM MChoice(1 TO 1) AS FLen2
  REDIM ARCust(1) AS ARCustRecType
  
  ARCustLen = LEN(ARCust(1))
  
  SearchLen = LEN(Search$)
  Match = False

  ARFile = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM SHARED AS #ARFile LEN = LEN(ARCust(1))
  'open data file
  
  NumOfCust& = LOF(ARFile) / ARCustLen
  
  ' Find matching record

  FOR Ccnt& = 1 TO NumOfCust&
    GET ARFile, Ccnt&, ARCust(1)
    IF UCASE$(ARCust(1).Deleted) = "Y" THEN
      GOTO DelSkip2:
    END IF
    WhatRec& = Ccnt&
    IF CNumFlag THEN
      UBSearchN$ = ARCust(1).CUSTNUMB
      IF INSTR(UBSearchN$, Search$) > 0 THEN
        GOSUB CustLoadEM2
      END IF

    ELSE
      UBSearchN$ = ARCust(1).SORTNAME
      IF INSTR(UBSearchN$, Search$) > 0 THEN
        GOSUB CustLoadEM2
      END IF
    END IF

    'UBSearchN$ = LEFT$(ARCust(1).SORTNAME, SearchLen)
    'IF (Search$ = UBSearchN$) THEN
    '  GOSUB CustLoadEM2
    'END IF
DelSkip2:
    ShowSearchPCT WhatRec&, NumOfCust&
  NEXT
  
  CLOSE ARFile
  
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
    LSET Title$ = "  Billing Name             City                  Cust No.  "
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
    
LoopRestart:
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
  
  ERASE ScrnArray, MChoice, ARCust
  
  EXIT SUB
  
CustLoadEM2:
  DCnt = DCnt + 1
  REDIM PRESERVE MChoice(1 TO DCnt) AS FLen2
  'RSET Acct$ = QPTrim$(STR$(WhatRec&))
  RSET Acct$ = QPTrim$(ARCust(1).CUSTNUMB)
  LSET MChoice(DCnt).V = LEFT$(QPTrim$(ARCust(1).BILLNAME), 24)
  MID$(MChoice(DCnt).V, 28, 11) = ARCust(1).CITY
  MID$(MChoice(DCnt).V, 48, 10) = Acct$
  MID$(MChoice(DCnt).V, 61) = MKL$(WhatRec&)
RETURN
  
  
END SUB

SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%) STATIC
  RSET PctC(1) = STR$(INT((RecNo / NumOfRecs) * 100))
  QPrintRC PctC(1), 13, 39, Cnf.HiLite
END SUB

SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&) STATIC
  RSET PctC(1) = STR$(INT((RecNo& / NumOfRecs&) * 100))
  QPrintRC PctC(1), 13, 40, Cnf.HiLite
END SUB

SUB ShowProcessingScrn (RptTitle$)
  
  TitleRow = 9
  TitleCol = 40 - (LEN(RptTitle$) \ 2) + 1
  CursorOff

  DisplayARScrn "PRORPT"

  HideCursor
  QPrintRC RptTitle$, TitleRow, TitleCol, 126
  QPrintRC "Processing:    % Completed.", 13, 28, Cnf.HiLite
  ShowCursor
  
END SUB

SUB ShowSearchPCT (BYVAL RecNo&, BYVAL NumOfRecs&) STATIC
  RSET PctC(1) = STR$(INT((RecNo& / NumOfRecs&) * 100))
  HideCursor
  QPrintRC PctC(1), 12, 38, 15
  ShowCursor
END SUB

'********** Unique.Bas - provides a unique file name
'Copyright (c) 1989 Ethan Winer
'NOTE: Although the manual shows no arguments to the Unique$ function, we
'have added the capability to specify a path name as an argument.  This lets
'you create a unique file name, and also be sure a file with that name does
'not exist in any given directory.
FUNCTION Unique$ (Path$)
  
  IF LEN(Path$) AND RIGHT$(Path$, 1) <> "\" THEN Path$ = Path$ + "\"
  Seed& = ABS(TIMER)            'use the TIMER as a seed
  DO
    TempName$ = Path$ + MID$(STR$(Seed&), 2)    'make a string out of it
    TempName$ = TempName$ + ".RPT"
    Seed& = Seed& + 1           'increment for next time
  LOOP UNTIL NOT Exist%(TempName$)              'loop and try another name
  Unique$ = TempName$           'this is the function output
  
END FUNCTION

