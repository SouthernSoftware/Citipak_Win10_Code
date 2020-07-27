DECLARE FUNCTION Rounds! (N!)
DEFINT A-Z
DECLARE FUNCTION GetCustCnt& ()
DECLARE SUB Search4Cust1 (Search$, RecNo&, ChkBalFlag%, CLSFlag%, SSNFlag%)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION DoesCustOwe% (TaxCustRec AS ANY)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FLof& (Handle%)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION GetCustBalance# (RecNo&)
DECLARE FUNCTION GetCustName$ (CustRec&)
DECLARE FUNCTION GetTaxCustCnt& ()
DECLARE FUNCTION IsCustDeleted% (AcctNum&)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION OK2UPDateCust% ()
DECLARE FUNCTION ParseBillNum$ (Text$)
DECLARE FUNCTION PromptSaveData% ()
DECLARE FUNCTION QPStrI$ (Num%)
DECLARE FUNCTION QPStrL$ (Num&)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValL& (Number$)
DECLARE FUNCTION Round# (N#)
DECLARE FUNCTION Unique$ (Path$)
DECLARE SUB ClearBack ()
DECLARE SUB CursorOff ()
DECLARE SUB DelPropAbstract (PropRecs() AS LONG, WhatProp%, CustRec&)
DECLARE SUB DisplayCaScrn (ScrnName$)
DECLARE SUB FClose (Handle%)
DECLARE SUB FGetA (Handle%, SEG Element AS ANY, NumBytes AS ANY)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FOpenS (FileName$, Handle)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB HideCursor ()
DECLARE SUB MPaintBox (UlRow%, UlCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB RestScrn (Array%())
DECLARE SUB SaveScrn (Array%())
DECLARE SUB Search4Cust (Search$, RecNo&, ChkBalFlag%, CLSFlag%, SSNFlag%)
DECLARE SUB ShowCursor ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB ShowSearchPCT (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice, MaxLen%, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB WaitForAction ()
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)


  CONST False = 0, True = NOT False

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'CaCust.BI'
  
  TYPE FLen2
    V AS STRING * 64
  END TYPE
  
  TYPE SortStruct
    who AS STRING * 14
    RecNum AS INTEGER
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

SUB DisplayCaScrn (ScrnName$)
  LibFile2Scrn "TAX", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  IF ErrorCode% <> 0 THEN

    PRINT "Screen Error: "; ScrnName$
    END
  END IF
END SUB

FUNCTION GetCustCnt&
  
  REDIM CaCust(1)  AS CACUSTType
  CaCustLen = LEN(CaCust(1))  'Length of Cust Record Structure
  
  CaFile = FREEFILE
  OPEN "CAMASTER.DAT" FOR RANDOM SHARED AS CaFile LEN = CaCustLen
  NumOfRecs& = LOF(CaFile) \ CaCustLen
  CLOSE CaFile
  
  ERASE CaCust
  
  GetCustCnt& = NumOfRecs&
  
END FUNCTION

FUNCTION IsCustDeleted (AcctNum&)
  IsCustDeleted = False         'assume they aren't deleted
  
  REDIM CaCust(1)  AS CACUSTType
  CaCustLen = LEN(CaCust(1))  'Length of Cust Record Structure
  CaFile = FREEFILE
  OPEN "CAMASTER.DAT" FOR RANDOM SHARED AS CaFile LEN = CaCustLen
  GET CaFile, AcctNum&, CaCust(1)
  CLOSE CaFile
  
  IF CaCust(1).Deleted <> 0 THEN
    IsCustDeleted = True
  END IF
  
  ERASE CaCust
  
END FUNCTION

SUB LookUp (RecNo&, Text$, ChkBalFlag%, CLSFlag%, SSNFlag%)

LookUpTop:



  REDIM Hlp$(1 TO 2)
  Hlp$(1) = "Enter an account number to look-up here."
  Hlp$(2) = "Enter all or part of the Customer Search Name here."
  REDIM CaCust(1) AS CACUSTType
  CaCustLen = LEN(CaCust(1))

  SName$ = ""
  AcctNum& = 0
  PIN$ = ""
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
  DisplayCaScrn ScrnName$
  
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
        AcctNum& = QPValL(Form$(1, 0))
        IF AcctNum& < 1 OR AcctNum& > GetCustCnt& THEN
          Ok = MsgBox%("TAX.QSL", "BADACCTN")
        ELSEIF IsCustDeleted(AcctNum&) THEN
          Ok = MsgBox%("TAX.QSL", "DELACCTN")
        ELSEIF ChkBalFlag THEN
          REDIM CaCust(1) AS CACUSTType
          CaCustLen = LEN(CaCust(1))
          FOpenS "CAMASTER.DAT", CaFile   'open data file
          FGetRTA CaFile, CaCust(1), AcctNum&, CaCustLen
          FClose CaFile
        ELSE
          RecNo& = AcctNum&
          OKFlag = True
        END IF
        Action = 1
      CASE 2    'Name lookup
        SName$ = LEFT$(QPTrim$(Form$(0, 0)), 10)
        IF LEN(SName$) = 0 THEN
          SName$ = SPACE$(10)
        END IF
        SaveScrn ScrnArray2()
        RestScrn ScrnArray()
        Search4Cust SName$, RecNo&, ChkBalFlag, CLSFlag, False
        IF RecNo& > 0 THEN
          OKFlag = True
        ELSEIF RecNo& = 0 THEN
          Ok = MsgBox%("TAX.QSL", "NOMATCH")
        END IF
        RestScrn ScrnArray2()
        Action = 1
      CASE 3
        SName$ = Form$(1, 0)
        IF LEN(SName$) = 0 THEN
          SName$ = SPACE$(10)
        END IF
        SaveScrn ScrnArray2()
        RestScrn ScrnArray()
        Search4Cust SName$, RecNo&, ChkBalFlag, CLSFlag, True
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
  
  ERASE CaCust, ScrnArray, ScrnArray2
  ERASE Frm, Form$, Fld, LText, Hlp$
  
  EXIT SUB
  
DisplayLookupText:
  QPrintRC LText(LScrn), 12, 15, -1
RETURN

END SUB

SUB MakeCustIndex (IdxType)
  
  ShowProcessingScrn "Creating Customer Name Index"
  REDIM CaCust(1) AS CACUSTType
  CaRecLen = LEN(CaCust(1))
  
  CustFile = FREEFILE
  OPEN "CAMASTER.DAT" FOR RANDOM AS CustFile LEN = CaRecLen
  NumOfCRecs = LOF(CustFile) / CaRecLen
  
  QPrintRC "Reading Customer Information", 11, 27, -1
  
  REDIM Array(1 TO NumOfCRecs) AS SortStruct
  FOR Cnt = 1 TO NumOfCRecs
    GET CustFile, Cnt, CaCust(1)
    SELECT CASE IdxType
    CASE 1, 2, 3
      Array(Cnt).who = (LEFT$(QPTrim$(CaCust(1).CustName), 12))
    END SELECT
    Array(Cnt).RecNum = Cnt
    ShowPctComp Cnt, NumOfCRecs
  NEXT
  CLOSE
  
  QPrintRC "Sorting Customer Information", 11, 27, -1
  'Sort Them Here
  SortT Array(1), NumOfCRecs, 0, LEN(Array(1)), 0, 14
  'SortT (Element, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
  
  QPrintRC "   Writing Customer Index   ", 11, 27, -1
  
  Idxfile = FREEFILE
  OPEN "CaTEMP.IDX" FOR RANDOM AS Idxfile LEN = 2
  FOR Cnt = 1 TO NumOfCRecs
    PUT #Idxfile, Cnt, Array(Cnt).RecNum
    ShowPctComp Cnt, NumOfCRecs
  NEXT
  CLOSE
  
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
  
  DisplayCaScrn FormName$
  
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

FUNCTION Rounds! (N!)
'****************************************************************************
'Rounds a Single Precision value to nearest hundreth
'****************************************************************************
'07-01-98
'corrected a bug which could cause certain numbers to round incorrectly
 Rounds! = INT(N! * 100 + .5000001#) / 100
END FUNCTION

SUB Search4Cust (Search$, RecNo&, ChkBalFlag%, CLSFlag%, SSNFlag%)
  
  STATIC Choice, LastSEARCH$
  ACCT$ = SPACE$(5)

  IF LastSEARCH$ <> Search$ THEN
    LastSEARCH$ = Search$
    Choice = 1
  END IF
  

  REDIM TScrnArray(0)
  REDIM ScrnArray(0)
  SaveScrn ScrnArray()
  
  WazzWind 10, 22, 14, 58, 10, 2, True
  QPrintRC "Searching:    % Completed.", 12, 28, 14
  
'091598 Found a bug that caused the lookup to get erroneous records
'       if this was changed to greater than 32???
  CustBlock = 32
  
  REDIM Mchoice(1 TO 1) AS FLen2
  REDIM CaCust(1 TO CustBlock) AS CACUSTType
  
  CaCustLen = LEN(CaCust(1))
  
  SearchLen = LEN(Search$)
  Match = False
  
  FOpenS "CAMASTER.DAT", CaFile   'open data file
  
  NumOfCust& = FLof&(CaFile) / CaCustLen
  NumChunks& = NumOfCust& / CustBlock
  
  OddRecs& = NumOfCust& MOD CustBlock
  
  BlockSize& = (0& + CaCustLen) * CustBlock
  '            ^^^^^ stops an overflow error
  'since CaCustLen is an integer, basic will try to multiply to
  'an integer result. the above "0&" causes basic to convert to a long
  'then multiply'

  ' Find matching record
  FOR CCnt& = 1 TO NumChunks&
    FGetA CaFile, CaCust(1), BlockSize&
    FOR RecCnt = 1 TO CustBlock
      WhatRec& = ((CCnt& - 1) * CustBlock) + RecCnt
        UBSearchN$ = LEFT$(CaCust(RecCnt).CustName, SearchLen)
        IF (Search$ = UBSearchN$) AND CaCust(RecCnt).Deleted = 0 THEN
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
      FGetRTA CaFile, CaCust(1), CCnt&, CaCustLen
      WhatRec& = CCnt&
        UBSearchN$ = LEFT$(CaCust(RecCnt).CustName, SearchLen)
        IF (Search$ = UBSearchN$ AND CaCust(RecCnt).Deleted = 0) THEN
          GOSUB CustLoadEM2
        END IF
      
DelSkip3:
      ShowSearchPCT WhatRec&, NumOfCust&
    NEXT
  END IF
  
  FClose CaFile
  
  IF DCnt = 0 THEN
    RecNo& = 0
    GOTO ExitSearch2
  ELSE
    
    SortT Mchoice(1), DCnt, Direction%, LEN(Mchoice(1).V), 0, 18
    
    MaxLen = 59 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something
    
    IF Choice < 1 THEN
      Choice = 1                'Pre-load choice to highlight
    END IF
    
    Title$ = SPACE$(MaxLen + 4)
    LSET Title$ = "  Name                                                   Acct"
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
      VertMenuT2 Mchoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        RecNo& = -1
        EXIT DO 'choice = 0
      END IF
      RecNo& = CVL(MID$(Mchoice(Choice).V, 61, 4))
      IF ChkBalFlag THEN
        FOpenS "CAMASTER.DAT", CaFile   'open data file
        FGetRTA CaFile, CaCust(1), RecNo&, CaCustLen
        FClose CaFile
      END IF
    LOOP UNTIL RecNo& > 0
  END IF
  
ExitSearch2:
  RestScrn ScrnArray()
  
  ERASE ScrnArray, Mchoice, CaCust
  
  EXIT SUB
  
CustLoadEM2:
  
  DCnt = DCnt + 1
  REDIM PRESERVE Mchoice(1 TO DCnt) AS FLen2
  RSET ACCT$ = QPTrim$(STR$(WhatRec&))
  LSET Mchoice(DCnt).V = LEFT$(QPTrim$(CaCust(RecCnt).CustName), 34)
  MID$(Mchoice(DCnt).V, 55, 5) = ACCT$
  MID$(Mchoice(DCnt).V, 61) = MKL$(WhatRec&)
RETURN
  
  
END SUB

SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
  RSET PctC(1) = QPStrI$(INT((RecNo / NumOfRecs) * 100))
  QPrintRC PctC(1), 13, 39, Cnf.HiLite
END SUB

SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&) STATIC
  RSET PctC(1) = QPStrL$(INT((RecNo& / NumOfRecs&) * 100))
  QPrintRC PctC(1), 13, 40, Cnf.HiLite
END SUB

SUB ShowProcessingScrn (RptTitle$)
  
  TitleRow = 9
  TitleCol = 40 - (LEN(RptTitle$) \ 2) + 1
  CursorOff
  'BlockClear
  DisplayCaScrn "PRORPT"
  
  HideCursor
  QPrintRC RptTitle$, TitleRow, TitleCol, 126
  QPrintRC "Processing:    % Completed.", 13, 28, Cnf.HiLite
  ShowCursor
  
END SUB

SUB ShowSearchPCT (BYVAL RecNo&, BYVAL NumOfRecs&) STATIC
  RSET PctC(1) = QPStrI$(INT((RecNo& / NumOfRecs&) * 100))
  HideCursor
  QPrintRC PctC(1), 12, 38, 15
  ShowCursor
END SUB

SUB SmallPause
  St# = TIMER + .6
  DO
  LOOP UNTIL TIMER > St#
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

SUB UpDateTicklerFile
  ThisMonth$ = LEFT$(DATE$, 2)
  TickFile = FREEFILE
  OPEN "CAINTCK.DAT" FOR OUTPUT AS #TickFile
  PRINT #TickFile, ThisMonth$
  CLOSE TickFile
END SUB

