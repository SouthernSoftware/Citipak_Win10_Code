DEFINT A-Z
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE FUNCTION FmtBook$ (Book$)
DECLARE FUNCTION FmtSeqN$ (SeqN$)
DECLARE FUNCTION FindRateTbl% (RateCode$, NumOfRates%, UBRateTbls() AS ANY)
DECLARE FUNCTION GetNumOfRevs% ()
DECLARE SUB AddEditLocation (RecNo&)
DECLARE SUB Search4LNum (LNum$, RecNo&, CLSFlag%)
DECLARE SUB Search4Meter (MeterNum$, RecNo&, CLSFlag%)
DECLARE SUB ShowSearchWheel (BYVAL Row%, BYVAL Col%)
DECLARE SUB Search4SAddr (SAddr$, RecNo&, CLSFlag%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB Search4Cust (SEARCH$, RecNo&, CLSFlag%, LocationFlag%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION GetNumOfAcct% ()
DECLARE FUNCTION ConvDateStr$ (D$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FudgeFactor& ()
DECLARE FUNCTION NovellThere% ()
DECLARE FUNCTION QPStrI$ (Num%)
DECLARE FUNCTION QPStrL$ (Num&)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Round# (N#)
DECLARE SUB BlockClear ()
DECLARE SUB ButtonPress (ButNo%, Down%, Presses%, X%, Y%)
DECLARE SUB ClearScrn ()
DECLARE SUB CursorOff ()
DECLARE SUB DisplayScrn (BTmp%(), Element%, MonoCode%, WipeType%)
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB ExplodBox (UlRow%, UlCol%, BrRow%, BrCol%, Scr(), El%)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB GetCursor (X, Y, Button)
DECLARE SUB HideCursor ()
DECLARE SUB BlockClear ()
DECLARE SUB MPaintBox (UlRow%, UlCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB MScrnRest (UlRow, UlCol, LRRow, LRCol, SEG Address)
DECLARE SUB MScrnSave (UlRow, UlCol, LRRow, LRCol, SEG Address)
DECLARE SUB Pause3 (MillaSecs%, ProcessorSpeed&)
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
DECLARE SUB ShowCursor ()
DECLARE SUB SaveScrn (Array())
DECLARE SUB RestScrn (Array())
DECLARE SUB StuffBuf (Ky$)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE SUB FOpenS (FileName$, Handle)
DECLARE SUB FClose (Handle%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE FUNCTION QPValL& (Number$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice, MaxLen%, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB WaitForAction ()
DECLARE SUB Get.Moose.OR.Key (Ky$, MooseButton%, MRow%, MCol%)

  TYPE FLen2
    V AS STRING * 64
  END TYPE
  
  '$INCLUDE: 'DefCnf.bi'      'Defines a TYPE for monitor and color info..
  '$INCLUDE: 'FORMEDIT.bi'
  '$INCLUDE: 'fieldinf.BI'
  
  '$INCLUDE: 'ARCUST.bi'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  
  
  CONST False = 0, True = NOT False
  CONST Wheel$ = "|/Ä\"
  
  DIM SHARED PctC(1) AS STRING * 3

'AskAbandonPrint

FUNCTION AskAbandonPrint%

  STATIC BeenHere, Escape$

  IF NOT BeenHere THEN
    BeenHere = True
    Escape$ = CHR$(27)
  END IF

  Ky$ = INKEY$              'ButNo,DnFlg,NoPresses,col,row
  ButtonPress 1, N, MooseButton, MCol, MRow  ' ----- Check Mouse activity
  IF N AND 2 THEN                'if it was the right moose button and
    DO                           'if they are still holding it down then
      GetCursor 0, 0, NewStatus  'wait till they let go of the button
    LOOP WHILE NewStatus
    ButtonPress 0, 0, 0, 0, 0    'this has the effect of clearing all
    Ky$ = Escape$                'moose buttons.
  END IF

  IF Ky$ <> Escape$ THEN
    EXIT FUNCTION
  END IF

  REDIM TempScrn(0)
  SaveScrn TempScrn()
  ActMRow = 14
  BlockClear
  DisplayUBScrn "ABORTJOB"

  DO

    Get.Moose.OR.Key Ky$, MooseButton, MRow, MCol

    IF MooseButton THEN
      MRow = (MRow \ 8) + 1     'Convert MCol, MRow to Row and Col.
      MCol = (MCol \ 8) + 1
      IF MRow = ActMRow THEN
        SELECT CASE MCol
        CASE 26 TO 39           'Cancel printing
          PressButton EscKey, ActMRow, 26, 39
        CASE 40 TO 55           'let it rip
          PressButton 13, ActMRow, 40, 55
        END SELECT
      END IF
    END IF

    IF LEN(Ky$) THEN
      IF Ky$ = Escape$ THEN
        AskAbandonPrint% = True
      ELSE
        AskAbandonPrint% = False
      END IF
      EXIT DO
    END IF
  LOOP

  RestScrn TempScrn()

END FUNCTION

SUB ClearScrn STATIC
  WazzWind 1, 1, 25, 80, 7, 0, 0
END SUB

SUB CursorOff STATIC
  LOCATE , , 0
END SUB

SUB DisplayUBScrn (ScrnName$)
  LibFile2Scrn "UB", ScrnName$, MonoCode, Attribute%, ErrCode
END SUB

FUNCTION FmtBook$ (Book$)
  Book$ = QPTrim$(Book$)
  BookLen = LEN(Book$)

  SELECT CASE BookLen
  CASE 0
    FmtBook$ = "00"
  CASE 1
    FmtBook$ = "0" + Book$
  CASE ELSE
    FmtBook$ = Book$
  END SELECT

END FUNCTION

FUNCTION FmtSeqN$ (SeqN$)

  SeqN$ = QPTrim$(SeqN$)
  SeqNLen = LEN(SeqN$)

  SELECT CASE SeqNLen
  CASE 0
    FmtSeqN$ = "000000"
  CASE 1 TO 5
    FmtSeqN$ = "000000"
    MID$(FmtSeqN$, (6 - SeqNLen) + 1) = SeqN$
  CASE ELSE
    FmtSeqN$ = SeqN$
  END SELECT

END FUNCTION

'This function returns the number of customer records
FUNCTION GetNumOfAcct
  REDIM TCustRec(1) AS UBCustRecType
  RecLen = LEN(TCustRec(1))
  CFileSize& = FileSize("UBCUST.DAT")
  GetNumOfAcct = CFileSize& \ RecLen
  ERASE TCustRec
END FUNCTION

SUB LookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, LocationFlag%)
  
  CursorOff
  
  STATIC SName$, AcctNum&, MeterNum$, SAddr$, LNum$
  
  REDIM ScrnArray(0)
  REDIM ScrnArray2(0)
  
  SaveScrn ScrnArray()
  
  REDIM LText$(5)
  
  LScrn = DefaultLook
  
  MScrn = 5
  
  LText$(1) = " Account Number:"
  LText$(2) = "    Search Name:"
  LText$(3) = "   Meter Number:"
  LText$(4) = "Service Address:"
  LText$(5) = "Location Number:"
  
  LibName$ = "UB"
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
  DisplayUBScrn ScrnName$
  
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
        'Fld(1).StorLen = 6
        'Action = 1
      CASE 2
        Form$(1, 0) = SName$
      CASE 3
        Form$(1, 0) = MeterNum$
      CASE 4
        Form$(1, 0) = SAddr$
      CASE 5
        Form$(1, 0) = LNum$
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
        IF AcctNum& < 1 OR AcctNum& > GetNumOfAcct THEN
          OK = MsgBox%("UB.QSL", "BADACCTN")

        ELSE
          RecNo& = AcctNum&
          OKFlag = True
        END IF
      CASE 2    'Name lookup
        SName$ = LEFT$(QPTrim$(Form$(0, 0)), 10)
        IF LEN(SName$) = 0 THEN
          OK = MsgBox%("UB.QSL", "BADSEARH")
          Action = 1
          Frm(1).FldNo = 1
        ELSE
          SaveScrn ScrnArray2()
          RestScrn ScrnArray()
          Search4Cust SName$, RecNo&, CLSFlag, LocationFlag
          IF RecNo& > 0 THEN
            OKFlag = True
          ELSEIF RecNo& = 0 THEN
            OK = MsgBox%("UB.QSL", "NOMATCH")
          END IF
          RestScrn ScrnArray2()
          Action = 1
        END IF
      CASE 3    'meter number
        MeterNum$ = QPTrim$(Form$(0, 0))
        IF LEN(MeterNum$) = 0 THEN
          OK = MsgBox%("UB.QSL", "BADSEARH")
          Action = 1
          Frm(1).FldNo = 1
        ELSE
          SaveScrn ScrnArray2()
          RestScrn ScrnArray()
          Search4Meter MeterNum$, RecNo&, CLSFlag
          IF RecNo& > 0 THEN
            OKFlag = True
          ELSEIF RecNo& = 0 THEN
            OK = MsgBox%("UB.QSL", "NOMATCH")
          END IF
          RestScrn ScrnArray2()
          Action = 1
        END IF
      CASE 4    'service address
        SAddr$ = QPTrim$(Form$(0, 0))
        IF LEN(SAddr$) = 0 THEN
          OK = MsgBox%("UB.QSL", "BADSEARH")
          Action = 1
          Frm(1).FldNo = 1
        ELSE
          SaveScrn ScrnArray2()
          RestScrn ScrnArray()
          Search4SAddr SAddr$, RecNo&, CLSFlag
          IF RecNo& > 0 THEN
            OKFlag = True
          ELSEIF RecNo& = 0 THEN
            OK = MsgBox%("UB.QSL", "NOMATCH")
          END IF
          RestScrn ScrnArray2()
          Action = 1
        END IF
      CASE 5    'Location lookup
        'CursorOff
        LNum$ = QPTrim$(Form$(0, 0))

        DashPos = INSTR(LNum$, "-")
        IF DashPos > 1 THEN
          Book$ = FmtBook$(LEFT$(LNum$, DashPos - 1))
          SeqN$ = FmtSeqN$(MID$(LNum$, DashPos + 1))
          LNum$ = Book$ + "-" + SeqN$
          LSET Form$(1, 0) = LNum$
          SaveField 0, Form$(), Fld(), BadField
          Action = 1
        END IF

        IF LEN(LNum$) = 0 OR DashPos <= 0 THEN
          'OK = MsgBox%("UB.QSL", "BADSEARH")
          OK = MsgBox%("UB.QSL", "BADACCTN")
          Action = 1
          Frm(1).FldNo = 1
        ELSE
          SaveScrn ScrnArray2()
          RestScrn ScrnArray()
          Search4LNum LNum$, RecNo&, CLSFlag
          IF RecNo& > 0 THEN
            OKFlag = True
          ELSEIF RecNo& = 0 THEN
            OK = MsgBox%("UB.QSL", "NOMATCH")
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

  EXIT SUB
  
DisplayLookupText:
  QPrintRC LText$(LScrn), 12, 15, -1
  RETURN
  
  
END SUB

  '****************************************************************************
  'Rounds a double precision value to nearest hundreth
  '****************************************************************************
FUNCTION Round# (N#)
  Round# = INT(N# * 100 + .5) / 100
END FUNCTION

SUB Search4Cust (SEARCH$, RecNo&, CLSFlag%, LocationFlag%)
  
  STATIC Choice, LastSEARCH$
  
  IF LastSEARCH$ <> SEARCH$ THEN
    LastSEARCH$ = SEARCH$
    Choice = 1
  END IF
  
  REDIM ScrnArray(0)
  SaveScrn ScrnArray()
  'Wheel$ = "|/Ä\"
  WPos = 1
  
  DisplayUBScrn "SHOWSCRH"
  
  REDIM UBCustRec(1) AS UBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBLocatRec(1) AS UBLocationRecType
  UBLocatRecLen = LEN(UBLocatRec(1))
  
  IdxRecLen = 4 'we are using a integer
  IdxFileSize& = FileSize("UBCUSTNM.IDX")
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  
  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  
  FGetAH "UBCUSTNM.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it
  
  SearchLen = LEN(SEARCH$)
  
  Match = False
  FirstRec = 1
  LastRec = IdxNumOfRecs
  
  BotOffSet = 1
  TopOffSet = IdxNumOfRecs
  
  FOpenS "UBCUST.DAT", C1Handle 'open data file
  FOpenS "UBLOCA.DAT", L1Handle 'open data file
  
  'Find matching record
  
  MidRec = (LastRec + FirstRec) \ 2
  
  DO
    IF LastSRec = MidRec THEN EXIT DO
    LastSRec = MidRec
    FGetRTA C1Handle, UBCustRec(1), CLNG(IdxBuff(MidRec).RecNum), UBCustRecLen
    
    UBSearchN$ = LEFT$(UBCustRec(1).SEARCH, SearchLen)
    ShowSearchWheel 12, 44
    IF SEARCH$ = UBSearchN$ THEN
      
      IF MidRec - BotOffSet > 1 THEN
        MidRec = MidRec - 1
      ELSE
        FirstMatchRec = MidRec
      END IF
    ELSEIF SEARCH$ < UBSearchN$ THEN            'lower
      TopOffSet = MidRec
      MidRec = TopOffSet - ((TopOffSet - BotOffSet) \ 2)
    ELSE        'higher
      BotOffSet = MidRec
      MidRec = BotOffSet + ((TopOffSet - BotOffSet) \ 2) + 1
      IF MidRec = IdxNumOfRecs + 1 THEN
        EXIT DO
      END IF
    END IF
    IF TopOffSet = BotOffSet THEN EXIT DO
    'Look into this
    'IF TopOffSet = BotOffSet THEN EXIT DO
  LOOP UNTIL FirstMatchRec
  
  IF FirstMatchRec > 0 THEN
    MidRec = FirstMatchRec
    LastMatchRec = IdxNumOfRecs
    FGetRTA C1Handle, UBCustRec(1), CLNG(IdxBuff(FirstMatchRec).RecNum), UBCustRecLen
    MidRec = MidRec + 1
    DO UNTIL MidRec > IdxNumOfRecs
      FGetRTA C1Handle, UBCustRec(1), CLNG(IdxBuff(MidRec).RecNum), UBCustRecLen
      UBSearchN$ = LEFT$(UBCustRec(1).SEARCH, SearchLen)
      IF SEARCH$ <> UBSearchN$ THEN
        LastMatchRec = MidRec - 1
        EXIT DO
      END IF
      ShowSearchWheel 12, 44
      MidRec = MidRec + 1
    LOOP

    MidRec = FirstMatchRec
    'LastMatchRec = IdxNumOfRecs
    FGetRTA C1Handle, UBCustRec(1), CLNG(IdxBuff(FirstMatchRec).RecNum), UBCustRecLen
    MidRec = MidRec - 1
    DO UNTIL MidRec = 0
      FGetRTA C1Handle, UBCustRec(1), CLNG(IdxBuff(MidRec).RecNum), UBCustRecLen
      UBSearchN$ = LEFT$(UBCustRec(1).SEARCH, SearchLen)
      IF SEARCH$ <> UBSearchN$ THEN
        FirstMatchRec = MidRec + 1
        EXIT DO
      ELSE
        FirstMatchRec = MidRec
      END IF
      ShowSearchWheel 12, 44
      MidRec = MidRec - 1
    LOOP


  ELSE
    RecNo& = 0
    FClose C1Handle
    FClose L1Handle
    
    GOTO ExitSearch
  END IF
  
  TotalMatch = (LastMatchRec - FirstMatchRec) + 1

  MemNeeded& = TotalMatch * 64&
  FreeMem& = FRE(-1)
  IF MemNeeded& > FreeMem& THEN
    FClose C1Handle
    FClose L1Handle
    'BEEP
    QPrintRC "Matched: " + QPStrI(TotalMatch), 24, 1, 15
    QPrintRC " Needed: " + QPStrL(MemNeeded&) + "  Free: " + QPStrL(FreeMem&), 25, 1, 15

    RecNo& = -1
    'TotalMatch = 0
    WaitForAction
    GOTO ExitSearch
  END IF


  IF TotalMatch > 1 THEN
    
    FOR Cnt = FirstMatchRec TO LastMatchRec
      FGetRTA C1Handle, UBCustRec(1), CLNG(IdxBuff(Cnt).RecNum), UBCustRecLen
      IF UBCustRec(1).CustFrstLocRec > 0 THEN

        DCnt = DCnt + 1
        StartSortOffSet = DCnt
        EndSortOffSet = DCnt

        REDIM PRESERVE MChoice(1 TO DCnt) AS FLen2
        
        FGetRTA L1Handle, UBLocatRec(1), UBCustRec(1).CustFrstLocRec, UBLocatRecLen
        IF DOSError THEN BEEP: STOP
        LSET MChoice(DCnt).V = LEFT$(QPTrim$(UBCustRec(1).CUSTNAME), 20)
        MID$(MChoice(DCnt).V, 22, 22) = LEFT$(QPTrim$(UBLocatRec(1).SERVADDR), 25)
        MID$(MChoice(DCnt).V, 48, 9) = UBLocatRec(1).Book + "-" + UBLocatRec(1).SeqNumb
        IF LocationFlag THEN
          MID$(MChoice(DCnt).V, 61) = MKL$(UBCustRec(1).CustFrstLocRec)
        ELSE
          MID$(MChoice(DCnt).V, 61) = MKL$(CLNG(IdxBuff(Cnt).RecNum))
        END IF
        MID$(MChoice(DCnt).V, 59, 1) = UBLocatRec(1).LocatStatus
        
        NextCustLoc& = UBLocatRec(1).NextLocRec
        DO WHILE NextCustLoc& > 0
          FGetRTA L1Handle, UBLocatRec(1), NextCustLoc&, UBLocatRecLen
          IF DOSError THEN BEEP: STOP
          DCnt = DCnt + 1
          EndSortOffSet = DCnt
          REDIM PRESERVE MChoice(1 TO DCnt) AS FLen2
          LSET MChoice(DCnt).V = LEFT$(QPTrim$(UBCustRec(1).CUSTNAME), 20)
          MID$(MChoice(DCnt).V, 22, 22) = LEFT$(QPTrim$(UBLocatRec(1).SERVADDR), 25)
          MID$(MChoice(DCnt).V, 48, 9) = UBLocatRec(1).Book + "-" + UBLocatRec(1).SeqNumb
          MID$(MChoice(DCnt).V, 61) = MKL$(NextCustLoc&)
          IF LocationFlag THEN
            MID$(MChoice(DCnt).V, 61) = MKL$(NextCustLoc&)
          ELSE
            MID$(MChoice(DCnt).V, 61) = MKL$(CLNG(IdxBuff(Cnt).RecNum))
          END IF

          MID$(MChoice(DCnt).V, 59, 1) = UBLocatRec(1).LocatStatus
          NextCustLoc& = UBLocatRec(1).NextLocRec
        LOOP
        IF StartSortOffSet < EndSortOffSet THEN
          ElmCnt = (EndSortOffSet - StartSortOffSet) + 1
          SortT MChoice(StartSortOffSet), ElmCnt, 0, 64, 47, 9
               '      (Element(1), NumElm, Dir, SSize%, MOffSet, MSize%)
        END IF
      ELSE
        IF NOT LocationFlag THEN
          DCnt = DCnt + 1
          REDIM PRESERVE MChoice(1 TO DCnt) AS FLen2
          LSET MChoice(DCnt).V = LEFT$(QPTrim$(UBCustRec(1).CUSTNAME), 20)
          MID$(MChoice(DCnt).V, 22, 22) = "  Inactive"
          MID$(MChoice(DCnt).V, 48, 9) = " Inactive"
          MID$(MChoice(DCnt).V, 59, 1) = "I"
          MID$(MChoice(DCnt).V, 61) = MKL$(IdxBuff(Cnt).RecNum)
        END IF
      END IF
      ShowSearchWheel 12, 44
    NEXT
    
    FClose C1Handle
    FClose L1Handle
    
    MaxLen = 59 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something
    
    IF Choice < 1 THEN
      Choice = 1                'Pre-load choice to highlight
    END IF
    
    Title$ = SPACE$(MaxLen + 4)
    LSET Title$ = "  Customer/Owner       Service Address        Location No.  S"
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
  ELSE
    RecNo& = IdxBuff(FirstMatchRec).RecNum
    FClose C1Handle
    FClose L1Handle

  END IF
  
ExitSearch:
  'SHELL
  RestScrn ScrnArray()
  
  
  
END SUB

SUB Search4LNum (LocNum$, RecNo&, CLSFlag%)

  REDIM ScrnArray(0)
  SaveScrn ScrnArray()

  DisplayUBScrn "SHOWSCRH"

  REDIM UBCustRec(1) AS UBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBLocatRec(1) AS UBLocationRecType
  UBLocatRecLen = LEN(UBLocatRec(1))

  IdxRecLen = 4 'we are using a integer
  IdxFileSize& = FileSize("UBCUSTBK.IDX")
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen

  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType

  FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it

  SearchLen = LEN(LocNum$)

  Match = False
  FirstRec = 1
  LastRec = IdxNumOfRecs

  BotOffSet = 0
  TopOffSet = IdxNumOfRecs

  FOpenS "UBCUST.DAT", C1Handle 'open data file
  FOpenS "UBLOCA.DAT", L1Handle 'open data file

  'Find matching record

  MidRec = (LastRec + FirstRec) \ 2

  DO
    IF LastSRec = MidRec THEN EXIT DO
    LastSRec = MidRec

    FGetRTA L1Handle, UBLocatRec(1), CLNG(IdxBuff(MidRec).RecNum), UBLocatRecLen

    UBSearchN$ = UBLocatRec(1).Book + "-" + UBLocatRec(1).SeqNumb
    ShowSearchWheel 12, 44

    IF LocNum$ = UBSearchN$ THEN

      IF MidRec - BotOffSet > 1 THEN
        MidRec = MidRec - 1
      ELSE
        FirstMatchRec = MidRec
      END IF
    ELSEIF LocNum$ < UBSearchN$ THEN            'lower
      TopOffSet = MidRec
      MidRec = TopOffSet - ((TopOffSet - BotOffSet) \ 2)
    ELSE        'higher
      BotOffSet = MidRec
      MidRec = BotOffSet + ((TopOffSet - BotOffSet) \ 2) + 1
      IF MidRec = IdxNumOfRecs + 1 THEN
        EXIT DO
      END IF
    END IF
    IF TopOffSet = BotOffSet THEN EXIT DO
    'Look into this
    'IF TopOffSet = BotOffSet THEN EXIT DO
  LOOP UNTIL FirstMatchRec

  FClose C1Handle
  FClose L1Handle

  IF FirstMatchRec = 0 THEN
    RecNo& = 0
  ELSE
    RecNo& = IdxBuff(FirstMatchRec).RecNum
  END IF

ExitLSearch:
  'SHELL
  RestScrn ScrnArray()

END SUB

SUB Search4Meter (MeterNum$, RecNo&, CLSFlag%)
  
  STATIC Choice, OMeterNum$
  
  IF OMeterNum$ <> MeterNum$ THEN
    Choice = 1
    OMeterNum$ = MeterNum$
  END IF
  
  REDIM ScrnArray(0)
  SaveScrn ScrnArray()
  
  WPos = 1
  
  DisplayUBScrn "SHOWSCRH"
  
  REDIM UBCustRec(1) AS UBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBLocatRec(1) AS UBLocationRecType
  UBLocatRecLen = LEN(UBLocatRec(1))
  
  LocatFileSize& = FileSize("UBLOCA.DAT")
  NumLocatRecs = LocatFileSize& \ UBLocatRecLen
  
  REDIM RecBuff(1 TO 1) AS LONG

  FOpenS "UBCUST.DAT", C1Handle 'open data file
  FOpenS "UBLOCA.DAT", L1Handle 'open data file

  'Find matching record

  MatchCnt = 0

  FOR Cnt = 1 TO NumLocatRecs
    ShowSearchWheel 12, 44
    FGetRTA L1Handle, UBLocatRec(1), CLNG(Cnt), UBLocatRecLen
    FOR MeterCnt = 1 TO 7
      IF INSTR(UBLocatRec(1).LocMeters(MeterCnt).MTRNum, MeterNum$) > 0 THEN
        MatchCnt = MatchCnt + 1
        REDIM PRESERVE RecBuff(1 TO MatchCnt) AS LONG
        RecBuff(MatchCnt) = Cnt
      END IF
    NEXT
  NEXT

  IF MatchCnt = 0 THEN
    RecNo& = 0
    FClose C1Handle
    FClose L1Handle
    GOTO ExitMeterSearch
  END IF

  FreeMem& = FRE(-1)
  MemNeeded& = MatchCnt * 64&
  IF MemNeeded& > FreeMem& THEN
    FClose C1Handle
    FClose L1Handle
    QPrintRC "Matched:>" + QPStrI(MatchCnt), 24, 1, 15
    QPrintRC " Needed: " + QPStrL(MemNeeded&) + "  Free: " + QPStrL(FreeMem&), 25, 1, 15
    RecNo& = -1
    WaitForAction
    GOTO ExitMeterSearch
  END IF

  REDIM MChoice(1 TO MatchCnt) AS FLen2

  FOR Cnt = 1 TO MatchCnt
    ShowSearchWheel 12, 44
    FGetRTA L1Handle, UBLocatRec(1), RecBuff(Cnt), UBLocatRecLen
    IF UBLocatRec(1).CustRec > 0 THEN
      FGetRTA C1Handle, UBCustRec(1), UBLocatRec(1).CustRec, UBCustRecLen
      TCustName$ = LEFT$(QPTrim$(UBCustRec(1).CUSTNAME), 30)
    ELSE
      TCustName$ = "NONE"
    END IF

    Cnt = Cnt - 1

    FOR MeterCnt = 1 TO 7
      IF INSTR(UBLocatRec(1).LocMeters(MeterCnt).MTRNum, MeterNum$) > 0 THEN
        Cnt = Cnt + 1
        LSET MChoice(Cnt).V = TCustName$
        Book$ = UBLocatRec(1).Book + "-" + UBLocatRec(1).SeqNumb
        MID$(MChoice(Cnt).V, 32, 12) = UBLocatRec(1).LocMeters(MeterCnt).MTRNum
        MID$(MChoice(Cnt).V, 50, 9) = UBLocatRec(1).Book + "-" + UBLocatRec(1).SeqNumb
        MID$(MChoice(Cnt).V, 61) = MKL$(RecBuff(Cnt))
      END IF
    NEXT

  NEXT
  
  FClose C1Handle
  FClose L1Handle
  
  IF MatchCnt = 0 THEN
    RecNo& = 0
  ELSEIF MatchCnt > 1 THEN
    
    QPrintRC "Sorting. . .  ", 12, 32, -1
    
    SortT MChoice(1), MatchCnt, 0, 64, 31, 12
    
    ' FOR PauCnt = 1 TO 100
    '   SmallPause
    ' NEXT
    
    MaxLen = 59 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something
    IF Choice = 0 THEN
      Choice = 1                'Pre-load choice to highlight
    ELSEIF Choice > MatchCnt THEN
      Choice = 1                'Pre-load choice to highlight
    END IF
    Title$ = SPACE$(MaxLen + 4)
    LSET Title$ = "  Customer/Owner               Meter No.         Location No."
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
  
ExitMeterSearch:
  'SHELL
  RestScrn ScrnArray()
  
END SUB

SUB Search4SAddr (SAddr$, RecNo&, CLSFlag%)
  
  STATIC Choice
  REDIM ScrnArray(0)
  SaveScrn ScrnArray()
  
  WPos = 1
  
  DisplayUBScrn "SHOWSCRH"
  
  REDIM UBCustRec(1) AS UBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBLocatRec(1) AS UBLocationRecType
  UBLocatRecLen = LEN(UBLocatRec(1))
  
  LocatFileSize& = FileSize("UBLOCA.DAT")
  NumLocatRecs = LocatFileSize& \ UBLocatRecLen
  
  REDIM RecBuff(1 TO 1) AS LONG
  
  FOpenS "UBCUST.DAT", C1Handle 'open data file
  FOpenS "UBLOCA.DAT", L1Handle 'open data file
  
  'Find matching record
  
  MatchCnt = 0
  FOR Cnt = 1 TO NumLocatRecs
    ShowSearchWheel 12, 44
    FGetRTA L1Handle, UBLocatRec(1), CLNG(Cnt), UBLocatRecLen
    IF INSTR(UBLocatRec(1).SERVADDR, SAddr$) > 0 THEN
      MatchCnt = MatchCnt + 1
      REDIM PRESERVE RecBuff(1 TO MatchCnt) AS LONG
      RecBuff(MatchCnt) = Cnt
    END IF
  NEXT
  
  FClose C1Handle
  FClose L1Handle

  MemNeeded& = MatchCnt * 64&
  FreeMem& = FRE(-1)
  IF MemNeeded& > FreeMem& THEN
    'BEEP
    QPrintRC "Matched: " + QPStrI(MatchCnt), 24, 1, 15
    QPrintRC " Needed: " + QPStrL(MemNeeded&) + "  Free: " + QPStrL(FreeMem&), 25, 1, 15
    'MatchCnt = 0
    RecNo& = -1
    WaitForAction
    GOTO ExitSAddrSearch
  END IF

  IF MatchCnt = 0 THEN
    
  ELSEIF MatchCnt > 1 THEN
    REDIM MChoice(1 TO MatchCnt) AS FLen2
    FOpenS "UBCUST.DAT", C1Handle               'open data file
    FOpenS "UBLOCA.DAT", L1Handle               'open data file
    FOR Cnt = 1 TO MatchCnt
      FGetRTA L1Handle, UBLocatRec(1), CLNG(RecBuff(Cnt)), UBLocatRecLen
      Addr$ = LEFT$(QPTrim$(UBLocatRec(1).SERVADDR), 25)
      Book$ = UBLocatRec(1).Book + "-" + UBLocatRec(1).SeqNumb
      IF UBLocatRec(1).CustRec > 0 THEN
        FGetRTA C1Handle, UBCustRec(1), UBLocatRec(1).CustRec, UBCustRecLen
        LSET MChoice(Cnt).V = LEFT$(QPTrim$(UBCustRec(1).CUSTNAME), 20)
      ELSE
        LSET MChoice(Cnt).V = "NONE"
      END IF
      MID$(MChoice(Cnt).V, 22, 25) = Addr$
      MID$(MChoice(Cnt).V, 50, 9) = Book$
      MID$(MChoice(Cnt).V, 61) = MKL$(RecBuff(Cnt))
      ShowSearchWheel 12, 44
    NEXT
    FClose C1Handle
    FClose L1Handle
    
    QPrintRC "Sorting. . .   ", 12, 32, -1
    SortT MChoice(1), MatchCnt, 0, 64, 21, 25

    ' FOR PauCnt = 1 TO 100
    '   SmallPause
    ' NEXT
    
    MaxLen = 59 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something
    IF Choice = 0 THEN
      Choice = 1                'Pre-load choice to highlight
    ELSEIF Choice > MatchCnt THEN
      Choice = 1                'Pre-load choice to highlight
    END IF
    Title$ = SPACE$(MaxLen + 4)
    LSET Title$ = "  Customer                      Address"
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
    RecNo& = RecBuff(MatchCnt)
  END IF
  
ExitSAddrSearch:
  'SHELL
  RestScrn ScrnArray()
  
  
END SUB

SUB ShowCustLocations (CustRecNo&)
  
  REDIM MChoice(1 TO 1) AS FLen2
  REDIM TempScrn(0)
  SaveScrn TempScrn()
  
  REDIM UBLocatRec(1) AS UBLocationRecType
  UBLocatRecLen = LEN(UBLocatRec(1))
  REDIM UBCustRec(1) AS UBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM AS UBFile LEN = UBCustRecLen
  GET UBFile, CustRecNo&, UBCustRec(1)
  CLOSE UBFile
  
Top:
  
  
  UBFile = FREEFILE
  OPEN "UBLOCA.DAT" FOR RANDOM AS UBFile LEN = UBLocatRecLen
  DCnt = 0
  IF UBCustRec(1).CustFrstLocRec > 0 THEN
    DCnt = DCnt + 1
    REDIM PRESERVE MChoice(1 TO DCnt) AS FLen2
    GET UBFile, UBCustRec(1).CustFrstLocRec, UBLocatRec(1)
    LSET MChoice(DCnt).V = UBLocatRec(1).Book + "-" + UBLocatRec(1).SeqNumb
    MID$(MChoice(DCnt).V, 22, 22) = LEFT$(QPTrim$(UBLocatRec(1).SERVADDR), 25)
    MID$(MChoice(DCnt).V, 59, 1) = UBLocatRec(1).LocatStatus
    MID$(MChoice(DCnt).V, 61) = MKL$(UBCustRec(1).CustFrstLocRec)
    
    NextCustLoc& = UBLocatRec(1).NextLocRec
    DO WHILE NextCustLoc& > 0
      DCnt = DCnt + 1
      REDIM PRESERVE MChoice(1 TO DCnt) AS FLen2
      GET UBFile, NextCustLoc&, UBLocatRec(1)
      LSET MChoice(DCnt).V = UBLocatRec(1).Book + "-" + UBLocatRec(1).SeqNumb
      MID$(MChoice(DCnt).V, 22, 22) = LEFT$(QPTrim$(UBLocatRec(1).SERVADDR), 25)
      
      MID$(MChoice(DCnt).V, 61) = MKL$(NextCustLoc&)
      MID$(MChoice(DCnt).V, 59, 1) = UBLocatRec(1).LocatStatus
      
      NextCustLoc& = UBLocatRec(1).NextLocRec
    LOOP
    
    CLOSE UBFile
    SortT MChoice(1), DCnt, 0, 64, 0, 9
         '(Element(1), NumElements%, Direction%, StructSize%, 0, MemberSize%)
    MPaintBox 3, 5, 22, 75, 8
    
    MaxLen = 59 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something
    
    IF Choice < 1 THEN
      Choice = 1                'Pre-load choice to highlight
    END IF
    
    Title$ = SPACE$(MaxLen + 4)
    'LSET Title$ = "  Customer/Owner       Service Address        Location No.  S"
    LSET Title$ = "  Location No.         Service Address        Location Status"
    '--Find max menu width
    
    '--Center Menu within Screen
    
    Row = 4
    Col = ((80 - 60) \ 2) - 1
    
    Row = 6
    BoxBot = 14 'limit the box length to go no lower than line 20
    
    LOCATE Row, Col, 0
    
    DO
      TitleBox BoxBot + 3, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select ", Cnf
      QPrintRC "Locations:" + STR$(DCnt), BoxBot + 4, Col + 2, 15
      QPrintRC Title$, Row - 1, Col, 112
      MPaintBox Row, Col + MaxLen + 4, Row, Col + MaxLen + 5, 8
      VertMenuT2 MChoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        EditLocRec& = -1
        RestScrn TempScrn()
        EXIT DO 'choice = 0
      END IF
      EditLocRec& = CVL(MID$(MChoice(Choice).V, 61, 4))
    LOOP UNTIL EditLocRec& > 0
  ELSE
    CLOSE UBFile
    EditLocRec& = 0
    OK = MsgBox%("UB.QSL", "NOCLOCAT")
    RestScrn TempScrn()
    'RecNo& = IdxBuff(FirstMatchRec).RecNum
  END IF
  IF EditLocRec& > 0 THEN
    REDIM TempScrn2(0)
    SaveScrn TempScrn2()
    BlockClear
    'AddEditLocation EditLocRec&
    RestScrn TempScrn2()
    UBFile = FREEFILE
    OPEN "UBCUST.DAT" FOR RANDOM AS UBFile LEN = UBCustRecLen
    GET UBFile, CustRecNo&, UBCustRec(1)
    CLOSE UBFile
    IF UBCustRec(1).CustFrstLocRec > 0 THEN
      GOTO Top
    ELSE
      RestScrn TempScrn()
    END IF
  END IF
  
END SUB

SUB ShowPctComp (BYVAL RecNo, BYVAL NumOfRecs) STATIC
  RSET PctC(1) = QPStrI$(INT((RecNo / NumOfRecs) * 100))
  QPrintRC PctC(1), 13, 40, Cnf.HiLite
  '  QPrintRC STR$(FRE("")), 25, 1, Cnf.HiLite
END SUB

SUB ShowProcessingScrn (RptTitle$)
  
  TitleRow = 9
  TitleCol = 40 - (LEN(RptTitle$) \ 2) + 1
  
  CursorOff
  BlockClear
  DisplayUBScrn "PRORPT"
  
  QPrintRC RptTitle$, TitleRow, TitleCol, 126
  QPrintRC "Processing:    % Completed.", 13, 28, Cnf.HiLite
  
END SUB

SUB ShowSearchWheel (BYVAL Row%, BYVAL Col%) STATIC
  WPos = WPos + 1
  IF WPos > 4 THEN WPos = 1
  QPrintRC MID$(Wheel$, WPos, 1), Row, Col, -1
  'smallpause
END SUB

SUB SmallPause
  st! = TIMER
  DO
  LOOP UNTIL TIMER > st! + .01
END SUB

