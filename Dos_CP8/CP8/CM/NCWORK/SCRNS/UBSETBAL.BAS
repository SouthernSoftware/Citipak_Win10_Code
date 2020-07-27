DEFINT A-Z
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB SearchGetCust (SEARCH$, RecNo%)
DECLARE SUB CursorOff ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB SmallPause ()
DECLARE SUB EditBalance ()
DECLARE SUB EnterNewBalance ()
DECLARE SUB PostBalanceTrans ()
DECLARE SUB PrintBalanceTrans ()
DECLARE SUB SaveNewBalanceRec (Form$(), CustRecNo%)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValI% (Number$)
DECLARE FUNCTION QPValL& (Number$)
DECLARE FUNCTION ASCII (Character$)
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION Round# (N#)
DECLARE SUB BlockClear ()

DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FClose (Handle%)
DECLARE SUB FCreate (FileName$)

DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB SortStr2 (BYVAL Address%, Size%, Direction%)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB WaitForAction ()
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE SUB ShowCursor ()
DECLARE SUB HideCursor ()

DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB Get.Moose.OR.Key (Ky$, MooseButton%, MRow%, MCol%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)


CONST False = 0, True = NOT False

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'Qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  ''$INCLUDE: 'ubfiles.BI'
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'ubdraft.BI'
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'PageInfo.bi'
  '$INCLUDE: 'ubcust.bi'
  '$INCLUDE: 'ubloc.bi'
  '$INCLUDE: 'ubrate.BI'
  '$INCLUDE: 'ubsetbal.BI'

SUB EditBalance
  
  REDIM BalanceRec(1) AS UBBalanceRecType
  BalRecLen = LEN(BalanceRec(1))
  
'  REDIM UBCustRec(1) AS UBCustRecType
'  CustRecLen = LEN(UBCustRec(1))
  
  NumBalRecs = FileSize("UBALANCE.DAT") \ BalRecLen
  
  Temp5$ = SPACE$(5)
  Temp8$ = SPACE$(8)
  
  REDIM MChoice$(1 TO NumBalRecs)
  UBFile = FREEFILE
  OPEN "UBALANCE.DAT" FOR RANDOM AS UBFile LEN = BalRecLen
  FOR Cnt = 1 TO NumBalRecs
    GET UBFile, Cnt, BalanceRec(1)
    MChoice$(Cnt) = SPACE$(55)
    Temp5$ = FUsing$(STR$(BalanceRec(1).CustRecNo), "#####")
    Temp8$ = FUsing$(STR$(BalanceRec(1).REVTOTAL), "#####.##")
    LSET MChoice$(Cnt) = Temp5$
    MID$(MChoice$(Cnt), 10, 30) = BalanceRec(1).CustName
    MID$(MChoice$(Cnt), 43) = Temp8$
    MID$(MChoice$(Cnt), 54) = MKI$(Cnt)
  NEXT
  CLOSE
  
  SortStr2 VARPTR(MChoice$(1)), NumBalRecs, 0
  
  MaxLen = 52   'Set menu width to
  BoxBot = 17   'limit the box length to go no lower than line 17
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  TText$ = SPACE$(MaxLen + 4)
  LSET TText$ = " Acct No.      Customer Name            Balance"
  
  '--Center Menu within Screen
  Row = 8
  Col = (80 - MaxLen) \ 2
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    'DisplayUBScrn "MENUBAK"
    BlockClear
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    ShowCursor
    QPrintRC TText$, Row - 1, Col, 112
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    IF Ky$ = CHR$(27) THEN      'user pressed escape (no choice)
      ExitFlag = True           'set exit flag
    ELSE
      OKFlag = True
    END IF
    
  LOOP UNTIL ExitFlag OR OKFlag
  IF ExitFlag THEN GOTO ExitBalanceEdit
  
  BalRecNo = CVI(RIGHT$(MChoice$(Choice), 2))
  
  
  LibName$ = "UB"
  ScrnName$ = "UBSETBAL"
  
  '--Initialize the form name array
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1
  
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
  
  '--display screen
  DisplayUBScrn ScrnName$
  
  ShowCursor
  
  Action = 1
  FirstTime = True
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False         'This can only be done after
      GOSUB GetEditBalRec       'first entry into editform
      Action = 2
    END IF
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE -68    'F10Key
      BegBalance# = Value(Form$(1, 0), ErrCode)
      TotalRev# = Value(Form$(13, 0), ErrCode)
      IF BegBalance# <> TotalRev# THEN          'if revenues are out of balance
        OK = MsgBox%("UB.QSL", "OUTOFBAL")      'then tell the user
        Action = 2              'Must reset Action after MsgBox function.
      ELSE      'else it's in balance
        GOSUB SaveEditBalanceRec                'go save it
        ExitFlag = True
      END IF
    CASE 27     'user pressed escape
      IF Frm(1).Edited THEN     'if the form was edited then ask
        OK = MsgBox%("UB.QSL", "ABANDON")       'user if they want to abandon
        IF OK = 1 THEN          'ok abandon edit
          ExitFlag = True
        END IF
      ELSE      'else nothing was edited let'em out
        ExitFlag = True
      END IF
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 21   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 43 TO 54           '--Save Button
          PressButton -68, 21, 43, 54
        CASE 62 TO 72           '--cancel button
          PressButton 27, 21, 57, 70
        END SELECT
      END SELECT                'row
    END IF
    
  LOOP UNTIL ExitFlag
  
ExitBalanceEdit:
  
  HideCursor
  EXIT SUB
  
GetEditBalRec:
  'NOTE:
  'EditForm must have been called once prior to doing this
  'so the Form(0,0) element will have the correct length.
  'else memory will be corrupted the system WILL CRASH!!!

  REDIM UBCustRec(1) AS UBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBLocaRec(1) AS UBLocationRecType
  UBLocaRecLen = LEN(UBLocaRec(1))
  
  FormLen = LEN(Form$(0, 0))
  
  FOpenS "UBALANCE.DAT", Handle 'open data file
  FGetRTA Handle, BalanceRec(1), CLNG(BalRecNo), BalRecLen
  FClose Handle
  
  CustRecNo = BalanceRec(1).CustRecNo
  FOpenS "UBCUST.DAT", Handle   'open data file
  FGetRTA Handle, UBCustRec(1), CLNG(CustRecNo), UBCustRecLen
  FClose Handle

  FOpenS "UBLOCA.DAT", Handle   'open data file
  FGetRTA Handle, UBLocaRec(1), UBCustRec(1).LocationRec, UBLocaRecLen
  FClose Handle
  
  BCopy VARSEG(BalanceRec(1)), VARPTR(BalanceRec(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), FormLen, 0
  
  CALL UnPackBuffer(0, 0, Form$(), Fld())
  
  GOSUB EditDisplayInfo
  RETURN
  
EditDisplayInfo:
  REDIM UBSetUp(1) AS UBSetupRecType
  UBSetupLen = LEN(UBSetUp(1))
  
  FOpenS "UBSETUP.DAT", Handle  'open data file
  FGetRTA Handle, UBSetUp(1), 1&, UBSetupLen
  FClose Handle
  
  StartOffSet = 6
  FOR Cnt = 1 TO 10
    RevenueText$ = QPTrim$(UBSetUp(1).Revenues(Cnt).REVNAME)
    IF LEN(RevenueText$) = 0 THEN
      EXIT FOR
    END IF
    QPrintRC RevenueText$, StartOffSet + Cnt, 47, -1
  NEXT
  'if all 10 rev's are used the cnt will = 11
  IF Cnt < 11 THEN              'if they don't have 10 rev sources
    StartOffSet = 2             'then we will protect the remaining
    Last = Cnt  'revenue fields on the form
    FOR Cnt = Last TO 10        '
      Fld(Cnt + StartOffSet).Protected = True
    NEXT
  END IF
  
  ERASE UBSetUp
  
  QPrintRC QPTrim$(STR$(BalanceRec(1).CustRecNo)), 5, 19, -1
  QPrintRC QPTrim$(UBLocaRec(1).BOOK + "-" + UBLocaRec(1).SEQNUMB), 6, 19, -1
  QPrintRC QPTrim$(UBCustRec(1).CustName), 7, 19, -1
  QPrintRC QPTrim$(UBCustRec(1).Status), 8, 19, -1
  
  Action = 2
  
  RETURN
  
SaveEditBalanceRec:
  
  FormLen = LEN(Form$(0, 0))
  FOpenS "UBCUST.DAT", Handle   'open data file
  FGetRTA Handle, UBCustRec(1), CLNG(CustRecNo), CustRecLen
  FClose Handle
  
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(BalanceRec(1)), VARPTR(BalanceRec(1)), FormLen, 0
  BalanceRec(1).CustRecNo = CustRecNo
  BalanceRec(1).CustName = UBCustRec(1).CustName
  
  FOpenS "UBALANCE.DAT", Handle 'open data file
  FPutRTA Handle, BalanceRec(1), CLNG(Choice), BalRecLen
  FClose Handle
  
  RETURN
  
END SUB

SUB EnterNewBalance

  REDIM UBCustRec(1) AS UBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBLocaRec(1) AS UBLocationRecType
  UBLocaRecLen = LEN(UBLocaRec(1))
  
  REDIM ScrnArray(0)
  
  REDIM LText$(4)
  
  LScrn = 1
  MScrn = 2
  
  LText$(1) = " Account Number:"
  LText$(2) = "    Search Name:"
  
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
  
  'DisplayUBScrn "MENUBAK"
  BlockClear
  DisplayUBScrn ScrnName$
  
  GOSUB DisplayLookupText
  
  ShowCursor
  
  Action = 1
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE -68, 13   'F10Key
      SEARCH$ = QPTrim$(Form$(0, 0))
      SearchAcct& = QPValL(SEARCH$)
      SELECT CASE LScrn
      CASE 1    'Customer account lookup
        IF SearchAcct& < 1 OR SearchAcct& > GetNumOfAcct THEN
          OK = MsgBox%("UB.QSL", "BADACCTN")
          Action = 2
        ELSE
          CustRecNo = SearchAcct&
          OKFlag = True
        END IF
      CASE 2    'Customer Name lookup
        IF LEN(SEARCH$) < 1 THEN
          OK = MsgBox%("UB.QSL", "BADSEARH")
          Action = 1
          Frm(1).FldNo = 1
        ELSE
          SearchGetCust SEARCH$, CustRecNo
          IF CustRecNo > 0 THEN
            OKFlag = True
          ELSEIF CustRecNo = 0 THEN
            OK = MsgBox%("UB.QSL", "NOMATCH")
            Action = 2
          END IF
        END IF
      END SELECT
    CASE -65    'F7Key     toggle thru the lookup selections
      IF LScrn < MScrn THEN
        LScrn = LScrn + 1
      ELSE
        LScrn = 1
      END IF
      LSET Form$(1, 0) = ""
      Action = 1
      SaveField 0, Form$(), Fld(), BadField
      GOSUB DisplayLookupText
    CASE 27
      CustRecNo = 0
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
  
  IF ExitFlag GOTO ExitSetBalance
  
  LibName$ = "UB"
  ScrnName$ = "UBSETBAL"
  
  '--Get the number of fields on this form
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1
  
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
  
  GOSUB GetCustInfo
  
  '--Set screen number to one and display screen
  DisplayUBScrn ScrnName$
  
  ShowCursor
  
  Action = 1
  FirstTime = True
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN           'if this is the first time
      FirstTime = False         'go and display the misc info
      GOSUB DisplayInfo         'not attached directly to the form
    END IF
    
    SELECT CASE Frm(1).KeyCode  '--Check for Key presses
    CASE -68    'F10Key
      BegBalance# = Value(Form$(1, 0), ErrCode)
      TotalRev# = Value(Form$(13, 0), ErrCode)
      'compare the total revenues and entered beginning balance
      IF BegBalance# <> TotalRev# THEN          'if not in balance then
        OK = MsgBox%("UB.QSL", "OUTOFBAL")      'ask user to abandon
        Action = 2              'Must reset formedits action variable
      ELSEIF BegBalance# <= 0 THEN              'after using the MsgBox function.
        OK = MsgBox%("UB.QSL", "NOBALNCE")      'ask user to abandon
        Action = 2              'Must reset formedits action variable
      ELSE
        SaveNewBalanceRec Form$(), CustRecNo
        ExitFlag = True
      END IF
    CASE 27     'Escape key pressed
      IF Frm(1).Edited THEN     'if the form has been edited then
        OK = MsgBox%("UB.QSL", "ABANDON")       'ask user if they want
        IF OK = 1 THEN          'to abandon edit
          ExitFlag = True
        ELSE    'Must reset formedits action variable
          Action = 2            'after using the MsgBox function.
        END IF
      ELSE
        ExitFlag = True
      END IF
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 21   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 43 TO 54           '--Save Button
          PressButton -68, 21, 43, 54
        CASE 62 TO 72           '--cancel button
          PressButton 27, 21, 57, 70
        END SELECT
      END SELECT
    END IF
    
  LOOP UNTIL ExitFlag
  
  HideCursor
  ERASE UBCustRec, Frm, Form$, Fld
  
  EXIT SUB
  
GetCustInfo:
  FOpenS "UBCUST.DAT", C1Handle 'open data file
  FGetRTA C1Handle, UBCustRec(1), CLNG(CustRecNo), UBCustRecLen
  FClose C1Handle

  FOpenS "UBLOCA.DAT", Handle   'open data file
  FGetRTA Handle, UBLocaRec(1), UBCustRec(1).LocationRec, UBLocaRecLen
  FClose Handle

  RETURN
  
DisplayInfo:
  
  REDIM UBSetUp(1) AS UBSetupRecType
  UBSetupLen = LEN(UBSetUp(1))
  
  FOpenS "UBSETUP.DAT", Handle  'open data file
  FGetRTA Handle, UBSetUp(1), 1&, UBSetupLen
  FClose Handle
  
  StartOffSet = 6
  FOR Cnt = 1 TO 10
    RevenueText$ = QPTrim$(UBSetUp(1).Revenues(Cnt).REVNAME)
    IF LEN(RevenueText$) = 0 THEN
      EXIT FOR
    END IF
    QPrintRC RevenueText$, StartOffSet + Cnt, 47, -1
  NEXT
  
  IF Cnt < 11 THEN              'if they don't have 10 rev sources
    StartOffSet = 2             'last field on form prior to revenues
    Last = Cnt  'start with next
    FOR Cnt = Last TO 10        'go from last revenue to end of revenues
      Fld(Cnt + StartOffSet).Protected = True   'protect the remainder
    NEXT
  END IF
  ERASE UBSetUp
  
  QPrintRC QPTrim$(STR$(CustRecNo)), 5, 19, -1
  QPrintRC QPTrim$(UBLocaRec(1).BOOK + "-" + UBLocaRec(1).SEQNUMB), 6, 19, -1
  QPrintRC QPTrim$(UBCustRec(1).CustName), 7, 19, -1
  Form$(2, 0) = DATE$
  SaveField 0, Form$(), Fld(), BadField
  Action = 2
  
  RETURN
  
ExitSetBalance:
  EXIT SUB
  
DisplayLookupText:
  QPrintRC LText$(LScrn), 12, 22, -1
  RETURN
  
END SUB

SUB PostBalanceTrans
  
  CursorOff
  
  ActMRow = 16
  BlockClear
  DisplayUBScrn "RDY2POST"
  
  DO
    
    Get.Moose.OR.Key Ky$, MooseButton, MRow, MCol
    
    IF MooseButton THEN
      MRow = (MRow \ 8) + 1     'Convert MCol, MRow to Row and Col.
      MCol = (MCol \ 8) + 1
      IF MRow = ActMRow THEN
        SELECT CASE MCol
        CASE 30 TO 41           'Cancel posting
          PressButton EscKey, ActMRow, 30, 41
        CASE 43 TO 53           'let it rip
          PressButton ASCII("D"), ActMRow, 43, 53
        END SELECT
      END IF
    END IF
    
    IF LEN(Ky$) THEN
      
      Choice = INSTR(1, CHR$(0) + "D" + CHR$(27), Ky$)
      SELECT CASE Choice
      CASE 1, 2
        OKFlag = True
      CASE 3
        BlockClear
        DisplayUBScrn "POSTCAN"
        WaitForAction
        ExitFlag = True
      END SELECT
    END IF
  LOOP UNTIL ExitFlag OR OKFlag
  
  IF ExitFlag THEN GOTO ExitPost
  
  ShowProcessingScrn "Posting Transactions"
  
  REDIM BalanceRec(1) AS UBBalanceRecType
  REDIM UBTransRec(1) AS UBTransRecType
  REDIM UBCustRec(1) AS UBCustRecType
  
  BalRecLen = LEN(BalanceRec(1))
  TransRecLen = LEN(UBTransRec(1))

  CustRecLen = LEN(UBCustRec(1))
  
  NumBalRecs = FileSize("UBALANCE.DAT") \ BalRecLen
  NextTransRec& = (FileSize("UBTRANS.DAT") \ TransRecLen) + 1
  
  BHandle = FREEFILE
  OPEN "UBALANCE.DAT" FOR RANDOM LOCK READ WRITE AS BHandle LEN = BalRecLen
  THandle = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM LOCK READ WRITE AS THandle LEN = TransRecLen
  CHandle = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CHandle LEN = CustRecLen
  
  FOR Cnt = 1 TO NumBalRecs
    GET BHandle, Cnt, BalanceRec(1)                    'get customers balance
    GET CHandle, BalanceRec(1).CustRecNo, UBCustRec(1) 'get customers Data
    
    'IF UBCustRec(1).LastTrans > 0 THEN    'if they have a previous trans
    '  GET THandle, UBCustRec(1).LastTrans, UBTransRec(1)
    'END IF
    
    REDIM UBTempTransRec(1) AS UBTransRecType   'Make Temp blank trans rec
    
    UBTempTransRec(1).TransDate = Date2Num(DATE$)
    UBTempTransRec(1).TransType = 1             'Beginning Balance trans Type
    UBTempTransRec(1).TransDesc = "Beginning Balance"
    UBTempTransRec(1).TransAmt = BalanceRec(1).REVTOTAL
    UBTempTransRec(1).CustAcctNo = BalanceRec(1).CustRecNo
    'UBTempTransRec(1).RunBalance = Round(UBTempTransRec(1).RunBalance + BalanceRec(1).REVTOTAL)
    FOR RCnt = 1 TO 10
      UBTempTransRec(1).RevAmt(RCnt) = BalanceRec(1).RevAmts(RCnt)
    NEXT
    'UBTempTransRec(1).CustLocation = UBCustRec(1).BOOK + "-" + UBCustRec(1).SEQNUMB
    UBTempTransRec(1).CustStatus = UBCustRec(1).Status
    'UBTempTransRec(1).CustZone = UBCustRec(1).ZONE
    
   ' IF UBCustRec(1).LastTrans > 0 THEN          'if they have a previous trans
   '   UBTempTransRec(1).PrevTrans = UBCustRec(1).LastTrans
   ' ELSE
   '   UBTempTransRec(1).PrevTrans = 0
   ' END IF
    'UBCustRec(1).LastTrans = NextTransRec&
    
    PUT CHandle, BalanceRec(1).CustRecNo, UBCustRec(1)
    'update the customers info
    PUT THandle, NextTransRec&, UBTempTransRec(1)
    'update transaction file
    
    SmallPause
    
    NextTransRec& = NextTransRec& + 1
    ShowPctComp Cnt, NumBalRecs 'show user percentage complete
    
  NEXT
  
  CLOSE BHandle, CHandle, THandle
  
  BlockClear
  DisplayUBScrn "UPDATEOK"
  WaitForAction
  
ExitPost:
END SUB

SUB PrintBalanceTrans
  ReportTitle$ = "Print Beginning Balance Report"
  ShowProcessingScrn ReportTitle$
  
  FF$ = CHR$(12)
  
  REDIM BalanceRec(1) AS UBBalanceRecType
  BalRecLen = LEN(BalanceRec(1))
  NumBalRecs = FileSize("UBALANCE.DAT") \ BalRecLen
  
  REDIM SortArray(1 TO NumBalRecs)  AS BalanceSortType
  
  Handle = FREEFILE
  OPEN "UBALANCE.DAT" FOR RANDOM AS Handle LEN = BalRecLen
  FOR Cnt = 1 TO NumBalRecs
    GET Handle, Cnt, BalanceRec(1)              'get this customers data
    SortArray(Cnt).AcctNo = BalanceRec(1).CustRecNo
    SortArray(Cnt).RecNo = Cnt
    ShowPctComp Cnt, NumBalRecs 'show user percentage complete
    IF NumBalRecs < 50 THEN
      SmallPause
    END IF
  NEXT
  CLOSE Handle
  
  SortT SortArray(1), NumBalRecs, 0, 4, 0, -1
  
  MaxLines = 60
  
  REDIM UBCustRec(1) AS UBCustRecType
  CustRecLen = LEN(UBCustRec(1))
  
  REDIM Dash(1) AS STRING * 80
  Dash$(1) = STRING$(80, "=")
  
  RptDate$ = DATE$
  ReportFile$ = "BBALANCE.RPT"
  
  RptHandle = FREEFILE
  
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  Handle = FREEFILE
  OPEN "UBALANCE.DAT" FOR RANDOM AS Handle LEN = BalRecLen
  
  Temp14$ = SPACE$(14)
  
  GOSUB PrintRptHeader
  
  FOR Cnt = 1 TO NumBalRecs
    GET Handle, SortArray(Cnt).RecNo, BalanceRec(1)             'get this customers data
    LSET Temp14$ = FUsing(STR$(BalanceRec(1).CustRecNo), "#######")
    PRINT #RptHandle, Temp14$; BalanceRec(1).CustName; TAB(50); USING "$####,.##"; BalanceRec(1).REVTOTAL
    
    LineCnt = LineCnt + 1       'Print data to report file and
    CustomerCnt = CustomerCnt + 1               'increment counters
    IF LineCnt >= MaxLines THEN 'If were at maxlines on this page
      PRINT #RptHandle, FF$     'formfeed to new page
      GOSUB PrintRptHeader      'print new page header info
    END IF
    IF NumBalRecs < 50 THEN
      SmallPause
    END IF
    ShowPctComp Cnt, NumBalRecs 'show user percentage complete
  NEXT          'until all are printed
  
  GOSUB PrintRtpEnd             'print report totals info
  
  CLOSE RptHandle, Handle       'Close files
  ERASE Dash, BalanceRec, SortArray             'Clean up
  
  EntryCode = 1 'Let user choose screen, printer or abort
  PrintRptFile ReportTitle$, ReportFile$, 1, RetCode, EntryCode
  
  GOTO ExitPrintListing         'DONE
  
PrintRptHeader:                 'prints report Title here
  page = page + 1               'Page Counter
  PRINT #RptHandle, TAB(27); "Beginning Balance Report"
  PRINT #RptHandle, "Report Date: "; RptDate$; TAB(70); "Page #"; page
  PRINT #RptHandle, "Account #"; TAB(14); "Customer Account Name"; TAB(50); "Balance"
  PRINT #RptHandle, Dash(1)
  LineCnt = 4   'reset number of lines printed counter
  RETURN
  
PrintRtpEnd:
  PRINT #RptHandle, Dash(1)
  PRINT #RptHandle, "Total Customers Listed = "; USING "######,"; CustomerCnt
  PRINT #RptHandle, FF$
  RETURN
  
ExitPrintListing:
  
END SUB

SUB SaveNewBalanceRec (Form$(), CustRecNo)
  
  REDIM BalanceRec(1) AS UBBalanceRecType
  REDIM UBCustRec(1) AS UBCustRecType
  
  FormLen = LEN(Form$(0, 0))
  BalRecLen = LEN(BalanceRec(1))
  CustRecLen = LEN(UBCustRec(1))
  
  FOpenS "UBCUST.DAT", Handle   'open data file
  FGetRTA Handle, UBCustRec(1), CLNG(CustRecNo), CustRecLen
  FClose Handle
  
  NumBalRecs = FileSize("UBALANCE.DAT") \ BalRecLen
  NextRec = NumBalRecs + 1
  IF NextRec = 1 THEN
    FCreate "UBALANCE.DAT"
  END IF
  
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(BalanceRec(1)), VARPTR(BalanceRec(1)), FormLen, 0
  BalanceRec(1).CustRecNo = CustRecNo
  BalanceRec(1).CustName = UBCustRec(1).CustName
  
  FOpenS "UBALANCE.DAT", Handle 'open data file
  FPutRTA Handle, BalanceRec(1), CLNG(NextRec), BalRecLen
  FClose Handle
  
END SUB

SUB SetBeginningBalance
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 4)
  
  MChoice$(1) = "Enter New Balance Transaction"
  MChoice$(2) = "Edit a Balance Transaction"
  MChoice$(3) = "Print Balance Transaction"
  MChoice$(4) = "Post Balance Transactions"
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 20   'limit the box length to go no lower than line 20
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  '--Find max menu width
  FOR Cnt = 1 TO UBOUND(MChoice$)
    TLen = LEN(MChoice$(Cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT
  
  '--Center Menu within Screen
  Row = ((25 - (UBOUND(MChoice$))) \ 2) - 1
  Col = ((80 - MaxLen) \ 2) - 2
  Help$ = "Misc Function Menu"
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    'LibFile2Scrn "GL.QSL", "MENUBAK", MonoCode, -1, ErrorCode
    'DisplayUBScrn "MENUBAK"
    BlockClear
    TitleBox 3, Col, MaxLen + 3, "Beginning Balance Menu ", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    'PrintTitle USER$
    'PrintHelp Help$
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      EnterNewBalance
    CASE 2
      EditBalance
    CASE 3
      PrintBalanceTrans
    CASE 4
      PostBalanceTrans
    END SELECT
  LOOP
  
END SUB

