DEFINT A-Z
DECLARE SUB ShowTransFixDate (RecNo&)
DECLARE FUNCTION nGetPassword$ (Row%, Col%, Max%, Abort%)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE SUB QPSound (Frequency%, Duration%)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (TDate$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE SUB BlockClear ()
DECLARE SUB CursorOff ()
DECLARE SUB HideCursor ()
DECLARE SUB LookUp (CustAcct&, Text$, Parm1, Parm2, ActiveOnly%)
DECLARE SUB PickList (Items() AS ANY, Picked%(), NPicked%, Cnf AS ANY)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB ShowTransAndWackEM (RecNo&)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB WaitForAction ()
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)

  '$INCLUDE: 'DefCnf.BI'

  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'ubdraft.BI'
  '$INCLUDE: 'ubsetup.bi'
  
  '$INCLUDE: 'PageInfo.bi'
  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubrate.BI'
  
  TYPE FLen2
    V AS STRING * 64
  END TYPE
  
  REDIM Items(1) AS FLen2
  'REDIM MChoice(1 TO 1)  AS FLen2
  REDIM Picked(1) AS INTEGER
  
  CONST False = 0, True = NOT False

'*****************************************
  CursorOff
  IF INSTR(COMMAND$, "TEST") > 0 THEN
    GOTO StartHere
  END IF

  DO
    BlockClear
    WazzWind 10, 15, 14, 65, 14, 1, 1
    QPrintRC "[ Transactions Edit ]", 10, 18, -1

    PW$ = nGetPassword(12, 32, 10, Abort)

    IF Abort THEN GOTO GetOut

    IF PW$ = "SOSOFT" THEN
      OKFlag = True
      EXIT DO
    END IF

    Try = Try + 1

    IF Try = 3 THEN
      ExitFlag = True
      BlockClear
      WazzWind 8, 15, 16, 65, 15, 1, 1
      QPrintRC "Get the password from the system operator.", 11, 20, 14
      QPrintRC "            Press any key.", 13, 20, 14
      WaitForAction
      EXIT DO
    END IF

    QPrintRC "Invalid Password. Press any key to try again.", 12, 18, 31

    QPSound 1000, 2
    QPSound 800, 2

    WaitForAction

    IF K = 27 THEN
      GOTO GetOut
    END IF
  LOOP

  IF NOT OKFlag THEN GOTO GetOut

StartHere:
'***************************************
  REDIM MChoice$(1 TO 3)

  MChoice$(1) = " Remove Transactions"
  MChoice$(2) = " Change Transaction Date "
  MChoice$(3) = " Exit to OS"

  MaxLen = 0    'Set menu width to zero
  BoxBot = 18   'limit the box length to go no lower than line 18
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight

  '--Find max menu width
  FOR cnt = 1 TO UBOUND(MChoice$)
    TLen = LEN(MChoice$(cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT

  '--Center Menu within Screen
  Row = ((25 - (UBOUND(MChoice$))) \ 2) - 1
  Col = ((80 - MaxLen) \ 2) - 2

  DO

    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear

    TitleBox 3, Col, MaxLen + 3, "Transaction Utility Menu ", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

    'ShowCursor

    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf

    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0

    SELECT CASE Choice
    CASE 1
      DO
        BlockClear
        LookUp RecNo&, "Trans Wack", 2, True, False
        IF RecNo& > 0 THEN
          ShowTransAndWackEM RecNo&
        END IF
      LOOP WHILE RecNo& > 0

    CASE 2
      DO
        BlockClear
        LookUp RecNo&, "Trans Date", 2, True, False

        IF RecNo& > 0 THEN
          ShowTransFixDate RecNo&
        END IF
      LOOP WHILE RecNo& > 0

    CASE 3
      EXIT DO
    END SELECT
  LOOP

GetOut:
  HideCursor
  CLS
  END

FUNCTION nGetPassword$ (Row%, Col%, Max%, Abort%)
  
Start:
  
  CurPos = Col

  QPrintRC "Password:", 12, 22, 15

  DO
    
    LOCATE Row, CurPos, 1
    P$ = INPUT$(1)
    
    SELECT CASE P$
      
    CASE CHR$(8)                'Backspace
      CurPos = CurPos - 1
      
      IF CurPos <= Col THEN     'Check for backspace key at
        LOCATE Row, Col         'start of input
        QPrintRC " ", Row, Col, 8
        PW$ = ""
        GOTO Start
      END IF
      
      LOCATE Row, CurPos
      QPrintRC " ", Row, CurPos, 8
      
    CASE CHR$(13)               'Enter
      EXIT DO
      
    CASE CHR$(27)               'Esc
      Abort = True
      EXIT FUNCTION
      
    CASE CHR$(32) TO CHR$(126)  'ignore all other keys
      QPrintRC "*", Row, CurPos, 8
      LOCATE Row, CurPos        'except keyboard chars
      PW$ = PW$ + P$
      CurPos = CurPos + 1
      
    END SELECT
    
  LOOP UNTIL LEN(PW$) = Max
  
  nGetPassword$ = UCASE$(PW$)
  CursorOff

END FUNCTION

SUB ShowTransAndWackEM (RecNo&)
  
  REDIM UBTranRec(1) AS UBTransRecType
  REDIM UBCustRec(1) AS NewUBCustRecType
  
  UBCustRecLen = LEN(UBCustRec(1))
  UBTranRecLen = LEN(UBTranRec(1))
  
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  GET UBFile, RecNo&, UBCustRec(1)
  CLOSE UBFile
  
  CurBal# = UBCustRec(1).CurrBalance
  PreBal# = UBCustRec(1).PrevBalance
  
Top:
  
  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen
  
  PrevTranRec& = UBCustRec(1).LastTrans
  
  IF PrevTranRec& > 0 THEN
    DO WHILE PrevTranRec& > 0
      DCnt = DCnt + 1
      REDIM PRESERVE MTChoice(1 TO DCnt) AS FLen2
      REDIM PRESERVE RecList(1 TO DCnt) AS LONG
      GET UBTran, PrevTranRec&, UBTranRec(1)
      LSET MTChoice(DCnt).V = Num2Date(UBTranRec(1).TransDate)
      MID$(MTChoice(DCnt).V, 15) = LEFT$(UBTranRec(1).TransDesc, 15)
'      MID$(MTChoice(DCnt).V, 30) = STR$(PrevTranRec&)
      MID$(MTChoice(DCnt).V, 40) = FUsing(STR$(UBTranRec(1).TransAmt), "#####.##")
      MID$(MTChoice(DCnt).V, 50) = FUsing(STR$(UBTranRec(1).RunBalance), "#####.##")
      RecList(DCnt) = PrevTranRec&
      PrevTranRec& = UBTranRec(1).PrevTrans
    LOOP
    
    CLOSE UBTran
    
    
HistTop:
    
    MaxLen = 59 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something
    
    IF Choice < 1 THEN
      Choice = 1                'Pre-load choice to highlight
    END IF
    
    Title$ = SPACE$(MaxLen + 9)
    LSET Title$ = "  Customer/Owner       Service Address        Location No.  S"
    LSET Title$ = " " + LEFT$(QPTrim$(UBCustRec(1).CustName), 20)
    MID$(Title$, 25) = LEFT$(QPTrim$(UBCustRec(1).SERVADDR), 25)
    MID$(Title$, 56, 9) = UBCustRec(1).BOOK + "-" + UBCustRec(1).SEQNUMB
    'MID$(Title$, 59, 1) = UBCustRec(1).STATUS
    
    '--Find max menu width
    '--Center Menu within Screen
    
    Col = ((80 - 60) \ 2) - 1
    
    Row = 4
    BoxBot = 14 'limit the box length to go no lower than line 20
    LOCATE 4, 1
    
    WazzWind 19, Col - 2, 23, Col + 65, Cnf.MenBox, 1, 1
    
    QPrintRC Title$, Row - 2, Col - 2, 112

    QPrintRC "PRESS <SPACE-BAR> to Select/UnSelect", 20, Col, 15
    QPrintRC "PRESS <F10> to Process Selections     <Arrow Keys> to Navigate.", 21, Col, 15
    QPrintRC "PRESS <ESCAPE> to Cancel Selections", 22, Col, 15

    REDIM Picked(DCnt)  AS INTEGER
    PickList MTChoice(), Picked(), DCnt, Cnf
  ELSE
    OK = MsgBox%("UB.QSL", "NOCTRANS")
  END IF

  IF DCnt > 0 THEN
    Count$ = QPTrim$(STR$(DCnt))
    IF DCnt < 10 THEN
      Count$ = " " + Count$
    END IF
'************************************
    LibName$ = "UBWACK"
    KillFlag = False

    FormName$ = "CANTUNDO"
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
    BlockClear
    LibFile2Scrn LibName$, FormName$, MonoCode%, Attribute%, ErrorCode%
    QPrintRC LEFT$(UBCustRec(1).CustName, 34), 10, 34, 78
    QPrintRC Count$, 8, 37, 78

    DO
      EditForm Form$(), Fld(), Frm(1), Cnf, Action
      SELECT CASE Frm(1).KeyCode
      CASE F0Key
        KillFlag = True
        EXIT DO
      CASE EscKey
        KillFlag = False
        EXIT DO
      END SELECT
    LOOP

    IF NOT KillFlag THEN
      CursorOff
      BlockClear
      LibFile2Scrn LibName$, "PRORPT", MonoCode%, Attribute%, ErrorCode%
      QPrintRC "REMOVE TRANSACTIONS ABORTED.", 10, 27, -1
      QPrintRC "Press any key to continue.", 13, 28, -1
      WaitForAction
    ELSE   'remove the transactions
      GOSUB WackEM
    END IF
'**************************
  
  END IF

  EXIT SUB

WackEM:

  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  GET UBFile, RecNo&, UBCustRec(1)
  UBCustRec(1).LastTrans = 0
  PUT UBFile, RecNo&, UBCustRec(1)

  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen

  FOR ZZ = 1 TO DCnt
    TransRec& = RecList(Picked(ZZ))
    GET UBTran, TransRec&, UBTranRec(1)
    UBTranRec(1).CustAcctNo = -UBTranRec(1).CustAcctNo
    UBTranRec(1).PrevTrans = 0
    PUT UBTran, TransRec&, UBTranRec(1)
    RecList(Picked(ZZ)) = -1
  NEXT

  LOCATE 1, 1
  FOR ZZ = UBOUND(RecList) TO 1 STEP -1
    IF RecList(ZZ) > 0 THEN
      GET UBFile, RecNo&, UBCustRec(1)
      GET UBTran, RecList(ZZ), UBTranRec(1)
      UBTranRec(1).PrevTrans = UBCustRec(1).LastTrans
      UBCustRec(1).LastTrans = RecList(ZZ)
      PUT UBFile, RecNo&, UBCustRec(1)
      PUT UBTran, RecList(ZZ), UBTranRec(1)
      'PRINT RecList(ZZ)
      'get this trans
    END IF
  NEXT
  CLOSE

  CursorOff
  BlockClear
  LibFile2Scrn LibName$, "PRORPT", MonoCode%, Attribute%, ErrorCode%
  QPrintRC "TRANSACTIONS REMOVED", 10, 30, -1
  QPrintRC "Press any key to continue.", 13, 28, -1
  WaitForAction

END SUB

SUB ShowTransFixDate (RecNo&)

  REDIM UBTranRec(1) AS UBTransRecType
  REDIM UBCustRec(1) AS NewUBCustRecType

  UBCustRecLen = LEN(UBCustRec(1))
  UBTranRecLen = LEN(UBTranRec(1))

  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  GET UBFile, RecNo&, UBCustRec(1)
  CLOSE UBFile

  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen

  PrevTranRec& = UBCustRec(1).LastTrans

  IF PrevTranRec& > 0 THEN
    DO WHILE PrevTranRec& > 0
      DCnt = DCnt + 1
      REDIM PRESERVE MTChoice(1 TO DCnt) AS FLen2
      GET UBTran, PrevTranRec&, UBTranRec(1)
      'IF PrevTranRec& = 11236 THEN
      '  FOR Rev = 1 TO 15
      '    UBTranRec(1).RevAmt(Rev) = ABS(UBTranRec(1).RevAmt(Rev))
      '  NEXT
      '  UBTranRec(1).TransAmt = ABS(UBTranRec(1).TransAmt)
      '  PUT UBTran, PrevTranRec&, UBTranRec(1)
      'END IF

      LSET MTChoice(DCnt).V = Num2Date(UBTranRec(1).TransDate)
      MID$(MTChoice(DCnt).V, 15) = LEFT$(UBTranRec(1).TransDesc, 15)
      MID$(MTChoice(DCnt).V, 40) = FUsing(STR$(UBTranRec(1).TransAmt), "#####.##")
      MID$(MTChoice(DCnt).V, 49) = FUsing(STR$(PrevTranRec&), "######")
      MID$(MTChoice(DCnt).V, 59) = MKL$(PrevTranRec&)
      PrevTranRec& = UBTranRec(1).PrevTrans
    LOOP

    CLOSE UBTran
    
    MaxLen = 57 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something

    IF Choice < 1 THEN
      Choice = 1                'Pre-load choice to highlight
    END IF

    Title$ = SPACE$(MaxLen + 4)
    LSET Title$ = " " + LEFT$(QPTrim$(UBCustRec(1).CustName), 20)
    MID$(Title$, 25) = LEFT$(QPTrim$(UBCustRec(1).SERVADDR), 25)
    MID$(Title$, 52, 9) = UBCustRec(1).BOOK + "-" + UBCustRec(1).SEQNUMB

    '--Find max menu width
    '--Center Menu within Screen

    Col = 11
    Row = 4
    BoxBot = 14 'limit the box length to go no lower than line 20

    DO
      LOCATE Row, Col, 0
      GOSUB DisplayStuff
      VertMenuT2 MTChoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        EXIT DO 'choice = 0
      ELSEIF Ky$ = CHR$(13) THEN
        GOSUB EditTransDate
        Ky$ = ""
        Action = 0
      END IF
    LOOP        'UNTIL EditLocRec& > 0
  ELSE
    OK = MsgBox%("UB.QSL", "NOCTRANS")
  END IF

EXIT SUB

DisplayStuff:
    BlockClear
    WazzWind 19, Col, 23, Col + 60, Cnf.MenBox, 1, 1
    QPrintRC Title$, Row - 1, Col, 112
    QPrintRC "PRESS <ENTER> to EDIT DATE", 20, Col + 2, 15
    QPrintRC "PRESS <Arrow Keys> to Navigate.", 21, Col + 2, 15
    QPrintRC "PRESS <ESCAPE> to Cancel ", 22, Col + 2, 15
RETURN

EditTransDate:

    LibName$ = "UBWACK"
    FormName$ = "EDTRDATE"
    NumFlds = LibNumberOfFields(LibName$, FormName$)

    REDIM Frm(1) AS FormInfo
    REDIM Form$(NumFlds, 2)       'DIM the form data array
    REDIM Fld(NumFlds) AS FieldInfo               'DIM the field information array
    StartEl = 0   'Load first form at array start
    LibGetFldDef LibName$, FormName$, StartEl, Fld(), Form$(), ErrCode

'    '----- Set the "Action" flag to force the editor to initialize itself and
'    '      display the data on the form.
    Action = 1

    '----- Setup TYPE for setting and reading form editing information.
    Frm(1).FldNo = 1              'Start editing on field #1
    Frm(1).InsStat = False        'Set insert state (True = Insert on)
    Frm(1).StartEl = 0            'Set form starting element to 0 and
    BlockClear
    GOSUB GetTransInfo
    LibFile2Scrn LibName$, FormName$, MonoCode%, Attribute%, ErrorCode%
    QPrintRC "  " + LEFT$(MTChoice(Choice).V, 58), 3, 11, 78

    FirstTime = True

    DO

      EditForm Form$(), Fld(), Frm(1), Cnf, Action

      IF FirstTime THEN
        FirstTime = False
        LSET Form$(1, 0) = Num2Date$(TranDate)
        Action = 1
      END IF

      SELECT CASE Frm(1).KeyCode
      CASE F0Key
        GOSUB ChangeDate
        EXIT DO
      CASE EscKey
        EXIT DO
      END SELECT
    LOOP

RETURN

ChangeDate:

  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen
  GET UBTran, TransRec&, UBTranRec(1)
  UBTranRec(1).TransDate = Date2Num%(Form$(1, 0))
  PUT UBTran, TransRec&, UBTranRec(1)
  CLOSE
  MID$(MTChoice(Choice).V, 1) = Form$(1, 0)
  BlockClear
  LibFile2Scrn LibName$, "PRORPT", MonoCode%, Attribute%, ErrorCode%
  QPrintRC "TRANSACTIONS UPDATED", 10, 30, -1
  QPrintRC "Press any key to continue.", 13, 28, -1
  WaitForAction

RETURN

GetTransInfo:
  TransRec& = CVL(MID$(MTChoice(Choice).V, 59, 4))
  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen
  GET UBTran, TransRec&, UBTranRec(1)
  CLOSE
  TranDate = UBTranRec(1).TransDate
RETURN

END SUB

