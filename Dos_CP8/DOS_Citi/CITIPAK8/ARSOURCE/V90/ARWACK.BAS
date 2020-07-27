DEFINT A-Z
DECLARE SUB ARLog (Text$)
DECLARE SUB RelinkTransactions ()
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB ShowCursor ()
DECLARE SUB OpenARCustFile (NumOfArRecs%, ARFile%)
DECLARE SUB OpenARCustIdxFile (NumOfARIdxRecs%, ARIdxFile%)
DECLARE SUB ShowNoCodes ()
DECLARE SUB SortARNameIndex ()
DECLARE SUB ARWLookUp (RecNo&)
DECLARE SUB ShowTransFixBal (RecNo&)
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
DECLARE SUB ShowTransSoWeCanWackIt (RecNo&)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB WaitForAction ()
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE FUNCTION Round# (DblNum#)

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'PageInfo.bi'
  '$INCLUDE: 'ARV90.BI'

  DIM SHARED ARCust(1) AS ARCustRecType
  DIM SHARED ARCustRec(1) AS ARCustRecType
  DIM SHARED ARCustIdxRec(1) AS ARCustIDXRecType

  TYPE Struct
    who AS STRING * 14
    RecNum AS INTEGER
  END TYPE

  TYPE FLen2
    V AS STRING * 64
  END TYPE
  
  REDIM Items(1) AS FLen2
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
  REDIM MChoice$(1 TO 4)

  MChoice$(1) = " Remove Transactions "
  MChoice$(2) = " Exit to OS"

  MaxLen = 0    'Set menu width to zero
  BoxBot = 18   'limit the box length to go no lower than line 18
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

  DO

    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear

    TitleBox 3, Col, MaxLen + 3, "Transaction Utility ", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

    ShowCursor

    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf

    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0

    SELECT CASE Choice
    CASE 1
      DO
        BlockClear
        ARWLookUp RecNo&
        IF RecNo& > 0 THEN
          ShowTransSoWeCanWackIt RecNo&
        END IF
      LOOP WHILE RecNo& > 0

    CASE 2
      HideCursor
      CLS
      END
    END SELECT
  LOOP

GetOut:
  RUN "ARCUST"

SUB ARLog (Text$) STATIC

  IF NOT BeenDone THEN
    BeenDone = True
    Today$ = DATE$
    Today$ = LEFT$(Today$, 2) + MID$(Today$, 4, 2) + RIGHT$(Today$, 2)
  END IF

  TheTime$ = TIME$
  IF LEFT$(TheTime$, 1) = "0" THEN
    Hour = VAL(MID$(TheTime$, 2, 1))
  ELSE
    Hour = VAL(MID$(TheTime$, 1, 2))
  END IF

  SELECT CASE Hour
  CASE IS > 11
    Hour = Hour - 12
    IF Hour = 0 THEN Hour = 12
    AmPm$ = "pm"
  CASE 1 TO 12
    AmPm$ = "am"
  CASE 0
    Hour = 12
    AmPm$ = "am"
  END SELECT
  SELECT CASE Hour
    CASE 1 TO 9
      Hour$ = "0" + QPTrim$(STR$(Hour))
    CASE ELSE
      Hour$ = QPTrim$(STR$(Hour))
  END SELECT
  TheTime$ = Hour$ + ":" + MID$(TheTime$, 4) + AmPm$
  LogFile = FREEFILE
  OPEN "ARLOG.DAT" FOR APPEND SHARED AS LogFile LEN = 255
  PRINT #LogFile, "AR: "; Today$; " @"; TheTime$
  PRINT #LogFile, "    "; Text$
  CLOSE #LogFile
  Text$ = "": TheTime$ = ""
END SUB

SUB ARWLookUp (RecNo&)

  CustomerGrabed = 0

  MaxLen = 50   'Set menu width to zero
  BoxBot = 17   'limit the box length to go no lower than line 20
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight

  TText$ = SPACE$(MaxLen + 4)
  LSET TText$ = " Cust #    Customer Sort Name"

  '--Center Menu within Screen
  Row = 8
  Col = 15

  OpenARCustIdxFile NumOfARIdxRecs, ARIdxFile

  REDIM CList(1 TO NumOfARIdxRecs) AS FLen2

  ChoiceCounter = 0
  FOR Cnt = 1 TO NumOfARIdxRecs
    GET ARIdxFile, Cnt, ARCustIdxRec(1)
    IF LEFT$(ARCustIdxRec(1).IDXNAME, 7) <> "DELETED" THEN
      ChoiceCounter = ChoiceCounter + 1
      LSET CList(ChoiceCounter).V = STR$(ARCustIdxRec(1).IDXRECORD)
      MID$(CList(ChoiceCounter).V, 10) = ARCustIdxRec(1).IDXNAME
    END IF
  NEXT Cnt

  DO

    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    LibFile2Scrn "AR.QSL", "MENUBAK", MonoCode, -1, ErrorCode
    ShowCursor
    QPrintRC TText$, Row - 1, Col, 112
    VertMenuT2 CList(), Choice, MaxLen%, BoxBot, Ky$, Action, Cnf

    IF Ky$ = CHR$(27) THEN
      RecNo& = 0
      ExitFlag = True
    ELSE
      RecNo& = VAL(LEFT$(CList(Choice).V, 8))
      ExitFlag = True
    END IF

  LOOP UNTIL ExitFlag

END SUB

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

SUB OpenARCustFile (NumOfArRecs, ARFile)
  CLOSE ARFile
  
  ARCustRecLen = LEN(ARCustRec(1))
  ARFile = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARFile LEN = ARCustRecLen
  NumOfArRecs = LOF(ARFile) \ ARCustRecLen
  'FOR x = 1 TO NumOfArRecs
  'GET ARFile, x, ARCust(1)
  'PRINT ARCust(1).Custnumb; TAB(15); ARCust(1).FirstTrans
  'SLEEP 1
  'NEXT x
  'STOP
END SUB

SUB OpenARCustIdxFile (NumOfARIdxRecs, ARIdxFile)
  CLOSE ARIdxFile
  ARCustIdxRecLen = LEN(ARCustIdxRec(1))
  ARIdxFile = FREEFILE
  OPEN "ARCUST.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS ARIdxFile LEN = ARCustIdxRecLen
  NumOfARIdxRecs = LOF(ARIdxFile) \ ARCustIdxRecLen
END SUB

SUB RelinkTransactions

  CLOSE

  LibFile2Scrn "AR.QSL", "MENUBAK", MonoCode, -1, ErrorCode
  'ClearBack
  LibFile2Scrn "ARWACK", "PRORPT", MonoCode%, Attribute%, ErrorCode%
  ShowProcessingScrn "Relinking Customer Transactions."
  
  CustRecLen = LEN(ARCust(1))

  REDIM ARTran(1 TO 2) AS ARTransRecType
  TranRecLen = LEN(ARTran(1))

  CustFile = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM SHARED AS CustFile LEN = CustRecLen
  NumCRec& = LOF(CustFile) / CustRecLen

  TranFile = FREEFILE
  OPEN "ARTRANS.DAT" FOR RANDOM SHARED AS TranFile LEN = TranRecLen
  NumTRec& = LOF(TranFile) / TranRecLen

  FOR Cnt& = 1 TO NumCRec&
    GET CustFile, Cnt&, ARCust(1)
    ARCust(1).FirstTrans = 0
    ARCust(1).LastTrans = 0
    PUT CustFile, Cnt&, ARCust(1)
    ShowPctComp Cnt&, NumCRec&
  NEXT
  
  FOR Cnt& = 1 TO NumTRec&
    GET TranFile, Cnt&, ARTran(1)
    CustRec& = VAL(ARTran(1).CustomerNumber)
    IF (CustRec& > 0) AND (CustRec& <= NumCRec&) THEN
      GET CustFile, CustRec&, ARCust(1)
      IF ARCust(1).LastTrans = 0 THEN
        ARCust(1).FirstTrans = Cnt&
        ARCust(1).LastTrans = Cnt&
        PUT CustFile, CustRec&, ARCust(1)
        ARTran(1).NextTrans = 0
        PUT TranFile, Cnt&, ARTran(1)
      ELSE
        GET TranFile, ARCust(1).LastTrans, ARTran(2)  'get old last tr
        ARTran(2).NextTrans = Cnt&                    'point it to next tr
        PUT TranFile, ARCust(1).LastTrans, ARTran(2)  'put it back
        ARCust(1).LastTrans = Cnt&                    'set new cust last TR
        PUT CustFile, CustRec&, ARCust(1)             'put it back
        ARTran(1).NextTrans = 0
        PUT TranFile, Cnt&, ARTran(1)
      END IF
      ShowPctComp Cnt&, NumTRec&
    ELSE
      BadTran = BadTran + 1
    END IF
  NEXT

  CLOSE


END SUB

SUB ShowTransSoWeCanWackIt (RecNo&)
  
  REDIM ARTranRec(1) AS ARTransRecType
  'REDIM ARCustRec(1) AS ARCustRecType
  
  ARCustRecLen = LEN(ARCustRec(1))
  ARTranRecLen = LEN(ARTranRec(1))
  
  ARFile = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM SHARED AS ARFile LEN = ARCustRecLen
  GET ARFile, RecNo&, ARCustRec(1)
  CLOSE ARFile
  
  'CurBal# = ARCustRec(1).CurrBalance
  'PreBal# = ARCustRec(1).PrevBalance
  
Top:
  
  ARTran = FREEFILE
  OPEN "ARTRANS.DAT" FOR RANDOM SHARED AS ARTran LEN = ARTranRecLen
  
  PrevTranRec& = ARCustRec(1).FirstTrans
  
  IF PrevTranRec& > 0 THEN
    DO WHILE PrevTranRec& > 0
      DCnt = DCnt + 1
      REDIM PRESERVE MTChoice(1 TO DCnt) AS FLen2
      REDIM PRESERVE RecList(1 TO DCnt) AS LONG
      GET ARTran, PrevTranRec&, ARTranRec(1)
      LSET MTChoice(DCnt).V = Num2Date(ARTranRec(1).TransDate)
      MID$(MTChoice(DCnt).V, 15) = LEFT$(ARTranRec(1).TransDesc, 15)
      MID$(MTChoice(DCnt).V, 30) = STR$(PrevTranRec&)
      MID$(MTChoice(DCnt).V, 40) = FUsing(STR$(ARTranRec(1).TransAmount), "#####.##")
      'MID$(MTChoice(DCnt).V, 50) = FUsing(STR$(ARTranRec(1).RunBalance), "#####.##")
      RecList(DCnt) = PrevTranRec&
      PrevTranRec& = ARTranRec(1).NextTrans
    LOOP
    
    CLOSE ARTran
    
    
HistTop:
    
    MaxLen = 59 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something
    
    IF Choice < 1 THEN
      Choice = 1                'Pre-load choice to highlight
    END IF
    
    Title$ = SPACE$(MaxLen + 9)
    LSET Title$ = "  Customer/Owner       Service Address        Location No.  S"
    LSET Title$ = " " + LEFT$(QPTrim$(ARCustRec(1).CustName), 20)
    'MID$(Title$, 25) = LEFT$(QPTrim$(ARCustRec(1).SERVADDR), 25)
    'MID$(Title$, 56, 9) = ARCustRec(1).BOOK + "-" + ARCustRec(1).SEQNUMB
    'MID$(Title$, 59, 1) = ARCustRec(1).STATUS
    
    '--Find max menu width
    '--Center Menu within Screen
    
    Col = ((80 - 60) \ 2) - 1
    
    Row = 6
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
    OK = MsgBox%("ARWACK.QSL", "NOCTRANS")
  END IF

  IF DCnt > 0 THEN
    Count$ = QPTrim$(STR$(DCnt))
    IF DCnt < 10 THEN
      Count$ = " " + Count$
    END IF
'************************************
    LibName$ = "ARWack"
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
    QPrintRC LEFT$(ARCustRec(1).CustName, 34), 10, 34, 78
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

  ARFile = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM SHARED AS ARFile LEN = ARCustRecLen

  ARTran = FREEFILE
  OPEN "ARTRANS.DAT" FOR RANDOM SHARED AS ARTran LEN = ARTranRecLen

  GET ARFile, RecNo&, ARCustRec(1)


  FOR ZZ = 1 TO DCnt
    TransRec& = RecList(Picked(ZZ))
    GET ARTran, TransRec&, ARTranRec(1)
    SELECT CASE ARTranRec(1).TransType
    CASE 1   '"Charge"
      ARCustRec(1).AcctBal = Round#(ARCustRec(1).AcctBal - ARTranRec(1).TransAmount)
      TrType$ = " Charge"
    CASE 2   '"Payment"
      ARCustRec(1).AcctBal = Round#(ARCustRec(1).AcctBal + ARTranRec(1).TransAmount)
      TrType$ = "Payment"
    CASE 9   '"Beg Bal"
      ARCustRec(1).AcctBal = Round#(ARCustRec(1).AcctBal - ARTranRec(1).TransAmount)
      TrType$ = "Beg Bal"
    END SELECT
    Log$ = "CUSTOMER:" + STR$(RecNo&) + " TR:" + STR$(TransRec&) + " TYPE:" + TrType$ + " AMT:" + FUsing$(STR$(ARTranRec(1).TransAmount), "#####.##")
    ARLog Log$
    LSET ARTranRec(1).CustomerNumber = "0"
    PUT ARTran, TransRec&, ARTranRec(1)
    RecList(Picked(ZZ)) = -1
  NEXT
  PUT ARFile, RecNo&, ARCustRec(1)
  CLOSE
  RelinkTransactions
  
  CursorOff
  BlockClear
  LibFile2Scrn LibName$, "PRORPT", MonoCode%, Attribute%, ErrorCode%
  QPrintRC "TRANSACTIONS REMOVED", 10, 30, -1
  QPrintRC "Press any key to continue.", 13, 28, -1
  WaitForAction

END SUB

SUB SortARNameIndex
  SHARED MChoice$
  
  
  size = 2500
  Start = 1     'start at array element 1
  Dir = 0       'sort direction - use anything else for descending
  SSize = 16    'total size of each TYPE element
  MOff = 0      'offset into the TYPE for the key element
  MSize = 7     'size of the key element - coded as follows:
  '   -1 = integer
  '   -2 = long integer
  '   -3 = single precision
  '   -4 = double precision
  '   +N = TYPE array/fixed-length string of length N
  
  DIM array(1 TO size)  AS Struct
  help$ = "Sorting Customer Index"
  'PrintHelp help$
  
  ARCustRecLen = LEN(ARCustRec(1))
  ARFile = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARFile LEN = ARCustRecLen
  NumOfArRecs = LOF(ARFile) \ ARCustRecLen
  
  ARCustIdxRecLen = LEN(ARCustIdxRec(1))
  ARIdxFile = FREEFILE
  OPEN "ARCUST.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS ARIdxFile LEN = ARCustIdxRecLen
  
  FOR Cnt = 1 TO NumOfArRecs
    GET ARFile, Cnt, ARCustRec(1)
    array(Cnt).who = ARCustRec(1).SORTNAME + "    "
    array(Cnt).RecNum = Cnt
  NEXT Cnt
  
  SortT array(Start), NumOfArRecs, Dir, SSize, MOff, MSize
  
  FOR Cnt = 1 TO NumOfArRecs
    ARCustIdxRec(1).IDXNAME = array(Cnt).who
    ARCustIdxRec(1).IDXRECORD = array(Cnt).RecNum
    PUT ARIdxFile, Cnt, ARCustIdxRec(1)
  NEXT Cnt
  CLOSE ARFile
  CLOSE ARIdxFile
END SUB

