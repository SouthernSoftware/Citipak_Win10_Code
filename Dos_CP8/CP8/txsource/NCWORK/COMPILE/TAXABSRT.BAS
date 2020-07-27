DEFINT A-Z
DECLARE SUB MovePropertyAbstract ()
DECLARE SUB SelectPers2Del (CustRec&)
DECLARE SUB DelPersAbstract (PersRecs() AS LONG, WhatPers%, CustRec&)
DECLARE SUB GetPersRecList (PersRecs() AS LONG, CustRec&)
DECLARE SUB DelPropAbstract (PropRecs() AS LONG, WhatProp%, CustRec&)
DECLARE SUB DisplayTaxScrn (ScrnName$)
DECLARE SUB GetPropRecList (PropRecs() AS LONG, CustRec&)
DECLARE SUB SelectProp2Del (RecNo&)
DECLARE SUB LookUp (RecNo&, Text$, ChkBalFlag%, CLSFlag%, SSNFlag%)
DECLARE SUB ClearBack ()
DECLARE SUB SendDist2GL ()
DECLARE SUB UBMiscMenu ()
DECLARE SUB UBBillMenu ()
DECLARE SUB UBCustomerMenu ()
DECLARE SUB ClearScrn ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION WEnvTest% ()

'$INCLUDE: 'DefCnf.BI'
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)
DECLARE SUB HideCursor ()
DECLARE SUB CursorOff ()
DECLARE SUB TextCursor (MouseFg%, MouseBg%)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice, MaxLen%, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB WaitForAction ()
DECLARE SUB MPaintBox (UlRow%, UlCol%, LRRow%, LRCol%, Colr%)

  CONST False = 0, True = NOT False

  TYPE FLen2
    V AS STRING * 64
  END TYPE
  
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TAXCONST.BI'
  '$INCLUDE: 'PROPAbst.BI'

  STACK 5000

  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 4)

  MChoice$(1) = "Delete Property Abstract"
  MChoice$(2) = "Delete Personal Abstract"
  MChoice$(3) = "Exit to OS"

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
  Row = ((23 - (UBOUND(MChoice$))) \ 2) + 1
  Col = ((80 - MaxLen) \ 2) - 1

  DO

    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0

    ClearBack

    TitleBox 2, Col, MaxLen + 3, "Abstract Maintenance ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

    ShowCursor

    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf

    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0

    SELECT CASE Choice
    'CASE 1
    '  MovePropertyAbstract
      
    CASE 1
      ClearBack
      LookUp RecNo&, "Delect Property Abstract Customer", False, True, False
      IF RecNo& > 0 THEN
        SelectProp2Del RecNo&
      END IF

    CASE 2
      ClearBack
      LookUp RecNo&, "Delect Personal Abstract Customer", False, True, False
      IF RecNo& > 0 THEN
        SelectPers2Del RecNo&
      END IF

    CASE 3
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP

  IF INSTR(COMMAND$, "TEST") THEN
    HideCursor
    ClearScrn
    END
  END IF

  IF WEnvTest THEN
    Ext$ = ".bas"
  ELSE
    Ext$ = ".exe"
  END IF
  IF Exist("Taxmenu" + Ext$) THEN
    RUN "TaxMenu"
  ELSE
    HideCursor
    ClearScrn
  END IF

END

SUB MovePropertyAbstract

  ClearBack
  LookUp MRecNo&, "Property to MOVE Customer", False, True, False
  IF MRecNo& <= 0 THEN
    GOTO ExitMoveProperty
  END IF
  GOSUB SelectProp2Move

  IF NOT GotAPropFlag THEN
    GOTO ExitMoveProperty
  END IF
  
  GOTO ExitMoveProperty


SelectProp2Move:
  GotAPropFlag = False
  REDIM PropRecs(0 TO 0) AS LONG
  GetPropRecList PropRecs(), MRecNo&

  REDIM PropRec(1) AS PropertyRecType
  PropRecLen = LEN(PropRec(1))

  PropCnt = PropRecs(0)
  IF PropCnt < 1 THEN   'Customer has no property to move!
    CursorOff
    DisplayTaxScrn "ERRSCRN1"
    QPrintRC "NO Property Abstracts to Move!", 10, 27, -1
    QPrintRC "Press any key to continue", 12, 28, -1
    WaitForAction
    GOTO SelProp2MoveExit
  END IF

  REDIM MChoice(1 TO PropCnt) AS FLen2

  PropFile = FREEFILE
  OPEN TaxPropFile FOR RANDOM SHARED AS PropFile LEN = PropRecLen
  FOR Cnt = 1 TO PropCnt
    GET PropFile, PropRecs(Cnt), PropRec(1)
    LSET MChoice(Cnt).V = LEFT$(PropRec(1).RealPin, 10)
    MID$(MChoice(Cnt).V, 12) = PropRec(1).Map          'len 6
    MID$(MChoice(Cnt).V, 20) = PropRec(1).Block        'len 6
    MID$(MChoice(Cnt).V, 28) = PropRec(1).LotNumb      'len 6
    MID$(MChoice(Cnt).V, 36) = FUsing$(STR$(PropRec(1).PropValu), Image$)
    MID$(MChoice(Cnt).V, 61) = MKL$(PropRecs(Cnt))
  NEXT
  CLOSE PersFile

  MaxLen = 59 'Set menu width to zero
  Action = 1  '0 means stay in the menu until they do something

  Choice = 1                'Pre-load choice to highlight

  Title$ = SPACE$(MaxLen + 4)
  LSET Title$ = "   PIN       Map   Block     Lot        Value            Acct"

  Row = 4
  Col = 10
  BoxBot = 17               'limit the box length

  TitleBox BoxBot + 3, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
  QPrintRC "Matched:" + STR$(PropCnt), BoxBot + 4, Col + 2, 15
  QPrintRC Title$, Row - 1, Col, 112
  MPaintBox Row, Col + MaxLen + 4, Row, Col + MaxLen + 5, 8

MoveLoopTop:
  LOCATE Row, Col, 0
  DO
    VertMenuT2 MChoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    IF Ky$ = CHR$(27) THEN
      PMRecNo& = -1
      EXIT DO 'choice = 0
    END IF
    PMRecNo& = CVL(MID$(MChoice(Choice).V, 61, 4))
  LOOP UNTIL Action = 4

  IF PMRecNo& > 0 THEN
    OK = MsgBox%("TAX.QSL", "DELCPROP")
    SELECT CASE OK
    CASE 1  'nope don't
      GOTO MoveLoopTop
    CASE 2  'yep go ahead kill it
      'DelPropAbstract PropRecs(), CINT(PMRecNo&), MRecNo&
    END SELECT
  END IF

SelProp2MoveExit:
RETURN

ExitMoveProperty:
END SUB

SUB SelectPers2Del (CustRec&)

  Image$ = "#######.##"

  REDIM PersRecs(0 TO 0) AS LONG
  GetPersRecList PersRecs(), CustRec&

  REDIM PersRec(1) AS PersonalRecType
  PersRecLen = LEN(PersRec(1))

  PersCnt = PersRecs(0)
  IF PersCnt < 1 THEN   'Show no properties message
    CursorOff
    DisplayTaxScrn "ERRSCRN1"
    QPrintRC "NO Personal Abstracts to Delete!", 10, 26, -1
    QPrintRC "Press any key to continue", 12, 28, -1
    WaitForAction
    GOTO Pers2DelExit
  END IF

  REDIM MChoice(1 TO PersCnt) AS FLen2

  PersFile = FREEFILE
  OPEN TaxPersFile FOR RANDOM SHARED AS PersFile LEN = PersRecLen
  FOR Cnt = 1 TO PersCnt
    GET PersFile, PersRecs(Cnt), PersRec(1)
    LSET MChoice(Cnt).V = LEFT$(PersRec(1).PropPin, 10)
    MID$(MChoice(Cnt).V, 11) = FUsing$(STR$(PersRec(1).PersVal), Image$)
    MID$(MChoice(Cnt).V, 21) = FUsing$(STR$(PersRec(1).MHValue), Image$)
    MID$(MChoice(Cnt).V, 31) = FUsing$(STR$(PersRec(1).MCValue), Image$)
    MID$(MChoice(Cnt).V, 41) = FUsing$(STR$(PersRec(1).CValue), Image$)
    MID$(MChoice(Cnt).V, 51) = FUsing$(STR$(PersRec(1).MTValue), Image$)
    MID$(MChoice(Cnt).V, 61) = MKL$(PersRecs(Cnt))
  NEXT
  CLOSE PersFile

  MaxLen = 59 'Set menu width to zero
  Action = 1  '0 means stay in the menu until they do something

  Choice = 1                'Pre-load choice to highlight

  Title$ = SPACE$(MaxLen + 4)
  LSET Title$ = "   PIN        Personal    Mobile  Merchant      Farm   Machine"

  Row = 4
  Col = 10
  BoxBot = 17               'limit the box length

  TitleBox BoxBot + 3, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
  QPrintRC "Matched:" + STR$(PersCnt), BoxBot + 4, Col + 2, 15
  QPrintRC Title$, Row - 1, Col, 112
  MPaintBox Row, Col + MaxLen + 4, Row, Col + MaxLen + 5, 8

PersLoopTop:
  LOCATE Row, Col, 0

  DO
    VertMenuT2 MChoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    IF Ky$ = CHR$(27) THEN
      RecNo& = -1
      EXIT DO 'choice = 0
    END IF
    RecNo& = CVL(MID$(MChoice(Choice).V, 61, 4))
  LOOP UNTIL Action = 4

  IF RecNo& > 0 THEN
    OK = MsgBox%("TAX.QSL", "DELCPERS")
    SELECT CASE OK
    CASE 1  'nope don't
      GOTO PersLoopTop
    CASE 2  'yep go ahead kill it
      DelPersAbstract PersRecs(), Choice, CustRec&
    END SELECT
  END IF

Pers2DelExit:


END SUB

SUB SelectProp2Del (CustRec&)

  Image$ = ",#######.##"

  REDIM PropRecs(0 TO 0) AS LONG
  GetPropRecList PropRecs(), CustRec&

  REDIM PropRec(1) AS PropertyRecType
  PropRecLen = LEN(PropRec(1))

  PropCnt = PropRecs(0)
  IF PropCnt < 1 THEN   'Show no properties message
    CursorOff
    DisplayTaxScrn "ERRSCRN1"
    QPrintRC "NO Property Abstracts to Delete!", 10, 26, -1
    QPrintRC "Press any key to continue", 12, 28, -1
    WaitForAction
    GOTO Prop2DelExit
  END IF

  REDIM MChoice(1 TO PropCnt) AS FLen2
  
  PropFile = FREEFILE
  OPEN TaxPropFile FOR RANDOM SHARED AS PropFile LEN = PropRecLen
  FOR Cnt = 1 TO PropCnt
    GET PropFile, PropRecs(Cnt), PropRec(1)
    LSET MChoice(Cnt).V = LEFT$(PropRec(1).RealPin, 10)
    MID$(MChoice(Cnt).V, 12) = PropRec(1).Map          'len 6
    MID$(MChoice(Cnt).V, 20) = PropRec(1).Block        'len 6
    MID$(MChoice(Cnt).V, 28) = PropRec(1).LotNumb      'len 6
    MID$(MChoice(Cnt).V, 36) = FUsing$(STR$(PropRec(1).PropValu), Image$)
    MID$(MChoice(Cnt).V, 61) = MKL$(PropRecs(Cnt))
  NEXT
  CLOSE PersFile

  MaxLen = 59 'Set menu width to zero
  Action = 1  '0 means stay in the menu until they do something

  Choice = 1                'Pre-load choice to highlight

  Title$ = SPACE$(MaxLen + 4)
  LSET Title$ = "   PIN       Map   Block     Lot        Value            Acct"

  Row = 4
  Col = 10
  BoxBot = 17               'limit the box length
  
  TitleBox BoxBot + 3, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
  QPrintRC "Matched:" + STR$(PropCnt), BoxBot + 4, Col + 2, 15
  QPrintRC Title$, Row - 1, Col, 112
  MPaintBox Row, Col + MaxLen + 4, Row, Col + MaxLen + 5, 8

LoopTop:
  LOCATE Row, Col, 0
  DO
    VertMenuT2 MChoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    IF Ky$ = CHR$(27) THEN
      RecNo& = -1
      EXIT DO 'choice = 0
    END IF
    RecNo& = CVL(MID$(MChoice(Choice).V, 61, 4))
  LOOP UNTIL Action = 4

  IF RecNo& > 0 THEN
    OK = MsgBox%("TAX.QSL", "DELCPROP")
    SELECT CASE OK
    CASE 1  'nope don't
      GOTO LoopTop
    CASE 2  'yep go ahead kill it
      DelPropAbstract PropRecs(), CINT(RecNo&), CustRec&
    END SELECT
  END IF

Prop2DelExit:

END SUB

