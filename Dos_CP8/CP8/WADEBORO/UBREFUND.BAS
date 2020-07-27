DEFINT A-Z
DECLARE SUB RefundDeposit ()
DECLARE SUB UBLog (Text$)
DECLARE SUB CustMessageSystem (RecNo&)
DECLARE SUB EditApplyDeposit (RecNo&)
DECLARE FUNCTION Date2Num% (WhatDate$)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION PromptSaveData% ()
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Round# (DoubleNum#)
DECLARE SUB BlockClear ()
DECLARE SUB ClearScrn ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB HideCursor ()
DECLARE SUB LookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, LocationFlag%)
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB ApplyRefund ()
DECLARE SUB RestScrn (ScrnArray())
DECLARE SUB SaveScrn (ScrnArray())
DECLARE SUB ShowCursor ()
DECLARE SUB ShowCustHistory (CustAcct&)
DECLARE SUB StuffBuf (Ky$)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB WaitForAction ()

  '$INCLUDE: 'defCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'PageInfo.bi'
  '$INCLUDE: 'newcust.bi'

  CONST False = 0, True = NOT False

  CrLf$ = CHR$(13) + CHR$(10)

SUB ApplyRefund

  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  DO
    BlockClear
    LookUp RecNo&, "Customer to Apply Refund", 4, True, False
    IF RecNo& > 0 THEN
      GOSUB CheckCust
      IF CustOK THEN
        EXIT DO
      END IF
    ELSE
      EXIT DO
    END IF
  LOOP

  IF RecNo& > 0 THEN
    EditApplyDeposit RecNo&
  END IF

  IF NOT OKFlag THEN
    GOTO ExitDeleteCust
  END IF

ExitDeleteCust:

  EXIT SUB

CheckCust:
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  GET UBCust, RecNo&, UBCustRec(1)
  CLOSE UBCust

  IF UBCustRec(1).DepositAmt <= 0 THEN
    OK = MsgBox%("UB", "NODPOSIT")
    CustOK = False
  ELSE
    CustOK = True
  END IF

  RETURN

END SUB

SUB EditApplyDeposit (RecNo&)
  
  CustAcct& = RecNo&

  REDIM TempScrn(0)
  REDIM UBCustRec(1)     AS NewUBCustRecType
  
  UBCustRecLen = LEN(UBCustRec(1))
  
  NumOfRevs = MaxRevsCnt
  
  REDIM RevText$(1 TO MaxRevsCnt)
  REDIM UBSetUpRec(1) AS UBSetupRecType
  
  REDIM DistArray(1 TO MaxRevsCnt) AS DistArrayType
  
  UBSetupLen = LEN(UBSetUpRec(1))
  FGetAH "UBSETUP.DAT", UBSetUpRec(1), UBSetupLen, 1            'load it
  
  FOR RevCnt = 1 TO MaxRevsCnt
    RevText$(RevCnt) = LEFT$(QPTrim$(UBSetUpRec(1).Revenues(RevCnt).REVNAME), 14)
    IF LEN(RevText$(RevCnt)) = 0 THEN
      NumOfRevs = RevCnt - 1
      EXIT FOR
    END IF
  NEXT

  TownName$ = UBSetUpRec(1).UTILNAME
  
  LibName$ = "UB"
  ScrnName$ = "UBAPPDEP"
  
  '--Get the total number of fields from all pages
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  Frm(1).StayOnField = True
  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  REDIM CurFlds(1 TO MaxRevsCnt)
  REDIM ActFlds(1 TO MaxRevsCnt)

  FOR Cnt = 1 TO MaxRevsCnt
    CurFlds(Cnt) = FldNum%("CUR" + LTRIM$(STR$(Cnt)), Fld())
    ActFlds(Cnt) = FldNum%("ACT" + LTRIM$(STR$(Cnt)), Fld())
  NEXT

  CurTotalFld = FldNum%("CURTOTAL", Fld())
  AppTotalFld = FldNum%("APPTOTAL", Fld())
  ActTotalFld = FldNum%("ACTTOTAL", Fld())

  IF NumOfRevs < MaxRevsCnt THEN
    FOR Cnt = NumOfRevs + 1 TO MaxRevsCnt
      Fld(CurFlds(Cnt) + 1).Protected = True
    NEXT
  END IF

  '--define the multi-choice fields
  
  Action = 1
  FirstTime = True
  
  DisplayUBScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    

    IF Frm(1).Edited AND Frm(1).PrevFld <> Frm(1).FldNo THEN
      BeenEditedFlag = True     'if the form has been edited
    END IF      'set the edited flag
    
    IF FirstTime THEN
      FirstTime = False         'if this is the first time
      GOSUB GetCustInfo
      GOSUB ShowRevSources      '
      FOR zz = 1 TO NumOfRevs
        CalcFields 0, CurFlds(zz), Form$(), Fld()
        CalcFields 0, ActFlds(zz), Form$(), Fld()
      NEXT
      CalcFields 0, AppTotalFld, Form$(), Fld()
      LSET Form$(1, 0) = DATE$
      Frm(1).FldNo = 2
      Action = 2
    END IF
    
    IF DistFlag THEN
      TempDepAmt# = UBCustRec(1).DepositAmt
      GOSUB AutoDistribute
      DistFlag = False
    END IF
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE EscKey
      IF BeenEditedFlag THEN
        SaveFlag = PromptSaveData
        SELECT CASE SaveFlag
        CASE True               'user wants to save
          StuffBuf CHR$(0) + CHR$(ABS(F10Key))
        CASE False              'user wants to abandon
          ExitFlag = True
        CASE ELSE               'continue editing
        END SELECT
        Action = 1
      ELSE
        ExitFlag = True
      END IF
      
    CASE F4KEY  'Customer History
      IF CustAcct& > 0 THEN
        ShowCustHistory CustAcct&
        Action = 2
      END IF

    CASE F7KEY
      SaveScrn TempScrn()
      CustMessageSystem CustAcct&
      Action = 1
      RestScrn TempScrn()

    CASE F9KEY
      TempDepAmt# = UBCustRec(1).DepositAmt
      IF TempDepAmt# > 0 THEN
        GOSUB AutoDistribute
      END IF
      
    CASE F10Key 'Save
      GOSUB CheckDepositInfo
      IF DepositDistOK THEN
        GOSUB SaveTransaction 'do the save routine
        ExitFlag = True
      END IF
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 22   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 5 TO 13            'f4 History
          PressButton F4KEY, 22, 5, 13
        CASE 31 TO 42           'f9 Save
          PressButton F7KEY, 22, 31, 42
        CASE 44 TO 52           'f9 Save
          PressButton F9KEY, 22, 44, 52
        CASE 54 TO 63           'f10 Save
          PressButton F10Key, 22, 54, 63
        CASE 65 TO 75           '--cancel button
          PressButton EscKey, 22, 65, 75
        END SELECT
      END SELECT                'row
    END IF
  LOOP UNTIL ExitFlag
  
  ERASE TempScrn, UBCustRec, RevText$, UBSetUpRec
  
  'HideCursor
  EXIT SUB
  
GetCustInfo:
  CustFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CustFile LEN = UBCustRecLen
  GET CustFile, CustAcct&, UBCustRec(1)
  CLOSE CustFile
  FOR Cnt = 1 TO NumOfRevs
    LSET Form$(CurFlds(Cnt), 0) = STR$(UBCustRec(1).CurrRevAmts(Cnt))
    SaveField CurFlds(Cnt), Form$(), Fld(), BadField
  NEXT
  TotalBalance# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
RETURN
  
ShowRevSources:
  FOR RevCnt = 1 TO NumOfRevs
    QPrintRC RevText$(RevCnt), RevCnt + 4, 6, -1
  NEXT
  QPrintRC LEFT$(UBCustRec(1).CUSTNAME, 25), 3, 33, -1
  QPrintRC STR$(CustAcct&), 3, 17, -1
  QPrintRC FUsing$(STR$(UBCustRec(1).DepositAmt), "$$####.##"), 20, 14, -1
RETURN
  
AutoDistribute:
  FOR Cnt = 1 TO NumOfRevs
    WhatRev = Cnt
    ThisAmt# = Value(Form$(CurFlds(WhatRev), 0), ECode)
    IF ThisAmt# <> 0 THEN
      IF TempDepAmt# >= ThisAmt# THEN
        LSET Form$(CurFlds(WhatRev) + 1, 0) = QPTrim$(STR$(ThisAmt#))
        TempDepAmt# = Round#(TempDepAmt# - ThisAmt#)
      ELSE
        ThisAmt# = TempDepAmt#
        LSET Form$(CurFlds(WhatRev) + 1, 0) = QPTrim$(STR$(ThisAmt#))
        TempDepAmt# = 0
      END IF
    ELSEIF TempDepAmt# = 0 THEN
      LSET Form$(CurFlds(WhatRev) + 1, 0) = STR$(0)
    END IF
    SaveField CurFlds(WhatRev) + 1, Form$(), Fld(), BadField
  NEXT
  
  IF TempDepAmt# > 0 THEN
    ThisAmt# = (Value(Form$(CurFlds(1), 0), ECode) + TempDepAmt#)
    LSET Form$(CurFlds(1) + 1, 0) = QPTrim$(STR$(ThisAmt#))
    SaveField CurFlds(1) + 1, Form$(), Fld(), BadField
  END IF

  FOR zz = 1 TO NumOfRevs
    LSET Form$(CurFlds(1) + 2, 0) = QPTrim$(STR$(Round#(Value(Form$(CurFlds(1), 0), ECode) - Value(Form$(CurFlds(1) + 1, 0), ECode))))
    SaveField CurFlds(1) + 2, Form$(), Fld(), BadField
  NEXT
  BeenEditedFlag = True

  CalcFields 0, AppTotalFld, Form$(), Fld()
  CalcFields 0, CurFlds(zz) + 1, Form$(), Fld()
  CalcFields 0, CurFlds(zz) + 2, Form$(), Fld()

  FirstTime = True

  Action = 2

RETURN
  
  
SaveTransaction:
  
  SaveScrn TempScrn()
  DisplayUBScrn "UPDATDSK"
  
  REDIM UBTransRec(1) AS UBTransRecType
  UBTransRecLen = LEN(UBTransRec(1))

  TransDate = Date2Num(Form$(1, 0))
  TransAmt# = Value#(Form$(AppTotalFld, 0), ECode)

  UBTransRec(1).TransDate = TransDate
  UBTransRec(1).CustLocation = CustAcct&
  UBTransRec(1).CustStatus = UBCustRec(1).Status
  UBTransRec(1).CustAcctNo = CustAcct&
  UBTransRec(1).TransAmt = TransAmt#

  FOR Cnt = 1 TO 15
    IF LEN(QPTrim$(Form$(CurFlds(Cnt) + 1, 0))) > 0 THEN
      UBTransRec(1).RevAmt(Cnt) = Value#(Form$(CurFlds(Cnt) + 1, 0), ECode)
    ELSE
      UBTransRec(1).RevAmt(Cnt) = 0
    END IF
  NEXT
  
  UBTransRec(1).TransDesc = "Applied Deposit"
  FOR RevCnt = 1 TO 15
    UBCustRec(1).CurrRevAmts(RevCnt) = Round#(Value#(Form$(ActFlds(RevCnt), 0), ECode))
  NEXT

  UBCustRec(1).CurrBalance = Round#(Value#(Form$(ActTotalFld, 0), ECode))
  SELECT CASE UBCustRec(1).PrevBalance
  CASE 0
    'don't do anything
  CASE IS > 0
    IF UBCustRec(1).PrevBalance < UBCustRec(1).DepositAmt THEN
      UBCustRec(1).PrevBalance = 0
    ELSE
      UBCustRec(1).PrevBalance = Round#(UBCustRec(1).PrevBalance - UBCustRec(1).DepositAmt)
      UBCustRec(1).CurrBalance = Round#(UBCustRec(1).CurrBalance - UBCustRec(1).PrevBalance)
    END IF
  CASE IS < 0
    UBCustRec(1).PrevBalance = 0
  END SELECT

  UBTransRec(1).TransType = TranAppliedDeposit
  UBTransRec(1).RunBalance = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
  UBCustRec(1).DepositAmt = 0

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTransRecLen

  NextTranRecs& = (LOF(UBTran) \ UBTransRecLen) + 1
  PrevLastTrans& = UBCustRec(1).LastTrans
  UBTransRec(1).PrevTrans = PrevLastTrans&
  UBCustRec(1).LastTrans = NextTranRecs&

  IF Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance) = 0 THEN
    IF UBCustRec(1).Status = "B" THEN
      CustChCnt = CustChCnt + 1
      UBLog "APPLY DEP: SET CUST STATUS to I. Acct:" + STR$(UBTransRec(1).CustAcctNo)
      UBCustRec(1).Status = "I"
    END IF
  END IF

  PUT UBCust, UBTransRec(1).CustAcctNo, UBCustRec(1)
  PUT UBTran, NextTranRecs&, UBTransRec(1)
  CLOSE UBTran, UBCust

  RestScrn TempScrn()
  DisplayUBScrn "UPDATEOK"
  WaitForAction
  RestScrn TempScrn()
  FirstTime = True
RETURN
  
CheckDepositInfo:

  WhatFld = Frm(1).PrevFld
  DepositDistOK = True           'assume all is well
  TAmtApp# = VAL(Form$(AppTotalFld, 0))
  IF TAmtApp# <> UBCustRec(1).DepositAmt THEN
    Frm(1).FldNo = WhatFld
    DepositDistOK = False
    Action = 2
    OK = MsgBox%("UB.QSL", "BADAPPLY")
    GOTO BadApply
  END IF

  FOR RCnt = 1 TO NumOfRevs
    TAmt# = Round#(VAL(Form$(CurFlds(RCnt), 0)) - VAL(Form$(CurFlds(RCnt) + 1, 0)))
    AAmt# = Round#(VAL(Form$(ActFlds(RCnt), 0)))
    IF TAmt# <> AAmt# THEN
      Frm(1).FldNo = WhatFld
      DepositDistOK = False
      Action = 2
      OK = MsgBox%("UB.QSL", "BADAPPLY")
      GOTO BadApply
    END IF
  NEXT

BadApply:
  RETURN
  
END SUB

SUB RefundDeposit

  LibName$ = "UB"
  ScrnName$ = "OK2REFND"
  '--define the multi-choice fields
  '--Initialize the form name array
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1

  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  '--for each screen, get first and last fields
  StartEl = 0

  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode

  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  DO
    BlockClear
    LookUp RecNo&, "Customer to Refund Deposit", 4, True, False
    IF RecNo& > 0 THEN
      GOSUB ChkRefCust
      IF CustOK THEN
        EXIT DO
      END IF
    ELSE
      EXIT DO
    END IF
  LOOP

  IF RecNo& > 0 THEN
    Action = 1
    '--Set screen number to one and display screen
    BlockClear
    DisplayUBScrn ScrnName$
    GOSUB DisplayInfo
    'FirstTime = True

    DO

      EditForm Form$(), Fld(), Frm(1), Cnf, Action
      '--Check for Key presses
      SELECT CASE Frm(1).KeyCode
      CASE F10Key
        OKFlag = True
        ExitFlag = True
      CASE EscKey
        OKFlag = False
        ExitFlag = True
      END SELECT

    LOOP UNTIL ExitFlag
  END IF

  IF NOT OKFlag THEN
    GOTO ExitRefDep
  ELSE
    GOSUB DoDepositRefund
  END IF


ExitRefDep:

  EXIT SUB

ChkRefCust:
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  GET UBCust, RecNo&, UBCustRec(1)
  CLOSE UBCust

  IF UBCustRec(1).DepositAmt <= 0 THEN
    OK = MsgBox%("UB", "NODPOSIT")
    CustOK = False
  ELSE
    CustOK = True
  END IF

RETURN

DoDepositRefund:
  BlockClear
  REDIM TempScrn(0)
  SaveScrn TempScrn()
  DisplayUBScrn "UPDATDSK"

  REDIM UBTransRec(1) AS UBTransRecType
  UBTransRecLen = LEN(UBTransRec(1))

  TransDate = Date2Num(DATE$)
  TransAmt# = -UBCustRec(1).DepositAmt

  UBTransRec(1).TransDate = TransDate
  UBTransRec(1).CustLocation = RecNo&
  UBTransRec(1).CustStatus = UBCustRec(1).Status
  UBTransRec(1).CustAcctNo = RecNo&
  UBTransRec(1).TransAmt = TransAmt#
  
  UBTransRec(1).TransDesc = "Refunded Deposit"
  
  UBTransRec(1).TransType = TranRefundDeposit
  UBTransRec(1).RunBalance = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
  UBCustRec(1).DepositAmt = 0

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTransRecLen

  NextTranRecs& = (LOF(UBTran) \ UBTransRecLen) + 1
  PrevLastTrans& = UBCustRec(1).LastTrans
  UBTransRec(1).PrevTrans = PrevLastTrans&
  UBCustRec(1).LastTrans = NextTranRecs&

  PUT UBCust, RecNo&, UBCustRec(1)
  PUT UBTran, NextTranRecs&, UBTransRec(1)
  CLOSE UBTran, UBCust

  RestScrn TempScrn()
  DisplayUBScrn "UPDATEOK"
  WaitForAction
  RestScrn TempScrn()
  FirstTime = True

RETURN


DisplayInfo:
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  GET UBCust, RecNo&, UBCustRec(1)
  CLOSE UBCust
  DepAmt$ = QPTrim$(FUsing$(STR$(UBCustRec(1).DepositAmt), "$$,####.##"))
  QPrintRC LEFT$(QPTrim$(UBCustRec(1).CUSTNAME), 30), 10, 30, -1
  QPrintRC DepAmt$, 11, 30, -1
RETURN

END SUB

SUB RefundMenu
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 3)

  MChoice$(1) = "Apply Deposit to Balance"
  MChoice$(2) = "Refund Customer Deposit."
  MChoice$(3) = "Exit to OS "

  MaxLen = 0    'Set menu width to zero
  BoxBot = 17   'limit the box length to go no lower than line 20
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

    TitleBox 3, Col, MaxLen + 3, "Customer Deposit Refunds ", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

    ShowCursor
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf

    IF Ky$ = CHR$(27) THEN
      Choice = 0
      ExitFlag = True
      EXIT DO
    END IF

    SELECT CASE Choice
    CASE 1
      ApplyRefund
      
    CASE 2
      RefundDeposit
    CASE 3
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP UNTIL ExitFlag

END SUB

