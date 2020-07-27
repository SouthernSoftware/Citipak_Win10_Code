DEFINT A-Z
DECLARE FUNCTION AdjPassWordOK% ()
DECLARE SUB UtilAdjustments ()
DECLARE SUB UBLog (Text$)
DECLARE SUB CustMessageSystem (RecNo&)
DECLARE SUB EditApplyDeposit (RecNo&)
DECLARE FUNCTION Monitor ()
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
DECLARE SUB QPrintRC (Text$, Row, Col, kolor)
DECLARE SUB ApplyRefund ()
DECLARE SUB RestScrn (ScrnArray())
DECLARE SUB SaveScrn (ScrnArray())
DECLARE SUB ShowCursor ()
DECLARE SUB ShowCustHistory (CustAcct&)
DECLARE SUB StuffBuf (Ky$)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB WaitForAction ()
DECLARE FUNCTION QPValL& (LongNum$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION IsDeleted% (CustAcctNo&)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetupLen%)
DECLARE SUB MPaintBox (trow, LCol, brow, RCol, kolor)
DECLARE FUNCTION WEnvTest ()
DECLARE FUNCTION nGetPassword$ (Row%, Col%, Max%, Abort%)
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)

  '$INCLUDE: 'defCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'PageInfo.bi'
  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'SetCnf.BI'

  CONST False = 0, True = NOT False

  CrLf$ = CHR$(13) + CHR$(10)

  BlockClear
  ShowCursor

  IF WEnvTest THEN
    UtilAdjustments
  ELSEIF AdjPassWordOK THEN
    UtilAdjustments
  END IF

  IF INSTR(COMMAND$, "TEST") THEN
    HideCursor
    ClearScrn
    END
  END IF

  RUN "UBBILLIN"

FUNCTION AdjPassWordOK

  UBLog " IN: Adjustments Password"

  Pz$ = ""
  z$ = "*+63>;5"    'Password is: UTILADJ
  FOR Cnt = 1 TO LEN(z$)
    Pz$ = Pz$ + CHR$(ASC(MID$(z$, Cnt, 1)) XOR 127)
  NEXT
  
  DO

    BlockClear
    WazzWind 8, 15, 16, 65, 14, 1, 1

    QPrintRC "[ UTILITY ADJUSTMENTS ]", 8, 18, 14
    QPrintRC "Enter password:", 12, 22, 15

    IF Try > 0 THEN
      QPrintRC "       Attempt:" + STR$(Try), 14, 22, 15
    END IF

    Pw$ = nGetPassword(12, 38, 8, Abort)

    IF Abort THEN
      UBLog "PASSWORD ABORTED: " + Pw$
      EXIT DO
    END IF

    IF Pw$ = Pz$ THEN
      UBLog "PASSWORD OK!"
      OkFlag = True
    ELSE
      UBLog "INVALID PASSWORD: " + Pw$
      QPrintRC "Invalid Password. Press any key to try again.", 14, 18, 31
      WaitForAction
    END IF

    Try = Try + 1

    IF Try = 4 THEN
      BlockClear
      WazzWind 8, 15, 16, 65, 15, 1, 1
      QPrintRC "Get the password from the system operator.", 11, 20, 14
      QPrintRC "Press any key to return to the Main Menu.", 13, 20, 14
      WaitForAction
      EXIT DO
    END IF

  LOOP UNTIL OkFlag

GetOut:
  IF OkFlag THEN
    AdjPassWordOK = True
  ELSE
    AdjPassWordOK = False
  END IF

END FUNCTION

SUB UtilAdjustments
  
  SHARED Choice$()
  AdjDate = Date2Num(DATE$)
  UBLog " IN: UTIL ADJUST"

  '**********************************************************
  REDIM TempScrn(0)
  REDIM UBCustRec(1)   AS NewUBCustRecType
  
  UBCustRecLen = LEN(UBCustRec(1))
  
  NumOfRevs = MaxRevsCnt
  
  REDIM RevText$(1 TO MaxRevsCnt)
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  
  TownName$ = UBSetUpRec(1).UTILNAME
  
  FOR RevCnt = 1 TO MaxRevsCnt
    RevText$(RevCnt) = LEFT$(QPTrim$(UBSetUpRec(1).Revenues(RevCnt).REVNAME), 14)
    IF LEN(RevText$(RevCnt)) = 0 THEN
      NumOfRevs = RevCnt - 1
      EXIT FOR
    END IF
  NEXT
  
  IF NumOfRevs < MaxRevsCnt THEN
    REDIM PRESERVE RevText$(1 TO NumOfRevs)
  END IF
  
  LibName$ = "UB"
  ScrnName$ = "UBADJUST"
  
  '--Get the total number of fields from all pages
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  Frm(1).StayOnField = True
  Frm(1).PageDisable = True
  
  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  CustAcctFld = FldNum%("CUSTACCT", Fld())
  AdjTypeFld = FldNum%("ADJTYPE", Fld())
  AdjDateFld = FldNum%("ADJDATE", Fld())
  TotAmtFld = FldNum%("TADJAMT", Fld())
  NoteFld = FldNum%("NOTE", Fld())
  
  REDIM ReadFlds(1 TO 7)
  REDIM CashFlds(1 TO MaxRevsCnt)
  
  FOR Cnt = 1 TO 7
    ReadFlds(Cnt) = FldNum%("READADJ" + LTRIM$(STR$(Cnt)), Fld())
  NEXT
  
  FOR Cnt = 1 TO MaxRevsCnt
    CashFlds(Cnt) = FldNum%("AMTADJ" + LTRIM$(STR$(Cnt)), Fld())
  NEXT
  
  FOR Cnt = NumOfRevs + 1 TO MaxRevsCnt
    Fld(CashFlds(Cnt)).Protected = True
  NEXT
  
  '--define the multi-choice fields
  
  REDIM Choice$(0 TO 2, 0 TO 0)
  
  Choice$(0, 0) = QPTrim$(STR$(AdjTypeFld))
  Choice$(1, 0) = "Upward"
  Choice$(2, 0) = "Downward"
  
  'GOSUB ClearForm
  
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
      GOSUB ShowRevSources      '
      Action = 2
    END IF
    
    IF Frm(1).FldNo > CustAcctFld AND Frm(1).PrevFld = CustAcctFld THEN
      CustAcct& = QPValL(Form$(CustAcctFld, 0))
      GOSUB GetCustInfo
    ELSEIF Frm(1).FldNo = CustAcctFld AND Frm(1).PrevFld <> CustAcctFld THEN
      MPaintBox 22, 37, 22, 41, 112
      MPaintBox 22, 35, 22, 36, 126
    END IF

    IF Frm(1).FldNo <> Frm(1).PrevFld THEN
      GOSUB CalcCashFlds
    END IF

    'CashFlds (Cnt)
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE EscKey
      IF BeenEditedFlag THEN
        SaveFlag = PromptSaveData
        SELECT CASE SaveFlag
        CASE True               'user wants to save
          EscapeFlag = True
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
      
    CASE F7KEY  'Lookup Customer
      IF Frm(1).FldNo = 1 THEN  'if user is on the Customer field
        SaveScrn TempScrn()     'and F7key then do lookup routine
        MPaintBox 4, 5, 22, 75, 8
        LookUp CustAcct&, "Payment", 2, False, False
        RestScrn TempScrn()
        IF CustAcct& > 0 THEN   'if this is a valid customer
          GOSUB GetCustInfo     'go get customer info
          UBLog "ADJUST:" + STR$(CustAcct&)
          Frm(1).FldNo = 4
          Action = 1
        ELSE
          GOSUB ClearForm
          Frm(1).FldNo = 1
          Action = 1
        END IF
      END IF
      
    CASE F10Key 'Save
      GOSUB CalcCashFlds
      GOSUB CheckAdjustInfo
      IF AdjustOKFlag THEN
        
        OFldNo = Frm(1).FldNo
        OK = MsgBox%("UB.QSL", "OK2ADJ")        'Ask if they are sure
        SELECT CASE OK
        CASE 1  'Yes do the adjustment
          GOSUB SaveTransaction 'do the save routine
          GOSUB ClearForm
          Frm(1).FldNo = 1
          Action = 1
        CASE ELSE               'Nope continue editing
          Frm(1).FldNo = OFldNo - 1
          Action = 1
        END SELECT
        IF EscapeFlag THEN
          IF CustAcct& > 0 THEN
            UBLog "ABORT ADJUST:" + STR$(CustAcct&)
          END IF
          ExitFlag = True
        END IF
      ELSE
        Action = 2
      END IF
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 22   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 5 TO 13            'f4 History
          PressButton F4KEY, 22, 5, 13
        CASE 34 TO 42           'f7 Look-Up
          PressButton F7KEY, 22, 34, 42
        CASE 54 TO 63           'f10 Save
          PressButton F10Key, 22, 54, 63
        CASE 65 TO 75           '--cancel button
          PressButton EscKey, 22, 65, 75
        END SELECT
      END SELECT                'row
    END IF
  LOOP UNTIL ExitFlag
  
  ERASE TempScrn, UBCustRec, RevText$, UBSetUpRec
  
  HideCursor
  
ExitAdjustments:
  UBLog "OUT: UTIL ADJUST" + CrLf$
EXIT SUB
  
GetCustInfo:
  NumOfCustRecs& = FileSize("UBCUST.DAT") \ UBCustRecLen
  IF CustAcct& > NumOfCustRecs& OR CustAcct& <= 0 THEN
    CustAcct& = 0
    OK = MsgBox%("UB.QSL", "BADACCTN")
    Frm(1).FldNo = 1
    GOSUB ClearForm
    GOTO SkipCustInfo
  END IF
  
  IF IsDeleted(CustAcct&) THEN
    CustAcct& = 0
    OK = MsgBox%("UB.QSL", "DELACCTN")
    Frm(1).FldNo = 1
    GOSUB ClearForm
    GOTO SkipCustInfo
  END IF
  
  GOSUB ClearForm
  
  CustFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CustFile LEN = UBCustRecLen
  GET CustFile, CustAcct&, UBCustRec(1)
  CLOSE CustFile
  
  FOR Cnt = 1 TO 7
    IF LEN(QPTrim$(UBCustRec(1).LocMeters(Cnt).MTRType)) = 0 THEN
      Fld(ReadFlds(Cnt)).Protected = True
    ELSE
      Fld(ReadFlds(Cnt)).Protected = False
    END IF
    QPrintRC UBCustRec(1).LocMeters(Cnt).MTRNUM, Cnt + 13, 8, -1
    QPrintRC UBCustRec(1).LocMeters(Cnt).MTRType, Cnt + 13, 24, -1
  NEXT
  
  LSET Form$(CustAcctFld, 0) = STR$(CustAcct&)
  LSET Form$(2, 0) = UBCustRec(1).CUSTNAME
  LSET Form$(3, 0) = UBCustRec(1).ADDR1
  SaveField 0, Form$(), Fld(), BadField
  CustAcct& = QPValL(Form$(CustAcctFld, 0))
  
  MPaintBox 22, 35, 22, 41, 115
  FirstTime = True
  
SkipCustInfo:
  Action = 1
  RETURN
  
ShowRevSources:
  FOR RevCnt = 1 TO NumOfRevs
    QPrintRC RevText$(RevCnt), RevCnt + 3, 42, -1
  NEXT
  RETURN
  
ClearForm:
  TempDate = Date2Num(Form$(AdjDateFld, 0))
  
  IF TempDate > 0 AND TempDate <> AdjDate THEN
    AdjDate = TempDate
  END IF
  
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""       '--Clear all fields
  NEXT
  
  LSET Form$(AdjDateFld, 0) = Num2Date(AdjDate)
  
  BeenEditedFlag = False        'clear the edited flag
  FOR Cnt = 1 TO 7
    QPrintRC "            ", Cnt + 13, 8, -1
    QPrintRC " ", Cnt + 13, 24, -1
  NEXT
  
  RETURN
  
SaveTransaction:
  
  SaveScrn TempScrn()
  DisplayUBScrn "UPDATDSK"
  REDIM UBTransRec(1) AS UBTransRecType
  UBTransRecLen = LEN(UBTransRec(1))
  
  TransDate = Date2Num(Form$(AdjDateFld, 0))
  TransAmt# = Value#(Form$(TotAmtFld, 0), ECode)
  
  UBTransRec(1).TransDate = TransDate
  'UBTransRec(1)CustLocation = CustAcct&
  UBTransRec(1).CustStatus = UBCustRec(1).STATUS
  UBTransRec(1).CustAcctNo = CustAcct&
  UBTransRec(1).BillMsg = QPTrim$(Form$(NoteFld, 0))
  UBTransRec(1).TransAmt = TransAmt#
  
  FOR Cnt = 1 TO NumOfRevs
    IF LEN(QPTrim$(Form$(CashFlds(Cnt), 0))) > 0 THEN
      UBTransRec(1).RevAmt(Cnt) = Value#(Form$(CashFlds(Cnt), 0), ECode)
    END IF
  NEXT
  
  FOR Cnt = 1 TO 7
    IF LEN(QPTrim$(Form$(ReadFlds(Cnt), 0))) > 0 THEN
      UBTransRec(1).CurRead(Cnt) = Value#(Form$(ReadFlds(Cnt), 0), ECode)
    END IF
  NEXT
  
  AdjType$ = UCASE$(LEFT$(Form$(AdjTypeFld, 0), 1))
  
  SELECT CASE AdjType$
    
  CASE "U"
    UBTransRec(1).TransDesc = "Upward Adjustment"
    FOR RevCnt = 1 TO MaxRevsCnt
      UBCustRec(1).CurrRevAmts(RevCnt) = Round#(UBCustRec(1).CurrRevAmts(RevCnt) + UBTransRec(1).RevAmt(RevCnt))
      UBCustRec(1).CurrBalance = Round#(UBCustRec(1).CurrBalance + UBTransRec(1).RevAmt(RevCnt))
    NEXT
    UBTransRec(1).TransType = TranUpwardAdjustment
    UBLog "UP ADJUSTED CUST:" + STR$(CustAcct&)
  CASE "D"
    UBTransRec(1).TransDesc = "Downward Adjustment"
    FOR RevCnt = 1 TO MaxRevsCnt
      UBCustRec(1).CurrRevAmts(RevCnt) = Round#(UBCustRec(1).CurrRevAmts(RevCnt) - UBTransRec(1).RevAmt(RevCnt))
      UBCustRec(1).CurrBalance = Round#(UBCustRec(1).CurrBalance - UBTransRec(1).RevAmt(RevCnt))
    NEXT
    UBTransRec(1).TransType = TranDownwardAdjustment
    UBLog "DOWN ADJUSTED CUST:" + STR$(CustAcct&)
  END SELECT
  
  UBTransRec(1).RunBalance = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)

  AdjDesc$ = QPTrim$(Form$(6, 0))
  IF LEN(AdjDesc$) > 0 THEN
    UBTransRec(1).TransDesc = AdjDesc$
  END IF
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTransRecLen
  
  NextTranRecs& = (LOF(UBTran) \ UBTransRecLen) + 1
  PrevLastTrans& = UBCustRec(1).LastTrans
  UBTransRec(1).PrevTrans = PrevLastTrans&
  UBCustRec(1).LastTrans = NextTranRecs&
  
  IF Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance) = 0 THEN
    IF UBCustRec(1).STATUS = "B" THEN
      CustChCnt = CustChCnt + 1
      UBLog "ADJUST: SET CUST STATUS to I. Acct:" + STR$(UBTransRec(1).CustAcctNo)
      UBCustRec(1).STATUS = "I"
    END IF
  END IF

  PUT UBCust, UBTransRec(1).CustAcctNo, UBCustRec(1)
  PUT UBTran, NextTranRecs&, UBTransRec(1)
  CLOSE UBTran, UBCust

  UBLog "ADJUSTED CUST:" + STR$(CustAcct&) + "  TRANS:" + STR$(NextTranRecs&)
  
  RestScrn TempScrn()
  DisplayUBScrn "UPDATEOK"
  WaitForAction
  RestScrn TempScrn()
  FirstTime = True
RETURN

CalcCashFlds:

  TAmt# = 0
  FOR Cnt = 1 TO MaxRevsCnt
    TAmt# = Round#(TAmt# + Value#(Form$(CashFlds(Cnt), 0), ECode))
  NEXT

  LSET Form$(TotAmtFld, 0) = STR$(TAmt#)
  SaveField TotAmtFld, Form$(), Fld(), BadField
  Action = 2

RETURN

CheckAdjustInfo:
  CurrFld = Frm(1).FldNo
  
  AdjustOKFlag = True
  ConsAdjFlag = False
  BillAdjFlag = False
  
  TestDate = Date2Num(Form$(AdjDateFld, 0))
  IF TestDate < 0 THEN
    OK = MsgBox%("UB.QSL", "UBBADTDT")
    Action = 2
    AdjustOKFlag = False
    Frm(1).FldNo = 4
    GOTO BadAdjustment
  END IF
  
  IF LEN(QPTrim$(Form$(AdjTypeFld, 0))) = 0 THEN
    OK = MsgBox%("UB.QSL", "UBBADTYP")
    Action = 2
    AdjustOKFlag = False
    Frm(1).FldNo = 5
    GOTO BadAdjustment
  END IF
  
  FOR Cnt = 1 TO 7
    IF LEN(QPTrim$(Form$(6 + Cnt, 0))) > 0 THEN
      ConsAdjFlag = True
      EXIT FOR
    END IF
  NEXT
  
  FOR Cnt = 1 TO MaxRevsCnt
    IF LEN(QPTrim$(Form$(13 + Cnt, 0))) > 0 THEN
      BillAdjFlag = True
      EXIT FOR
    END IF
  NEXT
  
  IF (NOT ConsAdjFlag AND NOT BillAdjFlag) THEN
    OK = MsgBox%("UB.QSL", "UBBADADJ")
    Action = 2
    AdjustOKFlag = False
    Frm(1).FldNo = CurrFld - 1
    GOTO BadAdjustment
  END IF
  
  
BadAdjustment:
  RETURN
  
END SUB

