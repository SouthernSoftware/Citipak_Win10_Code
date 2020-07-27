DEFINT A-Z
DECLARE SUB ApplyPenalty ()
DECLARE SUB ApplyReconnect ()
DECLARE SUB CursorOff ()
DECLARE SUB PenaltyProcess ()
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetUpLen%)
DECLARE SUB MakeSequenceIndex (IndexText$)
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE FUNCTION FindRateTbl% (RATECODE$, NumOfRates%, UBRateTbls() AS ANY)
DECLARE FUNCTION GetNumOfRevs% ()
DECLARE FUNCTION GetRevCharge# (RateTbl AS ANY, TMeterConsp&)
DECLARE FUNCTION Round# (N#)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB MakeZipCodeIndex ()
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB MakePostalIndex (IndexText$)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB MakeTempIndex (Indextype%)
DECLARE SUB PreBillReport ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB ManualMeterReading ()
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB EnterMeterReadingsMenu ()
DECLARE FUNCTION GetNumOfCust% ()
DECLARE FUNCTION GetNumRateRecs% ()
DECLARE FUNCTION Exist% (FileName$)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION FLof& (FileHandle%)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB HideCursor ()
DECLARE SUB ClearScrn ()
  '$INCLUDE: 'DefCnf.BI'
  
DECLARE FUNCTION WEnvTest% ()
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FClose (Handle%)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB SortStr2 (BYVAL Address%, Size%, Direction%)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION QPValI% (Number$)
DECLARE SUB WaitForAction ()
DECLARE SUB BlockClear ()
  
DECLARE SUB SaveScrn (TempScrn())
DECLARE SUB RestScrn (TempScrn())
  
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
  '$INCLUDE: 'ubsetbal.BI'
  
  CONST False = 0, True = NOT False
  
  'PenaltyProcess

  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 3)

  MChoice$(1) = "  Apply Late Penalties   "
  MChoice$(2) = "  Apply Reconnect Fees   "
  MChoice$(3) = "  Exit to OS "

  MaxLen = 0    'Set menu width to zero
  BoxBot = 18   'limit the box length to go no lower than line 20
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
  Col = ((80 - MaxLen) \ 2)

  DO

    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear

    TitleBox 2, Col, MaxLen + 3, "Penalty/Reconnect Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

    ShowCursor
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf

    IF Ky$ = CHR$(27) THEN
      Choice = 0
      ExitFlag = True
      EXIT DO
    END IF

    SELECT CASE Choice
    CASE 1
      ApplyPenalty
    CASE 2
      ApplyReconnect
    CASE 3
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP UNTIL ExitFlag

  IF WEnvTest AND INSTR(COMMAND$, "TEST") > 0 THEN
    HideCursor
    '   ClearScrn
    END
  ELSE
    RUN "UBBILLIN"
  END IF

SUB ApplyPenalty

  REDIM TempScrn(0)

  REDIM UBCustRec(1) AS NewUBCustRecType
  REDIM UBSetUpRec(1) AS UBSetupRecType
  REDIM UBTranRec(1 TO 2) AS UBTransRecType

  UBCustRecLen = LEN(UBCustRec(1))
  UBTranRecLen = LEN(UBTranRec(1))

  Today = Date2Num(DATE$)

  LibName$ = "UB"
  ScrnName$ = "UBPENREC"

  MenuTitle$ = " Apply Penalty Fees "
  MTLen = LEN(MenuTitle$)

  '--define the multi-choice fields
  NumFlds = -1
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1

  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  '--Set screen number to one and display screen
  DisplayUBScrn ScrnName$
  QPrintRC "[", 8, 20, -1
  QPrintRC MenuTitle$, 8, 21, 14
  QPrintRC "]", 8, 21 + MTLen, -1
  QPrintRC "Percent", 12, 23, -1

  Action = 1

  FirstTime = True

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF FirstTime THEN
      FirstTime = False
      Action = 1
      LSET Form$(1, 0) = DATE$
    END IF

    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      GOSUB CheckPFields
      Action = 1
    CASE EscKey
      ExitFlag = True
    END SELECT

  LOOP UNTIL ExitFlag OR OKFlag

  IF ExitFlag THEN GOTO ExitPenalty

  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen

  NumOfRec = LOF(UBCust) \ UBCustRecLen
  BlockClear
  ShowProcessingScrn "Applying Penalty Charges"
  FOR Cnt = 1 TO NumOfRec
    GET UBCust, Cnt, UBCustRec(1)
    IF UBCustRec(1).DelFlag <> 0 THEN
      GOTO PSkipEm
    END IF
    IF UBCustRec(1).LateFee <> "Y" THEN
      GOTO PSkipEm
    END IF
    IF UBCustRec(1).EPPFlag = "Y" THEN
      GOTO PSkipEm
    END IF

    IF Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance) > 0 THEN
      GOSUB FindLastBill
      IF OK4Penalty THEN
        GOSUB DoCustPenalty
        PenCnt = PenCnt + 1
      END IF
    END IF

PSkipEm:
    ShowPctComp Cnt, NumOfRec
  NEXT
  CLOSE
  BlockClear
  DisplayUBScrn "UPDATEOK"
  WaitForAction

EXIT SUB

FindLastBill:

  OK4Penalty = False  'assume no penalty

  ThisTrans& = UBCustRec(1).LastTrans

  DO WHILE ThisTrans& > 0
    GET UBTran, ThisTrans&, UBTranRec(1)
    IF (UBTranRec(1).TransType = TranUtilityBill) OR (UBTranRec(1).TransType = TranUtilityBill + 100) THEN
      IF LEN(QPTrim$(UBTranRec(1).PenalFlag)) = 0 THEN
        OK4Penalty = True
      END IF
      EXIT DO
    END IF
    ThisTrans& = UBTranRec(1).PrevTrans
  LOOP

RETURN

DoCustPenalty:

  REDIM Pens(1 TO 5) AS DOUBLE
  REDIM PUBTranRec(1) AS UBTransRecType

  TotPen# = 0

  IF UBTranRec(1).RevAmt(1) > 0 THEN
    Pens(1) = Round#(UBTranRec(1).RevAmt(1) * Pct#)
  END IF
  IF UBTranRec(1).RevAmt(2) > 0 THEN
    Pens(2) = Round#(UBTranRec(1).RevAmt(2) * Pct#)
  END IF
  IF UBTranRec(1).RevAmt(3) > 0 THEN
    Pens(3) = Round#(UBTranRec(1).RevAmt(3) * Pct#)
  END IF
  IF UBTranRec(1).RevAmt(9) > 0 THEN
    Pens(4) = Round#(UBTranRec(1).RevAmt(9) * Pct#)
  END IF
  IF UBTranRec(1).RevAmt(11) > 0 THEN
    Pens(5) = Round#(UBTranRec(1).RevAmt(11) * Pct#)
  END IF

  IF UBTranRec(1).RevAmt(4) = 0 THEN
    TotPen# = Round#(TotPen# + Pens(1))
  ELSE
    Pens(1) = 0
  END IF
  IF UBTranRec(1).RevAmt(5) = 0 THEN
    TotPen# = Round#(TotPen# + Pens(2))
  ELSE
    Pens(2) = 0
  END IF
  IF UBTranRec(1).RevAmt(6) = 0 THEN
    TotPen# = Round#(TotPen# + Pens(3) + Pens(5))
  ELSE
    Pens(3) = 0
    Pens(5) = 0
  END IF
  IF UBTranRec(1).RevAmt(10) = 0 THEN
    TotPen# = Round#(TotPen# + Pens(4))
  ELSE
    Pens(4) = 0
  END IF

  IF TotPen# <= 0 THEN
    GOTO SkipTrans
  END IF

  PUBTranRec(1).TransAmt = TotPen#
  PUBTranRec(1).RevAmt(4) = Pens(1)
  PUBTranRec(1).RevAmt(5) = Pens(2)
  PUBTranRec(1).RevAmt(6) = Round#(Pens(3) + Pens(5))
  PUBTranRec(1).RevAmt(10) = Pens(4)

  PUBTranRec(1).TransDate = PDate
  PUBTranRec(1).TransType = TranPenaltyCharge
  PUBTranRec(1).TransDesc = "LATE PENALTY"
  PUBTranRec(1).CustAcctNo = Cnt

  UBCustRec(1).CurrBalance = Round#(UBCustRec(1).CurrBalance + TotPen#)

  UBCustRec(1).CurrRevAmts(4) = Round#(UBCustRec(1).CurrRevAmts(4) + Pens(1))
  'water pen rev
  UBCustRec(1).CurrRevAmts(5) = Round#(UBCustRec(1).CurrRevAmts(5) + Pens(2))
  'sewer pen rev
  UBCustRec(1).CurrRevAmts(6) = Round#(UBCustRec(1).CurrRevAmts(6) + Pens(3))
  'Elec/Security lights pen rev
  UBCustRec(1).CurrRevAmts(10) = Round#(UBCustRec(1).CurrRevAmts(10) + Pens(4))
  'trash pen rev
  UBCustRec(1).CurrRevAmts(6) = Round#(UBCustRec(1).CurrRevAmts(6) + Pens(5))
  'Elec/Security lights pen rev

  PUBTranRec(1).RunBalance = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)

  UBTranRec(1).PenalFlag = "Y"
  PUT UBTran, ThisTrans&, UBTranRec(1)

  PUBTranRec(1).PrevTrans = UBCustRec(1).LastTrans
  NextTransRec& = (LOF(UBTran) / UBTranRecLen) + 1
  UBCustRec(1).LastTrans = NextTransRec&

  PUT UBTran, NextTransRec&, PUBTranRec(1)
  PUT UBCust, Cnt, UBCustRec(1)
  'END IF
SkipTrans:

RETURN

CheckPFields:
  OKFlag = False  'Assume all is NOT well

  PDate = Date2Num(Form$(1, 0))
  Pct# = Value#(Form$(2, 0), ECode)

  IF PDate <= (Today - 180) THEN
    OK = MsgBox%(LibName$, "UBPBDATE")
    Frm(1).FldNo = 1
    GOTO CheckPReturn
  ELSE
    DateOk = True
  END IF

  IF Pct# <= 0 THEN
    SaveScrn TempScrn()
    DisplayUBScrn "UBBADAMT"
    HideCursor
    QPrintRC "Invalid Penalty Percentage!", 11, 27, -1
    QPrintRC "Please Reenter Amount.", 13, 29, -1
    ShowCursor
    WaitForAction
    RestScrn TempScrn()
    Frm(1).FldNo = 2
    GOTO CheckPReturn
  ELSE
    PctOk = True
    Pct# = Round#(Pct# * .01)
  END IF
  OKFlag = True

CheckPReturn:
RETURN

ExitPenalty:
END SUB

SUB ApplyReconnect
  
  REDIM TempScrn(0)

  REDIM UBCustRec(1) AS NewUBCustRecType
  REDIM UBSetUpRec(1) AS UBSetupRecType
  REDIM UBTranRec(1 TO 2) AS UBTransRecType

  UBCustRecLen = LEN(UBCustRec(1))
  UBTranRecLen = LEN(UBTranRec(1))

  Today = Date2Num(DATE$)

  LibName$ = "UB"
  ScrnName$ = "UBPENREC"

  MenuTitle$ = " Apply Reconnect Fees "
  MTLen = LEN(MenuTitle$)

  '--define the multi-choice fields
  NumFlds = -1
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1

  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  '--Set screen number to one and display screen
  DisplayUBScrn ScrnName$
  QPrintRC "[", 8, 20, -1
  QPrintRC MenuTitle$, 8, 21, 14
  QPrintRC "]", 8, 21 + MTLen, -1
  QPrintRC "Amount", 12, 24, -1
  
  Action = 1

  FirstTime = True

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF FirstTime THEN
      FirstTime = False
      Action = 1
      LSET Form$(1, 0) = DATE$
    END IF

    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      GOSUB RCheckPFields
      Action = 1
    CASE EscKey
      ExitFlag = True
    END SELECT

  LOOP UNTIL ExitFlag OR OKFlag

  IF ExitFlag THEN GOTO RExitPenalty

  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen

  NumOfRec = LOF(UBCust) \ UBCustRecLen
  BlockClear

  ShowProcessingScrn "Applying Reconnect Fees"
  FOR Cnt = 1 TO NumOfRec
    OK2Apply = False
    GET UBCust, Cnt, UBCustRec(1)
    IF UBCustRec(1).DelFlag <> 0 THEN
      GOTO RSkipEm
    END IF
    IF UBCustRec(1).EPPFlag = "Y" THEN
      GOTO RSkipEm
    END IF
    
    IF UBCustRec(1).CurrBalance > 0 THEN
      IF UBCustRec(1).Status = "A" THEN
        GOSUB CheckCustRevs
        IF OK2Apply THEN
          GOSUB DOReconFee
          PenCnt = PenCnt + 1
        END IF
      END IF
    END IF

RSkipEm:
    ShowPctComp Cnt, NumOfRec
  NEXT
  CLOSE
  BlockClear
  DisplayUBScrn "UPDATEOK"
  WaitForAction

EXIT SUB

CheckCustRevs:

  OK2Apply = False  'assume no penalty
  DOWater = False
  DOElect = False

'        IF irev7 > 0 OR irev8 > 0 THEN RETURN
'        IF irev3 > 0 THEN IF irev7 > 0 THEN RETURN
'        IF irev3 = 0 AND irev1 > 0 THEN IF irev8 > 0 THEN RETURN
'        END IF

  'skip em if they currently have any reconnect
  IF (UBCustRec(1).CurrRevAmts(7) > 0) OR (UBCustRec(1).CurrRevAmts(8) > 0) THEN
    GOTO NoReconReturn
  END IF

  IF (UBCustRec(1).CurrRevAmts(1) > 7.5) AND (UBCustRec(1).CurrRevAmts(8) = 0) THEN
    DOWater = True
    OK2Apply = True
  END IF

  IF UBCustRec(1).CurrRevAmts(3) > 7.5 AND (UBCustRec(1).CurrRevAmts(7) = 0) THEN
    DOElect = True
    OK2Apply = True
  END IF

NoReconReturn:
RETURN

DOReconFee:

  REDIM PUBTranRec(1) AS UBTransRecType

  TotPen# = 0

  IF DOElect THEN
    PUBTranRec(1).RevAmt(7) = Pct#
    TotPen# = TotPen# + Pct#
    UBCustRec(1).CurrRevAmts(7) = PUBTranRec(1).RevAmt(7)
    GOTO Skip2
  END IF

  IF DOWater THEN
    PUBTranRec(1).RevAmt(8) = Pct#
    TotPen# = TotPen# + Pct#
    UBCustRec(1).CurrRevAmts(8) = PUBTranRec(1).RevAmt(8)
  END IF
  
Skip2:
  TotPen# = Round#(TotPen#)

  IF TotPen# <= 0 THEN
    GOTO RSkipTrans
  END IF

  PUBTranRec(1).TransAmt = TotPen#
  PUBTranRec(1).TransDate = PDate
  PUBTranRec(1).TransType = TranReconnectFee
  PUBTranRec(1).TransDesc = "RECONNECT FEE"
  PUBTranRec(1).CustAcctNo = Cnt

  UBCustRec(1).CurrBalance = Round#(UBCustRec(1).CurrBalance + TotPen#)
  PUBTranRec(1).RunBalance = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)

  PUBTranRec(1).PrevTrans = UBCustRec(1).LastTrans
  NextTransRec& = (LOF(UBTran) / UBTranRecLen) + 1
  UBCustRec(1).LastTrans = NextTransRec&

  PUT UBTran, NextTransRec&, PUBTranRec(1)
  PUT UBCust, Cnt, UBCustRec(1)

'  LPRINT UBCustRec(1).CustName, Cnt, ;
'  IF DOElect THEN
'    LPRINT "Elect"
'  ELSE
'    LPRINT "Water"
'  END IF

RSkipTrans:

RETURN

RCheckPFields:
  OKFlag = False  'Assume all is NOT well

  PDate = Date2Num(Form$(1, 0))
  Pct# = Value#(Form$(2, 0), ECode)

  IF PDate <= (Today - 180) THEN
    OK = MsgBox%(LibName$, "UBPBDATE")
    Frm(1).FldNo = 1
    GOTO RCheckPReturn
  END IF

  IF Pct# <= 0 THEN
    SaveScrn TempScrn()
    DisplayUBScrn "UBBADAMT"
    HideCursor
    QPrintRC "Invalid Reconnect Amount!", 11, 27, -1
    QPrintRC "Please Reenter Amount.", 13, 29, -1
    ShowCursor
    WaitForAction
    RestScrn TempScrn()
    Frm(1).FldNo = 2
    GOTO RCheckPReturn
  END IF

  OKFlag = True

RCheckPReturn:
RETURN

RExitPenalty:
'reconnect:
'        IF irev7 > 0 OR irev8 > 0 THEN RETURN
'        IF irev3 > 0 THEN
'         IF irev7 > 0 THEN RETURN
'        END IF
'        IF irev3 = 0 AND irev1 > 0 THEN
'         IF irev8 > 0 THEN RETURN
'        END IF
'
' IF irev3 > 7.5 THEN
'        IF irev7 = 0 THEN fee# = rFEE#: GOTO rec1
' END IF
' IF irev1 > 7.5 THEN
'       IF irev8 = 0 THEN fee# = rFEE#
' END IF
'rec1:
'fee# = INT((fee# * 100) + .5) / 100
'IF fee# = 0 THEN RETURN
'LSET PASTDUE$ = MKS$((CVS(PASTDUE$) + fee#))
'
'IF irev3 > 0 THEN
' IF irev7 = 0 THEN
'  LSET irevenue7$ = MKS$(CVS(irevenue7$) + fee#): PUT 1, CVS(mrec$): GOTO rec2
' END IF
'END IF
'IF irev1 > 0 THEN
' IF irev8 = 0 THEN
'  LSET irevenue8$ = MKS$(CVS(irevenue8$) + fee#): PUT 1, CVS(mrec$)
' END IF
'END IF
'
'rec2:
' GOSUB transaction
'RETURN
'

END SUB

