DECLARE SUB oPenaltyProcess ()
DEFINT A-Z
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
  
  oPenaltyProcess
  GOTO Here

  REDIM MChoice$(1 TO 4)

  MChoice$(1) = " Calculate Penalty Charges"
  MChoice$(2) = " Print Customer Penalty Report"
  MChoice$(3) = " Post Penalty Transactions "
  MChoice$(4) = " Exit to OS"

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
  Row = ((25 - (UBOUND(MChoice$))) \ 2)
  Col = ((80 - MaxLen) \ 2) - 2

  DO

    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear

    TitleBox 2, Col, MaxLen + 3, "Penalty Processing Menu  ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

    ShowCursor

    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0

    SELECT CASE Choice
    CASE 1
      oPenaltyProcess
    CASE 2
    CASE 3
    CASE 4
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP

  ' RUN "ubmenu"

Here:
  IF WEnvTest AND INSTR(COMMAND$, "TEST") > 0 THEN
    HideCursor
    '   ClearScrn
    END
  ELSE
    RUN "UBBILLIN"
  END IF

SUB oPenaltyProcess
  
  SHARED Choice$()
  
  REDIM TempScrn(0)
  REDIM Source$(15)
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  REDIM UBSetUpRec(1) AS UBSetupRecType
  REDIM UBTranRec(1 TO 3) AS UBTransRecType
  REDIM TaxAmt(1 TO 15) AS DOUBLE
  
  UBCustRecLen = LEN(UBCustRec(1))
  UBTranRecLen = LEN(UBTranRec(1))
  
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen

  TownName$ = UBSetUpRec(1).UtilName
  IF INSTR(TownName$, "TENN") > 0 AND INSTR(TownName$, "RIDGE") > 0 THEN
    TennFlag = True
  END IF

  FOR Cnt = 1 TO MaxRevsCnt
    Source$(Cnt) = UBSetUpRec(1).Revenues(Cnt).REVNAME
    TaxAmt(Cnt) = UBSetUpRec(1).Revenues(Cnt).TAXRATE
  NEXT
  UseCycle$ = UBSetUpRec(1).BILLCYCL
  
  LibName$ = "UB"
  ScrnName$ = "UBPENALT"
  
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
  
  '--Clear all fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT
  
  '--Set choices
  NumOfRevs = 0
  FOR Cnt = 1 TO MaxRevsCnt
    TempRev$ = QPTrim$(Source$(Cnt))
    IF LEN(TempRev$) = 0 THEN
      NumOfRevs = Cnt - 1
      EXIT FOR
    END IF
  NEXT
  
  REDIM Choice$(NumOfRevs, 1)
  Choice$(0, 0) = "3"
  FOR TCnt = 1 TO NumOfRevs
    Choice$(TCnt, 0) = Source$(TCnt)
  NEXT TCnt
  
  Choice$(0, 1) = "4"
  Choice$(1, 1) = "Current Balance"
  Choice$(2, 1) = "Previous Balance"
  Choice$(3, 1) = "Total Balance"
  
  ' USE CYCLE CHECK
  
  '--Set screen number to one and display screen
  DisplayUBScrn ScrnName$

  Action = 1
  FirstTime = True
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      FirstTime = False
      Action = 1
      LSET Form$(1, 0) = DATE$
      LSET Form$(7, 0) = "0"
      IF UseCycle$ = "N" THEN
        LSET Form$(8, 0) = "0"
        LSET Form$(9, 0) = "0"
        Fld(8).Protected = True
        Fld(9).Protected = True
        NoCycleFlag = True
      END IF
    END IF
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      GOSUB CheckPenaltyFlds
      'If FCheck for valid Data in Fields
      'Then Process the Penalties
    CASE EscKey
      ExitFlag = True
    END SELECT
    
  LOOP UNTIL ExitFlag OR OKFlag
  
  IF ExitFlag THEN GOTO ExitPenalty
  
  BlockClear
  ShowProcessingScrn "Applying Penalty Charges"
  
  MinBalance# = Value#(Form$(7, 0), ECode)
  
  IF MinBalance# < 0 THEN
    MinBalance# = 0
  END IF
  
  'Get if we are applying to Current, Previous or Both
  SELECT CASE LEFT$(QPTrim$(Form$(4, 0)), 1)
  CASE "C"
    UsePrevFlag = False
    UseCurrFlag = True
  CASE "P"
    UsePrevFlag = True
    UseCurrFlag = False
  CASE "T"
    UsePrevFlag = True
    UseCurrFlag = True
  END SELECT
  
  'Get percent or fixed amount
  PctAmt# = Value#(Form$(5, 0), ECode)
  FixAmt# = Value#(Form$(6, 0), ECode)
  IF PctAmt# > 0 THEN
    PctAmt# = Round#(PctAmt# * .01)
    FixAmt# = 0
    UsePctFlag = True
  ELSE
    PctAmt# = 0
    UsePctFlag = False
  END IF
  
  'Get the Rev source number
  FOR Cnt = 1 TO NumOfRevs
    ThisRev$ = QPTrim$(Form$(3, 0))
    IF INSTR(Choice$(Cnt, 0), ThisRev$) THEN
      RevSource = Cnt
      EXIT FOR
    END IF
  NEXT
  
  'Get Who/How to process
  '***********************
  IF Value#(Form$(8, 0), ECode) > 0 THEN
    FirstCycle = Value#(Form$(8, 0), ECode)
    LastCycle = Value#(Form$(9, 0), ECode)
    UsingCycle = True
  ELSE
    FirstBook = Value#(Form$(10, 0), ECode)
    LastBook = Value#(Form$(11, 0), ECode)
    UsingBook = True
  END IF
  
  TransDesc$ = QPTrim$(Form$(2, 0))
  
  FOpenS "UBCUST.DAT", CHandle  'open data file
  
  NumCustRecs& = FLof(CHandle) / UBCustRecLen
  
  FOR CCnt& = 1 TO NumCustRecs&
    FGetRTA CHandle, UBCustRec(1), CCnt&, UBCustRecLen
    'IF CCnt& = 202 THEN STOP
    IF NOT UBCustRec(1).DelFlag THEN
      IF UBCustRec(1).LateFee = "Y" THEN
        LSET UBTranRec(1) = UBTranRec(2)        'Transrec 2 is blank
        'Make a clean transaction record
        IF UBCustRec(1).STATUS = "A" THEN       'if they are not inactive
          '05-01-97 fixed bug where CurrBalance+PrevBalance is <= 0
          IF Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance) > 0 THEN
            IF UBCustRec(1).CurrBalance >= MinBalance# OR UBCustRec(1).PrevBalance > MinBalance# THEN
              'if they have any balance
              IF UsingBook THEN 'if they want it by Book
                ThisBook = QPValI(UBCustRec(1).BOOK)
                IF ThisBook >= FirstBook AND ThisBook <= LastBook THEN
                  'if this is in the correct book
                  IF UsePctFlag THEN            'if they want a percentage penalty
                    IF UsePrevFlag AND NOT UseCurrFlag THEN     'using prev not curr
'030398 Modified to consider a credit in cur or prev balances
                      IF UBCustRec(1).CurrBalance < 0 THEN
                        PenBal# = UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance
                      ELSE
                        PenBal# = UBCustRec(1).PrevBalance
                      END IF
                      IF PenBal# <= MinBalance# THEN            'if cust had prev bal
                        GOTO SkipEM
                      END IF
                      CustPenalty# = Round#(PenBal# * PctAmt#)
                      UBCustRec(1).CurrBalance = Round#(UBCustRec(1).CurrBalance + CustPenalty#)
                      UBCustRec(1).CurrRevAmts(RevSource) = Round#(UBCustRec(1).CurrRevAmts(RevSource) + CustPenalty#)
                      '-VVV- this is correct
                      GOSUB MakeTransaction
                    ELSEIF UseCurrFlag AND NOT UsePrevFlag THEN 'using curr not prev
'030398 Modified to consider a credit in cur or prev balances
                      IF UBCustRec(1).PrevBalance < 0 THEN
                        PenBal# = UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance
                      ELSE
                        PenBal# = UBCustRec(1).CurrBalance
                      END IF
'code added to exclude tax
'092898 Said they didn't take partial payments - Not!
                      IF TennFlag THEN 'AND UBCustRec(1).TaxExpt <> "Y" THEN
                        GOSUB GetTennRidgeLastBill
                      END IF

                      IF PenBal# <= MinBalance# THEN
                        GOTO SkipEM
                      END IF
                      CustPenalty# = Round#(PenBal# * PctAmt#)
                      UBCustRec(1).CurrBalance = Round#(UBCustRec(1).CurrBalance + CustPenalty#)
                      UBCustRec(1).CurrRevAmts(RevSource) = Round#(UBCustRec(1).CurrRevAmts(RevSource) + CustPenalty#)
                      GOSUB MakeTransaction
                    ELSEIF UsePrevFlag AND UseCurrFlag THEN     'use curr and prev
                      TotalBalance# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
                      IF TotalBalance# <= MinBalance# THEN
                        GOTO SkipEM
                      END IF
                      CustPenalty# = Round#(TotalBalance# * PctAmt#)
                      UBCustRec(1).CurrBalance = Round#(UBCustRec(1).CurrBalance + CustPenalty#)
                      UBCustRec(1).CurrRevAmts(RevSource) = Round#(UBCustRec(1).CurrRevAmts(RevSource) + CustPenalty#)
                      '-VVVV-  This was rem Why???
                      GOSUB MakeTransaction
                    END IF
                  ELSE          'Using a FIXED penalty amount
                    IF UsePrevFlag AND NOT UseCurrFlag THEN
'030398 Modified to consider a credit in cur or prev balances
                      IF UBCustRec(1).CurrBalance < 0 THEN
                        PenBal# = UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance
                      ELSE
                        PenBal# = UBCustRec(1).PrevBalance
                      END IF
                      IF PenBal# <= MinBalance# THEN            'if cust had prev bal
                        GOTO SkipEM
                      END IF
                      CustPenalty# = FixAmt#
                      UBCustRec(1).CurrBalance = Round#(UBCustRec(1).CurrBalance + FixAmt#)
                      UBCustRec(1).CurrRevAmts(RevSource) = Round#(UBCustRec(1).CurrRevAmts(RevSource) + CustPenalty#)
                      GOSUB MakeTransaction
                    ELSEIF UseCurrFlag AND NOT UsePrevFlag THEN
'030398 Modified to consider a credit in cur or prev balances
                      IF UBCustRec(1).PrevBalance < 0 THEN
                        PenBal# = UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance
                      ELSE
                        PenBal# = UBCustRec(1).CurrBalance
                      END IF
                      IF PenBal# <= MinBalance# THEN
                        GOTO SkipEM
                      END IF
                      CustPenalty# = FixAmt#
                      UBCustRec(1).CurrBalance = Round#(UBCustRec(1).CurrBalance + FixAmt#)
                      UBCustRec(1).CurrRevAmts(RevSource) = Round#(UBCustRec(1).CurrRevAmts(RevSource) + CustPenalty#)
                      GOSUB MakeTransaction
                    ELSEIF UsePrevFlag AND UseCurrFlag THEN
                      'do not need to check for prev >0 or curr>0 here!!
                      CustPenalty# = FixAmt#
                      TotalBalance# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
                      IF TotalBalance# > MinBalance# THEN
                        UBCustRec(1).CurrBalance = Round#(UBCustRec(1).CurrBalance + FixAmt#)
                        UBCustRec(1).CurrRevAmts(RevSource) = Round#(UBCustRec(1).CurrRevAmts(RevSource) + CustPenalty#)
                        GOSUB MakeTransaction
                      END IF
                    END IF
                  END IF
                END IF
              ELSEIF UsingCycle THEN            'they using cycles
                ThisCycle = UBCustRec(1).BILLCYCL
                IF ThisCycle >= FirstCycle AND ThisCycle <= LastCycle THEN
                  IF UsePctFlag THEN
                    IF UsePrevFlag AND NOT UseCurrFlag THEN     'use prev not curr
'030398 Modified to consider a credit in cur or prev balances
                      IF UBCustRec(1).CurrBalance < 0 THEN
                        PenBal# = UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance
                      ELSE
                        PenBal# = UBCustRec(1).PrevBalance
                      END IF
                      IF PenBal# <= MinBalance# THEN            'if cust had prev bal
                        GOTO SkipEM
                      END IF
                      CustPenalty# = Round#(PenBal# * PctAmt#)
                      UBCustRec(1).CurrBalance = Round#(UBCustRec(1).CurrBalance + CustPenalty#)
                      UBCustRec(1).CurrRevAmts(RevSource) = Round#(UBCustRec(1).CurrRevAmts(RevSource) + CustPenalty#)
                      GOSUB MakeTransaction
                    ELSEIF UseCurrFlag AND NOT UsePrevFlag THEN 'use curr not prev
'030398 Modified to consider a credit in cur or prev balances
                      IF UBCustRec(1).PrevBalance < 0 THEN
                        PenBal# = UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance
                      ELSE
                        PenBal# = UBCustRec(1).CurrBalance
                      END IF
                      IF PenBal# <= MinBalance# THEN
                        GOTO SkipEM
                      END IF
                      CustPenalty# = Round#(PenBal# * PctAmt#)
                      UBCustRec(1).CurrBalance = Round#(UBCustRec(1).CurrBalance + CustPenalty#)
                      UBCustRec(1).CurrRevAmts(RevSource) = Round#(UBCustRec(1).CurrRevAmts(RevSource) + CustPenalty#)
                      GOSUB MakeTransaction
                    ELSEIF UsePrevFlag AND UseCurrFlag THEN     'use curr and prev
                      TotalBalance# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
                      IF TotalBalance# <= MinBalance# THEN
                        GOTO SkipEM
                      END IF
                      CustPenalty# = Round#(TotalBalance# * PctAmt#)
                      UBCustRec(1).CurrBalance = Round#(UBCustRec(1).CurrBalance + CustPenalty#)
                      UBCustRec(1).CurrRevAmts(RevSource) = Round#(UBCustRec(1).CurrRevAmts(RevSource) + CustPenalty#)
                      GOSUB MakeTransaction
                    END IF
                  ELSE          'Using a FIXED penalty amount
                    IF UsePrevFlag AND NOT UseCurrFlag THEN
'030398 Modified to consider a credit in cur or prev balances
                      IF UBCustRec(1).CurrBalance < 0 THEN
                        PenBal# = UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance
                      ELSE
                        PenBal# = UBCustRec(1).PrevBalance
                      END IF
                      IF PenBal# <= MinBalance# THEN            'if cust had prev bal
                        GOTO SkipEM
                      END IF
                      CustPenalty# = FixAmt#
                      UBCustRec(1).CurrBalance = Round#(UBCustRec(1).CurrBalance + FixAmt#)
                      UBCustRec(1).CurrRevAmts(RevSource) = Round#(UBCustRec(1).CurrRevAmts(RevSource) + CustPenalty#)
                      GOSUB MakeTransaction
                    ELSEIF UseCurrFlag AND NOT UsePrevFlag THEN
'030398 Modified to consider a credit in cur or prev balances
                      IF UBCustRec(1).PrevBalance < 0 THEN
                        PenBal# = UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance
                      ELSE
                        PenBal# = UBCustRec(1).CurrBalance
                      END IF
                      IF PenBal# <= MinBalance# THEN
                        GOTO SkipEM
                      END IF
                      CustPenalty# = FixAmt#
                      UBCustRec(1).CurrBalance = Round#(UBCustRec(1).CurrBalance + FixAmt#)
                      UBCustRec(1).CurrRevAmts(RevSource) = Round#(UBCustRec(1).CurrRevAmts(RevSource) + CustPenalty#)
                      GOSUB MakeTransaction
                    ELSEIF UsePrevFlag AND UseCurrFlag THEN
                      CustPenalty# = FixAmt#
                      TotalBalance# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
                      IF TotalBalance# > MinBalance# THEN
                        UBCustRec(1).CurrBalance = Round#(UBCustRec(1).CurrBalance + CustPenalty#)
                        UBCustRec(1).CurrRevAmts(RevSource) = Round#(UBCustRec(1).CurrRevAmts(RevSource) + CustPenalty#)
                        GOSUB MakeTransaction
                      END IF
                    END IF
                  END IF
                END IF
              END IF
            END IF
          END IF
        END IF
      END IF
    END IF
SkipEM:
    ShowPctComp CCnt&, NumCustRecs&
  NEXT
  FClose CHandle
  '  FClose LHandle
  BlockClear
  DisplayUBScrn "UPDATEOK"
  WaitForAction
ExitPenalty:
  
  ERASE TempScrn, Source$
  ERASE UBCustRec, UBSetUpRec, UBTranRec
  ERASE Frm, Form$, Fld
  
  EXIT SUB
  
MakeTransaction:

'remove this
'  RETURN
  
  UBTranRec(1).TransAmt = CustPenalty#
  UBTranRec(1).RevAmt(RevSource) = CustPenalty#
  UBTranRec(1).TransDate = Date2Num%(Form$(1, 0))
  UBTranRec(1).TransType = TranPenaltyCharge
  UBTranRec(1).TransDesc = TransDesc$
  UBTranRec(1).CustAcctNo = CCnt&
  'UBTranRec(1)CustLocation = LocaRec&
  
  UBTranRec(1).RunBalance = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
  
  UBTranRec(1).PrevTrans = UBCustRec(1).LastTrans
  FOpenS "UBTRANS.DAT", THandle
  NextTransRec& = (FLof(THandle) / UBTranRecLen) + 1
  FPutRTA THandle, UBTranRec(1), NextTransRec&, UBTranRecLen
  FClose THandle
  
  UBCustRec(1).LastTrans = NextTransRec&
  FPutRTA CHandle, UBCustRec(1), CCnt&, UBCustRecLen
  
  RETURN
  
CheckPenaltyFlds:
  CursorOff
  SaveScrn TempScrn()
  IF Value#(Form$(1, 0), ECode) = 0 THEN
    DisplayUBScrn "UBNODATE"
    WaitForAction
    Frm(1).FldNo = 1
  ELSEIF LEN(QPTrim$(Form$(3, 0))) = 0 THEN
    DisplayUBScrn "UBNOREVS"
    WaitForAction
    Frm(1).FldNo = 3
  ELSEIF LEN(QPTrim$(Form$(4, 0))) = 0 THEN
    DisplayUBScrn "UBNOBSOU"
    WaitForAction
    Frm(1).FldNo = 4
  ELSEIF Value#(Form$(5, 0), ECode) <> 0 AND Value#(Form$(6, 0), ECode) <> 0 THEN
    DisplayUBScrn "UBINVPEN"
    WaitForAction
    Frm(1).FldNo = 5
  ELSEIF Value#(Form$(5, 0), ECode) = 0 AND Value#(Form$(6, 0), ECode) = 0 THEN
    DisplayUBScrn "UBNOPEN"
    WaitForAction
    Frm(1).FldNo = 5
  ELSEIF Value#(Form$(8, 0), ECode) = 0 AND Value#(Form$(9, 0), ECode) = 0 AND Value#(Form$(10, 0), ECode) = 0 AND VAL(Form$(11, 0)) = 0 OR VAL(Form$(9, 0)) < VAL(Form$(8, 0)) OR VAL(Form$(11, 0)) < VAL(Form$(10, 0)) THEN
    DisplayUBScrn "UBNOPRT"
    WaitForAction
    Frm(1).FldNo = 7
  ELSE
    OKFlag = True
  END IF
  Action = 1
  RestScrn TempScrn()
  
  RETURN

GetTennRidgeLastBill:

  'IF INSTR(UBCustRec(1).CustName, "DALE WATSON") > 0 THEN STOP

  FOpenS "UBTRANS.DAT", THandle
  PrevTranRec& = UBCustRec(1).LastTrans

  IF PrevTranRec& > 0 THEN
    DO WHILE PrevTranRec& > 0
      FGetRTA THandle, UBTranRec(3), PrevTranRec&, UBTranRecLen
      IF UBTranRec(3).TransType = TranUtilityBill THEN
        PenBal# = Round#(UBTranRec(3).RevAmt(1) + UBTranRec(3).RevAmt(2))
        EXIT DO
      END IF
      PrevTranRec& = UBTranRec(3).PrevTrans
    LOOP
  END IF

  FClose THandle

RETURN

END SUB

