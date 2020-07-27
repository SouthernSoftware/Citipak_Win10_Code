DEFINT A-Z
DECLARE FUNCTION AskAbandonPrint ()
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FLof& (FileHandle%)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Num2Date$ (DateNumb%)
DECLARE FUNCTION PromptSaveData ()
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValI% (Number$)
DECLARE FUNCTION Round# (N#)
DECLARE FUNCTION WEnvTest% ()
DECLARE SUB BlockClear ()
DECLARE SUB ClearScrn ()
DECLARE SUB CursorOff ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB EditPenalties ()
DECLARE SUB FClose (Handle%)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB HideCursor ()
DECLARE SUB KillFile (FileName$)
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetUpLen%)
DECLARE SUB PenaltyProcess ()
DECLARE SUB PenaltyReport ()
DECLARE SUB PostPenalties ()
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB RestScrn (TempScrn())
DECLARE SUB SaveScrn (TempScrn())
DECLARE SUB ShowCursor ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB UBLog (Text2Log$)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB WaitForAction ()
DECLARE SUB ShowCustHistory (RecNo&)
  
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'newcust.bi'
  
  CONST False = 0, True = NOT False
  
  TYPE PenaltyInfoType
    PenDate    AS INTEGER
    PenDesc    AS STRING * 21
    RevSource  AS INTEGER
    ChargeOn   AS STRING * 17
'032299 Changed to a double variable
    PctCharge  AS DOUBLE
    'PctCharge  AS INTEGER
    AmtCharge  AS DOUBLE
    GreatLess  AS STRING * 1
    MinBalance AS DOUBLE
    CycFirst   AS INTEGER
    CycLast    AS INTEGER
    BookFirst  AS INTEGER
    BookLast   AS INTEGER
    PenCnt     AS INTEGER
  END TYPE
  
  CrLf$ = CHR$(13) + CHR$(10)
  
  REDIM MChoice$(1 TO 5)
  
  MChoice$(1) = "Calculate Penalty Charges"
  MChoice$(2) = "Customer Penalty Report"
  MChoice$(3) = "Edit Penalty Transactions "
  MChoice$(4) = "Post Penalty Transactions "
  MChoice$(5) = "Exit to OS"
  
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
  Col = ((80 - MaxLen) \ 2)
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    
    TitleBox 2, Col, MaxLen + 3, "Penalty Processing Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      PenaltyProcess
    CASE 2
      PenaltyReport
    CASE 3
      EditPenalties
    CASE 4
      PostPenalties
    CASE 5
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP
  
  ' RUN "ubmenu"
  
  IF INSTR(COMMAND$, "TEST") > 0 THEN
    HideCursor
    '   ClearScrn
    END
  ELSE
    RUN "UBBILLIN"
  END IF

SUB EditPenalties
  
  UBLog " IN: Edit Penalty File (EPF)"
  
  REDIM TempScrn(0)
  
  IF NOT Exist%("UBPENINF.DAT") THEN
    CursorOff
    BlockClear
    DisplayUBScrn "ERRSCRN1"
    QPrintRC "NO UNPOSTED PENALTY TRANSACTIONS!", 10, 24, -1
    QPrintRC "NOTHING TO EDIT.", 12, 32, -1
    WaitForAction
    GOTO ExitEdPenalty
  END IF
  
  REDIM PenaltyInfo(1) AS PenaltyInfoType
  FGetAH "UBPENINF.DAT", PenaltyInfo(1), LEN(PenaltyInfo(1)), 1
  
  TrShow$ = SPACE$(15)
  
  PenFile$ = "UBPENTRN.DAT"
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  REDIM UBTranRec(1) AS UBTransRecType
  
  UBCustRecLen = LEN(UBCustRec(1))
  UBTranRecLen = LEN(UBTranRec(1))
  
  TransRec& = 1
  NumPTran& = FileSize&(PenFile$) / UBTranRecLen
  
  LibName$ = "UB"
  ScrnName$ = "UBPENED"
  
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
  
  ShowFlag = True
  Action = 1
  FirstTime = True
  
  BlockClear
  DisplayUBScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF Frm(1).FldEdited = True THEN
      EditedFlag = True
    END IF
    
    IF FirstTime THEN
      FirstTime = False
      GOSUB GetPenaltyTrans
    END IF
    IF ShowFlag THEN
      GOSUB ShowTransRec
    END IF
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F3Key
      IF NOT DelFlag THEN
        Ok = MsgBox%("UB", "UBDELPEN")
        SELECT CASE Ok
        CASE 1
          GOSUB MarkTransDeleted
        END SELECT
        Action = 1
      END IF

    CASE F4Key
      IF UBTranRec(1).CustAcctNo > 0 THEN
        ShowCustHistory UBTranRec(1).CustAcctNo
      END IF
      Action = 1
    CASE PgDnKey, -80
      IF EditedFlag THEN
        GOSUB AskAbandon
      ELSE
        Ok2GoFlag = True
      END IF
      IF Ok2GoFlag THEN
        IF TransRec& < NumPTran& THEN
          TransRec& = TransRec& + 1
          GOSUB GetPenaltyTrans
        END IF
      END IF
    CASE PgUpKey, -72
      IF EditedFlag THEN
        GOSUB AskAbandon
      ELSE
        Ok2GoFlag = True
      END IF
      IF Ok2GoFlag THEN
        IF TransRec& > 1 THEN
          TransRec& = TransRec& - 1
          GOSUB GetPenaltyTrans
        END IF
      END IF
    CASE F10Key
      IF EditedFlag THEN
        GOSUB SavePenTrans
      END IF
    CASE EscKey
      IF EditedFlag THEN
        GOSUB AskAbandon
      ELSE
        Ok2GoFlag = True
      END IF
      IF Ok2GoFlag THEN
        ExitFlag = True
      END IF
    CASE -119   'ctrl home
      IF EditedFlag THEN
        GOSUB AskAbandon
      ELSE
        Ok2GoFlag = True
      END IF
      IF Ok2GoFlag THEN
        TransRec& = 1
        GOSUB GetPenaltyTrans
      END IF
    CASE -117   'ctrl End
      IF EditedFlag THEN
        GOSUB AskAbandon
      ELSE
        Ok2GoFlag = True
      END IF
      IF Ok2GoFlag THEN
        TransRec& = NumPTran&
        GOSUB GetPenaltyTrans
      END IF
    CASE -118   'ctrl pgdn
      IF EditedFlag THEN
        GOSUB AskAbandon
      ELSE
        Ok2GoFlag = True
      END IF
      IF Ok2GoFlag THEN
        TransRec& = TransRec& + 50
        IF TransRec& > NumPTran& THEN
          TransRec& = NumPTran&
        END IF
        GOSUB GetPenaltyTrans
      END IF
    CASE -132   'ctrl Pgup
      IF EditedFlag THEN
        GOSUB AskAbandon
      ELSE
        Ok2GoFlag = True
      END IF
      IF Ok2GoFlag THEN
        TransRec& = TransRec& - 50
        IF TransRec& < 1 THEN
          TransRec& = 1
        END IF
        GOSUB GetPenaltyTrans
      END IF
    CASE IS <> 0
      'BEEP
      'STOP
    END SELECT
    
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 16   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 13 TO 22           '--Cancel button
          PressButton 27, 16, 13, 22
        CASE 23 TO 30           'f3
          PressButton F3Key, 16, 23, 30
        CASE 31 TO 40           'f10
          PressButton F10Key, 16, 31, 40
        CASE 41 TO 46           'PgDn
          PressButton PgDnKey, 16, 41, 46
        CASE 47 TO 52           'PgUp
          PressButton PgUpKey, 16, 47, 52
        CASE 53 TO 61           'Ctrl-PgDn
          PressButton -118, 16, 53, 61
        CASE 62 TO 70           'Ctrl-PgUp
          PressButton -132, 16, 62, 70
        END SELECT
      END SELECT                'row
    END IF
    
  LOOP UNTIL ExitFlag
  
  GOTO ExitEdPenalty
  
GetPenaltyTrans:
  EditedFlag = False
  DelFlag = False
  ShowFlag = True
  PHandle = FREEFILE
  OPEN PenFile$ FOR RANDOM SHARED AS PHandle LEN = UBTranRecLen
  CHandle = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CHandle LEN = UBCustRecLen
  
  GET PHandle, TransRec&, UBTranRec(1)
  GET CHandle, UBTranRec(1).CustAcctNo, UBCustRec(1)
  CLOSE
  IF UBTranRec(1).ActiveFlag = 0 THEN
    DelFlag = True
  END IF
  LSET Form$(1, 0) = FUsing$(STR$(UBTranRec(1).CustAcctNo), "#####")
  LSET Form$(2, 0) = UBCustRec(1).book + "-" + UBCustRec(1).SeqNumb
  LSET Form$(3, 0) = UBCustRec(1).CustName
  LSET Form$(4, 0) = FUsing$(STR$(UBTranRec(1).TransAmt), "######.##")
  Action = 1
RETURN
  
SavePenTrans:
  PHandle = FREEFILE
  OPEN PenFile$ FOR RANDOM SHARED AS PHandle LEN = UBTranRecLen
  GET PHandle, TransRec&, UBTranRec(1)
  
  TPenAmt# = Value#(Form$(4, 0), ECode)
  
  IF TPenAmt# = 0 THEN
    UBTranRec(1).ActiveFlag = False
  ELSE
    UBTranRec(1).ActiveFlag = True
  END IF
  IF Round#(TPenAmt#) <> Round#(UBTranRec(1).TransAmt) THEN
    EditLog$ = STR$(UBTranRec(1).CustAcctNo) + "   was " + QPTrim$(FUsing$(STR$(UBTranRec(1).TransAmt), "#####.##"))
    EditLog$ = EditLog$ + " to " + QPTrim$(FUsing$(STR$(TPenAmt#), "#####.##"))
    UBLog " EPF: Changed Acct:" + EditLog$
  END IF
  UBTranRec(1).RevAmt(PenaltyInfo(1).RevSource) = TPenAmt#
  UBTranRec(1).TransAmt = TPenAmt#
  
  PUT PHandle, TransRec&, UBTranRec(1)
  CLOSE
  SaveScrn TempScrn()
  DisplayUBScrn "UPDATEOK"
  WaitForAction
  RestScrn TempScrn()
  FirstTime = True
  EditedFlag = False
RETURN
  
ShowTransRec:
  LSET TrShow$ = QPTrim$(STR$(TransRec&)) + " of " + QPTrim$(STR$(NumPTran&))
  QPrintRC TrShow$, 7, 51, -1
  IF DelFlag THEN
    LSET TrShow$ = "DELETED!"
  ELSE
    LSET TrShow$ = ""
  END IF
  QPrintRC TrShow$, 12, 40, 10
  ShowFlag = False
RETURN
  
AskAbandon:
  SaveFlag = PromptSaveData
  SELECT CASE SaveFlag
  CASE True     'user wants to save
    GOSUB SavePenTrans
    Ok2GoFlag = True
    Action = 1
  CASE False    'user wants to abandon
    Ok2GoFlag = True
    Action = 1
  CASE ELSE     'continue editing
    Ok2GoFlag = False
    Action = 2
  END SELECT
RETURN
  
MarkTransDeleted:
  PHandle = FREEFILE
  OPEN PenFile$ FOR RANDOM SHARED AS PHandle LEN = UBTranRecLen
  GET PHandle, TransRec&, UBTranRec(1)
  UBTranRec(1).ActiveFlag = 0
  UBTranRec(1).RevAmt(PenaltyInfo(1).RevSource) = 0
  UBTranRec(1).TransAmt = 0
  PUT PHandle, TransRec&, UBTranRec(1)
  CLOSE
  SaveScrn TempScrn()
  DisplayUBScrn "UPDATEOK"
  WaitForAction
  RestScrn TempScrn()
  FirstTime = True
  EditedFlag = False
RETURN
  
ExitEdPenalty:
  UBLog "OUT: Edit Penalty File." + CrLf$
END SUB

SUB PenaltyProcess
  
  UBLog " IN: Create Penalty File (CPF)"

  Pct2ChgFld = 5

  SHARED Choice$()
  
  REDIM TempScrn(0)
  REDIM Source$(15)
  
  REDIM PenaltyInfo(1) AS PenaltyInfoType
  REDIM UBCustRec(1) AS NewUBCustRecType
  REDIM UBSetUpRec(1) AS UBSetupRecType
  REDIM UBTranRec(1 TO 3) AS UBTransRecType
  REDIM TaxAmt(1 TO 15) AS DOUBLE
  
  UBCustRecLen = LEN(UBCustRec(1))
  UBTranRecLen = LEN(UBTranRec(1))
  
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen
  TownName$ = UBSetUpRec(1).UTILNAME
  IF INSTR(TownName$, "TENN") > 0 AND INSTR(TownName$, "RIDGE") > 0 THEN
    TennFlag = True
  END IF
  
  IF INSTR(TownName$, "HAMLET") > 0 THEN
    HamFlag = True
  END IF
  
  IF UBSetUpRec(1).BILLCYCL = "Y" THEN
    CycleFlag = True
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
  
  REDIM Choice$(NumOfRevs, 2)
  Choice$(0, 0) = "3"
  FOR TCnt = 1 TO NumOfRevs
    Choice$(TCnt, 0) = Source$(TCnt)
  NEXT TCnt
  
  Choice$(0, 1) = "4"
  Choice$(1, 1) = "Current Balance"
  Choice$(2, 1) = "Previous Balance"
  Choice$(3, 1) = "Total Balance"
  
  Choice$(0, 2) = "7"
  Choice$(1, 2) = "LESS"
  Choice$(2, 2) = "GREATER"
  
  PenFile$ = "UBPENTRN.DAT"
  
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
      LSET Form$(8, 0) = "0"
      IF CycleFlag THEN
        LSET Form$(11, 0) = "0"
        LSET Form$(12, 0) = "0"
        Fld(11).Protected = True
        Fld(12).Protected = True
      ELSE
        LSET Form$(9, 0) = "0"
        LSET Form$(10, 0) = "0"
        Fld(9).Protected = True
        Fld(10).Protected = True
      END IF
    END IF
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      GOSUB CheckPenaltyFlds
      'If valid Data in Fields
      'Then Process the Penalties
    CASE EscKey
      ExitFlag = True
    END SELECT
    
  LOOP UNTIL ExitFlag OR OKFlag
  
  IF ExitFlag THEN
    UBLog " CPF: ABORTED Create Penalty File"
    GOTO ExitPenalty
  END IF
  BlockClear
  ShowProcessingScrn "Calculating Penalty Charges"
  
  KillFile PenFile$
  PHandle = FREEFILE
  OPEN PenFile$ FOR RANDOM SHARED AS PHandle LEN = UBTranRecLen
  
  CHandle = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CHandle LEN = UBCustRecLen
  
  NumCustRecs& = LOF(CHandle) / UBCustRecLen
  
  FOR CCnt& = 1 TO NumCustRecs&
    GET CHandle, CCnt&, UBCustRec(1)

    'IF CCnt& = 2 THEN STOP

    IF NOT UBCustRec(1).DelFlag THEN
      IF UBCustRec(1).LateFee = "Y" THEN
        LSET UBTranRec(1) = UBTranRec(2)        'Transrec 2 is blank
        'Make a clean transaction record
        IF HamFlag THEN
          'IF UBCustRec(1).Status = "I" THEN
          '  GOTO SkipEm
          'END IF
          'IF UBCustRec(1).Status <> "A" THEN
          '  NOPenFlag = False
          '  GOSUB ChkLastTrans
          '  IF NOPenFlag THEN
          '    GOTO SkipEm
          '  END IF
          'END IF
        ELSEIF UBCustRec(1).Status <> "A" THEN      'if they are not inactive
          GOTO SkipEm
        END IF
        '05-01-97 fixed bug where CurrBalance+PrevBalance is <= 0
        IF Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance) > 0 THEN
          IF UBCustRec(1).CurrBalance >= MinBalance# OR UBCustRec(1).PrevBalance > MinBalance# THEN
            'if they have any balance
            IF UsingBook THEN   'if they want it by Book
              ThisBook = QPValI(UBCustRec(1).book)
              IF ThisBook >= FirstBook AND ThisBook <= LastBook THEN
                'if this is in the correct book
                IF UseBothFlag THEN             'both an amount and percent
                  IF UsePrevFlag AND NOT UseCurrFlag THEN       'use prev not curr
                    IF UBCustRec(1).CurrBalance < 0 THEN
                      PenBal# = UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance
                    ELSE
                      PenBal# = UBCustRec(1).PrevBalance
                    END IF
                  ELSEIF UseCurrFlag AND NOT UsePrevFlag THEN   'use curr not prev
                    IF UBCustRec(1).PrevBalance < 0 THEN
                      PenBal# = UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance
                    ELSE
                      PenBal# = UBCustRec(1).CurrBalance
                    END IF
                  ELSEIF UsePrevFlag AND UseCurrFlag THEN       'use curr and prev
                    PenBal# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
                  END IF
                  CustPctPenalty# = Round#(PenBal# * PctAmt#)
                  CustFixPenalty# = FixAmt#
                  IF PenBal# <= MinBalance# THEN                'if cust had prev bal
                    GOTO SkipEm
                  END IF
                  IF GreaterFlag THEN
                    IF CustPctPenalty# >= CustFixPenalty# THEN
                      CustPenalty# = CustPctPenalty#
                    ELSE
                      CustPenalty# = CustFixPenalty#
                    END IF
                  ELSE          'nope want whichever is less
                    IF CustPctPenalty# >= CustFixPenalty# THEN
                      CustPenalty# = CustFixPenalty#
                    ELSE
                      CustPenalty# = CustPctPenalty#
                    END IF
                  END IF
                  GOSUB MakeTransaction
                ELSEIF UsePctFlag THEN          'if they want a percent penalty
                  IF UsePrevFlag AND NOT UseCurrFlag THEN       'using prev not curr
                    '030398 Modified to consider a credit in cur or prev balances
                    IF UBCustRec(1).CurrBalance < 0 THEN
                      PenBal# = UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance
                    ELSE
                      PenBal# = UBCustRec(1).PrevBalance
                    END IF
                    IF PenBal# <= MinBalance# THEN    'if cust had prev bal
                      GOTO SkipEm
                    END IF
                    CustPenalty# = Round#(PenBal# * PctAmt#)
                    GOSUB MakeTransaction
                  ELSEIF UseCurrFlag AND NOT UsePrevFlag THEN   'using curr not prev
                    '030398 Modified to consider a credit in cur or prev balances
                    IF UBCustRec(1).PrevBalance < 0 THEN
                      PenBal# = UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance
                    ELSE
                      PenBal# = UBCustRec(1).CurrBalance
                    END IF
                    'code added to exclude tax
                    '092898 Said they didn't take partial payments - Not!
                    IF TennFlag THEN            'AND UBCustRec(1).TaxExpt <> "Y" THEN
                      GOSUB GetTennRidgeLastBill
                    END IF
                    IF PenBal# <= MinBalance# THEN
                      GOTO SkipEm
                    END IF
                    CustPenalty# = Round#(PenBal# * PctAmt#)
                    GOSUB MakeTransaction
                  ELSEIF UsePrevFlag AND UseCurrFlag THEN       'use curr and prev
                    PenBal# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
                    IF PenBal# <= MinBalance# THEN
                      GOTO SkipEm
                    END IF
                    CustPenalty# = Round#(PenBal# * PctAmt#)
                    GOSUB MakeTransaction
                  END IF
                ELSE            'Using a FIXED penalty amount
                  IF UsePrevFlag AND NOT UseCurrFlag THEN
                    '030398 Modified to consider a credit in cur or prev balances
                    IF UBCustRec(1).CurrBalance < 0 THEN
                      PenBal# = UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance
                    ELSE
                      PenBal# = UBCustRec(1).PrevBalance
                    END IF
                  ELSEIF UseCurrFlag AND NOT UsePrevFlag THEN
                    '030398 Modified to consider a credit in cur or prev balances
                    IF UBCustRec(1).PrevBalance < 0 THEN
                      PenBal# = UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance
                    ELSE
                      PenBal# = UBCustRec(1).CurrBalance
                    END IF
                  ELSEIF UsePrevFlag AND UseCurrFlag THEN
                    'do not need to check for prev >0 or curr>0 here!!
                    PenBal# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
                  END IF
                  IF PenBal# <= MinBalance# THEN
                    GOTO SkipEm
                  END IF
                  CustPenalty# = FixAmt#
                  GOSUB MakeTransaction
                END IF
              END IF
            ELSEIF UsingCycle THEN              'they are using cycles
              ThisCycle = UBCustRec(1).BILLCYCL
              IF ThisCycle >= FirstCycle AND ThisCycle <= LastCycle THEN
                IF UseBothFlag THEN             'both an amount and percent
                  IF UsePrevFlag AND NOT UseCurrFlag THEN       'use prev not curr
                    IF UBCustRec(1).CurrBalance < 0 THEN
                      PenBal# = UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance
                    ELSE
                      PenBal# = UBCustRec(1).PrevBalance
                    END IF
                  ELSEIF UseCurrFlag AND NOT UsePrevFlag THEN   'use curr not prev
                    IF UBCustRec(1).PrevBalance < 0 THEN
                      PenBal# = UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance
                    ELSE
                      PenBal# = UBCustRec(1).CurrBalance
                    END IF
                  ELSEIF UsePrevFlag AND UseCurrFlag THEN       'use curr and prev
                    PenBal# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
                  END IF
                  IF PenBal# <= MinBalance# THEN                'if cust had prev bal
                    GOTO SkipEm
                  END IF
                  CustPctPenalty# = Round#(PenBal# * PctAmt#)
                  CustFixPenalty# = FixAmt#
                  IF GreaterFlag THEN
                    IF CustPctPenalty# >= CustFixPenalty# THEN
                      CustPenalty# = CustPctPenalty#
                    ELSE
                      CustPenalty# = CustFixPenalty#
                    END IF
                  ELSE          'nope want whichever is less
                    IF CustPctPenalty# >= CustFixPenalty# THEN
                      CustPenalty# = CustFixPenalty#
                    ELSE
                      CustPenalty# = CustPctPenalty#
                    END IF
                  END IF
                  GOSUB MakeTransaction

                ELSEIF UsePctFlag THEN   '*** Percentage method
                  IF UsePrevFlag AND NOT UseCurrFlag THEN       'use prev not curr
                    '030398 Modified to consider a credit in cur or prev balances
                    IF UBCustRec(1).CurrBalance < 0 THEN
                      PenBal# = UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance
                    ELSE
                      PenBal# = UBCustRec(1).PrevBalance
                    END IF
                    IF PenBal# <= MinBalance# THEN              'if cust had prev bal
                      GOTO SkipEm
                    END IF
                    CustPenalty# = Round#(PenBal# * PctAmt#)
                    GOSUB MakeTransaction
                  ELSEIF UseCurrFlag AND NOT UsePrevFlag THEN   'use curr not prev
                    '030398 Modified to consider a credit in cur or prev balances
                    IF UBCustRec(1).PrevBalance < 0 THEN
                      PenBal# = UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance
                    ELSE
                      PenBal# = UBCustRec(1).CurrBalance
                    END IF
                    IF PenBal# <= MinBalance# THEN
                      GOTO SkipEm
                    END IF
                    CustPenalty# = Round#(PenBal# * PctAmt#)
                    GOSUB MakeTransaction
                  ELSEIF UsePrevFlag AND UseCurrFlag THEN       'use curr and prev
                    TotalBalance# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
                    IF TotalBalance# <= MinBalance# THEN
                      GOTO SkipEm
                    END IF
                    CustPenalty# = Round#(TotalBalance# * PctAmt#)
                    GOSUB MakeTransaction
                  END IF
                ELSE            'Using a FIXED penalty amount
                  IF UsePrevFlag AND NOT UseCurrFlag THEN
                    '030398 Modified to consider a credit in cur or prev balances
                    IF UBCustRec(1).CurrBalance < 0 THEN
                      PenBal# = UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance
                    ELSE
                      PenBal# = UBCustRec(1).PrevBalance
                    END IF
                    IF PenBal# <= MinBalance# THEN              'if cust had prev bal
                      GOTO SkipEm
                    END IF
                    CustPenalty# = FixAmt#
                    GOSUB MakeTransaction
                  ELSEIF UseCurrFlag AND NOT UsePrevFlag THEN
                    '030398 Modified to consider a credit in cur or prev balances
                    IF UBCustRec(1).PrevBalance < 0 THEN
                      PenBal# = UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance
                    ELSE
                      PenBal# = UBCustRec(1).CurrBalance
                    END IF
                    IF PenBal# <= MinBalance# THEN
                      GOTO SkipEm
                    END IF
                    CustPenalty# = FixAmt#
                    GOSUB MakeTransaction
                  ELSEIF UsePrevFlag AND UseCurrFlag THEN
                    CustPenalty# = FixAmt#
                    TotalBalance# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
                    IF TotalBalance# > MinBalance# THEN
                      GOSUB MakeTransaction
                    END IF
                  END IF        'if current, previous or both
                END IF          'if fixed amount or percent
              END IF            'if this cycle
            END IF              'if using cycle
          END IF                'if balance >= minbalance
        END IF                  'if balance >0
      END IF                    'if allow late fee
    END IF                      'if not deleted
SkipEm:
    ShowPctComp CCnt&, NumCustRecs&
  NEXT
  
  CLOSE
  
  PenaltyInfo(1).PenCnt = PCnt
  FPutAH "UBPENINF.DAT", PenaltyInfo(1), LEN(PenaltyInfo(1)), 1
  
  BlockClear
  DisplayUBScrn "UPDATEOK"
  WaitForAction

ExitPenalty:
  
  ERASE TempScrn, Source$
  ERASE UBCustRec, UBSetUpRec, UBTranRec
  ERASE Frm, Form$, Fld
  
  IF NOT ExitFlag THEN
    UBLog " CPF: Created" + STR$(PenaltyInfo(1).PenCnt) + " work transactions."
  END IF
  UBLog "OUT: Create Penalty File." + CrLf$
  
  EXIT SUB
  
MakeTransaction:
  '011499 Corrected to check for a penalty amount of less than .01
  IF Round#(CustPenalty#) > 0 THEN
    PCnt = PCnt + 1
    UBTranRec(1).TransAmt = CustPenalty#
    UBTranRec(1).RevAmt(RevSource) = CustPenalty#
    UBTranRec(1).TransDate = PenaltyDate
    UBTranRec(1).TransType = TranPenaltyCharge
    UBTranRec(1).TransDesc = TransDesc$
    UBTranRec(1).CustAcctNo = CCnt&
    'UBTranRec(1).RunBalance = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
    UBTranRec(1).ActiveFlag = True
    PUT PHandle, PCnt, UBTranRec(1)
  END IF
  
  RETURN
  
CheckPenaltyFlds:

  CursorOff
  SaveScrn TempScrn()
  IF Value#(Form$(1, 0), ECode) = 0 THEN
    DisplayUBScrn "ERRSCRN1"
    QPrintRC "You Have Not Entered a Proper", 10, 26, -1
    QPrintRC "Transaction Date!", 12, 32, -1
    Frm(1).FldNo = 1
  ELSEIF LEN(QPTrim$(Form$(3, 0))) = 0 THEN
    DisplayUBScrn "ERRSCRN1"
    QPrintRC "No Revenue Source for applied penalty.", 10, 22, -1
    QPrintRC "Correct and try again.", 12, 29, -1
    Frm(1).FldNo = 3
  ELSEIF LEN(QPTrim$(Form$(4, 0))) = 0 THEN
    DisplayUBScrn "ERRSCRN1"
    QPrintRC "Must identify Balance source to apply.", 10, 22, -1
    QPrintRC "penalty on.", 11, 22, -1
    Frm(1).FldNo = 4

  ELSEIF Value#(Form$(Pct2ChgFld, 0), ECode) <> 0 AND Value#(Form$(6, 0), ECode) <> 0 AND LEN(QPTrim$(Form$(7, 0))) = 0 THEN
    DisplayUBScrn "ERRSCRN1"
    QPrintRC "MAY NOT have Both a percentage", 10, 25, -1
    QPrintRC "and a fixed amount WITHOUT the", 11, 25, -1
    QPrintRC "'Whichever is:' parameter!", 12, 27, -1
    Frm(1).FldNo = 7

  ELSEIF LEN(QPTrim$(Form$(7, 0))) > 0 AND (Value#(Form$(Pct2ChgFld, 0), ECode) = 0 OR Value#(Form$(6, 0), ECode) = 0) THEN
    DisplayUBScrn "ERRSCRN1"
    QPrintRC "Invalid percentage, or Fixed amount", 10, 23, -1
    QPrintRC "with the 'Whichever is:' parameter!", 12, 23, -1
    Frm(1).FldNo = 5
  ELSEIF Value#(Form$(Pct2ChgFld, 0), ECode) = 0 AND Value#(Form$(6, 0), ECode) = 0 THEN
    DisplayUBScrn "ERRSCRN1"
    QPrintRC "No Penalty Would Calculate Because", 10, 24, -1
    QPrintRC "BOTH the Percentage and the Amount are", 11, 22, -1
    QPrintRC "SET TO ZERO!", 13, 35, -1
    Frm(1).FldNo = 5
  ELSEIF Value#(Form$(9, 0), ECode) = 0 AND Value#(Form$(10, 0), ECode) = 0 AND Value#(Form$(11, 0), ECode) = 0 AND VAL(Form$(12, 0)) = 0 OR VAL(Form$(10, 0)) < VAL(Form$(9, 0)) OR VAL(Form$(12, 0)) < VAL(Form$(11, 0)) THEN
    DisplayUBScrn "ERRSCRN1"
    QPrintRC "Invalid Book or Cycle range!", 10, 26, -1
    QPrintRC "Correct and try again.", 12, 29, -1
    IF CycleFlag THEN
      Frm(1).FldNo = 9
    ELSE
      Frm(1).FldNo = 11
    END IF
  ELSE
    OKFlag = True
  END IF
  IF OKFlag THEN
    PenaltyDate = Date2Num%(Form$(1, 0))
    MinBalance# = Value#(Form$(8, 0), ECode)
    IF MinBalance# < 0 THEN
      MinBalance# = 0
    END IF
    PenaltyInfo(1).PenDate = PenaltyDate
    PenaltyInfo(1).MinBalance = MinBalance#
    SELECT CASE LEFT$(QPTrim$(Form$(4, 0)), 1)
    CASE "C"    'Applying to Current
      UsePrevFlag = False
      UseCurrFlag = True
    CASE "P"    'Applying to Previous
      UsePrevFlag = True
      UseCurrFlag = False
    CASE "T"    'Applying to Both
      UsePrevFlag = True
      UseCurrFlag = True
    END SELECT
    PenaltyInfo(1).ChargeOn = QPTrim$(Form$(4, 0))
    'Get percent or fixed amount
    PctAmt# = Value#(Form$(Pct2ChgFld, 0), ECode)
    FixAmt# = Value#(Form$(6, 0), ECode)
    PenaltyInfo(1).PctCharge = PctAmt#
    PenaltyInfo(1).AmtCharge = FixAmt#
    
    IF LEN(QPTrim$(Form$(7, 0))) > 0 THEN
      PenaltyInfo(1).GreatLess = LEFT$(QPTrim$(Form$(7, 0)), 1)
      IF PenaltyInfo(1).GreatLess = "G" THEN
        GreaterFlag = True
      ELSE
        GreaterFlag = False
      END IF
      PctAmt# = PctAmt# / 100
      UseBothFlag = True
    ELSE
      IF PctAmt# > 0 THEN
        PctAmt# = PctAmt# / 100
        FixAmt# = 0
        UsePctFlag = True
      ELSE
        PctAmt# = 0
        UsePctFlag = False
      END IF
      UseBothFlag = False
    END IF
    
    'Get the Rev source number
    FOR Cnt = 1 TO NumOfRevs
      ThisRev$ = QPTrim$(Form$(3, 0))
      IF INSTR(Choice$(Cnt, 0), ThisRev$) THEN
        RevSource = Cnt
        EXIT FOR
      END IF
    NEXT
    PenaltyInfo(1).RevSource = RevSource
    'Get Who/How to process
    '***********************
    IF Value#(Form$(9, 0), ECode) > 0 OR Value#(Form$(10, 0), ECode) > 0 THEN
      FirstCycle = Value#(Form$(9, 0), ECode)
      LastCycle = Value#(Form$(10, 0), ECode)
      UsingCycle = True
    ELSE
      FirstBook = Value#(Form$(11, 0), ECode)
      LastBook = Value#(Form$(12, 0), ECode)
      UsingBook = True
    END IF
    TransDesc$ = QPTrim$(Form$(2, 0))
    PenaltyInfo(1).PenDesc = TransDesc$
    PenaltyInfo(1).CycFirst = FirstCycle
    PenaltyInfo(1).CycLast = LastCycle
    PenaltyInfo(1).BookFirst = FirstBook
    PenaltyInfo(1).BookLast = LastBook
  ELSE          'there is an error. Scrn already displayed, wait for input
    WaitForAction
  END IF
  
  Action = 1
  RestScrn TempScrn()
  
  RETURN
  
GetTennRidgeLastBill:
  
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

ChkLastTrans:
  'IF UBCustRec(1).Status = "B" THEN

  FOpenS "UBTRANS.DAT", THandle
  PrevTranRec& = UBCustRec(1).LastTrans
  IF PrevTranRec& > 0 THEN
    FGetRTA THandle, UBTranRec(3), PrevTranRec&, UBTranRecLen
    IF UBTranRec(3).TransType = TranPenaltyCharge THEN
      NOPenFlag = True
    END IF
  END IF
  FClose THandle
RETURN

END SUB

SUB PenaltyReport
  
  PenFile$ = "UBPENTRN.DAT"
  
  FF$ = CHR$(12)
  REDIM Source$(15)
  
  Fmt$ = STRING$(80, "-")
  MaxLines = 55
  
  REDIM PenaltyInfo(1) AS PenaltyInfoType
  
  IF NOT Exist%("UBPENINF.DAT") THEN
    CursorOff
    BlockClear
    DisplayUBScrn "ERRSCRN1"
    QPrintRC "NO UNPOSTED PENALTY TRANSACTIONS!", 10, 24, -1
    QPrintRC "CAN NOT PRINT PENALTY REPORT", 12, 26, -1
    WaitForAction
    GOTO ExitPenReport
  END IF
  
  FGetAH "UBPENINF.DAT", PenaltyInfo(1), LEN(PenaltyInfo(1)), 1
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  REDIM UBSetUpRec(1) AS UBSetupRecType
  REDIM UBTranRec(1) AS UBTransRecType
  
  UBCustRecLen = LEN(UBCustRec(1))
  UBTranRecLen = LEN(UBTranRec(1))
  
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen
  TownName$ = QPTrim$(UBSetUpRec(1).UTILNAME)
  
  FOR Cnt = 1 TO 15
    Source$(Cnt) = UBSetUpRec(1).Revenues(Cnt).REVNAME
  NEXT
  
  BlockClear
  ShowProcessingScrn "Penalty Transaction Report"
  
  PHandle = FREEFILE
  OPEN PenFile$ FOR RANDOM SHARED AS PHandle LEN = UBTranRecLen
  
  CHandle = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CHandle LEN = UBCustRecLen
  
  UBRpt = FREEFILE
  OPEN "UBPENTRN.RPT" FOR OUTPUT AS UBRpt
  
  NumPenRec& = LOF(PHandle) / UBTranRecLen
  
  GOSUB PenRptHeader
  
  FOR Cnt& = 1 TO NumPenRec&
    GET PHandle, Cnt&, UBTranRec(1)
    'IF UBTranRec(1).ActiveFlag <> 0 THEN
    PCnt& = PCnt& + 1
    GET CHandle, UBTranRec(1).CustAcctNo, UBCustRec(1)
    PRINT #UBRpt, FUsing$(STR$(UBTranRec(1).CustAcctNo), "######");
    PRINT #UBRpt, TAB(10); UBCustRec(1).Status; TAB(15); UBCustRec(1).book; "-"; UBCustRec(1).SeqNumb;
    PRINT #UBRpt, TAB(30); UBCustRec(1).CustName;
    IF UBTranRec(1).TransAmt = 0 THEN
      PRINT #UBRpt, TAB(65); " DELETED"
      DCnt& = DCnt& + 1
    ELSE
      PRINT #UBRpt, TAB(65); FUsing$(STR$(UBTranRec(1).TransAmt), "#####.##")
    END IF
    PenTotal# = Round#(PenTotal# + UBTranRec(1).TransAmt)
    LineCnt = LineCnt + 1
    'END IF
    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, FF$
      GOSUB PenRptHeader
    END IF
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
    ShowPctComp Cnt&, NumPenRec&
  NEXT
  
  GOSUB PenRptFooter
  PRINT #UBRpt, FF$
  GOSUB PenRptParms
  PRINT #UBRpt, FF$
  CLOSE
  
  IF NOT AbortFlag THEN
    PrintRptFile "Penalty Transaction Report", "UBPENTRN.RPT", LPTPort, RetCode, EntryPoint
  END IF
  
  EXIT SUB
  
PenRptHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, "Penalty Transaction Report.  "; TAB(70); "Page: "; PageNo
  PRINT #UBRpt, TownName$
  PRINT #UBRpt, " Acct    Location       Customer                             Amount"
  PRINT #UBRpt, Fmt$
  LineCnt = 4
  RETURN
  
PenRptFooter:
  PRINT #UBRpt, Fmt$
  PRINT #UBRpt, " Transactions:"; PCnt&; TAB(22); "Deleted:"; DCnt&; TAB(42); USING "Penalty Total: $$######.##"; PenTotal#
  LineCnt = 4
  RETURN
  
PenRptParms:
  PageNo = PageNo + 1
  PRINT #UBRpt, "Penalty Calculation Parameters.  "; TAB(70); "Page: "; PageNo
  PRINT #UBRpt, TownName$
  PRINT #UBRpt, Fmt$
  PRINT #UBRpt, "    Penalty Date: "; Num2Date$(PenaltyInfo(1).PenDate)
  PRINT #UBRpt, "    Balance Type: "; PenaltyInfo(1).ChargeOn
  PRINT #UBRpt, "  Revenue Source: "; Source$(PenaltyInfo(1).RevSource)
  PRINT #UBRpt, " Minimum Balance: "; USING "####.##"; PenaltyInfo(1).MinBalance
  IF LEN(QPTrim$(PenaltyInfo(1).GreatLess)) > 0 THEN
    PRINT #UBRpt, " Penalty Percent: "; USING "######"; PenaltyInfo(1).PctCharge
    PRINT #UBRpt, "     Flat Amount: "; USING "####.##"; PenaltyInfo(1).AmtCharge
    PRINT #UBRpt, "    Whichever is: ";
    SELECT CASE PenaltyInfo(1).GreatLess
    CASE "L"
      PRINT #UBRpt, "LESS"
    CASE "G"
      PRINT #UBRpt, "GREATER"
    END SELECT
  ELSE
    IF PenaltyInfo(1).PctCharge > 0 THEN
      PRINT #UBRpt, " Penalty Percent: "; USING "######"; PenaltyInfo(1).PctCharge
    ELSEIF PenaltyInfo(1).AmtCharge > 0 THEN
      PRINT #UBRpt, "     Flat Amount: "; USING "####.##"; PenaltyInfo(1).AmtCharge
    END IF
  END IF
  IF PenaltyInfo(1).CycLast > 0 THEN
    PRINT #UBRpt, "      From Cycle: "; USING "######"; PenaltyInfo(1).CycFirst
    PRINT #UBRpt, "      Thru Cycle: "; USING "######"; PenaltyInfo(1).CycLast
  ELSEIF PenaltyInfo(1).BookLast > 0 THEN
    PRINT #UBRpt, "       From Book: "; USING "######"; PenaltyInfo(1).BookFirst
    PRINT #UBRpt, "       Thru Book: "; USING "######"; PenaltyInfo(1).BookLast
  END IF
  PRINT #UBRpt, "     Description: "; PenaltyInfo(1).PenDesc
  'PenDesc    AS STRING * 21
  'CycFirst   AS INTEGER
  'CycLast    AS INTEGER
  'BookFirst  AS INTEGER
  'BookLast   AS INTEGER
  'PenCnt     AS INTEGER
  
  RETURN
  
ExitPenReport:
END SUB

SUB PostPenalties
  '01-15-99 Added penalty processing as seperate parts
  
  UBLog "IN: Post Penalty Transactions (PPT)"
  
  PenFile$ = "UBPENTRN.DAT"
  
  IF NOT Exist%("UBPENINF.DAT") THEN
    CursorOff
    BlockClear
    DisplayUBScrn "ERRSCRN1"
    QPrintRC "NO UNPOSTED PENALTY TRANSACTIONS!", 10, 24, -1
    QPrintRC "NOTHING TO POST.", 12, 32, -1
    UBLog "ERROR: NO UNPOSTED PENALTY TRANSACTIONS."
    WaitForAction
    GOTO ExitPenPost:
  END IF
  
  REDIM Source$(15)
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen
  FOR Cnt = 1 TO 15
    Source$(Cnt) = UBSetUpRec(1).Revenues(Cnt).REVNAME
  NEXT
  
  LibName$ = "UB"
  ScrnName$ = "PSTBILLS"
  
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
  Action = 1
  Frm(1).StayOnField = True
  
  '--Set screen number to one and display screen
  
  BlockClear
  DisplayUBScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      OKFlag = True
    CASE EscKey
      ExitFlag = True
    END SELECT
    
  LOOP UNTIL OKFlag OR ExitFlag
  
  IF ExitFlag THEN
    BlockClear
    CursorOff
    DisplayUBScrn "POSTCAN"
    WaitForAction
    UBLog "ABORTED:  Penalty Posting."
    GOTO ExitPenPost
  END IF
  
  UBLog "START: Posting Penalty Transactions."
  
  REDIM PenaltyInfo(1) AS PenaltyInfoType
  FGetAH "UBPENINF.DAT", PenaltyInfo(1), LEN(PenaltyInfo(1)), 1
  
  REDIM UBTranRec(1) AS UBTransRecType
  REDIM UBCustRec(1) AS NewUBCustRecType
  
  UBTranRecLen = LEN(UBTranRec(1))
  UBCustRecLen = LEN(UBCustRec(1))
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  PHandle = FREEFILE
  OPEN PenFile$ FOR RANDOM SHARED AS PHandle LEN = UBTranRecLen
  
  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen
  
  NumOfTranRecs& = LOF(UBTran) \ UBTranRecLen
  NumPenRec& = LOF(PHandle) \ UBTranRecLen
  
  ShowProcessingScrn "Posting Penalty Transactions"
  FOR PenCnt& = 1 TO NumPenRec&
    GET PHandle, PenCnt&, UBTranRec(1)
    IF (UBTranRec(1).ActiveFlag AND UBTranRec(1).TransAmt > 0) THEN
      PostedCnt& = PostedCnt& + 1
      NumOfTranRecs& = NumOfTranRecs& + 1       'point to next trans to write
      GET UBCust, UBTranRec(1).CustAcctNo, UBCustRec(1)
      IF UBCustRec(1).Status = "B" THEN
        UBCustRec(1).LateFee = "N"
      END IF
      '020399 Changed this back to the way it used to be
      UBCustRec(1).CurrBalance = Round#(UBCustRec(1).CurrBalance + UBTranRec(1).TransAmt)
      'UBCustRec(1).PrevBalance = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
      UBTranRec(1).RunBalance = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
      FOR RevCnt = 1 TO MaxRevsCnt
        '        UBCustRec(1).PrevRevAmts(RevCnt) = Round#(UBCustRec(1).CurrRevAmts(RevCnt) + UBCustRec(1).PrevRevAmts(RevCnt))
        UBCustRec(1).CurrRevAmts(RevCnt) = Round#(UBCustRec(1).CurrRevAmts(RevCnt) + UBTranRec(1).RevAmt(RevCnt))
      NEXT
      UBTranRec(1).TransType = TranPenaltyCharge
      UBTranRec(1).TransDesc = PenaltyInfo(1).PenDesc
      UBTranRec(1).TransDate = PenaltyInfo(1).PenDate
      PrevLastTrans& = UBCustRec(1).LastTrans
      UBTranRec(1).PrevTrans = PrevLastTrans&
      UBCustRec(1).LastTrans = NumOfTranRecs&
      PUT UBCust, UBTranRec(1).CustAcctNo, UBCustRec(1)
      PUT UBTran, NumOfTranRecs&, UBTranRec(1)
    END IF
    ShowPctComp PenCnt&, NumPenRec&
  NEXT
  CLOSE
  UBLog "  DONE: Posting Penalty Transactions."
  UBLog "POSTED:" + STR$(PostedCnt&) + " New Penalty Transactions."
  UBLog " Parameters:"
  UBRpt = FREEFILE
  OPEN "UBLOG.DAT" FOR APPEND SHARED AS UBRpt
  PRINT #UBRpt, "   Penalty Date: "; Num2Date$(PenaltyInfo(1).PenDate)
  PRINT #UBRpt, "   Balance Type: "; PenaltyInfo(1).ChargeOn
  PRINT #UBRpt, " Revenue Source: "; Source$(PenaltyInfo(1).RevSource)
  PRINT #UBRpt, "Minimum Balance: "; USING "####.##"; PenaltyInfo(1).MinBalance
  IF PenaltyInfo(1).PctCharge > 0 THEN
    PRINT #UBRpt, "Penalty Percent: "; USING "######"; PenaltyInfo(1).PctCharge
  ELSEIF PenaltyInfo(1).AmtCharge > 0 THEN
    PRINT #UBRpt, "    Flat Amount: "; USING "####.##"; PenaltyInfo(1).AmtCharge
  END IF
  IF PenaltyInfo(1).CycLast > 0 THEN
    PRINT #UBRpt, "     From Cycle: "; USING "######"; PenaltyInfo(1).CycFirst
    PRINT #UBRpt, "     Thru Cycle: "; USING "######"; PenaltyInfo(1).CycLast
  ELSEIF PenaltyInfo(1).BookLast > 0 THEN
    PRINT #UBRpt, "      From Book: "; USING "######"; PenaltyInfo(1).BookFirst
    PRINT #UBRpt, "      Thru Book: "; USING "######"; PenaltyInfo(1).BookLast
  END IF
  CLOSE
  
  KillFile "UBPENINF.DAT"
  KillFile PenFile$
  
  UBLog "KILLED: UBPENINF.DAT  & " + PenFile$
  BlockClear
  DisplayUBScrn "UPDATEOK"
  WaitForAction
  
ExitPenPost:
  UBLog "OUT: Penalty Posting." + CrLf$
  
END SUB

