DEFINT A-Z
DECLARE SUB EnterPPAdjustments ()
DECLARE SUB EnterREAdjustments ()
DECLARE SUB ShowCustHistory (CustRec&, TaxType%)
DECLARE FUNCTION ParseBillNum$ (Text$)
DECLARE SUB EnterAdjustments ()
DECLARE SUB DelList (Items() AS ANY, Picked%(), NPicked%)
DECLARE SUB DeletePayments ()
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB DeleteT (SEG Element, ElSize, NumEls)
DECLARE SUB PostTaxPayments ()
DECLARE SUB PrintPaymentJournal ()
DECLARE FUNCTION AskSavePayment% ()
DECLARE SUB WPickList (Items() AS ANY, Picked%(), NPicked%, DspRow%, UpDateFlag%, Cnf AS ANY)
DECLARE SUB DisplayAPScrn (ScrnName$)
DECLARE SUB AddEditOperator ()
DECLARE SUB AuthorizeOperator (oktoadd%)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB BlockClear ()
DECLARE SUB ClearBack ()
DECLARE SUB ClearScrn ()
DECLARE SUB CursorOff ()
DECLARE SUB DisplayOperEntry ()
DECLARE SUB DisplayTaxScrn (ScrnName$)
DECLARE SUB EnterPayments ()
DECLARE SUB FClose (Handle%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB HideCursor ()
DECLARE SUB KillFile (FileName$)
DECLARE SUB LookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, ActiveOnly%, TaxType%)
DECLARE SUB OperatorEntry ()
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB ButtonPress (WhatBtn, N, MooseButton, MCol, MRow)   ' ----- Check Mouse activity
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPORT%, RetCode%, EntryPoint%)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB RestScrn (Array%())
DECLARE SUB SaveScrn (Array%())
DECLARE SUB ShowCursor ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB StuffBuf (Ky$)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB ValidateOperator ()
DECLARE SUB WaitForAction ()
DECLARE SUB MPaintBox (UlRow%, UlCol%, LRRow%, LRCol%, Colr%)
DECLARE FUNCTION Unique$ (Path$)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION PromptSaveData% ()
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValL& (Number$)
DECLARE FUNCTION WEnvTest ()
DECLARE FUNCTION Round# (DblNum#)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION IsCustDeleted% (AcctNum&, TaxType%)
DECLARE FUNCTION DoesCustOwe% (TaxCustRec AS ANY)
DECLARE FUNCTION Exist% (FileName$)
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)

  TYPE FLen2
    V AS STRING * 64
  END TYPE
  
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TaxCONST.BI'
  ''$INCLUDE: 'taxajust.BI'

  CONST False = 0, True = NOT False

  STACK 5000
  
  CrLf$ = CHR$(13) + CHR$(10)
  
  ClearBack
  ShowCursor

  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 3)

  MChoice$(1) = "Real Estate Tax Abatement"
  MChoice$(2) = "Personal Property Tax Abatement    "
  MChoice$(3) = "Exit to OS"

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
  Col = ((80 - MaxLen) \ 2) - 1

  DO

    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0

    ClearBack

    TitleBox 2, Col, MaxLen + 3, "Virginia Tax Billing Abatement Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

    ShowCursor

    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf

    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0

    SELECT CASE Choice
    CASE 1
     EnterREAdjustments
    
    CASE 2
     EnterPPAdjustments
    CASE 3
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP

  IF WEnvTest THEN
    Ext$ = ".bas"
  ELSE
    Ext$ = ".exe"
  END IF
  IF Exist("cpmenu" + Ext$) THEN
    RUN "cpmenu"
  'ELSE
  '  HideCursor
  '  ClearScrn
  END IF


  
  IF INSTR(COMMAND$, "TEST") THEN
    HideCursor
    CLS
    END
  END IF

  RUN "TaxMenu"

  END

FUNCTION AskSavePayment%
  
  REDIM TempScrn(0)
  SaveScrn TempScrn()
  
  LibName$ = "TAX"
  ScrnName$ = "ASKR2PST"
  
  '--Initialize the form name array
  '--Get the total number of fields from all pages
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  Frm(1).StayOnField = True
  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  DisplayTaxScrn ScrnName$
  QPrintRC SPACE$(30), 11, 22, 68

  ShowCursor
  
  Action = 1
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE EscKey
      AskSavePayment% = False
      EXIT DO
    CASE F10Key
      AskSavePayment% = True
      EXIT DO
    END SELECT
    
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 14
        SELECT CASE Frm(1).MCol
        CASE 27 TO 40           'ESC Cancel button
          PressButton 27, 14, 27, 40
        CASE 41 TO 56           'F10 Save Button
          PressButton -68, 14, 41, 56
        END SELECT
      END SELECT
    END IF
  LOOP
  
  RestScrn TempScrn()
  
END FUNCTION

SUB EnterPPAdjustments
  
  SHARED Choice$()

  BB$ = SPACE$(5)
  u$ = CHR$(24)
  d$ = CHR$(25)
  
  REDIM TempScrn(0)
  REDIM TaxCustRec(1)        AS TaxCustType
  REDIM TaxTrans(1)          AS TaxTransactionType
  REDIM AmtFlds(1 TO 7)      AS INTEGER

  TaxTranRecLen = LEN(TaxTrans(1))
  TaxCustRecLen = LEN(TaxCustRec(1))
  
  LibName$ = "TAX"
  ScrnName$ = "VAPPABAT"
  TaxType% = 2
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
  
  '--define the multi-choice fields
  CustAcctFld = FldNum%("CUSTACCT", Fld())
  BillFld = FldNum%("BILLS", Fld())
  
  Action = 1
  FirstTime = True
  
  DisplayTaxScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF Frm(1).Edited AND Frm(1).PrevFld <> Frm(1).FldNo THEN
      BeenEditedFlag = True     'if the form has been edited
    END IF      'set the edited flag
    
    IF FirstTime THEN
      FirstTime = False         'if this is the first time
      GOSUB SetupFormInfo
    END IF
    
    IF Frm(1).FldNo > CustAcctFld AND Frm(1).PrevFld = CustAcctFld THEN
      CustAcct& = QPValL(Form$(CustAcctFld, 0))
      IF CustAcct& <> LastCust& THEN
        LastCust& = CustAcct&
        GOSUB GetCustInfo
      ELSEIF CustAcct& = 0 THEN
        Frm(1).FldNo = 1
      END IF
      MPaintBox 17, 20, 17, 21, 112
    ELSEIF Frm(1).FldNo = CustAcctFld AND Frm(1).PrevFld <> CustAcctFld THEN
      MPaintBox 17, 20, 17, 21, 126
      MPaintBox 17, 32, 17, 33, 112
    ELSEIF Frm(1).FldNo = BillFld AND Frm(1).PrevFld <> BillFld THEN
      MPaintBox 17, 32, 17, 33, 126
    ELSEIF Frm(1).PrevFld = BillFld AND Frm(1).FldNo <> BillFld THEN
      MPaintBox 17, 32, 17, 33, 112
    END IF

    IF Frm(1).FldNo > CustAcctFld THEN
      IF CustAcct& = 0 THEN
        Frm(1).FldNo = 1
      END IF
    END IF

    IF Frm(1).FldNo > BillFld THEN
      IF TransRecNum& = 0 THEN
        Frm(1).FldNo = BillFld
      END IF
    END IF

    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F4Key
    IF CustAcct& > 0 THEN
      ShowCustHistory -CustAcct&, 2      'DON'T CHANGE THIS
      Action = 1
    END IF
    
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
    CASE F7KEY  'Lookup Customer
      IF Frm(1).FldNo = CustAcctFld THEN  'if user is on the Customer field
        SaveScrn TempScrn()     'and F7key then do lookup routine
        MPaintBox 6, 4, 18, 76, 8
        LastCust& = CustAcct&
        LookUp CustAcct&, "Payment", 0, False, False, TaxType%
        RestScrn TempScrn()
        IF CustAcct& > 0 THEN   'if this is a valid customer
          GOSUB GetCustInfo     'go get customer info
          Frm(1).FldNo = 4
          Action = 1
        ELSEIF LastCust& = CustAcct& THEN   ' don't do anything
          Frm(1).FldNo = 1
          Action = 1
        ELSE
          GOSUB ClearForm
          Frm(1).FldNo = 1
          Action = 1
        END IF
      END IF
    CASE F8KEY           'Select the bills being paid
      IF Frm(1).FldNo = BillFld THEN
        GOSUB SelectBills2Pay
      END IF
    CASE F10Key 'Save
      GOSUB ChkAdjInfo
      IF AdjustOkFlag THEN
        SELECT CASE AskSavePayment
        CASE True               'Save trans no receipt
          GOSUB SaveTransaction
          GOSUB ClearForm
          Frm(1).FldNo = 1
          Action = 1
          EditFlag = False
        CASE False              'oops, just keep editing
          Action = 2
        END SELECT
      END IF
    'CASE IS <> 0
    'STOP
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow   'See if moose was on row 17
      CASE 17                   'it was
        SELECT CASE Frm(1).MCol 'See what column
        CASE 6 TO 17           'f4 History
          PressButton F4Key, 17, 6, 17
        CASE 19 TO 29           'f7 Look-Up
          PressButton F7KEY, 17, 19, 29
        CASE 31 TO 39           'f8 Bill select
          PressButton F8KEY, 17, 31, 39
        CASE 53 TO 62           'f10 Save
          PressButton F10Key, 17, 53, 62
        CASE 64 TO 74           '--cancel button
          PressButton EscKey, 17, 64, 74
        END SELECT
      END SELECT                'row
    END IF
  LOOP UNTIL ExitFlag
  
  ERASE TempScrn
  
  HideCursor
EXIT SUB
  
GetCustInfo:
  GOSUB ClearForm
  NumOfCustRecs& = FileSize("PPTXCUST.DAT") \ TaxCustRecLen
  IF CustAcct& > NumOfCustRecs& OR CustAcct& = 0 THEN
    LastCust& = 0
    CustAcct& = 0
    Ok = MsgBox%("TAX.QSL", "BADACCTN")
    Frm(1).FldNo = 1
    GOTO SkipCustInfo
  ELSEIF IsCustDeleted(CustAcct&, TaxType%) THEN
    LastCust& = 0
    CustAcct& = 0
    Ok = MsgBox%("TAX.QSL", "DELACCTN")
    Frm(1).FldNo = 1
    GOTO SkipCustInfo
  END IF
  
  CustFile = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM SHARED AS CustFile LEN = TaxCustRecLen
  GET CustFile, CustAcct&, TaxCustRec(1)
  CLOSE CustFile
  IF NOT DoesCustOwe%(TaxCustRec(1)) THEN
    SaveScrn TempScrn()
    DisplayTaxScrn "ERRSCRN1"
    QPrintRC "This customer has NO BALANCE!", 10, 26, -1
    QPrintRC "Press any key to continue.", 13, 28, -1
    WaitForAction
    RestScrn TempScrn()
    Frm(1).FldNo = 1
    LastCust& = 0
    CustAcct& = 0
    TransRecNum& = 0
    GOTO SkipCustInfo
  END IF
  LSET Form$(CustAcctFld, 0) = STR$(CustAcct&)
  First$ = QPTrim$(TaxCustRec(1).FNAME)
  IF LEN(First$) > 0 THEN
    CustName$ = First$ + " " + QPTrim$(TaxCustRec(1).LName)
  ELSE
    CustName$ = First$ + " " + QPTrim$(TaxCustRec(1).LName)
  END IF

  LSET Form$(2, 0) = CustName$
  LSET Form$(3, 0) = TaxCustRec(1).Addr1
  CustAcct& = QPValL(Form$(CustAcctFld, 0))
  FirstTime = True
  
SkipCustInfo:
  Action = 1
RETURN
  
SetupFormInfo:
  LSET BB$ = ""
  QPrintRC BB$, 13, 15, -1
  LSET Form$(4, 0) = DATE$
  LSET Form$(14, 0) = "0.00"
  LSET Form$(15, 0) = "0.00"
  LSET Form$(16, 0) = "0.00"
  LSET Form$(17, 0) = "0.00"
  LSET Form$(18, 0) = "0.00"
  LSET Form$(19, 0) = "0.00"
  LSET Form$(20, 0) = "0.00"
  CalcFields 0, 10, Form$(), Fld()
  Action = 2
RETURN
  
ClearForm:
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""       '--Clear all fields
  NEXT
  BeenEditedFlag = False        'clear the edited flag
RETURN
  
SaveTransaction:
  
  SaveScrn TempScrn()
  DisplayTaxScrn "UPDATDSK"
  REDIM TaxAdjTrans(1) AS TaxTransactionType

  CustFile = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM SHARED AS CustFile LEN = TaxCustRecLen

  TranFile = FREEFILE
  OPEN "TAXTRANS.DAT" FOR RANDOM SHARED AS TranFile LEN = TaxTranRecLen

  GET #CustFile, CustAcct&, TaxCustRec(1)      'get cust rec
  GET #TranFile, TransRecNum&, TaxTrans(1)     'get bill we are adjusting

  BillNum$ = ParseBillNum(TaxTrans(1).Description)

  TotalAdj# = Round#(PrinceAdj1# + PrinceAdj2# + PrinceAdj3# + PrinceAdj4# + PrinceAdj5# + InterestAdj# + PenaltyAdj#)
  TaxAdjTrans(1).TransDate = TransDate
  TaxAdjTrans(1).TranType = 3              'Adjustment transaction
  TaxAdjTrans(1).Revenue.Principle1Pd = PrinceAdj1#
  TaxAdjTrans(1).Revenue.Principle2Pd = PrinceAdj2#
  TaxAdjTrans(1).Revenue.Principle3Pd = PrinceAdj3#
  TaxAdjTrans(1).Revenue.Principle4Pd = PrinceAdj4#
  TaxAdjTrans(1).Revenue.Principle5Pd = PrinceAdj5#
  TaxAdjTrans(1).Revenue.InterestPd = InterestAdj#
  TaxAdjTrans(1).Revenue.PenaltyPd = PenaltyAdj#
  TaxAdjTrans(1).Amount = TotalAdj#
  TaxAdjTrans(1).Description = "Tax Abate to Bill #" + BillNum$
  TaxAdjTrans(1).CustomerRec = CustAcct&
  TaxAdjTrans(1).LastTrans = TaxCustRec(1).LastTrans
  TaxAdjTrans(1).BelongTo = TransRecNum&

  TaxTrans(1).Revenue.Principle1Pd = Round#(TaxTrans(1).Revenue.Principle1Pd + PrinceAdj1#)
  TaxTrans(1).Revenue.Principle2Pd = Round#(TaxTrans(1).Revenue.Principle2Pd + PrinceAdj2#)
  TaxTrans(1).Revenue.Principle3Pd = Round#(TaxTrans(1).Revenue.Principle3Pd + PrinceAdj3#)
  TaxTrans(1).Revenue.Principle4Pd = Round#(TaxTrans(1).Revenue.Principle4Pd + PrinceAdj4#)
  TaxTrans(1).Revenue.Principle5Pd = Round#(TaxTrans(1).Revenue.Principle5Pd + PrinceAdj5#)
  TaxTrans(1).Revenue.InterestPd = Round#(TaxTrans(1).Revenue.InterestPd + InterestAdj#)
  TaxTrans(1).Revenue.PenaltyPd = Round#(TaxTrans(1).Revenue.PenaltyPd + PenaltyAdj#)

  PUT #TranFile, TransRecNum&, TaxTrans(1)

  NextTransRec& = (LOF(TranFile) \ TaxTranRecLen) + 1

  TaxCustRec(1).LastTrans = NextTransRec&

  PUT #TranFile, NextTransRec&, TaxAdjTrans(1)
  PUT #CustFile, CustAcct&, TaxCustRec(1)

  CLOSE

SaveExit:
  DisplayTaxScrn "UPDATEOK"
  WaitForAction
  LastCust& = 0
  CustAcct& = 0
  TransRecNum& = 0
  RestScrn TempScrn()
  FirstTime = True
RETURN
  
UPDateListOfPayments:
RETURN

ChkAdjInfo:
  'Parse and move data to Paylist records here
  AdjustOkFlag = True
  ErrCode = 0
  TransDate = Date2Num(Form$(4, 0))
  PrinceAdj1# = Value#(Form$(14, 0), ErrCode%)
  PrinceAdj2# = Value#(Form$(15, 0), ErrCode%)
  PrinceAdj3# = Value#(Form$(16, 0), ErrCode%)
  PrinceAdj4# = Value#(Form$(17, 0), ErrCode%)
  PrinceAdj5# = Value#(Form$(18, 0), ErrCode%)
  PenaltyAdj# = Value#(Form$(19, 0), ErrCode%)
  InterestAdj# = Value#(Form$(20, 0), ErrCode%)
  
  
  IF (PrinceAdj1# > TPrince1#) OR (PrinceAdj2# > TPrince2#) OR (PrinceAdj3# > TPrince3#) OR (PrinceAdj4# > TPrince4#) OR (PrinceAdj5# > TPrince5#) OR (InterestAdj# > TInterest#) OR (CollectAdj# > TCollect#) THEN
    ErrCode = 1
    GOSUB ShowAdjError
    GOTO BadPayment
  ELSEIF (PrinceAdj1# < 0) OR (PrinceAdj2# < 0) OR (PrinceAdj3# < 0) OR (PrinceAdj4# < 0) OR (PrinceAdj5# < 0) OR (InterestAdj# < 0) OR (PenaltyAdj# < 0) THEN
    ErrCode = 2
    GOSUB ShowAdjError
    GOTO BadPayment
  ELSEIF (PrinceAdj1# = 0) AND (PrinceAdj2# = 0) AND (PrinceAdj3# = 0) AND (PrinceAdj4# = 0) AND (PrinceAdj5# = 0) AND (InterestAdj# = 0) AND (PenaltyAdj# = 0) THEN
    ErrCode = 3
    GOSUB ShowAdjError
    GOTO BadPayment
  ELSEIF TransDate <= 0 THEN
    ErrCode = 4
    GOSUB ShowAdjError
    GOTO BadPayment
  END IF
  
BadPayment:
  
RETURN

ShowAdjError:
  SaveScrn TempScrn()
  DisplayTaxScrn "ERRSCRN1"
  SELECT CASE ErrCode
  CASE 1
    QPrintRC "Adjustment greater than Balance.", 10, 25, -1
    Frm(1).FldNo = 10
  CASE 2
    QPrintRC "Adjustment can not be less than zero.", 10, 22, -1
    Frm(1).FldNo = 10
  CASE 3
    QPrintRC "No adjustment amount entered.", 10, 26, -1
    Frm(1).FldNo = 10
  CASE 4
    QPrintRC "Invalid Adjustment Date.", 10, 29, -1
    Frm(1).FldNo = 4
  END SELECT
  QPrintRC "Correct and Save transaction Again.", 12, 23, -1
  WaitForAction
  RestScrn TempScrn()
  Action = 2
  AdjustOkFlag = False
  
RETURN

SelectBills2Pay:
  Title$ = SPACE$(63)
  LSET Title$ = " TR Date          Tax Year     BillNo.        Amt     Balance"

  BillCnt = 0
  REDIM TempScrn(1)
  SaveScrn TempScrn()
  REDIM TaxTrans(1) AS TaxTransactionType
  IF TaxCustRec(1).LastTrans > 0 THEN
    TransFile = FREEFILE
    OPEN "TaxTrans.dat" FOR RANDOM SHARED AS TransFile LEN = TaxTranRecLen
    TransRecord& = TaxCustRec(1).LastTrans
    DO WHILE TransRecord& <> 0
      GET TransFile, TransRecord&, TaxTrans(1)
      IF TaxTrans(1).TranType = 1 THEN     'only bill transactions
'  Balance# = Round#(TaxTrans(1).Revenue.Principle1 + TaxTrans(1).Revenue.Principle2 + TaxTrans(1).Revenue.Principle3 + TaxTrans(1).Revenue.Principle4 + TaxTrans(1).Revenue.Principle5)
'  Balance# = Round#(Balance# + TaxTrans(1).Revenue.Interest + TaxTrans(1).Revenue.Penalty + TaxTrans(1).Revenue.Penalty)
'  Balance# = Round#(Balance# - (TaxTrans(1).Revenue.Principle1Pd + TaxTrans(1).Revenue.Principle2Pd + TaxTrans(1).Revenue.Principle3Pd + TaxTrans(1).Revenue.Principle4Pd + TaxTrans(1).Revenue.Principle5Pd))
'  Balance# = Round#(Balance# - (TaxTrans(1).Revenue.InterestPd + TaxTrans(1).Revenue.PenaltyPd + TaxTrans(1).Revenue.PenaltyPd))
'^^^?????
        Balance# = Round#(TaxTrans(1).Revenue.Principle1 + TaxTrans(1).Revenue.Principle2 + TaxTrans(1).Revenue.Principle3 + TaxTrans(1).Revenue.Principle4 + TaxTrans(1).Revenue.Principle5)
        Balance# = Round#(Balance# + TaxTrans(1).Revenue.Interest + TaxTrans(1).Revenue.Penalty)
        Credit# = Round#(TaxTrans(1).Revenue.Principle1Pd + TaxTrans(1).Revenue.Principle2Pd + TaxTrans(1).Revenue.Principle3Pd + TaxTrans(1).Revenue.Principle4Pd + TaxTrans(1).Revenue.Principle5Pd + TaxTrans(1).Revenue.InterestPd + TaxTrans(1). _
Revenue.PenaltyPd)
        Balance# = Round#(Balance# - Credit#)


        IF Balance# > 0 THEN               'only bills with a balance
          BillCnt = BillCnt + 1
          REDIM PRESERVE Items(1 TO BillCnt) AS FLen2
          LSET Items(BillCnt).V = Num2Date(TaxTrans(1).TransDate)
          MID$(Items(BillCnt).V, 20) = FUsing$(STR$(TaxTrans(1).TaxYear), "####")
          BillNum& = Value(ParseBillNum(TaxTrans(1).Description), ECode)
          MID$(Items(BillCnt).V, 30) = FUsing$(STR$(BillNum&), "#####")
          'MID$(Items(BillCnt).V, 30) = FUsing$(STR$(TransRecord&), "#####")
          MID$(Items(BillCnt).V, 40) = FUsing$(STR$(TaxTrans(1).Amount), "######.##")
          MID$(Items(BillCnt).V, 50) = FUsing$(STR$(Balance#), "######.##")
          MID$(Items(BillCnt).V, 61) = MKL$(TransRecord&)
        END IF
      END IF
      TransRecord& = TaxTrans(1).LastTrans
    LOOP
    CLOSE TransFile
  ELSE
    Ok = MsgBox%("Tax.QSL", "NOCTRANS")
    RestScrn TempScrn()
  END IF
  
  MPaintBox 6, 4, 18, 76, 8     'to add shadow to list
  QPrintRC Title$, 3, 9, 112
  MaxLen = 59 'Set menu width to zero
  Action = 0  '0 means stay in the menu until they select something
  IF Choice < 1 THEN
    Choice = 1                'Pre-load choice to highlight
  END IF
  Col = ((80 - 60) \ 2) - 1
  Row = 4
  BoxBot = 15 'limit the box length to go no lower than line 20
  WazzWind BoxBot + 2, Col, BoxBot + 4, MaxLen + 3 + Col, 10, 4, True
  QPrintRC "   Use: " + u$ + "-" + d$ + " to select.", BoxBot + 3, Col + 3, 15
  QPrintRC u$, BoxBot + 3, Col + 11, 14
  QPrintRC d$, BoxBot + 3, Col + 13, 14
  QPrintRC "Press: [ESC] to Cancel.", BoxBot + 3, Col + 33, 15
  QPrintRC "ESC", BoxBot + 3, Col + 41, 14

  DO
    LOCATE Row, Col, 0
    VertMenuT2 Items(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    IF Ky$ = CHR$(27) THEN
      Choice = 0
      EXIT DO 'choice = 0
    ELSEIF Ky$ = CHR$(13) THEN
      GOSUB ParseTransInfo
      Frm(1).FldNo = Frm(1).FldNo + 1
      'Action = 2
      EXIT DO
    END IF
  LOOP        'UNTIL EditLocRec& > 0
  RestScrn TempScrn()

  IF Choice > 0 THEN
    LSET BB$ = FUsing$(STR$(BillNum&), "#####")
  ELSE
    LSET BB$ = ""
  END IF

  QPrintRC BB$, 13, 15, -1

  Action = 1
  ERASE Items
  

RETURN
  
ParseTransInfo:
  TransFile = FREEFILE
  OPEN "TaxTrans.dat" FOR RANDOM SHARED AS TransFile LEN = TaxTranRecLen
  TransRecNum& = CVL(RIGHT$(Items(Choice).V, 4))
  GET TransFile, TransRecNum&, TaxTrans(1)
  CLOSE TransFile
  BillNum& = Value(ParseBillNum(TaxTrans(1).Description), ECode)

  TPrince1# = Round#(TaxTrans(1).Revenue.Principle1)
  TPrince1# = Round#(TPrince1# - TaxTrans(1).Revenue.Principle1Pd)
  TPrince2# = Round#(TaxTrans(1).Revenue.Principle2)
  TPrince2# = Round#(TPrince2# - TaxTrans(1).Revenue.Principle2Pd)
  TPrince3# = Round#(TaxTrans(1).Revenue.Principle3)
  TPrince3# = Round#(TPrince3# - TaxTrans(1).Revenue.Principle3Pd)
  TPrince4# = Round#(TaxTrans(1).Revenue.Principle4)
  TPrince4# = Round#(TPrince4# - TaxTrans(1).Revenue.Principle4Pd)
  TPrince5# = Round#(TaxTrans(1).Revenue.Principle5)
  TPrince5# = Round#(TPrince5# - TaxTrans(1).Revenue.Principle5Pd)
  TInterest# = Round#(TaxTrans(1).Revenue.Interest - TaxTrans(1).Revenue.InterestPd)
  TPenalty# = Round#(TaxTrans(1).Revenue.Penalty - TaxTrans(1).Revenue.PenaltyPd)
  
  LSET Form$(7, 0) = STR$(TPrince1#)
  LSET Form$(8, 0) = STR$(TPrince2#)
  LSET Form$(9, 0) = STR$(TPrince3#)
  LSET Form$(10, 0) = STR$(TPrince4#)
  LSET Form$(11, 0) = STR$(TPrince5#)
  LSET Form$(12, 0) = STR$(TPenalty#)
  LSET Form$(13, 0) = STR$(TInterest#)
  Action = 1
  RETURN

END SUB

SUB EnterREAdjustments

  SHARED Choice$()

  BB$ = SPACE$(5)
  u$ = CHR$(24)
  d$ = CHR$(25)

  REDIM TempScrn(0)
  REDIM TaxCustRec(1)        AS TaxCustType
  REDIM TaxTrans(1)          AS TaxTransactionType
  REDIM AmtFlds(1 TO 3)      AS INTEGER

  TaxTranRecLen = LEN(TaxTrans(1))
  TaxCustRecLen = LEN(TaxCustRec(1))

  LibName$ = "TAX"
  ScrnName$ = "VAREABAT"
  TaxType% = 1                  'Real Estate

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

  '--define the multi-choice fields
  CustAcctFld = FldNum%("CUSTACCT", Fld())
  BillFld = FldNum%("BILLS", Fld())

  Action = 1
  FirstTime = True

  DisplayTaxScrn ScrnName$

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF Frm(1).Edited AND Frm(1).PrevFld <> Frm(1).FldNo THEN
      BeenEditedFlag = True     'if the form has been edited
    END IF      'set the edited flag

    IF FirstTime THEN
      FirstTime = False         'if this is the first time
      GOSUB SetupRealFormInfo
    END IF

    IF Frm(1).FldNo > CustAcctFld AND Frm(1).PrevFld = CustAcctFld THEN
      CustAcct& = QPValL(Form$(CustAcctFld, 0))
      IF CustAcct& <> LastCust& THEN
        LastCust& = CustAcct&
        GOSUB GetRealCustInfo
      ELSEIF CustAcct& = 0 THEN
        Frm(1).FldNo = 1
      END IF
      MPaintBox 17, 20, 17, 21, 112
    ELSEIF Frm(1).FldNo = CustAcctFld AND Frm(1).PrevFld <> CustAcctFld THEN
      MPaintBox 17, 20, 17, 21, 126
      MPaintBox 17, 32, 17, 33, 112
    ELSEIF Frm(1).FldNo = BillFld AND Frm(1).PrevFld <> BillFld THEN
      MPaintBox 17, 32, 17, 33, 126
    ELSEIF Frm(1).PrevFld = BillFld AND Frm(1).FldNo <> BillFld THEN
      MPaintBox 17, 32, 17, 33, 112
    END IF

    IF Frm(1).FldNo > CustAcctFld THEN
      IF CustAcct& = 0 THEN
        Frm(1).FldNo = 1
      END IF
    END IF

    IF Frm(1).FldNo > BillFld THEN
      IF TransRecNum& = 0 THEN
        Frm(1).FldNo = BillFld
      END IF
    END IF

    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F4Key
    IF CustAcct& > 0 THEN
      ShowCustHistory -CustAcct&, 1      'DON'T CHANGE THIS
      Action = 1
    END IF

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
    CASE F7KEY  'Lookup Customer
      IF Frm(1).FldNo = CustAcctFld THEN  'if user is on the Customer field
        SaveScrn TempScrn()     'and F7key then do lookup routine
        MPaintBox 6, 4, 18, 76, 8
        LastCust& = CustAcct&
        LookUp CustAcct&, "Payment", 0, False, False, TaxType%
        RestScrn TempScrn()
        IF CustAcct& > 0 THEN   'if this is a valid customer
          GOSUB GetRealCustInfo     'go get customer info
          Frm(1).FldNo = 4
          Action = 1
        ELSEIF LastCust& = CustAcct& THEN   ' don't do anything
          Frm(1).FldNo = 1
          Action = 1
        ELSE
          GOSUB ClearRealForm
          Frm(1).FldNo = 1
          Action = 1
        END IF
      END IF
    CASE F8KEY           'Select the bills being paid
      IF Frm(1).FldNo = BillFld THEN
        GOSUB SelectRealBills2Pay
      END IF
    CASE F10Key 'Save
      GOSUB ChkRealAdjInfo
      IF AdjustOkFlag THEN
        SELECT CASE AskSavePayment
        CASE True               'Save trans no receipt
          GOSUB SaveRealTransaction
          GOSUB ClearRealForm
          Frm(1).FldNo = 1
          Action = 1
          EditFlag = False
        CASE False              'oops, just keep editing
          Action = 2
        END SELECT
      END IF
    'CASE IS <> 0
    'STOP
    END SELECT

    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow   'See if moose was on row 17
      CASE 17                   'it was
        SELECT CASE Frm(1).MCol 'See what column
        CASE 6 TO 17           'f4 History
          PressButton F4Key, 17, 6, 17
        CASE 19 TO 29           'f7 Look-Up
          PressButton F7KEY, 17, 19, 29
        CASE 31 TO 39           'f8 Bill select
          PressButton F8KEY, 17, 31, 39
        CASE 53 TO 62           'f10 Save
          PressButton F10Key, 17, 53, 62
        CASE 64 TO 74           '--cancel button
          PressButton EscKey, 17, 64, 74
        END SELECT
      END SELECT                'row
    END IF
  LOOP UNTIL ExitFlag

  ERASE TempScrn

  HideCursor
EXIT SUB

GetRealCustInfo:
  GOSUB ClearRealForm
  NumOfCustRecs& = FileSize("RETXCUST.DAT") \ TaxCustRecLen
  IF CustAcct& > NumOfCustRecs& OR CustAcct& = 0 THEN
    LastCust& = 0
    CustAcct& = 0
    Ok = MsgBox%("TAX.QSL", "BADACCTN")
    Frm(1).FldNo = 1
    GOTO SkipRealCustInfo
  ELSEIF IsCustDeleted(CustAcct&, TaxType%) THEN
    LastCust& = 0
    CustAcct& = 0
    Ok = MsgBox%("TAX.QSL", "DELACCTN")
    Frm(1).FldNo = 1
    GOTO SkipRealCustInfo
  END IF

  CustFile = FREEFILE
  OPEN "RETXCUST.DAT" FOR RANDOM SHARED AS CustFile LEN = TaxCustRecLen
  GET CustFile, CustAcct&, TaxCustRec(1)
  CLOSE CustFile
  IF NOT DoesCustOwe%(TaxCustRec(1)) THEN
    SaveScrn TempScrn()
    DisplayTaxScrn "ERRSCRN1"
    QPrintRC "This customer has NO BALANCE!", 10, 26, -1
    QPrintRC "Press any key to continue.", 13, 28, -1
    WaitForAction
    RestScrn TempScrn()
    Frm(1).FldNo = 1
    LastCust& = 0
    CustAcct& = 0
    TransRecNum& = 0
    GOTO SkipRealCustInfo
  END IF
  LSET Form$(CustAcctFld, 0) = STR$(CustAcct&)
  First$ = QPTrim$(TaxCustRec(1).FNAME)
  IF LEN(First$) > 0 THEN
    CustName$ = First$ + " " + QPTrim$(TaxCustRec(1).LName)
  ELSE
    CustName$ = First$ + " " + QPTrim$(TaxCustRec(1).LName)
  END IF

  LSET Form$(2, 0) = CustName$
  LSET Form$(3, 0) = TaxCustRec(1).Addr1
  CustAcct& = QPValL(Form$(CustAcctFld, 0))
  FirstTime = True

SkipRealCustInfo:
  Action = 1
RETURN

SetupRealFormInfo:
  LSET BB$ = ""
  QPrintRC BB$, 13, 15, -1
  LSET Form$(4, 0) = DATE$
  LSET Form$(10, 0) = "0.00"
  LSET Form$(11, 0) = "0.00"
  LSET Form$(12, 0) = "0.00"
  CalcFields 0, 10, Form$(), Fld()
  Action = 2
RETURN

ClearRealForm:
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""       '--Clear all fields
  NEXT
  BeenEditedFlag = False        'clear the edited flag
RETURN

SaveRealTransaction:

  SaveScrn TempScrn()
  DisplayTaxScrn "UPDATDSK"
  REDIM TaxAdjTrans(1) AS TaxTransactionType

  CustFile = FREEFILE
  OPEN "RETXCUST.DAT" FOR RANDOM SHARED AS CustFile LEN = TaxCustRecLen
  TranFile = FREEFILE
  OPEN "TAXTRANS.DAT" FOR RANDOM SHARED AS TranFile LEN = TaxTranRecLen
  GET #CustFile, CustAcct&, TaxCustRec(1)      'get cust rec
  GET #TranFile, TransRecNum&, TaxTrans(1)    'get bill we are adjusting

  BillNum$ = ParseBillNum(TaxTrans(1).Description)

  TotalAdj# = Round#(PrinceAdj# + InterestAdj# + PenaltyAdj#)
  TaxAdjTrans(1).TransDate = TransDate
  TaxAdjTrans(1).TranType = 3              'Adjustment transaction

  TaxAdjTrans(1).Revenue.Principle1Pd = PrinceAdj#
  TaxAdjTrans(1).Revenue.InterestPd = InterestAdj#
  TaxAdjTrans(1).Revenue.PenaltyPd = PenaltyAdj#
  TaxAdjTrans(1).Amount = TotalAdj#
  TaxAdjTrans(1).Description = "Tax Abate to Bill #" + BillNum$
  TaxAdjTrans(1).CustomerRec = CustAcct&
  TaxAdjTrans(1).LastTrans = TaxCustRec(1).LastTrans
  TaxAdjTrans(1).BelongTo = TransRecNum&

  TaxTrans(1).Revenue.Principle1Pd = Round#(TaxTrans(1).Revenue.Principle1Pd + PrinceAdj#)
  TaxTrans(1).Revenue.InterestPd = Round#(TaxTrans(1).Revenue.InterestPd + InterestAdj#)
  TaxTrans(1).Revenue.PenaltyPd = Round#(TaxTrans(1).Revenue.PenaltyPd + PenaltyAdj#)

  PUT #TranFile, TransRecNum&, TaxTrans(1)

  NextTransRec& = (LOF(TranFile) \ TaxTranRecLen) + 1

  TaxCustRec(1).LastTrans = NextTransRec&

  PUT #TranFile, NextTransRec&, TaxAdjTrans(1)
  PUT #CustFile, CustAcct&, TaxCustRec(1)

  CLOSE

SaveRealExit:
  DisplayTaxScrn "UPDATEOK"
  WaitForAction
  LastCust& = 0
  CustAcct& = 0
  TransRecNum& = 0
  RestScrn TempScrn()
  FirstTime = True
RETURN

UPDateRealListOfPayments:
RETURN

ChkRealAdjInfo:
  'Parse and move data to Paylist records here
  AdjustOkFlag = True
  ErrCode = 0
  TransDate = Date2Num(Form$(4, 0))
  PrinceAdj# = Value#(Form$(10, 0), ErrCode%)
  PenaltyAdj# = Value#(Form$(11, 0), ErrCode%)
  InterestAdj# = Value#(Form$(12, 0), ErrCode%)

  IF (PrinceAdj# > TPrince#) OR (InterestAdj# > TInterest#) OR (PenaltyAdj# > TPenalty#) THEN
    ErrCode = 1
    GOSUB ShowRealAdjError
    GOTO BadRealPayment
  ELSEIF (PrinceAdj# < 0) OR (InterestAdj# < 0) OR (PenaltyAdj# < 0) THEN
    ErrCode = 2
    GOSUB ShowRealAdjError
    GOTO BadRealPayment
  ELSEIF (PrinceAdj# = 0) AND (InterestAdj# = 0) AND (PenaltyAdj# = 0) THEN
    ErrCode = 3
    GOSUB ShowRealAdjError
    GOTO BadRealPayment
  ELSEIF TransDate <= 0 THEN
    ErrCode = 4
    GOSUB ShowRealAdjError
    GOTO BadRealPayment
  END IF

BadRealPayment:

RETURN

ShowRealAdjError:
  SaveScrn TempScrn()
  DisplayTaxScrn "ERRSCRN1"
  SELECT CASE ErrCode
  CASE 1
    QPrintRC "Adjustment greater than Balance.", 10, 25, -1
    Frm(1).FldNo = 10
  CASE 2
    QPrintRC "Adjustment can not be less than zero.", 10, 22, -1
    Frm(1).FldNo = 10
  CASE 3
    QPrintRC "No adjustment amount entered.", 10, 26, -1
    Frm(1).FldNo = 10
  CASE 4
    QPrintRC "Invalid Adjustment Date.", 10, 29, -1
    Frm(1).FldNo = 4
  END SELECT
  QPrintRC "Correct and Save transaction Again.", 12, 23, -1
  WaitForAction
  RestScrn TempScrn()
  Action = 2
  AdjustOkFlag = False

RETURN

SelectRealBills2Pay:
  Title$ = SPACE$(63)
  LSET Title$ = " TR Date          Tax Year     BillNo.        Amt     Balance"

  BillCnt = 0
  REDIM TempScrn(1)
  SaveScrn TempScrn()
  REDIM TaxTrans(1) AS TaxTransactionType
  IF TaxCustRec(1).LastTrans > 0 THEN
    TransFile = FREEFILE
    OPEN "TaxTrans.dat" FOR RANDOM SHARED AS TransFile LEN = TaxTranRecLen
    TransRecord& = TaxCustRec(1).LastTrans
    DO WHILE TransRecord& <> 0
      GET TransFile, TransRecord&, TaxTrans(1)
      IF TaxTrans(1).TranType = 1 THEN     'only bill transactions
        Balance# = Round#(TaxTrans(1).Revenue.Principle1 + TaxTrans(1).Revenue.Principle2 + TaxTrans(1).Revenue.Principle3 + TaxTrans(1).Revenue.Principle4 + TaxTrans(1).Revenue.Principle5)
        Balance# = Round#(Balance# + TaxTrans(1).Revenue.Interest + TaxTrans(1).Revenue.Penalty + TaxTrans(1).Revenue.Collection)
        Balance# = Round#(Balance# - (TaxTrans(1).Revenue.Principle1Pd + TaxTrans(1).Revenue.Principle2Pd + TaxTrans(1).Revenue.Principle3Pd + TaxTrans(1).Revenue.Principle4Pd + TaxTrans(1).Revenue.Principle5Pd))
        Balance# = Round#(Balance# - (TaxTrans(1).Revenue.InterestPd + TaxTrans(1).Revenue.PenaltyPd + TaxTrans(1).Revenue.CollectionPd))
        IF Balance# > 0 THEN               'only bills with a balance
          BillCnt = BillCnt + 1
          REDIM PRESERVE Items(1 TO BillCnt) AS FLen2
          LSET Items(BillCnt).V = Num2Date(TaxTrans(1).TransDate)
          MID$(Items(BillCnt).V, 20) = FUsing$(STR$(TaxTrans(1).TaxYear), "####")
          BillNum& = Value(ParseBillNum(TaxTrans(1).Description), ECode)
          MID$(Items(BillCnt).V, 30) = FUsing$(STR$(BillNum&), "#####")
          MID$(Items(BillCnt).V, 40) = FUsing$(STR$(TaxTrans(1).Amount), "######.##")
          MID$(Items(BillCnt).V, 50) = FUsing$(STR$(Balance#), "######.##")
          MID$(Items(BillCnt).V, 61) = MKL$(TransRecord&)
        END IF
      END IF
      TransRecord& = TaxTrans(1).LastTrans
    LOOP
    CLOSE TransFile
  ELSE
    Ok = MsgBox%("Tax.QSL", "NOCTRANS")
    RestScrn TempScrn()
  END IF

  MPaintBox 6, 4, 18, 76, 8     'to add shadow to list
  QPrintRC Title$, 3, 9, 112
  MaxLen = 59 'Set menu width to zero
  Action = 0  '0 means stay in the menu until they select something
  IF Choice < 1 THEN
    Choice = 1                'Pre-load choice to highlight
  END IF
  Col = ((80 - 60) \ 2) - 1
  Row = 4
  BoxBot = 15 'limit the box length to go no lower than line 20
  WazzWind BoxBot + 2, Col, BoxBot + 4, MaxLen + 3 + Col, 10, 4, True
  QPrintRC "   Use: " + u$ + "-" + d$ + " to select.", BoxBot + 3, Col + 3, 15
  QPrintRC u$, BoxBot + 3, Col + 11, 14
  QPrintRC d$, BoxBot + 3, Col + 13, 14
  QPrintRC "Press: [ESC] to Cancel.", BoxBot + 3, Col + 33, 15
  QPrintRC "ESC", BoxBot + 3, Col + 41, 14

  DO
    LOCATE Row, Col, 0
    VertMenuT2 Items(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    IF Ky$ = CHR$(27) THEN
      Choice = 0
      EXIT DO 'choice = 0
    ELSEIF Ky$ = CHR$(13) THEN
      GOSUB ParseRealTransInfo
      Frm(1).FldNo = Frm(1).FldNo + 1
      'Action = 2
      EXIT DO
    END IF
  LOOP        'UNTIL EditLocRec& > 0
  RestScrn TempScrn()

  IF Choice > 0 THEN
    LSET BB$ = FUsing$(STR$(BillNum&), "#####")
  ELSE
    LSET BB$ = ""
  END IF

  QPrintRC BB$, 13, 15, -1

  Action = 1
  ERASE Items


RETURN

ParseRealTransInfo:
  TransFile = FREEFILE
  OPEN "TaxTrans.dat" FOR RANDOM SHARED AS TransFile LEN = TaxTranRecLen
  TransRecNum& = CVL(RIGHT$(Items(Choice).V, 4))
  GET TransFile, TransRecNum&, TaxTrans(1)
  CLOSE TransFile
  BillNum& = Value(ParseBillNum(TaxTrans(1).Description), ECode)

  TPrince# = Round#(TaxTrans(1).Revenue.Principle1 + TaxTrans(1).Revenue.Principle2 + TaxTrans(1).Revenue.Principle3 + TaxTrans(1).Revenue.Principle4 + TaxTrans(1).Revenue.Principle5)
  TPrince# = Round#(TPrince# - (TaxTrans(1).Revenue.Principle1Pd + TaxTrans(1).Revenue.Principle2Pd + TaxTrans(1).Revenue.Principle3Pd + TaxTrans(1).Revenue.Principle4Pd + TaxTrans(1).Revenue.Principle5Pd))
  TInterest# = Round#(TaxTrans(1).Revenue.Interest - TaxTrans(1).Revenue.InterestPd)
  TPenalty# = Round#(TaxTrans(1).Revenue.Penalty - TaxTrans(1).Revenue.PenaltyPd)
  

  LSET Form$(7, 0) = STR$(TPrince#)
  LSET Form$(8, 0) = STR$(TPenalty#)
  LSET Form$(9, 0) = STR$(TInterest#)

  Action = 1

RETURN

END SUB

