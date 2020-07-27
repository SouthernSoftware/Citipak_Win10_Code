DEFINT A-Z
DECLARE SUB OPENDrawer ()
DECLARE SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE FUNCTION AskSavePayment% (DefaultFlag$)
DECLARE SUB TLookUp (RecNo&, Text$, ChkBalFlag%, CLSFlag%, SSNFlag%)
DECLARE SUB DelList (Items() AS ANY, Picked%(), NPicked%)
DECLARE SUB DeletePayments ()
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB DeleteT (SEG Element, ElSize, NumEls)
DECLARE SUB PostTaxPayments (OperNum, PostDate$)
DECLARE SUB PrintPaymentJournal ()
DECLARE FUNCTION AskSavePayment% (DefaultFlag$)
DECLARE FUNCTION GetDEFPort% ()
DECLARE SUB WPickList (Items() AS ANY, Picked%(), NPicked%, DspRow%, UpDateFlag%, Cnf AS ANY)
DECLARE SUB DisplayAPScrn (ScrnName$)
DECLARE SUB AddEditOperator ()
DECLARE SUB AuthorizeOperator (oktoadd%)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB BlockClear ()
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
DECLARE SUB LookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, ActiveOnly%)
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
DECLARE FUNCTION IsCustDeleted% (AcctNum&)
DECLARE FUNCTION DoesCustOwe% (TaxCustRec AS ANY)
DECLARE FUNCTION Exist% (FileName$)

       
  TYPE FLen2
    V AS STRING * 64
  END TYPE
  
  TYPE CustPayListType
    CustAcct     AS LONG
    LastPayRec  AS LONG
    NumPayRec   AS LONG
  END TYPE
  
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TaxCONST.BI'
  '$INCLUDE: 'txpaymnt.BI'
  ''$INCLUDE: 'PROPAbst.BI'
  '$INCLUDE: 'cmfiles.BI'

  CONST False = 0, True = NOT False

  STACK 5000
  
  CrLf$ = CHR$(13) + CHR$(10)
  
  DIM SHARED OperNum AS INTEGER, OperOKFlag AS INTEGER
  DIM SHARED PostDate AS STRING, OperPassword AS STRING

SUB EnterTaxPayments (OperNum, PostDate$)
  
  SHARED Choice$()
  
  REDIM TempScrn(0)
  REDIM TaxPaymentRec(1)     AS TaxPaymentRecType
  REDIM PayList(1 TO 1)      AS PayListType
  REDIM TaxCustRec(1)        AS TaxCustType
  REDIM AmtFlds(1 TO 3)      AS INTEGER
  
  TaxPayRecLen = LEN(TaxPaymentRec(1))
  PayListLen = LEN(PayList(1))
  TaxCustRecLen = LEN(TaxCustRec(1))
  
  REDIM TaxSetUp(1) AS TaxMasterType
  TaxSetupLen = LEN(TaxSetUp(1))
  FGetAH "TAXSETUP.DAT", TaxSetUp(1), TaxSetupLen, 1            'load it

  RcptPort = TaxSetUp(1).RcptPort

  RcptPort = GetDEFPort

  IF RcptPort < 1 THEN
    RcptPort = 1
  ELSEIF RcptPort > 2 THEN
    RcptPort = 2
  END IF
  GOSUB LoadCustPayList
  
  IF RecpPort < 1 OR RecpPort > 2 THEN
    RecpPort = 1
  END IF
  
  TOWNNAME$ = TaxSetUp(1).Name
  
  LibName$ = "TAX"
  ScrnName$ = "TXPAYMNT"
  
  '--Get the total number of fields from all pages
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode

  Frm(1).StayOnField = True
  Frm(1).PageDisable = True
  
  '--define the multi-choice fields
  BillsFld = FldNum%("BILLS", Fld())
  AmtOwedFld = FldNum%("AMTOWED", Fld())        'these get field numbers
  TenderFld = FldNum%("TENDERTY", Fld())        'so we can track what field
  CashAmtFld = FldNum%("CASHAMT", Fld())        'the user is currently on
  ChkAmtFld = FldNum%("CHKAMT", Fld())
  AmtRecvFld = FldNum%("AMTRECD", Fld())
  ChangeFld = FldNum%("CHANGE", Fld())
  TotalFld = FldNum%("TOTOWED", Fld())
  AmtPaidFld = FldNum%("AMTPAID", Fld())
  DescFld = FldNum%("DESC", Fld())
  CustAcctFld = FldNum%("CUSTACCT", Fld())
  
  FOR zz = 1 TO 3
    AmtFlds(zz) = FldNum%("AMTOWE" + QPTrim$(STR$(zz)), Fld())
  NEXT
  
  REDIM Choice$(0 TO 3, 0 TO 0)
  
  Choice$(0, 0) = QPTrim$(STR$(TenderFld))
  Choice$(1, 0) = "Cash"
  Choice$(2, 0) = "Check"
  Choice$(3, 0) = "Cash & Check"
  
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
      GOSUB SetOperInfo
    END IF
    
    IF Frm(1).FldNo > CustAcctFld AND Frm(1).PrevFld = CustAcctFld THEN
      CustAcct& = QPValL(Form$(CustAcctFld, 0))
      IF CustAcct& <> LastCust& THEN
        LastCust& = CustAcct&
        GOSUB ChkCustList
        GOSUB GetCustInfo
      ELSEIF CustAcct& = 0 THEN
        Frm(1).FldNo = 1
      END IF
      MPaintBox 22, 20, 22, 21, 112
    ELSEIF Frm(1).FldNo = CustAcctFld AND Frm(1).PrevFld <> CustAcctFld THEN
      MPaintBox 22, 20, 22, 21, 126
      MPaintBox 22, 32, 22, 33, 112
    ELSEIF Frm(1).FldNo = BillsFld AND Frm(1).PrevFld <> BillsFld THEN
      MPaintBox 22, 32, 22, 33, 126
    ELSEIF Frm(1).PrevFld = BillsFld AND Frm(1).FldNo <> BillsFld THEN
      MPaintBox 22, 32, 22, 33, 112
    END IF

    IF Frm(1).FldNo <> TenderFld AND Frm(1).PrevFld = TenderFld THEN
      GOSUB FixCashChkFlds
    END IF

    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE AltO
      OPENDrawer
      
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
      IF Frm(1).FldNo = 3 THEN  'if user is on the Customer field
        SaveScrn TempScrn()     'and F7key then do lookup routine
        MPaintBox 4, 5, 22, 75, 8
        LastCust& = CustAcct&
        TLookUp CustAcct&, "Payment", 0, False, False
        RestScrn TempScrn()
        IF CustAcct& > 0 THEN   'if this is a valid customer
          GOSUB ChkCustList
          GOSUB GetCustInfo     'go get customer info
          IF CustAcct& > 0 THEN   'if this is a valid customer
            Frm(1).FldNo = 4
          END IF
          Action = 1
        ELSEIF LastCust& = CustAcct& THEN
          Frm(1).FldNo = 1
          Action = 1
          ' don't do anything
        ELSE
          GOSUB ClearForm
          Frm(1).FldNo = 1
          Action = 1
        END IF
      END IF
      
    CASE F8KEY  'Select the bills being paid
      IF Frm(1).FldNo = BillsFld THEN
        GOSUB SelectBills2Pay
      END IF
      
    CASE F9KEY
      TempAmtRecv# = Value#(Form$(AmtRecvFld, 0), Ecode)
      IF TempAmtRecv# > 0 THEN
        GOSUB AutoDistribute
      END IF
      
    CASE F10Key 'Save
      GOSUB CheckPaymentInfo
      IF PaymentOKFlag THEN
        SELECT CASE AskSavePayment("N")
        CASE 1  'Save trans print receipt
          ReceiptFlag = True
          GOSUB SaveTransaction
          GOSUB ClearForm
          Frm(1).FldNo = 1
          Action = 1
          EditFlag = False
          GOSUB LoadCustPayList
        CASE -1               'Save trans no receipt
          ReceiptFlag = False
          GOSUB SaveTransaction
          GOSUB ClearForm
          Frm(1).FldNo = 1
          Action = 1
          EditFlag = False
          GOSUB LoadCustPayList
        CASE False              'oops, just keep editing
          Action = 2
        END SELECT
      END IF
    CASE IS <> 0
      'STOP
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 22   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 19 TO 29           'f7 Look-Up
          PressButton F7KEY, 22, 19, 29
        CASE 31 TO 40           'f8 Bill select
          PressButton F8KEY, 22, 31, 40
        CASE 42 TO 50           'f9 Distrubt
          PressButton F9KEY, 22, 42, 50
        CASE 54 TO 63           'f10 Save
          PressButton F10Key, 22, 54, 63
        CASE 65 TO 75           '--cancel button
          PressButton EscKey, 22, 65, 75
        END SELECT
      END SELECT                'row
    END IF
  LOOP UNTIL ExitFlag
  
  ERASE TempScrn, TaxPaymentRec, TaxSetUp
  
  HideCursor
  EXIT SUB
  
GetCustInfo:
  GOSUB ClearForm
  NumOfCustRecs& = FileSize("TAXCUST.DAT") \ TaxCustRecLen
  IF CustAcct& > NumOfCustRecs& OR CustAcct& = 0 THEN
    CustAcct& = 0
    OK = MsgBox%("TAX.QSL", "BADACCTN")
    Frm(1).FldNo = 1
    GOSUB SetOperInfo
    GOTO SkipCustInfo
  ELSEIF IsCustDeleted(CustAcct&) THEN
    CustAcct& = 0
    OK = MsgBox%("TAX.QSL", "DELACCTN")
    Frm(1).FldNo = 1
    GOSUB SetOperInfo
    GOTO SkipCustInfo
  END IF
  
  CustFile = FREEFILE
  OPEN "TAXCUST.DAT" FOR RANDOM SHARED AS CustFile LEN = TaxCustRecLen
  GET CustFile, CustAcct&, TaxCustRec(1)
  CLOSE CustFile
  
  IF NOT EditFlag THEN
    IF NOT DoesCustOwe%(TaxCustRec(1)) THEN
      CustAcct& = 0
      LastCust& = 0
      SaveScrn TempScrn()
      DisplayTaxScrn "ERRSCRN1"
      QPrintRC "This customer has NO BALANCE!", 10, 26, -1
      QPrintRC "Press any key to continue.", 13, 28, -1
      WaitForAction
      RestScrn TempScrn()
      Frm(1).FldNo = 1
      GOSUB SetOperInfo
      GOTO SkipCustInfo
    END IF
    LSET Form$(CustAcctFld, 0) = STR$(CustAcct&)
    CustName$ = QPTrim$(TaxCustRec(1).FNAME) + " " + QPTrim$(TaxCustRec(1).LName)
    LSET Form$(4, 0) = CustName$
    LSET Form$(5, 0) = TaxCustRec(1).Addr1
  ELSE
    PayRecFile = FREEFILE
    OPEN TaxCPRFileName$ FOR RANDOM SHARED AS PayRecFile LEN = TaxPayRecLen
    GET PayRecFile, CustPayRec&, TaxPaymentRec(1)
    CLOSE PayRecFile
    BCopy VARSEG(TaxPaymentRec(1)), VARPTR(TaxPaymentRec(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), LEN(Form$(0, 0)), 0
    UnPackBuffer 0, 0, Form$(), Fld()
  END IF

  CustAcct& = QPValL(Form$(CustAcctFld, 0))
  FirstTime = True
  
SkipCustInfo:
  Action = 1
RETURN
  
SetOperInfo:
  LSET Form$(1, 0) = FUsing$(STR$(OperNum), "##")
  LSET Form$(2, 0) = PostDate$
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
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(TaxPaymentRec(1)), VARPTR(TaxPaymentRec(1)), LEN(Form$(0, 0)), 0
  FOR Cnt = 1 TO 3              'make sure there are no blanks
    IF TaxPaymentRec(1).PaidOwed(Cnt).AMTPAID < -100000# THEN
      TaxPaymentRec(1).PaidOwed(Cnt).AMTPAID = 0
    END IF
    IF TaxPaymentRec(1).PaidOwed(Cnt).AmtOwed < -100000# THEN
      TaxPaymentRec(1).PaidOwed(Cnt).AmtOwed = 0
    END IF
  NEXT
  PayRecFile = FREEFILE
  OPEN TaxCPRFileName$ FOR RANDOM SHARED AS PayRecFile LEN = TaxPayRecLen
  IF NOT EditFlag THEN
    NumOfRecs& = (LOF(PayRecFile) \ TaxPayRecLen) + 1
    CustPayRec& = NumOfRecs&
  END IF
  PUT #PayRecFile, CustPayRec&, TaxPaymentRec(1)
  GOSUB UPDateListOfPayments
  TaxPaymentRec(1).CustPin = TaxCustRec(1).PIN
  PUT #PayRecFile, CustPayRec&, TaxPaymentRec(1)
  EditFlag = False
  CLOSE PayRecFile
  NPicked = 0
  LastCust& = 0
  CustAcct& = 0

  GOSUB PostTaxPayments

  IF ReceiptFlag = True THEN
    'RestScrn TempScrn()
    GOSUB PrintReceipt
  END IF

  KillFile TaxCPRFileName$
  KillFile TaxLOPFileName$

  BlockClear
  DisplayTaxScrn "UPDATEOK"
  WaitForAction

  'PostTaxPayments OperNum, PostDate$

  RestScrn TempScrn()
  FirstTime = True
RETURN
  
UPDateListOfPayments:
'  IF EditFlag THEN
'    REDIM TPayList(1 TO 2) AS PayListType
'    TempName$ = Unique$("")
'    NAME TaxLOPFileName$ AS TempName$
'    NPayListFile = FREEFILE
'    OPEN TaxLOPFileName$ FOR RANDOM AS NPayListFile LEN = PayListLen
'    PayListFile = FREEFILE
'    OPEN TempName$ FOR RANDOM AS PayListFile LEN = PayListLen
'    NumListRec& = (LOF(PayListFile) \ PayListLen)
'    FOR Cnt& = 1 TO NumListRec&
'      GET #PayListFile, Cnt&, TPayList(1)
'      IF TPayList(1).CustRec <> CustAcct& THEN  'if this rec isn't this cust
'        PUT #NPayListFile, , TPayList(1)        'write it to disk
'      ELSE
'        PUT #NPayListFile, , TPayList(2)        'write it to disk
'      END IF
'    NEXT
'    CLOSE PayListFile, NPayListFile
'    KILL TempName$              'erase the old temp file
'    'Paylist file is now clean of cust old records.
'    'Common code below will handle putting new ones back
'  END IF
'
  PayListFile = FREEFILE
  OPEN TaxLOPFileName$ FOR RANDOM AS PayListFile LEN = PayListLen
  NextListRec& = (LOF(PayListFile) \ PayListLen)
  FOR Cnt = 1 TO NPicked
    NextListRec& = NextListRec& + 1
    IF Cnt = 1 THEN
      PayList(Cnt).PrevListRec = 0
    ELSE
      PayList(Cnt).PrevListRec = NextListRec& - 1
    END IF
    PUT #PayListFile, NextListRec&, PayList(Cnt)
  NEXT
  'point to last list rec that belongs to this customer
  TaxPaymentRec(1).LastPayRec = NextListRec&
  'Count of pay list record this customer has
  TaxPaymentRec(1).NumPayRec = NPicked
  CLOSE NPayListFile
RETURN
  
LoadCustPayList:
  Oper$ = QPTrim$(STR$(OperNum))
  PayRecpName$ = "C:\TAXRCP" + Oper$ + ".RPT"
  TaxCPRFileName$ = "TAXCPR" + Oper$ + ".DAT"   'Customers Payment Record file
  TaxLOPFileName$ = "TAXLOP" + Oper$ + ".DAT"   'List Of Payments customers
  PayRecFile = FREEFILE
  OPEN TaxCPRFileName$ FOR RANDOM SHARED AS PayRecFile LEN = TaxPayRecLen
  NumOfRecs& = LOF(PayRecFile) \ TaxPayRecLen
  IF NumOfRecs& > 0 THEN
    REDIM CustList(1 TO NumOfRecs&)    AS CustPayListType
    FOR Cnt& = 1 TO NumOfRecs&
      GET #PayRecFile, Cnt&, TaxPaymentRec(1)
      CustList(Cnt&).CustAcct = TaxPaymentRec(1).CustAcct
      CustList(Cnt&).LastPayRec = TaxPaymentRec(1).LastPayRec
      CustList(Cnt&).NumPayRec = TaxPaymentRec(1).NumPayRec
    NEXT
  END IF
  CLOSE PayRecFile
  CustListCnt& = NumOfRecs&
RETURN
'
ChkCustList:
  EditFlag = False
  IF CustListCnt& > 0 THEN
    FOR Cnt = 1 TO CustListCnt&
      IF CustList(Cnt).CustAcct = CustAcct& THEN
        CustPayRec& = Cnt
        NPicked = CustList(Cnt).NumPayRec
        LastPayRec& = CustList(Cnt).LastPayRec
        GOSUB LoadEditCustPayList
        EditFlag = True
        EXIT FOR
      END IF
    NEXT
  END IF
RETURN
  
LoadEditCustPayList:
  TPrinciple# = 0
  TInterest# = 0
  TCollection# = 0
  LCnt = 0
  REDIM TPayList(1) AS PayListType
  REDIM PayList(1 TO 1)      AS PayListType
  PayListFile = FREEFILE
  OPEN TaxLOPFileName$ FOR RANDOM AS PayListFile LEN = PayListLen
  ThisPayRec& = LastPayRec&
  DO WHILE ThisPayRec& > 0
    LCnt = LCnt + 1
    REDIM PRESERVE PayList(1 TO LCnt)      AS PayListType
    GET #PayListFile, ThisPayRec&, TPayList(1)
    PayList(LCnt).BillRec = TPayList(1).BillRec
    PayList(LCnt).CustRec = TPayList(1).CustRec
    PayList(LCnt).Principle1 = TPayList(1).Principle1
    TPrinciple# = Round#(TPrinciple# + TPayList(1).Principle1)
    PayList(LCnt).Interest1 = TPayList(1).Interest1
    TInterest# = Round#(TInterest# + TPayList(1).Interest1)
    PayList(LCnt).Collection = TPayList(1).Collection
    TCollection# = Round#(TCollection# + TPayList(1).Collection)
    ThisPayRec& = TPayList(1).PrevListRec
  LOOP
  NPicked = LCnt
  CLOSE PayListFile
RETURN
  
CheckPaymentInfo:
  'Parse and move data to Paylist records here
  PaymentOKFlag = True
  PrinceOw# = Value#(Form$(AmtFlds(1), 0), Ecode)
  PrincePD# = Value#(Form$(AmtFlds(1) + 1, 0), Ecode)
  InterestOw# = Value#(Form$(AmtFlds(2), 0), Ecode)
  InterestPD# = Value#(Form$(AmtFlds(2) + 1, 0), Ecode)
  CollectOw# = Value#(Form$(AmtFlds(3), 0), Ecode)
  CollectPd# = Value#(Form$(AmtFlds(3) + 1, 0), Ecode)
  
  TAmtRecv# = Value#(Form$(AmtRecvFld, 0), Ecode)
  TAmtPaid# = Value#(Form$(AmtPaidFld, 0), Ecode)
  ChangeAmt# = Value#(Form$(ChangeFld, 0), Ecode)
  
  IF TAmtPaid# = 0 THEN
    OK = MsgBox%("TAX.QSL", "BADPYTOT")         'show bad scrn
    Action = 2
    PaymentOKFlag = False
    Frm(1).FldNo = Frm(1).PrevFld
    GOTO BadPayment
  END IF
  IF TAmtRecv# = Round#(TAmtPaid# + ChangeAmt#) AND TAmtRecv# > 0 AND ChangeAmt# >= 0 THEN
    PaymentOKFlag = True
  ELSE
    OK = MsgBox%("TAX.QSL", "BADPYTOT")         'show bad scrn
    Action = 2
    PaymentOKFlag = False
    Frm(1).FldNo = Frm(1).PrevFld
    GOTO BadPayment
  END IF
  TenderType$ = QPTrim$(Form$(TenderFld, 0))
  IF LEN(TenderType$) = 0 THEN
    OK = MsgBox%("TAX.QSL", "BADTENDR")
    Action = 2
    PaymentOKFlag = False
    Frm(1).FldNo = TenderFld
    GOTO BadPayment
  END IF
  
  IF (PrincePD# > PrinceOw#) OR (InterestPD# > InterestOw#) OR (CollectPd# > CollectOw#) THEN
    SaveScrn TempScrn()
    DisplayTaxScrn "ERRSCRN1"
    QPrintRC "Can not overpay Tax Payments.", 10, 27, -1
    QPrintRC "Correct and Save transaction Again.", 12, 24, -1
    WaitForAction
    RestScrn TempScrn()
    Action = 2
    PaymentOKFlag = False
    Frm(1).FldNo = AmtFlds(1)
    GOTO BadPayment
  END IF
  
  FOR Cnt = 1 TO NPicked
    PPrinciple# = PayList(Cnt).Principle1
    IF (PrincePD# >= PPrinciple#) AND (PrincePD# > 0) THEN
      PrincePD# = Round#(PrincePD# - PPrinciple#)
    ELSE
      IF PrincePD# > 0 THEN
        PayList(Cnt).Principle1 = PrincePD#
        PrincePD# = 0
      ELSE
        PayList(Cnt).Principle1 = 0
      END IF
    END IF
    PInterest# = PayList(Cnt).Interest1
    IF (InterestPD# >= PInterest#) AND (InterestPD# > 0) THEN
      InterestPD# = Round#(InterestPD# - PInterest#)
    ELSE
      IF InterestPD# > 0 THEN
        PayList(Cnt).Interest1 = InterestPD#
        InterestPD# = 0
      ELSE
        PayList(Cnt).Interest1 = 0
      END IF
    END IF
    PCollect# = PayList(Cnt).Collection
    IF (CollectPd# >= PCollect#) AND (CollectPd# > 0) THEN
      CollectPd# = Round#(CollectPd# - PCollect#)
    ELSE
      IF CollectPd# > 0 THEN
        PayList(Cnt).Collection = CollectPd#
        CollectPd# = 0
      ELSE
        PayList(Cnt).Collection = 0
      END IF
    END IF
  NEXT
BadPayment:
  
RETURN
  
SelectBills2Pay:
  
  BillCnt = 0
  REDIM TempScrn(1)
  SaveScrn TempScrn()
  REDIM TaxTrans(1) AS TaxTransactionType
  IF TaxCustRec(1).LastTrans > 0 THEN
    TransFile = FREEFILE
    OPEN "TaxTrans.dat" FOR RANDOM SHARED AS TransFile LEN = LEN(TaxTrans(1))
    TransRecord& = TaxCustRec(1).LastTrans
    DO WHILE TransRecord& <> 0
      GET TransFile, TransRecord&, TaxTrans(1)
      IF TaxTrans(1).TranType = 1 THEN
        Balance# = Round#(TaxTrans(1).Revenue.Principle1 + TaxTrans(1).Revenue.Principle2 + TaxTrans(1).Revenue.Principle3 + TaxTrans(1).Revenue.Principle4 + TaxTrans(1).Revenue.Principle5)
        Balance# = Round#(Balance# + TaxTrans(1).Revenue.Interest + TaxTrans(1).Revenue.Penalty + TaxTrans(1).Revenue.Collection)
        Balance# = Round#(Balance# - (TaxTrans(1).Revenue.Principle1PD + TaxTrans(1).Revenue.Principle2Pd + TaxTrans(1).Revenue.Principle3Pd + TaxTrans(1).Revenue.Principle4Pd + TaxTrans(1).Revenue.Principle5Pd))
        Balance# = Round#(Balance# - (TaxTrans(1).Revenue.InterestPD + TaxTrans(1).Revenue.PenaltyPd + TaxTrans(1).Revenue.CollectionPD))
        IF Balance# > 0 THEN
          BillCnt = BillCnt + 1
          REDIM PRESERVE Items(1 TO BillCnt) AS FLen2
          LSET Items(BillCnt).V = Num2Date(TaxTrans(1).TransDate)
          MID$(Items(BillCnt).V, 20) = FUsing$(STR$(TaxTrans(1).TaxYear), "####")
          MID$(Items(BillCnt).V, 30) = FUsing$(STR$(TransRecord&), "#####")
          MID$(Items(BillCnt).V, 40) = FUsing$(STR$(TaxTrans(1).Amount), "######.##")
          MID$(Items(BillCnt).V, 50) = FUsing$(STR$(Balance#), "######.##")
          'MID$(Items(BillCnt).V, 61) = MKL$(TransRecord&)
        END IF
      END IF
      TransRecord& = TaxTrans(1).LastTrans
    LOOP
    CLOSE TransFile
  END IF
  
  DspRow = 5
  DidCnt = 0
  REDIM Picked(1 TO BillCnt)
  FOR Cnt = 1 TO BillCnt        'retag any previous selections
    TNum$ = MID$(Items(Cnt).V, 30, 5)
    TransNum& = QPValL&(TNum$)
    FOR zz = 1 TO NPicked
      IF PayList(zz).BillRec = TransNum& THEN
        DidCnt = DidCnt + 1
        Picked(DidCnt) = Cnt
        EXIT FOR
      END IF
    NEXT
  NEXT
  
  MPaintBox 2, 4, 23, 76, 8     'to add shadow to list
  DisplayTaxScrn "PICKOPTI"
  Title$ = SPACE$(68)
  LSET Title$ = " TR Date          Tax Year      TR No.        Amt     Balance"
  QPrintRC Title$, 4, 7, 112
  
  WPickList Items(), Picked(), NPicked, DspRow%, UpDateFlag, Cnf
  
  IF NPicked > 0 AND UpDateFlag THEN
    TPrinciple# = 0
    TInterest# = 0
    TCollection# = 0
    REDIM PayList(1 TO NPicked)      AS PayListType
    TransFile = FREEFILE
    OPEN "TaxTrans.dat" FOR RANDOM SHARED AS TransFile LEN = LEN(TaxTrans(1))
    FOR zz = 1 TO NPicked
      TNum$ = MID$(Items(Picked(zz)).V, 30, 5)
      TransNum& = QPValL&(TNum$)
      GET #TransFile, TransNum&, TaxTrans(1)
      PayList(zz).BillRec = TransNum&
      PayList(zz).CustRec = CustAcct&
      PayList(zz).Principle1 = Round#(TaxTrans(1).Revenue.Principle1 - TaxTrans(1).Revenue.Principle1PD)
      TPrinciple# = Round#(TPrinciple# + PayList(zz).Principle1)
      PayList(zz).Interest1 = Round#(TaxTrans(1).Revenue.Interest - TaxTrans(1).Revenue.InterestPD)
      TInterest# = Round#(TInterest# + PayList(zz).Interest1)
      PayList(zz).Collection = Round#(TaxTrans(1).Revenue.Collection - TaxTrans(1).Revenue.CollectionPD)
      TCollection# = Round#(TCollection# + PayList(zz).Collection)
      Frm(1).FldNo = TenderFld
    NEXT
    CLOSE
    LSET Form$(AmtFlds(1), 0) = STR$(TPrinciple#)
    LSET Form$(AmtFlds(1) + 1, 0) = STR$(0)
    LSET Form$(AmtFlds(2), 0) = STR$(TInterest#)
    LSET Form$(AmtFlds(2) + 1, 0) = STR$(0)
    LSET Form$(AmtFlds(3), 0) = STR$(TCollection#)
    LSET Form$(AmtFlds(3) + 1, 0) = STR$(0)
    TAmtOwed# = Round#(TPrinciple# + TInterest# + TCollection#)
    LSET Form$(AmtOwedFld, 0) = STR$(TAmtOwed#)

    LSET Form$(ChkAmtFld, 0) = STR$(0)
    LSET Form$(CashAmtFld, 0) = STR$(0)
    LSET Form$(TotalFld, 0) = STR$(0)
    LSET Form$(AmtRecvFld, 0) = STR$(0)
    LSET Form$(ChangeFld, 0) = STR$(0)
    LSET Form$(AmtPaidFld, 0) = STR$(0)
    
    FOR zz = 1 TO 3
      SaveField AmtFlds(zz), Form$(), Fld(), BadField
      SaveField AmtFlds(zz) + 1, Form$(), Fld(), BadField
    NEXT

    SaveField ChangeFld, Form$(), Fld(), BadField
    SaveField ChkAmtFld, Form$(), Fld(), BadField
    SaveField CashAmtFld, Form$(), Fld(), BadField
    SaveField TotalFld, Form$(), Fld(), BadField
    SaveField AmtRecvFld, Form$(), Fld(), BadField
    SaveField AmtPaidFld, Form$(), Fld(), BadField

    CalcFields 0, AmtFlds(1), Form$(), Fld()
    CalcFields 0, AmtFlds(1) + 1, Form$(), Fld()
    CalcFields 0, AmtRecvFld, Form$(), Fld()
    CalcFields 0, AmtPaidFld, Form$(), Fld()
    CalcFields 0, TotalFld, Form$(), Fld()
  END IF
  
  Action = 1
  RestScrn TempScrn()
  
  ERASE Items
  
RETURN
  
AutoDistribute:
  
  FOR Cnt = 3 TO 1 STEP -1
    ThisAmt# = Value(Form$(AmtFlds(Cnt), 0), Ecode)
    IF ThisAmt# < 0 THEN
      TempAmtRecv# = Round#(TempAmtRecv# - ThisAmt#)
    END IF
  NEXT
  
  FOR Cnt = 3 TO 1 STEP -1
    ThisAmt# = Value(Form$(AmtFlds(Cnt), 0), Ecode)
    IF ThisAmt# > 0 THEN
      IF TempAmtRecv# >= ThisAmt# THEN
        LSET Form$(AmtFlds(Cnt) + 1, 0) = QPTrim$(Form$(AmtFlds(Cnt), 0))
        TempAmtRecv# = Round#(TempAmtRecv# - ThisAmt#)
      ELSE
        ThisAmt# = TempAmtRecv#
        LSET Form$(AmtFlds(Cnt) + 1, 0) = STR$(ThisAmt#)
        TempAmtRecv# = 0
      END IF
    ELSEIF TempAmtRecv# = 0 THEN
      LSET Form$(AmtFlds(Cnt) + 1, 0) = STR$(0)
    ELSEIF ThisAmt# = 0 THEN
      LSET Form$(AmtFlds(Cnt) + 1, 0) = STR$(0)
    END IF
    SaveField AmtFlds(Cnt) + 1, Form$(), Fld(), BadField
  NEXT
  
  SaveField ChkAmtFld, Form$(), Fld(), BadField
  SaveField CashAmtFld, Form$(), Fld(), BadField
  
  CalcFields 0, AmtFlds(1) + 1, Form$(), Fld()
  CalcFields 0, CashAmtFld, Form$(), Fld()
  
  PrintArray 1, NumFlds - 1, Form$(), Fld()
  
RETURN

FixCashChkFlds:
  SELECT CASE QPTrim$(Form$(TenderFld, 0))
  CASE Choice$(1, 0)            'CASH               this sets the cash or check
    Fld(TenderFld + 1).Protected = False        'amount fields protected or
    Fld(TenderFld + 2).Protected = True         'unprotected based of the
    LSET Form$(ChkAmtFld, 0) = "0"
  CASE Choice$(2, 0)            'CHECK              tender type field selection
    Fld(TenderFld + 1).Protected = True
    Fld(TenderFld + 2).Protected = False
    LSET Form$(CashAmtFld, 0) = "0"
  CASE ELSE     'BOTH
    Fld(TenderFld + 1).Protected = False
    Fld(TenderFld + 2).Protected = False
  END SELECT

  SaveField ChkAmtFld, Form$(), Fld(), BadField
  SaveField CashAmtFld, Form$(), Fld(), BadField

  CalcFields 0, AmtFlds(1), Form$(), Fld()
  CalcFields 0, CashAmtFld, Form$(), Fld()

  PrintArray 1, NumFlds - 1, Form$(), Fld()
  Action = 2
  'GOSUB AutoDistribute:
RETURN
  
PrintReceipt:
'  SaveScrn TempScrn()
  PayRecFile = FREEFILE

  OPEN TaxCPRFileName$ FOR RANDOM SHARED AS PayRecFile LEN = TaxPayRecLen
  GET #PayRecFile, CustPayRec&, TaxPaymentRec(1)
  CLOSE
  
  RHandle = FREEFILE
  OPEN PayRecpName$ FOR OUTPUT AS RHandle
  PRINT #RHandle, CHR$(27); "p"; CHR$(0); CHR$(25); CHR$(250)
  PRINT #RHandle, CHR$(7)
  PRINT #RHandle, TOWNNAME$
  PRINT #RHandle, "TAX PAYMENT"
  PRINT #RHandle, "Date: "; PostDate$
  PRINT #RHandle,
  PRINT #RHandle, "CUSTOMER NAME & DESC. OF PAYMENT"
  PRINT #RHandle, TaxPaymentRec(1).CustName
  PRINT #RHandle, TaxPaymentRec(1).CustAddr
  PRINT #RHandle, TaxPaymentRec(1).Desc
  PRINT #RHandle, "Acct. No. "; TaxPaymentRec(1).CustAcct
  PRINT #RHandle,
  PRINT #RHandle, "Total Owed: "; USING "$$####,#.##"; TaxPaymentRec(1).AmtOwed
  PRINT #RHandle, "Total Paid: "; USING "$$####,#.##"; TaxPaymentRec(1).AMTPAID
  PRINT #RHandle, "Change Due: "; USING "$$####,#.##"; TaxPaymentRec(1).Change
  PRINT #RHandle,
  PRINT #RHandle,
  PRINT #RHandle, "Operator: "; OperNum
  PRINT #RHandle, "Receipt#: "; USING "######"; FileSize(TaxCPRFileName$) \ TaxPayRecLen
  PRINT #RHandle,
  PRINT #RHandle, "       T H A N K   Y O U !"
  PRINT #RHandle,
  PRINT #RHandle,
  PRINT #RHandle,
  PRINT #RHandle,
  PRINT #RHandle,
  PRINT #RHandle,
  CLOSE RHandle

  'Shell$ = "type " + PayRecpName$ + " > com2:"
  'SHELL Shell$
  
  PrintRptFile Header$, PayRecpName$, RcptPort, RetCode%, 5

  'KillFile PayRecpName$
 ' RestScrn TempScrn()
RETURN

PostTaxPayments:
  CursorOff

  Oper$ = QPTrim$(STR$(OperNum))

  TaxCPRFileName$ = "TAXCPR" + Oper$ + ".DAT"   'Customers Payment Record file
  TaxLOPFileName$ = "TAXLOP" + Oper$ + ".DAT"   'List Of Payments customers

  REDIM TaxPaymentRec(1)     AS TaxPaymentRecType
  REDIM PayListRec(1 TO 1)   AS PayListType
  REDIM TaxCustRec(1)        AS TaxCustType
  REDIM TaxTranRec(1)        AS TaxTransactionType
  REDIM PayTranRec(1)        AS TaxTransactionType
  REDIM CMTRRec(1)           AS CMTransRecType

  TaxPayRecLen = LEN(TaxPaymentRec(1))
  PayListLen = LEN(PayListRec(1))
  TaxCustRecLen = LEN(TaxCustRec(1))
  TaxTranRecLen = LEN(TaxTranRec(1))
  CMTrRecLen = LEN(CMTRRec(1))

  'BlockClear
  ShowProcessingScrn "Posting Payment Transactions"

  PayRecFile = FREEFILE
  OPEN TaxCPRFileName$ FOR RANDOM SHARED AS PayRecFile LEN = TaxPayRecLen

  PayListFile = FREEFILE
  OPEN TaxLOPFileName$ FOR RANDOM AS PayListFile LEN = PayListLen

  CustFile = FREEFILE
  OPEN "TAXCUST.DAT" FOR RANDOM SHARED AS CustFile LEN = TaxCustRecLen

  TranFile = FREEFILE
  OPEN "TAXTRANS.DAT" FOR RANDOM SHARED AS TranFile LEN = TaxTranRecLen

  CHandle = FREEFILE
  OPEN "CMTRANS.DAT" FOR RANDOM SHARED AS CHandle LEN = CMTrRecLen

  NumOfRec& = LOF(PayRecFile) \ TaxPayRecLen
  'ShowPctComp 1, NumOfRec&

  FOR Cnt& = 1 TO NumOfRec&
    GET #PayRecFile, Cnt&, TaxPaymentRec(1)
    REDIM CMTRRec(1)     AS CMTransRecType
    ThisListRec& = TaxPaymentRec(1).LastPayRec
    DO WHILE ThisListRec& > 0
      GET #PayListFile, ThisListRec&, PayListRec(1)          'get paylist rec
      GET #CustFile, TaxPaymentRec(1).CustAcct, TaxCustRec(1)'get cust rec
      GET #TranFile, PayListRec(1).BillRec, TaxTranRec(1)
      'get bill trans this payrec is for
      REDIM PayTranRec(1)  AS TaxTransactionType
      'make new clean payment trans
      TotalPaid# = Round#(PayListRec(1).Principle1 + PayListRec(1).Interest1 + PayListRec(1).Collection)
      IF TotalPaid# = 0 THEN
        GOTO SkipThisRec
      END IF

      PayTranRec(1).TransDate = TaxPaymentRec(1).PayDate
      PayTranRec(1).TranType = 2
      PayTranRec(1).Revenue.Principle1PD = PayListRec(1).Principle1
      PayTranRec(1).Revenue.InterestPD = PayListRec(1).Interest1
      PayTranRec(1).Revenue.CollectionPD = PayListRec(1).Collection
      PayTranRec(1).Amount = TotalPaid#
      PayTranRec(1).Description = TaxPaymentRec(1).Desc
      PayTranRec(1).CustomerRec = TaxPaymentRec(1).CustAcct
      '  PayTranRec(1).LastTrans = TaxTranRec(1).LastTrans 'prev tran pointer
      PayTranRec(1).LastTrans = TaxCustRec(1).LastTrans
      PayTranRec(1).BelongTo = PayListRec(1).BillRec

      TaxTranRec(1).Revenue.Principle1PD = Round#(TaxTranRec(1).Revenue.Principle1PD + PayListRec(1).Principle1)
      TaxTranRec(1).Revenue.InterestPD = Round#(TaxTranRec(1).Revenue.InterestPD + PayListRec(1).Interest1)
      TaxTranRec(1).Revenue.CollectionPD = Round#(TaxTranRec(1).Revenue.CollectionPD + PayListRec(1).Collection)

      CMTRRec(1).TransRevAmt(1) = Round#(CMTRRec(1).TransRevAmt(1) + PayListRec(1).Principle1)
      CMTRRec(1).TransRevAmt(2) = Round#(CMTRRec(1).TransRevAmt(2) + PayListRec(1).Interest1)
      CMTRRec(1).TransRevAmt(3) = Round#(CMTRRec(1).TransRevAmt(3) + PayListRec(1).Collection)

      PUT #TranFile, PayListRec(1).BillRec, TaxTranRec(1)
      NextTransRec& = (LOF(TranFile) \ TaxTranRecLen) + 1
      PUT #TranFile, NextTransRec&, PayTranRec(1)
      TaxCustRec(1).LastTrans = NextTransRec&
      PUT #CustFile, TaxPaymentRec(1).CustAcct, TaxCustRec(1)
SkipThisRec:
      ThisListRec& = PayListRec(1).PrevListRec
    LOOP
    CMTRRec(1).TransDate = TaxPaymentRec(1).PayDate
    CMTRRec(1).TransAmount = TaxPaymentRec(1).AmtRecd
    CMTRRec(1).TransAmtOwed = TaxPaymentRec(1).AmtOwed
    CMTRRec(1).TransCash = TaxPaymentRec(1).CashAmt
    CMTRRec(1).TransCheck = TaxPaymentRec(1).ChkAmt
    CMTRRec(1).TransDesc = "Tax Payment"
    CMTRRec(1).TransSource = 33
    CMTRRec(1).TransName = TaxPaymentRec(1).CustName
    CMTRRec(1).TransAcctNum = TaxPaymentRec(1).CustAcct
    CMTRRec(1).TransDetNum = 0
    CMTRRec(1).TransOperNum = OperNum
    CMTRRec(1).TransPad = ""
    NextCMTransRec& = (LOF(CHandle) \ CMTrRecLen) + 1
    PUT #CHandle, NextCMTransRec&, CMTRRec(1)
  NEXT

  CLOSE



RETURN
  
END SUB

