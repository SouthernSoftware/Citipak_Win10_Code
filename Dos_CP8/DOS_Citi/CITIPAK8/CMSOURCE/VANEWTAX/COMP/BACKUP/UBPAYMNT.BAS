DEFINT A-Z
DECLARE SUB CustomerInquiry (RecNo&)
DECLARE FUNCTION IsDeleted (CustAcct&)
DECLARE FUNCTION AskSavePayment% (ReceiptFlag$)
DECLARE FUNCTION Date2Num% (WhatDate$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FLof& (Handle%)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Num2Date$ (WhatDate%)
DECLARE FUNCTION PromptSaveData% ()
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValL& (Numb$)
DECLARE FUNCTION Round# (DoubleNum#)
DECLARE FUNCTION WEnvTest% ()
DECLARE SUB AddEditOperator ()
DECLARE SUB AuthorizeOperator (oktoadd%)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB BlockClear ()
DECLARE SUB UBLog (Text$)
DECLARE SUB DisplayOperEntry ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB FClose (FileHandle)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB HideCursor ()
DECLARE SUB KillFile (FileName$)
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB LookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, LocationFlag%)
DECLARE SUB MPaintBox (UlRow%, UlCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB OperatorEntry (OperOKFlag%, OperNum%, PostingDate$)
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPORT%, RetCode%, EntryPoint%)
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
DECLARE SUB RestScrn (Array())
DECLARE SUB SaveScrn (Array())
DECLARE SUB ShowCursor ()
DECLARE SUB ShowCustHistory (CustAcct&)
DECLARE SUB StuffBuf (Ky$)
DECLARE SUB ValidateOperator (OperNum%, OperPassword$, OperOKFlag%)
DECLARE SUB WaitForAction ()
DECLARE FUNCTION GetDEFPort% ()
DECLARE SUB OPENDrawer ()

'---- ASM subs

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.bi'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'newCUST.bi'
  '$INCLUDE: 'ubsetup.BI'
  '$INCLUDE: 'ubpaymnt.BI'
  '$INCLUDE: 'UBtrans.bi'
  '$INCLUDE: 'cmfiles.BI'
  
  TYPE FLen2
    V AS STRING * 64
  END TYPE
  
  CONST False = 0, True = NOT False
  
  DIM SHARED OperNum AS INTEGER, OperOKFlag AS INTEGER

SUB AddEditDeposit (OperNum, PostDate$)
  
  SHARED Choice$()

  City$ = SPACE$(20)
  Deps$ = SPACE$(12)
  Fmt$ = "$$####.##"
  RecpRev$ = SPACE$(15)

  REDIM TempScrn(0)
  REDIM UBCustRec(1)      AS NewUBCustRecType
  
  REDIM UBPaymentRec(1)   AS UBPaymentRecType
  REDIM PayList(1 TO 1)   AS PayListType
  
  UBCustRecLen = LEN(UBCustRec(1))
  UBPayRecLen = LEN(UBPaymentRec(1))
  
  PayFileName$ = "UBDEP" + QPTrim$(STR$(OperNum)) + ".DAT"
  PayRecpName$ = "UBRCP" + QPTrim$(STR$(OperNum)) + ".RPT"
  
  NumOfRevs = MaxRevsCnt
  
  REDIM RevText$(1 TO MaxRevsCnt)
  REDIM UBSetUpRec(1) AS UBSetupRecType
  
  UBSetupLen = LEN(UBSetUpRec(1))
  FGetAH "UBSETUP.DAT", UBSetUpRec(1), UBSetupLen, 1            'load it

  IF INSTR(UBSetUpRec(1).UTILNAME, "SPENCER") THEN
    SpencerFlag = True
  END IF
  Work$ = SPACE$(19)
  
  FOR RevCnt = 1 TO MaxRevsCnt
    RevText$(RevCnt) = LEFT$(QPTrim$(UBSetUpRec(1).Revenues(RevCnt).RevName), 14)
    IF LEN(RevText$(RevCnt)) = 0 THEN
      NumOfRevs = RevCnt - 1
      EXIT FOR
    END IF
  NEXT
  
  TOWNNAME$ = UBSetUpRec(1).UTILNAME
  
  IF NumOfRevs < MaxRevsCnt THEN
    REDIM PRESERVE RevText$(1 TO NumOfRevs)
  END IF
  
  LibName$ = "UB"
  ScrnName$ = "UBPAYDEP"
  
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
  
  GOSUB DepClearForm
  
  REDIM AmtOweFlds(1 TO NumOfRevs)
  REDIM AmtPadFlds(1 TO NumOfRevs)
  
  FOR Cnt = 1 TO NumOfRevs
    AmtOweFlds(Cnt) = FldNum%("AMTOWE" + LTRIM$(STR$(Cnt)), Fld())
    AmtPadFlds(Cnt) = FldNum%("AMTPD" + LTRIM$(STR$(Cnt)), Fld())
  NEXT
  
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
  
  '--define the multi-choice fields
  
  REDIM Choice$(0 TO 3, 0 TO 0)
  
  Choice$(0, 0) = QPTrim$(STR$(TenderFld))
  Choice$(1, 0) = "Cash"
  Choice$(2, 0) = "Check"
  Choice$(3, 0) = "Cash & Check"
  
  Action = 1
  FirstTime = True
  
  DisplayUBScrn ScrnName$
  
  GOSUB DepSetUpDeposit
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF Frm(1).Edited AND Frm(1).PrevFld <> Frm(1).FldNo THEN
      BeenEditedFlag = True     'if the form has been edited
    END IF      'set the edited flag
    
    IF FirstTime THEN
      FirstTime = False         'if this is the first time
      GOSUB DepShowRevSources   '
      GOSUB DepSetOperInfo
      QPrintRC City$, 8, 15, -1
      QPrintRC Deps$, 9, 15, -1
    END IF
    
    IF Frm(1).FldNo > CustAcctFld AND Frm(1).PrevFld = CustAcctFld THEN
      CustAcct& = QPValL(Form$(CustAcctFld, 0))
      'GOSUB DepCheckPayList
      GOSUB DepGetCustInfo
    ELSEIF Frm(1).FldNo = CustAcctFld AND Frm(1).PrevFld <> CustAcctFld THEN
      MPaintBox 22, 37, 22, 41, 112
      MPaintBox 22, 35, 22, 36, 126
    END IF
    
    IF Frm(1).PrevFld = TenderFld AND Frm(1).FldNo <> TenderFld THEN
      GOSUB DepFixCashChkFlds
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
        SaveScrn TempScrn()
        'CustomerInquiry CustAcct&
        RestScrn TempScrn()
        'ShowCustHistory CustAcct&
        Action = 2
      END IF
      
    CASE F7KEY  'Lookup Customer
      IF Frm(1).FldNo = 3 THEN  'if user is on the Customer field
        SaveScrn TempScrn()     'and F7key then do lookup routine
        MPaintBox 4, 5, 22, 75, 8
        LookUp CustAcct&, "Payment", 2, False, False
        RestScrn TempScrn()
        IF CustAcct& > 0 THEN   'if this is a valid customer
          'GOSUB DepCheckPayList
          GOSUB DepGetCustInfo  'go get customer info
          Frm(1).FldNo = 4
          Action = 1
        ELSE
          GOSUB DepClearForm
          Frm(1).FldNo = 1
          Action = 1
        END IF
      END IF

    CASE AltO
      OPENDrawer

    CASE F10Key 'Save
      GOSUB DepCheckPaymentInfo
      IF PaymentOKFlag THEN
        SELECT CASE AskSavePayment(UBSetUpRec(1).RECPDEFT)
        CASE 1
          'here
          UBLog "CM DEPOSIT POSTED ACCT:" + STR$(CustAcct&)
          SaveScrn TempScrn()
          GOSUB DepSaveTransaction              'do the save routine
          GOSUB DepPrintReceipt
          RestScrn TempScrn()
          DisplayUBScrn "PRESSKEY"
          WaitForAction
          RestScrn TempScrn()
          GOSUB DepClearForm
          GOSUB DClearCity
          Frm(1).FldNo = 1
          Action = 1
        CASE True
          UBLog "CM DEPOSIT POSTED ACCT:" + STR$(CustAcct&)
          ReceiptFlag = False
          SaveScrn TempScrn()
          GOSUB DepSaveTransaction              'do the save routine
          DisplayUBScrn "PRESSKEY"
          WaitForAction
          RestScrn TempScrn()
          GOSUB DepClearForm
          GOSUB DClearCity
          Frm(1).FldNo = 1
          Action = 1
        CASE False
          Action = 2
        END SELECT
      END IF
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 22   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
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
  
  ERASE TempScrn, UBCustRec, UBPaymentRec, RevText$, UBSetUpRec
  
  HideCursor
  EXIT SUB
  
DepSetUpDeposit:
  FOR Cnt = 1 TO NumOfRevs
    IF UBSetUpRec(1).Revenues(Cnt).UseDep <> "Y" THEN
      Fld(AmtPadFlds(Cnt)).Protected = True
    END IF
  NEXT
  FOR Cnt = AmtPadFlds(NumOfRevs) + 1 TO NumFlds
    Fld(Cnt).Protected = True
  NEXT
RETURN

DClearCity:
  LSET City$ = ""
  LSET Deps$ = ""
  QPrintRC City$, 8, 15, -1
  QPrintRC Deps$, 9, 15, -1
RETURN
  
DepGetCustInfo:
  NumOfCustRecs& = FileSize("UBCUST.DAT") \ UBCustRecLen
  IF CustAcct& > NumOfCustRecs& OR CustAcct& <= 0 THEN
    CustAcct& = 0
    OK = MsgBox%("UB.QSL", "BADACCTN")
    Frm(1).FldNo = 1
    GOSUB DepClearForm
    GOSUB DepSetOperInfo
    GOSUB DClearCity
    GOTO DepSkipCustInfo
  END IF

'here
  IF IsDeleted(CustAcct&) THEN
    UBLog "CM ERROR: Deleted Account:" + STR$(CustAcct&)
    CustAcct& = 0
    OK = MsgBox%("UB.QSL", "DELACCTN")
    Frm(1).FldNo = 1
    GOSUB DepClearForm
    GOSUB DepSetOperInfo
    GOSUB DClearCity
    GOTO DepSkipCustInfo
  END IF
  
  GOSUB DepClearForm
  CustFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CustFile LEN = UBCustRecLen
  GET CustFile, CustAcct&, UBCustRec(1)
  CLOSE CustFile
  FOR Cnt = 1 TO NumOfRevs
    LSET Form$(AmtOweFlds(Cnt), 0) = STR$(UBCustRec(1).CurrRevAmts(Cnt))
  NEXT
  TotalBalance# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
  LSET Form$(CustAcctFld, 0) = STR$(CustAcct&)
  LSET Form$(4, 0) = UBCustRec(1).CustName
  LSET Form$(5, 0) = UBCustRec(1).ADDR1
  LSET Form$(6, 0) = UBCustRec(1).PAYCMNT
  LSET Form$(TotalFld, 0) = STR$(TotalBalance#)
  LSET Form$(AmtOwedFld, 0) = STR$(TotalBalance#)
  SaveField 0, Form$(), Fld(), BadField
  CustAcct& = QPValL(Form$(CustAcctFld, 0))
  LSET Deps$ = QPTrim$(FUsing$(STR$(UBCustRec(1).DepositAmt), Fmt$))
  LSET City$ = UBCustRec(1).City

  MPaintBox 22, 35, 22, 41, 115
  FirstTime = True
  
DepSkipCustInfo:
  Action = 1
RETURN
  
DepShowRevSources:
  FOR RevCnt = 1 TO NumOfRevs
    QPrintRC RevText$(RevCnt), RevCnt + 4, 42, -1
  NEXT
RETURN
  
DepSetOperInfo:
  LSET Form$(1, 0) = FUsing$(STR$(OperNum), "##")
  LSET Form$(2, 0) = PostDate$
  Action = 2
RETURN
  
DepClearForm:
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""       '--Clear all fields
  NEXT
  BeenEditedFlag = False        'clear the edited flag
RETURN
  
DepFixCashChkFlds:
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
  
  CalcFields 0, AmtPadFlds(1), Form$(), Fld()
  CalcFields 0, CashAmtFld, Form$(), Fld()
  
  PrintArray 1, NumFlds - 1, Form$(), Fld()
  
RETURN
  
DepSaveTransaction:

  DisplayUBScrn "UPDATDSK"

  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(UBPaymentRec(1)), VARPTR(UBPaymentRec(1)), UBPayRecLen, 0

  FOR Cnt = 1 TO 15
    IF UBPaymentRec(1).PaidOwed(Cnt).AmtPd1 < -100000# THEN
      UBPaymentRec(1).PaidOwed(Cnt).AmtPd1 = 0
    END IF
    UBPaymentRec(1).PaidOwed(Cnt).AmtOwe1 = UBPaymentRec(1).PaidOwed(Cnt).AmtPd1
  NEXT
  
  REDIM UBTransRec(1)  AS UBTransRecType

  UBCustRecLen = LEN(UBCustRec(1))
  UBTransRecLen = LEN(UBTransRec(1))

  UBTransRec(1).TransDate = UBPaymentRec(1).PAYDATE
  UBTransRec(1).TransType = TranDepositPayment
  UBTransRec(1).TransDesc = "DEPOSIT PAYMENT"
  UBTransRec(1).OperatorNumber = OperNum
  UBTransRec(1).CustAcctNo = UBPaymentRec(1).CustAcct
  UBTransRec(1).CustStatus = UBCustRec(1).Status
  UBTransRec(1).TransAmt = UBPaymentRec(1).AMTPAID
  UBTransRec(1).CheckAmount = UBPaymentRec(1).ChkAmt
  UBTransRec(1).CashAmount = UBPaymentRec(1).CashAmt

  IF UBTransRec(1).CheckAmount > 0 AND UBTransRec(1).CashAmount > 0 THEN
    UBTransRec(1).PayTypeCode = 3
  ELSEIF UBTransRec(1).CashAmount > 0 THEN
    UBTransRec(1).PayTypeCode = 1
  ELSEIF UBTransRec(1).CheckAmount > 0 THEN
    UBTransRec(1).PayTypeCode = 2
  END IF

  FOR RevAmts = 1 TO MaxRevsCnt
    UBTransRec(1).RevAmt(RevAmts) = UBPaymentRec(1).PaidOwed(RevAmts).AmtPd1
  NEXT

  CHandle = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS #CHandle LEN = UBCustRecLen
  GET #CHandle, UBPaymentRec(1).CustAcct, UBCustRec(1)
  UBCustRec(1).DepositAmt = Round#(UBCustRec(1).DepositAmt + UBTransRec(1).TransAmt)
  UBTransRec(1).PrevTrans = UBCustRec(1).LastTrans

  THandle = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS #THandle LEN = UBTransRecLen
  NextTransRec& = (LOF(THandle) \ UBTransRecLen) + 1

  PUT #THandle, NextTransRec&, UBTransRec(1)
  UBCustRec(1).LastTrans = NextTransRec&

  PUT #CHandle, UBPaymentRec(1).CustAcct, UBCustRec(1)
  CLOSE

  REDIM CMTRRec(1) AS CMTransRecType
  CMTrRecLen = LEN(CMTRRec(1))
  CMTRRec(1).TransDate = UBTransRec(1).TransDate
  CMTRRec(1).TransAmount = UBTransRec(1).TransAmt
  CMTRRec(1).TransAmtOwed = TAmtPaid#
  CMTRRec(1).TransCash = UBTransRec(1).CashAmount
  CMTRRec(1).TransCheck = UBTransRec(1).CheckAmount
  CMTRRec(1).TransDesc = "Deposit Payment"
  CMTRRec(1).TransSource = TranDepositPayment + 20
  CMTRRec(1).TransName = UBCustRec(1).CustName
  CMTRRec(1).TransAcctNum = CustAcct&
  CMTRRec(1).TransDetNum = 0
  CMTRRec(1).TransOperNum = OperNum
  CMTRRec(1).TransPad = ""
  CMTRRec(1).TransAmtOwed = CMTRRec(1).TransAmount

  FldFactor = 0
  FOR Cnt = 1 TO 15
    CMTRRec(1).TransRevAmt(Cnt) = UBTransRec(1).RevAmt(Cnt)
  NEXT Cnt

  CHandle = FREEFILE
  OPEN "CMTRANS.DAT" FOR RANDOM SHARED AS CHandle LEN = CMTrRecLen
  PUT CHandle, (LOF(CHandle) / CMTrRecLen) + 1, CMTRRec(1)
  CLOSE CHandle

  FirstTime = True
RETURN
  
DepPrintReceipt:
'  SaveScrn TempScrn()

  PayRecpName$ = "CMUBDEP.PRN"
  'FCreate PayRecpName$
  RHandle = FREEFILE
  OPEN PayRecpName$ FOR OUTPUT SHARED AS RHandle

  PRINT #RHandle, CHR$(27); "p"; CHR$(0); CHR$(25); CHR$(250)
  PRINT #RHandle, CHR$(7)
  PRINT #RHandle,
  PRINT #RHandle, TOWNNAME$
  PRINT #RHandle, "UTILITY DEPOSIT PAYMENT"
  PRINT #RHandle, "Date: "; PostDate$
  PRINT #RHandle,
  PRINT #RHandle, "CUSTOMER NAME & DESC. OF PAYMENT"
  PRINT #RHandle, UBPaymentRec(1).CustName
  PRINT #RHandle, UBPaymentRec(1).CustAddr
  PRINT #RHandle, UBPaymentRec(1).DESC
  PRINT #RHandle,
  PRINT #RHandle,
  PRINT #RHandle, "Total Paid: "; USING "$$####,#.##"; UBPaymentRec(1).AMTPAID
  PRINT #RHandle, "Change Due: "; USING "$$####,#.##"; UBPaymentRec(1).Change
  PRINT #RHandle,
  FOR RevCnt = 1 TO NumOfRevs
    IF UBPaymentRec(1).PaidOwed(RevCnt).AmtPd1 > 0 THEN
      LSET RecpRev$ = RevText$(RevCnt)
      PRINT #RHandle, RecpRev$; USING "$$#####.##"; UBPaymentRec(1).PaidOwed(RevCnt).AmtPd1
    END IF
  NEXT
  PRINT #RHandle,
  PRINT #RHandle, "Operator: "; OperNum
  PRINT #RHandle, "Receipt#: "; USING "#######"; FileSize("UBTRANS.DAT") \ UBTransRecLen
'put spencer trap here
  IF SpencerFlag THEN
    PRINT #RHandle, CHR$(27); "="; CHR$(2)   'only to display
    LSET Work$ = "THANK YOU!"
    PRINT #RHandle, Work$
    LSET Work$ = "Change:" + FUsing$(STR$(UBPaymentRec(1).Change), "$$####.##")
    PRINT #RHandle, Work$; CHR$(10)
    PRINT #RHandle, CHR$(27); "="; CHR$(1)  'only to printer
  END IF
''
  PRINT #RHandle,
  PRINT #RHandle, "       T H A N K   Y O U !"
  PRINT #RHandle,
  PRINT #RHandle,
  PRINT #RHandle,
  PRINT #RHandle,
  PRINT #RHandle,
  CLOSE RHandle

  PrintRptFile Header$, PayRecpName$, GetDEFPort%, RetCode%, 5

  KillFile PayRecpName$
RETURN

  
DepCheckPaymentInfo:
  PaymentOKFlag = True
  TAmtRecv# = Value#(Form$(AmtRecvFld, 0), ECode)
  TAmtPaid# = Value#(Form$(AmtPaidFld, 0), ECode)
  ChangeAmt# = Value#(Form$(ChangeFld, 0), ECode)
  
  IF TAmtRecv# = Round#(TAmtPaid# + ChangeAmt#) AND TAmtRecv# > 0 THEN
    PaymentOKFlag = True
  ELSE
    OK = MsgBox%("UB.QSL", "BADPYTOT")          'show bad scrn
    Action = 2
    PaymentOKFlag = False
    Frm(1).FldNo = Frm(1).PrevFld
    GOTO DepBadPayment
  END IF
  
  TenderType$ = QPTrim$(Form$(TenderFld, 0))
  IF LEN(TenderType$) = 0 THEN
    OK = MsgBox%("UB.QSL", "BADTENDR")
    Action = 2
    PaymentOKFlag = False
    Frm(1).FldNo = TenderFld
    GOTO DepBadPayment
  END IF
  '  CashAmtFld
  '  ChkAmtFld
  
DepBadPayment:
RETURN
  
END SUB

SUB AddEditPayment (OperNum, PostDate$)
  
'look into keeping date on payments edited on a different day

  SHARED Choice$()
  City$ = SPACE$(20)
  Deps$ = SPACE$(12)
  Fmt$ = "$$####.##"
  RecpRev$ = SPACE$(15)

  REDIM TempScrn(0)
  REDIM UBCustRec(1)    AS NewUBCustRecType
  REDIM UBPaymentRec(1) AS UBPaymentRecType
  REDIM PayList(1 TO 1)           AS PayListType
  
  UBCustRecLen = LEN(UBCustRec(1))
  UBPayRecLen = LEN(UBPaymentRec(1))

  NumOfRevs = MaxRevsCnt
  
  REDIM RevText$(1 TO MaxRevsCnt)
  REDIM UBSetUpRec(1) AS UBSetupRecType
  
  REDIM DistArray(1 TO MaxRevsCnt) AS DistArrayType
  
  UBSetupLen = LEN(UBSetUpRec(1))
  FGetAH "UBSETUP.DAT", UBSetUpRec(1), UBSetupLen, 1            'load it
  IF INSTR(UBSetUpRec(1).UTILNAME, "SPENCER") THEN
    SpencerFlag = True
  END IF
  Work$ = SPACE$(19)
  
  FOR RevCnt = 1 TO MaxRevsCnt
    RevText$(RevCnt) = LEFT$(QPTrim$(UBSetUpRec(1).Revenues(RevCnt).RevName), 14)
    DistArray(RevCnt).DistOrder = UBSetUpRec(1).Revenues(RevCnt).DistOr
    DistArray(RevCnt).DistCnt = RevCnt
    IF LEN(RevText$(RevCnt)) = 0 THEN
      NumOfRevs = RevCnt - 1
      EXIT FOR
    END IF
  NEXT

  REDIM PRESERVE DistArray(1 TO NumOfRevs) AS DistArrayType
  
  DO
    OutOfOrder = False          'assume it's sorted
    FOR X = 1 TO NumOfRevs - 1
      IF DistArray(X).DistOrder > DistArray(X + 1).DistOrder THEN
        SWAP DistArray(X), DistArray(X + 1)     'if we had to swap
        OutOfOrder = True       'we're not done yet
      END IF
    NEXT
  LOOP WHILE OutOfOrder
  
  TOWNNAME$ = UBSetUpRec(1).UTILNAME
  
  IF NumOfRevs < MaxRevsCnt THEN
    REDIM PRESERVE RevText$(1 TO NumOfRevs)
  END IF
  
  LibName$ = "UB"
  ScrnName$ = "UBPAYMNT"
  
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
  
  GOSUB ClearForm
  
  REDIM AmtOweFlds(1 TO NumOfRevs)
  REDIM AmtPadFlds(1 TO NumOfRevs)
  
  FOR Cnt = 1 TO NumOfRevs
    AmtOweFlds(Cnt) = FldNum%("AMTOWE" + LTRIM$(STR$(Cnt)), Fld())
    AmtPadFlds(Cnt) = FldNum%("AMTPD" + LTRIM$(STR$(Cnt)), Fld())
  NEXT
  
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
  
  '--define the multi-choice fields
  
  REDIM Choice$(0 TO 3, 0 TO 0)
  
  Choice$(0, 0) = QPTrim$(STR$(TenderFld))
'  Choice$(1, 0) = "Cash"
'  Choice$(2, 0) = "Check"
'  Choice$(3, 0) = "Cash & Check"
  
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
      GOSUB SetOperInfo
      QPrintRC City$, 8, 15, -1
      QPrintRC Deps$, 9, 15, -1
    END IF

    IF DistFlag THEN
      TempAmtRecv# = Value#(Form$(AmtRecvFld, 0), ECode)
      GOSUB AutoDistribute:
      DistFlag = False
    END IF

    IF Frm(1).FldNo > CustAcctFld AND Frm(1).PrevFld = CustAcctFld THEN
      CustAcct& = QPValL(Form$(CustAcctFld, 0))
      GOSUB GetCustInfo
    ELSEIF Frm(1).FldNo = CustAcctFld AND Frm(1).PrevFld <> CustAcctFld THEN
      MPaintBox 22, 37, 22, 41, 112
      MPaintBox 22, 35, 22, 36, 126
    END IF
    IF Frm(1).FldNo = TenderFld AND Frm(1).PrevFld <> TenderFld THEN
      MPaintBox 22, 18, 22, 22, 112             'this paints the cash and check
      MPaintBox 22, 28, 22, 31, 112             'buttons based on whether user
      MPaintBox 22, 16, 22, 17, 126             'buttons based on whether user
      MPaintBox 22, 26, 22, 27, 126             'buttons based on whether user
      GOSUB FixCashChkFlds
    ELSEIF Frm(1).PrevFld = TenderFld AND Frm(1).FldNo <> TenderFld THEN
      MPaintBox 22, 16, 22, 22, 115             'is on tender type field or
      MPaintBox 22, 26, 22, 31, 115             'on any another field
      GOSUB FixCashChkFlds
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
        SaveScrn TempScrn()
        CustomerInquiry CustAcct&
        RestScrn TempScrn()
        'ShowCustHistory CustAcct&
        Action = 2
      END IF
      
    CASE F5KEY  'Cash Payment
      IF Frm(1).FldNo = TenderFld THEN          'if user is on tender field
        GOSUB DoCashPayment     'and F5key then go do the
      END IF    'cash payment routine
      DistFlag = True
      
    CASE F6KEY  'Check Payment
      IF Frm(1).FldNo = TenderFld THEN          'if user is on tender field
        GOSUB DoCheckPayment    'and F6key then go do the
      END IF    'check payment routine
      DistFlag = True
      
    CASE F7KEY  'Lookup Customer
      IF Frm(1).FldNo = 3 THEN  'if user is on the Customer field
        SaveScrn TempScrn()     'and F7key then do lookup routine
        MPaintBox 4, 5, 22, 75, 8
        LookUp CustAcct&, "Payment", 2, False, False
        RestScrn TempScrn()
        IF CustAcct& > 0 THEN   'if this is a valid customer
          'GOSUB CheckPayList
          GOSUB GetCustInfo     'go get customer info
          Frm(1).FldNo = 4
          Action = 1
        ELSE
          GOSUB ClearForm
          Frm(1).FldNo = 1
          Action = 1
        END IF
      END IF

    CASE AltO
      OPENDrawer

    CASE F9KEY
      TempAmtRecv# = Value#(Form$(AmtRecvFld, 0), ECode)
      IF TempAmtRecv# > 0 THEN
        GOSUB AutoDistribute
      END IF
      
    CASE F10Key 'Save
      GOSUB CheckPaymentInfo
      IF PaymentOKFlag THEN
        SELECT CASE AskSavePayment(UBSetUpRec(1).RECPDEFT)
        CASE 1
          UBLog "CM PAYMENT POSTED ACCT:" + STR$(CustAcct&)
          SaveScrn TempScrn()
          GOSUB SaveTransaction 'do the save routine
          GOSUB PrintReceipt
          RestScrn TempScrn()
          DisplayUBScrn "PRESSKEY"
          WaitForAction
          RestScrn TempScrn()
          GOSUB ClearForm
          GOSUB ClearCity
          Frm(1).FldNo = 1
          Action = 1
        CASE True
          UBLog "CM PAYMENT POSTED ACCT:" + STR$(CustAcct&)
          ReceiptFlag = False
          SaveScrn TempScrn()
          GOSUB SaveTransaction 'do the save routine
          DisplayUBScrn "PRESSKEY"
          WaitForAction
          RestScrn TempScrn()
          GOSUB ClearForm
          GOSUB ClearCity
          Frm(1).FldNo = 1
          Action = 1
        CASE False
          Action = 2
        END SELECT
      END IF
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 22   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 5 TO 13            'f4 History
          PressButton F4KEY, 22, 5, 13
        CASE 15 TO 23           'f5 cash
          IF Frm(1).FldNo = TenderFld THEN
            PressButton F5KEY, 22, 15, 23
          END IF
        CASE 25 TO 32           'f6 check
          IF Frm(1).FldNo = TenderFld THEN
            PressButton F6KEY, 22, 25, 32
          END IF
        CASE 34 TO 42           'f7 Look-Up
          PressButton F7KEY, 22, 34, 42
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
  
  ERASE TempScrn, UBCustRec, UBPaymentRec, RevText$, UBSetUpRec
  
  HideCursor
  EXIT SUB
  
GetCustInfo:
  NumOfCustRecs& = FileSize("UBCUST.DAT") \ UBCustRecLen
  IF CustAcct& > NumOfCustRecs& OR CustAcct& <= 0 THEN
    CustAcct& = 0
    OK = MsgBox%("UB.QSL", "BADACCTN")
    Frm(1).FldNo = 1
    GOSUB ClearForm
    GOSUB SetOperInfo
    GOSUB ClearCity
    GOTO SkipCustInfo
  END IF
  
  CustFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CustFile LEN = UBCustRecLen
  GET CustFile, CustAcct&, UBCustRec(1)
  CLOSE CustFile

  CashFlag = False
  IF UBCustRec(1).CASHONLY = "Y" THEN
    CashFlag = True
  END IF
  GOSUB SetTenderFlds:

'here
  IF IsDeleted(CustAcct&) THEN
    UBLog "CM ERROR: Deleted Account:" + STR$(CustAcct&)
    CustAcct& = 0
    OK = MsgBox%("UB.QSL", "DELACCTN")
    Frm(1).FldNo = 1
    GOSUB ClearForm
    GOSUB SetOperInfo
    GOSUB ClearCity
    GOTO SkipCustInfo
  END IF
  
  GOSUB ClearForm

  FOR Cnt = 1 TO NumOfRevs
    LSET Form$(AmtOweFlds(Cnt), 0) = STR$(UBCustRec(1).CurrRevAmts(Cnt))
  NEXT
  TotalBalance# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
  LSET Form$(CustAcctFld, 0) = STR$(CustAcct&)
  LSET Form$(4, 0) = UBCustRec(1).CustName
  LSET Form$(5, 0) = UBCustRec(1).ADDR1
  LSET Form$(6, 0) = UBCustRec(1).PAYCMNT
  LSET Form$(TotalFld, 0) = STR$(TotalBalance#)
  LSET Form$(AmtOwedFld, 0) = STR$(TotalBalance#)
  SaveField 0, Form$(), Fld(), BadField
  CustAcct& = QPValL(Form$(CustAcctFld, 0))
  LSET Deps$ = QPTrim$(FUsing$(STR$(UBCustRec(1).DepositAmt), Fmt$))
  LSET City$ = UBCustRec(1).City

  MPaintBox 22, 35, 22, 41, 115
  FirstTime = True
  
SkipCustInfo:
  Action = 1
RETURN
  
ShowRevSources:
  FOR RevCnt = 1 TO NumOfRevs
    QPrintRC RevText$(RevCnt), RevCnt + 4, 40, -1
  NEXT
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

ClearCity:
  LSET City$ = ""
  LSET Deps$ = ""
  QPrintRC City$, 8, 15, -1
  QPrintRC Deps$, 9, 15, -1
RETURN

DoCashPayment:
  LSET Form$(TenderFld, 0) = Choice$(1, 0)
  LSET Form$(ChkAmtFld, 0) = "0"
  LSET Form$(CashAmtFld, 0) = Form$(AmtOwedFld, 0)
  GOSUB PaymentCommon
RETURN
  
DoCheckPayment:
  LSET Form$(TenderFld, 0) = Choice$(2, 0)
  LSET Form$(ChkAmtFld, 0) = Form$(AmtOwedFld, 0)
  LSET Form$(CashAmtFld, 0) = "0"
  GOSUB PaymentCommon
RETURN
  
PaymentCommon:
  SaveField TenderFld, Form$(), Fld(), BadField
  SaveField ChkAmtFld, Form$(), Fld(), BadField
  SaveField CashAmtFld, Form$(), Fld(), BadField
  
  LSET Form$(ChangeFld, 0) = "0"
  SaveField ChangeFld, Form$(), Fld(), BadField
  
  FOR Cnt = 1 TO NumOfRevs
    LSET Form$(AmtOweFlds(Cnt) + 1, 0) = Form$(AmtOweFlds(Cnt), 0)
    SaveField AmtOweFlds(Cnt) + 1, Form$(), Fld(), BadField
  NEXT
  
  MPaintBox 22, 16, 22, 22, 115 'is on tender type field or
  MPaintBox 22, 26, 22, 31, 115 'on any another field
  
  Frm(1).FldNo = DescFld
  
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
  
  CalcFields 0, AmtPadFlds(1), Form$(), Fld()
  CalcFields 0, CashAmtFld, Form$(), Fld()
  PrintArray 1, NumFlds - 1, Form$(), Fld()
RETURN
  
AutoDistribute:
  FOR Cnt = 1 TO NumOfRevs
    WhatRev = DistArray(Cnt).DistCnt
    ThisAmt# = Value(Form$(AmtOweFlds(WhatRev), 0), ECode)
    IF ThisAmt# < 0 THEN
      TempAmtRecv# = Round#(TempAmtRecv# - ThisAmt#)
    END IF
  NEXT

  FOR Cnt = 1 TO NumOfRevs
    WhatRev = DistArray(Cnt).DistCnt
    ThisAmt# = Value(Form$(AmtOweFlds(WhatRev), 0), ECode)
    IF ThisAmt# <> 0 THEN
      IF TempAmtRecv# >= ThisAmt# THEN
        LSET Form$(AmtOweFlds(WhatRev) + 1, 0) = QPTrim$(Form$(AmtOweFlds(WhatRev), 0))
        TempAmtRecv# = Round#(TempAmtRecv# - ThisAmt#)
      ELSE
        ThisAmt# = TempAmtRecv#
        LSET Form$(AmtOweFlds(WhatRev) + 1, 0) = STR$(ThisAmt#)
        TempAmtRecv# = 0
      END IF
    ELSEIF TempAmtRecv# = 0 THEN
      LSET Form$(AmtOweFlds(WhatRev) + 1, 0) = STR$(0)
    ELSEIF ThisAmt# = 0 THEN
      LSET Form$(AmtOweFlds(WhatRev) + 1, 0) = STR$(0)
    END IF
    SaveField AmtOweFlds(WhatRev) + 1, Form$(), Fld(), BadField
  NEXT

  SaveField ChkAmtFld, Form$(), Fld(), BadField
  SaveField CashAmtFld, Form$(), Fld(), BadField

  CalcFields 0, AmtPadFlds(1), Form$(), Fld()
  CalcFields 0, CashAmtFld, Form$(), Fld()

  PrintArray 1, NumFlds - 1, Form$(), Fld()
RETURN


'
'  FOR Cnt = 1 TO NumOfRevs
'    WhatRev = DistArray(Cnt).DistCnt
'    ThisAmt# = Value(Form$(AmtOweFlds(WhatRev), 0), ECode)
'    IF ThisAmt# <> 0 THEN
'      IF TempAmtRecv# >= ThisAmt# THEN
'        LSET Form$(AmtOweFlds(WhatRev) + 1, 0) = QPTrim$(Form$(AmtOweFlds(WhatRev), 0))
'        TempAmtRecv# = Round#(TempAmtRecv# - ThisAmt#)
'      ELSE
'        ThisAmt# = TempAmtRecv#
'        LSET Form$(AmtOweFlds(WhatRev) + 1, 0) = STR$(ThisAmt#)
'        TempAmtRecv# = 0
'      END IF
'    ELSEIF TempAmtRecv# = 0 THEN
'      LSET Form$(AmtOweFlds(WhatRev) + 1, 0) = STR$(0)
'    ELSEIF ThisAmt# = 0 THEN
'      LSET Form$(AmtOweFlds(WhatRev) + 1, 0) = STR$(0)
'    END IF
'    SaveField AmtOweFlds(WhatRev) + 1, Form$(), Fld(), BadField
'  NEXT
'
'  SaveField ChkAmtFld, Form$(), Fld(), BadField
'  SaveField CashAmtFld, Form$(), Fld(), BadField
'
'  CalcFields 0, AmtPadFlds(1), Form$(), Fld()
'  CalcFields 0, CashAmtFld, Form$(), Fld()
'  PrintArray 1, NumFlds - 1, Form$(), Fld()
'RETURN

SaveTransaction:
  
  'DisplayUBScrn "UPDATDSK"
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(UBPaymentRec(1)), VARPTR(UBPaymentRec(1)), UBPayRecLen, 0

  FOR Cnt = 1 TO 15
    IF UBPaymentRec(1).PaidOwed(Cnt).AmtPd1 < -100000# THEN
      UBPaymentRec(1).PaidOwed(Cnt).AmtPd1 = 0
    END IF
    IF UBPaymentRec(1).PaidOwed(Cnt).AmtOwe1 < -100000# THEN
      UBPaymentRec(1).PaidOwed(Cnt).AmtOwe1 = 0
    END IF
  NEXT

  REDIM UBTransRec(1)  AS UBTransRecType

  UBCustRecLen = LEN(UBCustRec(1))
  UBTransRecLen = LEN(UBTransRec(1))

  CHandle = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CHandle LEN = UBCustRecLen

  THandle = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS THandle LEN = UBTransRecLen

  GET #CHandle, UBPaymentRec(1).CustAcct, UBCustRec(1)
'added for wadesboro
'more & more & MORE!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
  OldTotBalance# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)

  UBTransRec(1).TransDate = UBPaymentRec(1).PAYDATE
  UBTransRec(1).TransType = TranBillPayment
  UBTransRec(1).TransDesc = "BILLING PAYMENT"
  'UBTransRec(1).CustLocation = CustAcct&
  UBTransRec(1).OperatorNumber = OperNum
  UBTransRec(1).CustAcctNo = UBPaymentRec(1).CustAcct
  UBTransRec(1).CustStatus = UBCustRec(1).Status
  UBTransRec(1).TransAmt = UBPaymentRec(1).AMTPAID
  UBTransRec(1).CheckAmount = UBPaymentRec(1).ChkAmt
  UBTransRec(1).CashAmount = UBPaymentRec(1).CashAmt

  IF UBTransRec(1).CheckAmount > 0 AND UBTransRec(1).CashAmount > 0 THEN
    UBTransRec(1).PayTypeCode = 3
  ELSEIF UBTransRec(1).CashAmount > 0 THEN
    UBTransRec(1).PayTypeCode = 1
  ELSEIF UBTransRec(1).CheckAmount > 0 THEN
    UBTransRec(1).PayTypeCode = 2
  END IF

  IF UBCustRec(1).PrevBalance > 0 THEN
    IF UBTransRec(1).TransAmt >= UBCustRec(1).PrevBalance THEN
      UBCustRec(1).PrevBalance = 0
    ELSEIF UBTransRec(1).TransAmt < UBCustRec(1).PrevBalance THEN
      UBCustRec(1).PrevBalance = Round#(UBCustRec(1).PrevBalance - UBTransRec(1).TransAmt)
    END IF
  END IF

  FOR RevAmts = 1 TO MaxRevsCnt
    UBTransRec(1).RevAmt(RevAmts) = UBPaymentRec(1).PaidOwed(RevAmts).AmtPd1
    UBCustRec(1).CurrRevAmts(RevAmts) = Round#(UBCustRec(1).CurrRevAmts(RevAmts) - UBTransRec(1).RevAmt(RevAmts))
  NEXT
  
  TotalCustBalance# = 0
  FOR RevAmts = 1 TO MaxRevsCnt
    TotalCustBalance# = Round#(TotalCustBalance# + UBCustRec(1).CurrRevAmts(RevAmts))
  NEXT

  UBCustRec(1).CurrBalance = Round#(TotalCustBalance# - UBCustRec(1).PrevBalance)
  UBTransRec(1).RunBalance = TotalCustBalance#
  UBTransRec(1).PrevTrans = UBCustRec(1).LastTrans

'changing to regular basic file I/O

  'find me

  IF Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance) = 0 THEN
    IF UBCustRec(1).Status = "B" THEN
      CustChCnt = CustChCnt + 1
      UBLog "CM PAYMENT: SET CUST STATUS to I. Acct:" + STR$(UBTransRec(1).CustAcctNo)
      UBCustRec(1).Status = "I"
    END IF
  END IF
  
  NextTransRec& = (LOF(THandle) \ UBTransRecLen) + 1
  PUT #THandle, NextTransRec&, UBTransRec(1)

  UBCustRec(1).LastTrans = NextTransRec&
  PUT #CHandle, UBPaymentRec(1).CustAcct, UBCustRec(1)

  CLOSE

  REDIM CMTRRec(1) AS CMTransRecType
  CMTrRecLen = LEN(CMTRRec(1))
  CMTRRec(1).TransDate = UBTransRec(1).TransDate
  CMTRRec(1).TransAmount = UBTransRec(1).CashAmount + UBTransRec(1).CheckAmount
  CMTRRec(1).TransCash = UBTransRec(1).CashAmount
  CMTRRec(1).TransAmtOwed = TAmtPaid#
  CMTRRec(1).TransCheck = UBTransRec(1).CheckAmount
  CMTRRec(1).TransDesc = "Utility Payment"
  CMTRRec(1).TransSource = TranBillPayment + 20
  CMTRRec(1).TransName = UBCustRec(1).CustName
  CMTRRec(1).TransAcctNum = CustAcct&
  CMTRRec(1).TransDetNum = 0
  CMTRRec(1).TransOperNum = OperNum
  CMTRRec(1).TransPad = ""

  FldFactor = 0
  FOR Cnt = 1 TO 15
    CMTRRec(1).TransRevAmt(Cnt) = UBTransRec(1).RevAmt(Cnt)
  NEXT Cnt

  CHandle = FREEFILE
  OPEN "CMTRANS.DAT" FOR RANDOM SHARED AS CHandle LEN = CMTrRecLen
  PUT CHandle, (LOF(CHandle) / CMTrRecLen) + 1, CMTRRec(1)
  CLOSE CHandle

RETURN
  
  
PrintReceipt:
  TheTime$ = TIME$
  IF LEFT$(TheTime$, 1) = "0" THEN
    Hour = VAL(MID$(TheTime$, 2, 1))
  ELSE
    Hour = VAL(MID$(TheTime$, 1, 2))
  END IF

  SELECT CASE Hour
  CASE IS > 12
    Hour = Hour - 12
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

'  SaveScrn TempScrn()
  PayRecpName$ = "CMUBDEP.PRN"

  RHandle = FREEFILE
  OPEN PayRecpName$ FOR OUTPUT SHARED AS RHandle
  PRINT #RHandle, CHR$(27); "p"; CHR$(0); CHR$(25); CHR$(250)
  PRINT #RHandle, CHR$(7)
  PRINT #RHandle,
  PRINT #RHandle, TOWNNAME$
  PRINT #RHandle, "UTILITY PAYMENT"
  PRINT #RHandle, "Date: "; PostDate$
  PRINT #RHandle,
  PRINT #RHandle, "CUSTOMER NAME & DESC. OF PAYMENT"
  PRINT #RHandle, UBPaymentRec(1).CustName
  PRINT #RHandle, UBPaymentRec(1).CustAddr
  PRINT #RHandle, UBPaymentRec(1).DESC
  PRINT #RHandle,
  PRINT #RHandle, "Acct No:"; UBPaymentRec(1).CustAcct
  PRINT #RHandle,
  PRINT #RHandle, "  Cash Amt: "; USING "$$####,#.##"; UBPaymentRec(1).CashAmt
  PRINT #RHandle, " Check Amt: "; USING "$$####,#.##"; UBPaymentRec(1).ChkAmt
  PRINT #RHandle, "Total Paid: "; USING "$$####,#.##"; UBPaymentRec(1).AMTPAID
  PRINT #RHandle, "    Change: "; USING "$$####,#.##"; UBPaymentRec(1).Change
  PRINT #RHandle, "  Acct Bal: "; USING "$$####,#.##"; TotalCustBalance#

'put spencer trap here
  IF SpencerFlag THEN
    PRINT #RHandle, CHR$(27); "="; CHR$(2)   'only to display
    LSET Work$ = "THANK YOU!"
    PRINT #RHandle, Work$
    LSET Work$ = "Change:" + FUsing$(STR$(UBPaymentRec(1).Change), "$$####.##")
    PRINT #RHandle, Work$; CHR$(10)
    PRINT #RHandle, CHR$(27); "="; CHR$(1)  'only to printer
  END IF
''
  PRINT #RHandle,
  FOR RevCnt = 1 TO NumOfRevs
    IF UBPaymentRec(1).PaidOwed(RevCnt).AmtPd1 <> 0 OR UBPaymentRec(1).PaidOwed(RevCnt).AmtOwe1 <> 0 THEN
      LSET RecpRev$ = RevText$(RevCnt)
      PRINT #RHandle, RecpRev$; USING "$$#####.##"; UBPaymentRec(1).PaidOwed(RevCnt).AmtPd1
    END IF
  NEXT
  PRINT #RHandle,
  PRINT #RHandle, "Operator: "; OperNum
  PRINT #RHandle, "Receipt#: "; USING "#######"; FileSize("UBTRANS.DAT") \ UBTransRecLen
  PRINT #RHandle,
  PRINT #RHandle, "       T H A N K   Y O U !"
  PRINT #RHandle,
  PRINT #RHandle, "DATE: "; DATE$
  PRINT #RHandle, "TIME: "; TheTime$
  PRINT #RHandle,
  PRINT #RHandle,
  PRINT #RHandle,
  PRINT #RHandle,
  PRINT #RHandle,
  CLOSE RHandle

  PrintRptFile Header$, PayRecpName$, GetDEFPort%, RetCode%, 5

  KillFile PayRecpName$
'  RestScrn TempScrn()
RETURN

  
CheckPaymentInfo:
  PaymentOKFlag = True
  
  TAmtRecv# = Value#(Form$(AmtRecvFld, 0), ErCode)
  TAmtPaid# = Value#(Form$(AmtPaidFld, 0), ErCode)

  ChangeAmt# = Value#(Form$(ChangeFld, 0), ErCode)

  IF TAmtPaid# = 0 THEN
    OK = MsgBox%("UB.QSL", "BADPYTOT")          'show bad scrn
    Action = 2
    PaymentOKFlag = False
    Frm(1).FldNo = Frm(1).PrevFld
    GOTO BadPayment
  END IF

  IF TAmtRecv# = Round#(TAmtPaid# + ChangeAmt#) AND TAmtRecv# > 0 THEN
    PaymentOKFlag = True
  ELSE
    OK = MsgBox%("UB.QSL", "BADPYTOT")          'show bad scrn
    Action = 2
    PaymentOKFlag = False
    Frm(1).FldNo = Frm(1).PrevFld
    GOTO BadPayment
  END IF
  
  TenderType$ = QPTrim$(Form$(TenderFld, 0))
  IF LEN(TenderType$) = 0 THEN
    OK = MsgBox%("UB.QSL", "BADTENDR")
    Action = 2
    PaymentOKFlag = False
    Frm(1).FldNo = TenderFld
    GOTO BadPayment
  END IF
  
BadPayment:
RETURN

SetTenderFlds:
  IF CashFlag THEN
    Choice$(1, 0) = "Cash"
    Choice$(2, 0) = ""
    Choice$(3, 0) = ""
  ELSE
    Choice$(1, 0) = "Cash"
    Choice$(2, 0) = "Check"
    Choice$(3, 0) = "Cash & Check"
  END IF
RETURN
  
END SUB

