DEFINT A-Z
DECLARE SUB CustomerInquiry (RecNo&)
DECLARE SUB DeletePayTrans (OPERNUM%)
DECLARE SUB DeleteDepTrans (OPERNUM%)
DECLARE SUB DeleteTransMenu (OPERNUM%)
DECLARE FUNCTION IsDeleted% (AcctNo&)
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB DeleteT (SEG Element, ElSize, NumEls)
DECLARE SUB DelList (Items() AS ANY, Picked%(), NPicked%)
DECLARE SUB AddEditDeposit (OPERNUM%, PostDate$)
DECLARE SUB OperatorEntry (OperOKFlag%, OPERNUM%, PostDate$)
DECLARE SUB AddEditOperator ()
DECLARE SUB AuthorizeOperator (oktoadd%)
DECLARE SUB DisplayOperEntry ()
DECLARE SUB ValidateOperator (OPERNUM%, OperPassword$, OperOKFlag%)
DECLARE FUNCTION GetOperator% ()
DECLARE SUB StuffBuf (Ky$)
DECLARE SUB SmallPause ()
DECLARE FUNCTION AskSavePayment% (ReceiptFlag$)
DECLARE FUNCTION PromptSaveData% ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB LookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, LocationFlag%)
DECLARE SUB WaitForAction ()
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB AddEditPayment (OPERNUM, PostDate$)
DECLARE SUB PrintTransJournal (OPERNUM, PostDate$)
DECLARE SUB PostTransEntries (OPERNUM, PostDate$)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB BlockClear ()
DECLARE SUB CursorOff ()
DECLARE SUB MPaintBox (UlRow%, UlCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB MScrnRest (UlRow, UlCol, LRRow, LRCol, SEG Address)
DECLARE SUB MScrnSave (UlRow, UlCol, LRRow, LRCol, SEG Address)
DECLARE SUB Pause3 (MillaSecs%, ProcessorSpeed&)
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
DECLARE SUB SaveScrn (Array())
DECLARE SUB RestScrn (Array())
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValL& (Numb$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
  'DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB ClearScrn ()
DECLARE SUB HideCursor ()
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION WEnvTest% ()
DECLARE FUNCTION Date2Num% (WhatDate$)
DECLARE FUNCTION Num2Date$ (WhatDate%)
DECLARE FUNCTION Exist% (FileName$)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE FUNCTION FLof& (Handle%)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Round# (DoubleNum#)
DECLARE SUB FPutRTA (Handle%, SEG Source AS ANY, RecNo&, RecSize%)
DECLARE SUB FClose (FileHandle)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB UBLog (Text$)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB ShowCustHistory (CUSTACCT&)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPORT%, RetCode%, EntryPoint%)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB KillFile (FileName$)
DECLARE SUB ButtonPress (WhatBtn, N, MooseButton, MCol, MRow)   ' ----- Check Mouse activity
DECLARE SUB GetCursor (a, b, NewStatus)
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetUpRecLen%)
DECLARE SUB OPENDrawer (RecpPort)
DECLARE SUB QPSound (Freq, Dur)
  
  '---- ASM subs
  
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.bi'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'newCUST.bi'
  '$INCLUDE: 'UBtrans.bi'
  '$INCLUDE: 'ubsetup.BI'
  '$INCLUDE: 'ubpaymnt.BI'
  '$INCLUDE: 'cmfiles.BI'
  
  ''$INCLUDE: 'init.BI'
  
  STACK 8000
  
  TYPE FLen2
    V AS STRING * 64
  END TYPE
  
  CONST False = 0, True = NOT False
  
  DIM SHARED OPERNUM AS INTEGER, OperOKFlag AS INTEGER
  
  CrLf$ = CHR$(13) + CHR$(10)
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 6)
  
  MChoice$(1) = "Payment Transaction Entry/Edit "
  MChoice$(2) = "Deposit Transaction Entry/Edit "
  MChoice$(3) = "Delete Payment/Deposit Transaction"
  MChoice$(4) = "Print Transaction Journal"
  MChoice$(5) = "Post Transaction Entries"
  MChoice$(6) = "Exit to OS"
  
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
  Row = ((25 - (UBOUND(MChoice$))) \ 2)
  Col = ((80 - MaxLen) \ 2)
  
  BlockClear
  
  IF NOT WEnvTest% OR COMMAND$ = "TEST" THEN
    DO
      OperatorEntry OperOKFlag, OPERNUM, PostDate$
    LOOP WHILE OperOKFlag = 0
  ELSE
    OPERNUM = 1
    PostDate$ = DATE$
  END IF
  
  IF OperOKFlag < 0 THEN GOTO ExitPaymentProg
  
  '  DIM LB AS LockBoxRecType
  '  PRINT LEN(LB)
  '  END
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    TitleBox 2, Col, MaxLen + 3, "Payments, Deposits Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      AddEditPayment OPERNUM, PostDate$
    CASE 2
      AddEditDeposit OPERNUM, PostDate$
    CASE 3
      DeleteTransMenu OPERNUM
    CASE 4
      PrintTransJournal OPERNUM, PostDate$
    CASE 5
      PostTransEntries OPERNUM, PostDate$
    CASE 6
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP
  
ExitPaymentProg:
  RUN "UBmenu"
  '  HideCursor
  '  ClearScrn
  '  END

SUB AddEditDeposit (OPERNUM, PostDate$)
  
  UBLog " IN: Deposit Payments,  OPER:" + STR$(OPERNUM)
  SHARED Choice$()

  CutPaper$ = CHR$(29) + CHR$(86) + CHR$(66) + CHR$(64)

  CITY$ = SPACE$(20)
  Deps$ = SPACE$(12)
  Fmt$ = "$$####.##"
  RecpRev$ = SPACE$(15)
  
  REDIM TempScrn(0)
  REDIM UBCustRec(1)      AS NewUBCustRecType
  
  REDIM UBPaymentRec(1)   AS UBPaymentRecType
  REDIM PayList(1 TO 1)   AS PayListType
  
  UBCustRecLen = LEN(UBCustRec(1))
  UBPayRecLen = LEN(UBPaymentRec(1))
  
  PayFileName$ = "UBDEP" + QPTrim$(STR$(OPERNUM)) + ".DAT"
  PayRecpName$ = "UBRCP" + QPTrim$(STR$(OPERNUM)) + ".RPT"
  
  GOSUB DepLoadPayList
  
  NumOfRevs = MaxRevsCnt
  
  REDIM RevText$(1 TO MaxRevsCnt)
  REDIM UBSetUpRec(1) AS UBSetupRecType
  
  UBSetUpLen = LEN(UBSetUpRec(1))
  FGetAH "UBSETUP.DAT", UBSetUpRec(1), UBSetUpLen, 1            'load it
  
  FOR RevCnt = 1 TO MaxRevsCnt
    RevText$(RevCnt) = LEFT$(QPTrim$(UBSetUpRec(1).Revenues(RevCnt).RevName), 14)
    IF LEN(RevText$(RevCnt)) = 0 THEN
      NumOfRevs = RevCnt - 1
      EXIT FOR
    END IF
  NEXT
  
  TOWNNAME$ = UBSetUpRec(1).UTILNAME
  IF INSTR(TOWNNAME$, "CALDWELL") > 0 THEN
    CaldFlag = True
  END IF
  
  RecpPort = VAL(UBSetUpRec(1).RecpPort)
  
  IF RecpPort < 1 OR RecpPort > 2 THEN
    RecpPort = 1
  END IF
  
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
      QPrintRC CITY$, 8, 15, -1
      QPrintRC Deps$, 9, 15, -1
    END IF
    
    IF Frm(1).FldNo > CustAcctFld AND Frm(1).PrevFld = CustAcctFld THEN
      CUSTACCT& = QPValL(Form$(CustAcctFld, 0))
      GOSUB DepCheckPayList
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
      IF CUSTACCT& > 0 THEN
        ShowCustHistory CUSTACCT&
        Action = 2
      END IF
      
    CASE F7KEY  'Lookup Customer
      IF Frm(1).FldNo = 3 THEN  'if user is on the Customer field
        SaveScrn TempScrn()     'and F7key then do lookup routine
        MPaintBox 4, 5, 22, 75, 8
        LookUp CUSTACCT&, "Payment", 2, False, False
        RestScrn TempScrn()
        IF CUSTACCT& > 0 THEN   'if this is a valid customer
          GOSUB DepCheckPayList
          GOSUB DepGetCustInfo  'go get customer info
          Frm(1).FldNo = 4
          Action = 1
        ELSE
          GOSUB DepClearForm
          Frm(1).FldNo = 1
          Action = 1
        END IF
      END IF
      
    CASE F8KEY
      OPENDrawer RecpPort
      
    CASE F10Key 'Save
      GOSUB DepCheckPaymentInfo
      IF PaymentOKFlag THEN
        SELECT CASE AskSavePayment(UBSetUpRec(1).RECPDEFT)
        CASE 1
          GOSUB DepSaveTransaction              'do the save routine
          IF CaldFlag THEN
            GOSUB CWDepPrintReceipt
          ELSE
            GOSUB DepPrintReceipt
          END IF
          
          GOSUB DepClearForm
          GOSUB DClearCity
          Frm(1).FldNo = 1
          Action = 1
        CASE True
          ReceiptFlag = False
          GOSUB DepSaveTransaction              'do the save routine
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
  UBLog "OUT: Deposit Payments" + CrLf$
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
  
DepGetCustInfo:
  NumOfCustRecs& = FileSize("UBCUST.DAT") \ UBCustRecLen
  IF CUSTACCT& > NumOfCustRecs& OR CUSTACCT& = 0 THEN
    UBLog "ERROR: Invalid Account:" + STR$(CUSTACCT&)
    CUSTACCT& = 0
    OK = MsgBox%("UB.QSL", "BADACCTN")
    Frm(1).FldNo = 1
    GOSUB DepClearForm
    GOSUB DepSetOperInfo
    GOSUB DClearCity
    GOTO DepSkipCustInfo
  END IF
  
  IF IsDeleted(CUSTACCT&) THEN
    UBLog "ERROR: Deleted Account:" + STR$(CUSTACCT&)
    CUSTACCT& = 0
    OK = MsgBox%("UB.QSL", "DELACCTN")
    Frm(1).FldNo = 1
    GOSUB DepClearForm
    GOSUB DepSetOperInfo
    GOSUB DClearCity
    GOTO DepSkipCustInfo
  END IF
  
  GOSUB DepClearForm
  
  IF NOT EditFlag THEN
    UBLog "Entering deposit for Account:" + STR$(CUSTACCT&)
    CustFile = FREEFILE
    OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CustFile LEN = UBCustRecLen
    GET CustFile, CUSTACCT&, UBCustRec(1)
    CLOSE CustFile
    FOR Cnt = 1 TO NumOfRevs
      LSET Form$(AmtOweFlds(Cnt), 0) = STR$(UBCustRec(1).CurrRevAmts(Cnt))
    NEXT
    TotalBalance# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
    LSET Form$(CustAcctFld, 0) = STR$(CUSTACCT&)
    LSET Form$(4, 0) = UBCustRec(1).CUSTNAME
    LSET Form$(5, 0) = UBCustRec(1).ADDR1
    LSET Form$(6, 0) = UBCustRec(1).PAYCMNT
    LSET Form$(TotalFld, 0) = STR$(TotalBalance#)
    LSET Form$(AmtOwedFld, 0) = STR$(TotalBalance#)
    SaveField 0, Form$(), Fld(), BadField
  ELSE
    UBLog "Editing deposit for Account:" + STR$(CUSTACCT&)
    ListFile = FREEFILE
    OPEN PayFileName$ FOR RANDOM SHARED AS ListFile LEN = UBPayRecLen
    GET ListFile, PayListRec&, UBPaymentRec(1)
    CLOSE ListFile
    BCopy VARSEG(UBPaymentRec(1)), VARPTR(UBPaymentRec(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), UBPayRecLen, 0
    UnPackBuffer 0, 0, Form$(), Fld()
  END IF
  CUSTACCT& = QPValL(Form$(CustAcctFld, 0))
  LSET Deps$ = QPTrim$(FUsing$(STR$(UBCustRec(1).DepositAmt), Fmt$))
  LSET CITY$ = UBCustRec(1).CITY
  
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
  LSET Form$(1, 0) = FUsing$(STR$(OPERNUM), "##")
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
  SaveScrn TempScrn()
  DisplayUBScrn "UPDATDSK"
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(UBPaymentRec(1)), VARPTR(UBPaymentRec(1)), UBPayRecLen, 0
  FOR Cnt = 1 TO 15
    IF UBPaymentRec(1).PaidOwed(Cnt).AMTPD1 < -100000# THEN
      UBPaymentRec(1).PaidOwed(Cnt).AMTPD1 = 0
    END IF
    IF UBPaymentRec(1).PaidOwed(Cnt).AMTOWE1 < -100000# THEN
      UBPaymentRec(1).PaidOwed(Cnt).AMTOWE1 = 0
    END IF
  NEXT
  
  FOpenS PayFileName$, PHandle
  IF EditFlag THEN
    FPutRTA PHandle, UBPaymentRec(1), PayListRec&, UBPayRecLen
    EditFlag = False
  ELSE
    NumOfRecs& = (FLof(PHandle) \ UBPayRecLen) + 1
    FPutRTA PHandle, UBPaymentRec(1), NumOfRecs&, UBPayRecLen
    PayListCnt& = NumOfRecs&
    REDIM PRESERVE PayList(1 TO PayListCnt& + 1)    AS PayListType
    PayList(NumOfRecs&).CustRec = UBPaymentRec(1).CUSTACCT
    PayList(NumOfRecs&).ListRec = PayListCnt&
  END IF
  FClose PHandle
  UBLog "Updated Paylist for Account:" + STR$(UBPaymentRec(1).CUSTACCT)
  RestScrn TempScrn()
  DisplayUBScrn "UPDATEOK"
  WaitForAction
  RestScrn TempScrn()
  FirstTime = True
  RETURN
  
DClearCity:
  LSET CITY$ = ""
  LSET Deps$ = ""
  QPrintRC CITY$, 8, 15, -1
  QPrintRC Deps$, 9, 15, -1
  RETURN
  
CWDepPrintReceipt:
  SaveScrn TempScrn()
  PCnt = 0
  RHandle = FREEFILE
  OPEN PayRecpName$ FOR OUTPUT AS RHandle
  
  PRINT #RHandle, CHR$(27); "p"; CHR$(0); CHR$(25); CHR$(250); CHR$(7)
  PRINT #RHandle, TAB(7); TOWNNAME$; TAB(43); TOWNNAME$
  PRINT #RHandle, TAB(7); "UTILITY DEPOSIT PAYMENT"; TAB(43); "UTILITY DEPOSIT PAYMENT"
  PRINT #RHandle, TAB(7); "Date: "; PostDate$; TAB(43); "Date: "; PostDate$
  PRINT #RHandle,
  PRINT #RHandle, TAB(7); "CUSTOMER NAME & DESC."; TAB(43); "CUSTOMER NAME & DESC."
  PRINT #RHandle, TAB(7); UBPaymentRec(1).CUSTNAME; TAB(43); UBPaymentRec(1).CUSTNAME
  PRINT #RHandle, TAB(7); UBPaymentRec(1).CUSTADDR; TAB(43); UBPaymentRec(1).CUSTADDR
  PRINT #RHandle, TAB(7); UBPaymentRec(1).DESC; TAB(43); UBPaymentRec(1).DESC
  PRINT #RHandle,
  PRINT #RHandle,
  PRINT #RHandle, USING "Total Paid:   $$####,#.##"; TAB(7); UBPaymentRec(1).AMTPAID; TAB(43); UBPaymentRec(1).AMTPAID
  PRINT #RHandle, USING "Change Due:   $$####,#.##"; TAB(7); UBPaymentRec(1).CHANGE; TAB(43); UBPaymentRec(1).CHANGE
  PRINT #RHandle,
  '14 to here
  FOR RevCnt = 1 TO NumOfRevs
    IF UBPaymentRec(1).PaidOwed(RevCnt).AMTPD1 <> 0 THEN
      LSET RecpRev$ = RevText$(RevCnt)
      PRINT #RHandle, USING RecpRev$ + "$$#####.##"; TAB(7); UBPaymentRec(1).PaidOwed(RevCnt).AMTPD1; TAB(43); UBPaymentRec(1).PaidOwed(RevCnt).AMTPD1
      PCnt = PCnt + 1
    END IF
  NEXT
  IF PCnt < 8 THEN
    FOR ZZ = PCnt TO 8
      PRINT #RHandle,
    NEXT
  END IF
  PRINT #RHandle, TAB(7); "Operator: "; OPERNUM; TAB(43); "Operator: "; OPERNUM
  RecptNo& = FileSize(PayFileName$) / UBPayRecLen
  PRINT #RHandle, USING "Receipt: ######"; TAB(7); RecptNo&; TAB(43); RecptNo&
  PRINT #RHandle,
  PRINT #RHandle, TAB(7); " T H A N K   Y O U !"; TAB(43); " T H A N K   Y O U !"
  PRINT #RHandle,
  
  CLOSE RHandle
  
  GOSUB ReceiptCommon
  
  RETURN
  
  
DepPrintReceipt:
  SaveScrn TempScrn()
  RecptNo& = FileSize(PayFileName$) / UBPayRecLen
  
  RHandle = FREEFILE
  OPEN PayRecpName$ FOR OUTPUT AS RHandle
  'PRINT #RHandle, CHR$(27); "p"; CHR$(0); CHR$(25); CHR$(250)
  PRINT #RHandle, 'CHR$(7)
  PRINT #RHandle,
  PRINT #RHandle, TOWNNAME$
  PRINT #RHandle, "UTILITY DEPOSIT PAYMENT"
  PRINT #RHandle, "Date: "; PostDate$
  PRINT #RHandle,
  PRINT #RHandle, "CUSTOMER NAME & DESC. OF PAYMENT"
  PRINT #RHandle, UBPaymentRec(1).CUSTNAME
  PRINT #RHandle, UBPaymentRec(1).CUSTADDR
  PRINT #RHandle, UBPaymentRec(1).DESC
  PRINT #RHandle,
  PRINT #RHandle,
  PRINT #RHandle, "Total Paid: "; USING "$$####,#.##"; UBPaymentRec(1).AMTPAID
  PRINT #RHandle, "Change Due: "; USING "$$####,#.##"; UBPaymentRec(1).CHANGE
  PRINT #RHandle,
  FOR RevCnt = 1 TO NumOfRevs
    IF UBPaymentRec(1).PaidOwed(RevCnt).AMTPD1 <> 0 THEN
      LSET RecpRev$ = RevText$(RevCnt)
      PRINT #RHandle, RecpRev$; USING "$$#####.##"; UBPaymentRec(1).PaidOwed(RevCnt).AMTPD1
    END IF
  NEXT
  PRINT #RHandle,
  PRINT #RHandle, "Operator: "; OPERNUM
  
  PRINT #RHandle, "Receipt#: "; USING "######"; RecptNo&
  PRINT #RHandle,
  PRINT #RHandle, "       T H A N K   Y O U !"
  PRINT #RHandle,
  PRINT #RHandle,
  PRINT #RHandle,
  PRINT #RHandle,
  PRINT #RHandle,
  PRINT #RHandle,
  PRINT #RHandle, 'CutPaper$
  CLOSE RHandle
  GOSUB ReceiptCommon
  
  RETURN
  
ReceiptCommon:
  'Shell$ = "type " + PayRecpName$ + " > com2:"
  'SHELL Shell$
  
  PrintRptFile Header$, PayRecpName$, RecpPort, RetCode%, 5
  
  KillFile PayRecpName$
  RestScrn TempScrn()
  
  RETURN
  
DepLoadPayList:
  
  PHandle = FREEFILE
  OPEN PayFileName$ FOR RANDOM SHARED AS PHandle LEN = UBPayRecLen
  NumOfRecs& = LOF(PHandle) \ UBPayRecLen
  IF NumOfRecs& > 0 THEN
    REDIM PayList(1 TO NumOfRecs&)    AS PayListType
    FOR Cnt& = 1 TO NumOfRecs&
      GET #PHandle, Cnt&, UBPaymentRec(1)
      PayList(Cnt&).CustRec = UBPaymentRec(1).CUSTACCT
      PayList(Cnt&).ListRec = Cnt&
    NEXT
  END IF
  CLOSE PHandle
  PayListCnt& = NumOfRecs&
  RETURN
  
DepCheckPayList:
  EditFlag = False
  IF PayListCnt& > 0 THEN
    FOR Cnt = 1 TO PayListCnt&
      IF PayList(Cnt).CustRec = CUSTACCT& THEN
        PayListRec& = PayList(Cnt).ListRec
        EditFlag = True
        EXIT FOR
      END IF
    NEXT
  END IF
  RETURN
  
DepCheckPaymentInfo:
  PaymentOKFlag = True
  'RETURN
  TAmtRecv# = VAL(Form$(AmtRecvFld, 0))
  TAmtPaid# = VAL(Form$(AmtPaidFld, 0))
  ChangeAmt# = VAL(Form$(ChangeFld, 0))
  
  IF TAmtRecv# = Round#(TAmtPaid# + ChangeAmt#) AND TAmtPaid# > 0 THEN
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

SUB AddEditOperator
  
  AuthorizeOperator oktoadd
  
  IF oktoadd = 0 THEN EXIT SUB
  
  LibName$ = "UB"
  ScrnName$ = "UBNEWOP"
  BlockClear
  DisplayUBScrn ScrnName$
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  Frm(1).StayOnField = True
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  Action = 1
  REDIM CMOperRec(1) AS CMOperRecType
  CMOperRecLen = LEN(CMOperRec(1))
  CMFile = FREEFILE
  OPEN "CMOPER.DAT" FOR RANDOM SHARED AS CMFile LEN = CMOperRecLen
  NumRecs = LOF(CMFile) \ CMOperRecLen
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF Frm(1).PrevFld = 2 AND edit <> True THEN
      FOR Cnt = 1 TO NumRecs
        GET CMFile, Cnt, CMOperRec(1)
        IF VAL(Form$(1, 0)) = CMOperRec(1).OperatorNumber AND VAL(Form$(1, 0)) > 0 THEN
          Form$(2, 0) = CMOperRec(1).OperatorName
          Form$(3, 0) = CMOperRec(1).OperatorPassword
          Fld(1).Protected = True
          RecNo = Cnt
          edit = True
          Action = 1
          EXIT FOR
        END IF
      NEXT Cnt
    END IF
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      
      CMOperRec(1).OperatorNumber = VAL(Form$(1, 0))
      CMOperRec(1).OperatorName = Form$(2, 0)
      CMOperRec(1).OperatorPassword = Form$(3, 0)
      IF RecNo = 0 THEN
        PUT CMFile, LOF(CMFile) \ CMOperRecLen + 1, CMOperRec(1)
      ELSE
        PUT CMFile, RecNo, CMOperRec(1)
      END IF
      
      CLOSE CMFile
      Done = True
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  
END SUB

SUB AddEditPayment (OPERNUM, PostDate$)
  UBLog " IN: Bill Payments,  OPER:" + STR$(OPERNUM)

  CutPaper$ = CHR$(29) + CHR$(86) + CHR$(66) + CHR$(64)

  CITY$ = SPACE$(20)
  Deps$ = SPACE$(12)
  Fmt$ = "$$####.##"
  RecpRev$ = SPACE$(15)
  
  'look into keeping date on payments edited on a different day
  SHARED Choice$()
  
  REDIM TempScrn(0)
  REDIM UBCustRec(1)        AS NewUBCustRecType
  REDIM UBPaymentRec(1)     AS UBPaymentRecType
  REDIM PayList(1 TO 1)     AS PayListType
  
  UBCustRecLen = LEN(UBCustRec(1))
  UBPayRecLen = LEN(UBPaymentRec(1))
  
  GOSUB LoadPayList
  
  NumOfRevs = MaxRevsCnt
  
  REDIM RevText$(1 TO MaxRevsCnt)
  REDIM UBSetUpRec(1) AS UBSetupRecType
  
  REDIM DistArray(1 TO MaxRevsCnt) AS DistArrayType
  
  UBSetUpLen = LEN(UBSetUpRec(1))
  FGetAH "UBSETUP.DAT", UBSetUpRec(1), UBSetUpLen, 1            'load it
  
  RecpPort = VAL(UBSetUpRec(1).RecpPort)
  IF RecpPort < 1 OR RecpPort > 2 THEN
    RecpPort = 1
  END IF
  
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
  IF INSTR(TOWNNAME$, "CALDWELL") > 0 THEN
    CaldFlag = True
  END IF
  
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
  'REDIM PrevAmtOwe(1 TO NumOfRevs) AS DOUBLE
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
  StatFld = FldNum%("STATUS", Fld())
  TaxExemptFld = FldNum%("TAXEXPT", Fld())
  
  '--define the multi-choice fields
  
  REDIM Choice$(0 TO 3, 0 TO 0)
  
  Choice$(0, 0) = QPTrim$(STR$(TenderFld))
  'Choice$(1, 0) = "Cash"
  'Choice$(2, 0) = "Check"
  'Choice$(3, 0) = "Cash & Check"
  
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
      QPrintRC CITY$, 8, 15, -1
      QPrintRC Deps$, 9, 15, -1
    END IF
    
    IF DistFlag AND NOT PrevFlag THEN
      TempAmtRecv# = Value#(Form$(AmtRecvFld, 0), ECode)
      GOSUB AutoDistribute:
    END IF
    DistFlag = False
    
    IF Frm(1).FldNo > CustAcctFld AND Frm(1).PrevFld = CustAcctFld THEN
      PrevFlag = False
      CUSTACCT& = QPValL(Form$(CustAcctFld, 0))
      GOSUB CheckPayList
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
      IF CUSTACCT& > 0 THEN
        SaveScrn TempScrn()
        CustomerInquiry CUSTACCT&
        RestScrn TempScrn()
        Action = 2
      END IF
    CASE -88    'Shift-F5 Previous cash payment
      IF Frm(1).FldNo = TenderFld THEN          'if user is on tender field
        PrevFlag = True
        GOSUB DoCashPayment     'and F5key then go do the
      END IF    'cash payment routine
      DistFlag = True
    CASE F5KEY
      IF Frm(1).FldNo = TenderFld THEN          'if user is on tender field
        PrevFlag = False
        GOSUB DoCashPayment     'and F5key then go do the
      END IF    'cash payment routine
      DistFlag = True
    CASE -89    'ShiftF6 Previous Check payment
      IF Frm(1).FldNo = TenderFld THEN          'if user is on tender field
        PrevFlag = True
        GOSUB DoCheckPayment    'and F6key then go do the
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
        LookUp CUSTACCT&, "Payment", 2, False, False
        RestScrn TempScrn()
        IF CUSTACCT& > 0 THEN   'if this is a valid customer
          GOSUB CheckPayList
          GOSUB GetCustInfo     'go get customer info
          Frm(1).FldNo = 4
          Action = 1
        ELSE
          GOSUB ClearForm
          Frm(1).FldNo = 1
          Action = 1
        END IF
      END IF
      
      '    CASE -92    'Shift F9
      '      TempAmtRecv# = Value#(Form$(AmtRecvFld, 0), ECode)
      '      IF TempAmtRecv# > 0 THEN
      '        GOSUB AutoDistributeOLD
      '      END IF
      
    CASE F8KEY
      OPENDrawer RecpPort
      
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
          SaveScrn TempScrn()
          GOSUB SaveTransaction 'do the save routine
          IF CaldFlag THEN
            GOSUB CWPrintReceipt
          ELSE
            GOSUB PrintReceipt
          END IF
          RestScrn TempScrn()
          DisplayUBScrn "PRESSKEY"
          WaitForAction
          RestScrn TempScrn()
          GOSUB ClearForm
          GOSUB ClearCity
          Frm(1).FldNo = 1
          Action = 1
        CASE True
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
    CASE IS <> 0
      'STOP
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
  UBLog "OUT: Bill Payments" + CrLf$
  HideCursor
  
  EXIT SUB
  
GetCustInfo:
  CashFlag = False
  NumOfCustRecs& = FileSize("UBCUST.DAT") \ UBCustRecLen
  IF CUSTACCT& > NumOfCustRecs& OR CUSTACCT& = 0 THEN
    UBLog "ERROR: Invalid Account:" + STR$(CUSTACCT&)
    CUSTACCT& = 0
    OK = MsgBox%("UB.QSL", "BADACCTN")
    Frm(1).FldNo = 1
    GOSUB ClearForm
    GOSUB SetOperInfo
    GOSUB ClearCity
    GOTO SkipCustInfo
  END IF
  IF IsDeleted(CUSTACCT&) THEN
    UBLog "ERROR: Deleted Account:" + STR$(CUSTACCT&)
    CUSTACCT& = 0
    OK = MsgBox%("UB.QSL", "DELACCTN")
    Frm(1).FldNo = 1
    GOSUB ClearForm
    GOSUB SetOperInfo
    GOSUB ClearCity
    GOTO SkipCustInfo
  END IF
  
  GOSUB ClearForm
  CustFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CustFile LEN = UBCustRecLen
  GET CustFile, CUSTACCT&, UBCustRec(1)
  'FOR Cnt = 1 TO 15
  '  UBCustRec(1).CurrRevAmts(Cnt) = 0
  'NEXT
  'PUT CustFile, CUSTACCT&, UBCustRec(1)
  CLOSE CustFile
  
  IF NOT EditFlag THEN
    UBLog "Entering payment for Account:" + STR$(CUSTACCT&)
    IF UBCustRec(1).CASHONLY = "Y" THEN
      CashFlag = True
    END IF
    FOR Cnt = 1 TO NumOfRevs
      LSET Form$(AmtOweFlds(Cnt), 0) = STR$(UBCustRec(1).CurrRevAmts(Cnt))
      'PrevAmtOwe(Cnt) = UBCustRec(1).PrevRevAmts(Cnt)
    NEXT
    TotalBalance# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
    LSET Form$(CustAcctFld, 0) = STR$(CUSTACCT&)
    LSET Form$(4, 0) = UBCustRec(1).CUSTNAME
    LSET Form$(5, 0) = UBCustRec(1).ADDR1
    LSET Form$(6, 0) = UBCustRec(1).PAYCMNT
    LSET Form$(TaxExemptFld, 0) = UBCustRec(1).TAXEXPT
    LSET Form$(TotalFld, 0) = STR$(TotalBalance#)
    LSET Form$(AmtOwedFld, 0) = STR$(TotalBalance#)
    LSET Form$(StatFld, 0) = UBCustRec(1).Status
    SaveField 0, Form$(), Fld(), BadField
  ELSE
    UBLog "Editing payment for Account:" + STR$(CUSTACCT&)
    ListFile = FREEFILE
    OPEN PayFileName$ FOR RANDOM SHARED AS ListFile LEN = UBPayRecLen
    GET ListFile, PayListRec&, UBPaymentRec(1)
    CLOSE ListFile
    BCopy VARSEG(UBPaymentRec(1)), VARPTR(UBPaymentRec(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), UBPayRecLen, 0
    UnPackBuffer 0, 0, Form$(), Fld()
  END IF
  CUSTACCT& = QPValL(Form$(CustAcctFld, 0))
  
  LSET Deps$ = QPTrim$(FUsing$(STR$(UBCustRec(1).DepositAmt), Fmt$))
  LSET CITY$ = UBCustRec(1).CITY
  
  MPaintBox 22, 35, 22, 41, 115
  FirstTime = True
  
SkipCustInfo:
  GOSUB SetTenderFlds
  Action = 1
RETURN
  
ShowRevSources:
  FOR RevCnt = 1 TO NumOfRevs
    QPrintRC RevText$(RevCnt), RevCnt + 4, 40, -1
  NEXT
RETURN
  
SetOperInfo:
  LSET Form$(1, 0) = FUsing$(STR$(OPERNUM), "##")
  LSET Form$(2, 0) = PostDate$
  Action = 2
RETURN
  
ClearForm:
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""       '--Clear all fields
  NEXT
  BeenEditedFlag = False        'clear the edited flag
RETURN
  
DoCashPayment:
  'IF PrevFlag THEN
  '  ThisAmt# = 0
  '  FOR Cnt = 1 TO NumOfRevs
  '    ThisAmt# = Round#(ThisAmt# + PrevAmtOwe(Cnt))
  '  NEXT
  '  IF ThisAmt# > 0 THEN
  '    LSET Form$(TenderFld, 0) = Choice$(1, 0)
  '    LSET Form$(ChkAmtFld, 0) = "0"
  '    LSET Form$(CashAmtFld, 0) = QPTrim$(STR$(ThisAmt#))
  '    GOSUB PaymentCommon
  '  END IF
  'ELSE
  LSET Form$(TenderFld, 0) = Choice$(1, 0)
  LSET Form$(ChkAmtFld, 0) = "0"
  LSET Form$(CashAmtFld, 0) = Form$(AmtOwedFld, 0)
  GOSUB PaymentCommon
  'END IF
RETURN
  
DoCheckPayment:
  '  IF PrevFlag THEN
  '    ThisAmt# = 0
  '    FOR Cnt = 1 TO NumOfRevs
  '      ThisAmt# = Round#(ThisAmt# + PrevAmtOwe(Cnt))
  '    NEXT
  '    IF ThisAmt# > 0 THEN
  '      LSET Form$(TenderFld, 0) = Choice$(2, 0)
  '      LSET Form$(ChkAmtFld, 0) = QPTrim$(STR$(ThisAmt#))
  '      LSET Form$(CashAmtFld, 0) = "0"
  '      GOSUB PaymentCommon
  '    END IF
  '  ELSE
  LSET Form$(TenderFld, 0) = Choice$(2, 0)
  LSET Form$(ChkAmtFld, 0) = Form$(AmtOwedFld, 0)
  LSET Form$(CashAmtFld, 0) = "0"
  GOSUB PaymentCommon
  '  END IF
RETURN
  
PaymentCommon:
  SaveField TenderFld, Form$(), Fld(), BadField
  SaveField ChkAmtFld, Form$(), Fld(), BadField
  SaveField CashAmtFld, Form$(), Fld(), BadField
  
  LSET Form$(ChangeFld, 0) = "0"
  SaveField ChangeFld, Form$(), Fld(), BadField
  
  FOR Cnt = 1 TO NumOfRevs
    'IF PrevFlag THEN
    '  LSET Form$(AmtOweFlds(Cnt) + 1, 0) = QPTrim$(STR$(PrevAmtOwe(Cnt)))
    'ELSE
    LSET Form$(AmtOweFlds(Cnt) + 1, 0) = Form$(AmtOweFlds(Cnt), 0)
    'END IF
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
  
  'GOSUB AutoDistribute:
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
  
SaveTransaction:
  
  'DisplayUBScrn "UPDATDSK"
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(UBPaymentRec(1)), VARPTR(UBPaymentRec(1)), UBPayRecLen, 0
  FOR Cnt = 1 TO 15
    IF UBPaymentRec(1).PaidOwed(Cnt).AMTPD1 < -100000# THEN
      UBPaymentRec(1).PaidOwed(Cnt).AMTPD1 = 0
    END IF
    IF UBPaymentRec(1).PaidOwed(Cnt).AMTOWE1 < -100000# THEN
      UBPaymentRec(1).PaidOwed(Cnt).AMTOWE1 = 0
    END IF
  NEXT
  ListFile = FREEFILE
  OPEN PayFileName$ FOR RANDOM SHARED AS ListFile LEN = UBPayRecLen
  IF EditFlag THEN
    PUT #ListFile, PayListRec&, UBPaymentRec(1)
    EditFlag = False
  ELSE
    NumOfRecs& = (LOF(ListFile) \ UBPayRecLen) + 1
    PayListRec& = NumOfRecs&
    PUT #ListFile, PayListRec&, UBPaymentRec(1)
  END IF
  UBLog "Updated Paylist for Account:" + STR$(UBPaymentRec(1).CUSTACCT)
  CLOSE ListFile
  
  GOSUB LoadPayList
  
  '  RestScrn TempScrn()
  FirstTime = True
RETURN
  
PrintReceipt:
  
  'SaveScrn TempScrn()
  
  ListFile = FREEFILE
  OPEN PayFileName$ FOR RANDOM SHARED AS ListFile LEN = UBPayRecLen
  RecptNum& = LOF(ListFile) / UBPayRecLen
  GET #ListFile, PayListRec&, UBPaymentRec(1)
  CLOSE
  
  RHandle = FREEFILE
  OPEN PayRecpName$ FOR OUTPUT AS RHandle
  'PRINT #RHandle, CHR$(27); "p"; CHR$(0); CHR$(25); CHR$(250)
  PRINT #RHandle, 'CHR$(7)
  PRINT #RHandle, TOWNNAME$
  PRINT #RHandle, "UTILITY PAYMENT"
  PRINT #RHandle, "Date: "; PostDate$
  PRINT #RHandle,
  PRINT #RHandle, "CUSTOMER NAME & DESC. OF PAYMENT"
  PRINT #RHandle, UBPaymentRec(1).CUSTNAME
  PRINT #RHandle, UBPaymentRec(1).CUSTADDR
  PRINT #RHandle, UBPaymentRec(1).DESC
  PRINT #RHandle, "Acct. No. "; UBPaymentRec(1).CUSTACCT
  PRINT #RHandle,
  PRINT #RHandle, "Total Owed: "; USING "$$####,#.##"; UBPaymentRec(1).TOTOWED
  PRINT #RHandle, "Total Paid: "; USING "$$####,#.##"; UBPaymentRec(1).AMTPAID
  PRINT #RHandle, "Change Due: "; USING "$$####,#.##"; UBPaymentRec(1).CHANGE
  PRINT #RHandle, "   Balance: "; USING "$$####,#.##"; UBPaymentRec(1).TOTOWED - UBPaymentRec(1).AMTPAID
  PRINT #RHandle,
  FOR RevCnt = 1 TO NumOfRevs
    IF UBPaymentRec(1).PaidOwed(RevCnt).AMTPD1 <> 0 OR UBPaymentRec(1).PaidOwed(RevCnt).AMTOWE1 <> 0 THEN
      LSET RecpRev$ = RevText$(RevCnt)
      'PRINT #RHandle, RecpRev$; USING "$$####,#.##"; UBPaymentRec(1).PaidOwed(RevCnt).AmtOwe1; UBPaymentRec(1).PaidOwed(RevCnt).AmtPd1
      PRINT #RHandle, RecpRev$; USING "$$#####.##"; UBPaymentRec(1).PaidOwed(RevCnt).AMTPD1
    END IF
  NEXT
  PRINT #RHandle,
  PRINT #RHandle, "Operator: "; OPERNUM
  PRINT #RHandle, "Receipt#: "; USING "######"; RecptNum&
  PRINT #RHandle,
  PRINT #RHandle, "       T H A N K   Y O U !"
  PRINT #RHandle,
  PRINT #RHandle,
  PRINT #RHandle,
  PRINT #RHandle,
  PRINT #RHandle,
  PRINT #RHandle, 'CutPaper$
  CLOSE RHandle
  
  'Shell$ = "type " + PayRecpName$ + " > com2:"
  'SHELL Shell$
  
  PrintRptFile Header$, PayRecpName$, RecpPort, RetCode%, 5
  
  KillFile PayRecpName$
  
RETURN
  
CWPrintReceipt:
  PCnt = 0
  ListFile = FREEFILE
  OPEN PayFileName$ FOR RANDOM SHARED AS ListFile LEN = UBPayRecLen
  GET #ListFile, PayListRec&, UBPaymentRec(1)
  CLOSE
  
  RHandle = FREEFILE
  OPEN PayRecpName$ FOR OUTPUT AS RHandle
  PRINT #RHandle, CHR$(27); "p"; CHR$(0); CHR$(25); CHR$(250); CHR$(7)
  PRINT #RHandle, TAB(7); TOWNNAME$; TAB(43); TOWNNAME$
  PRINT #RHandle, TAB(7); "UTILITY PAYMENT"; TAB(43); "UTILITY PAYMENT"
  PRINT #RHandle, TAB(7); "Date: "; PostDate$; TAB(43); "Date: "; PostDate$
  PRINT #RHandle,
  PRINT #RHandle, TAB(7); "CUSTOMER NAME & DESC."; TAB(43); "CUSTOMER NAME & DESC."
  PRINT #RHandle, TAB(7); UBPaymentRec(1).CUSTNAME; TAB(43); UBPaymentRec(1).CUSTNAME
  PRINT #RHandle, TAB(7); UBPaymentRec(1).CUSTADDR; TAB(43); UBPaymentRec(1).CUSTADDR
  PRINT #RHandle, TAB(7); UBPaymentRec(1).DESC; TAB(43); UBPaymentRec(1).DESC
  PRINT #RHandle, TAB(7); "Acct. No. "; UBPaymentRec(1).CUSTACCT; TAB(43); "Acct. No. "; UBPaymentRec(1).CUSTACCT
  PRINT #RHandle,
  PRINT #RHandle, USING "Total Owed:   $$####,#.##"; TAB(7); UBPaymentRec(1).TOTOWED; TAB(43); UBPaymentRec(1).TOTOWED
  PRINT #RHandle, USING "Total Paid:   $$####,#.##"; TAB(7); UBPaymentRec(1).AMTPAID; TAB(43); UBPaymentRec(1).AMTPAID
  PRINT #RHandle, USING "Change Due:   $$####,#.##"; TAB(7); UBPaymentRec(1).CHANGE; TAB(43); UBPaymentRec(1).CHANGE
  PRINT #RHandle, USING "   Balance:   $$####,#.##"; TAB(7); UBPaymentRec(1).TOTOWED - UBPaymentRec(1).AMTPAID; TAB(43); UBPaymentRec(1).TOTOWED - UBPaymentRec(1).AMTPAID
  PRINT #RHandle,
  '16 to here
  FOR RevCnt = 1 TO NumOfRevs
    IF UBPaymentRec(1).PaidOwed(RevCnt).AMTPD1 <> 0 OR UBPaymentRec(1).PaidOwed(RevCnt).AMTOWE1 <> 0 THEN
      LSET RecpRev$ = RevText$(RevCnt)
      PRINT #RHandle, USING RecpRev$ + "$$#####.##"; TAB(7); UBPaymentRec(1).PaidOwed(RevCnt).AMTPD1; TAB(43); UBPaymentRec(1).PaidOwed(RevCnt).AMTPD1
      PCnt = PCnt + 1
    END IF
  NEXT
  IF PCnt < 6 THEN
    FOR ZZ = PCnt TO 6
      PRINT #RHandle,
    NEXT
  END IF
  
  PRINT #RHandle, TAB(7); "Operator: "; OPERNUM; TAB(43); "Operator: "; OPERNUM
  RecptNo& = FileSize(PayFileName$) \ UBPayRecLen
  PRINT #RHandle, TAB(7); USING "Receipt:  ######"; RecptNo&; TAB(43); RecptNo&
  PRINT #RHandle,
  PRINT #RHandle, TAB(7); " T H A N K   Y O U !"; TAB(43); " T H A N K   Y O U !"
  PRINT #RHandle,
  CLOSE RHandle
  
  'Shell$ = "type " + PayRecpName$ + " > com2:"
  'SHELL Shell$
  
  PrintRptFile Header$, PayRecpName$, RecpPort, RetCode%, 5
  
  KillFile PayRecpName$
RETURN
  
ClearCity:
  LSET CITY$ = ""
  LSET Deps$ = ""
  QPrintRC CITY$, 8, 15, -1
  QPrintRC Deps$, 9, 15, -1
RETURN
  
LoadPayList:
  
  Oper$ = QPTrim$(STR$(OPERNUM))
  
  PayFileName$ = "UBPAY" + Oper$ + ".DAT"
  PayRecpName$ = "UBRCP" + Oper$ + ".RPT"
  
  ListFile = FREEFILE
  OPEN PayFileName$ FOR RANDOM SHARED AS ListFile LEN = UBPayRecLen
  NumOfRecs& = LOF(ListFile) \ UBPayRecLen
  IF NumOfRecs& > 0 THEN
    REDIM PayList(1 TO NumOfRecs&)    AS PayListType
    FOR Cnt& = 1 TO NumOfRecs&
      GET #ListFile, Cnt&, UBPaymentRec(1)
      PayList(Cnt&).CustRec = UBPaymentRec(1).CUSTACCT
      PCustAcct = UBPaymentRec(1).CUSTACCT
      PayList(Cnt&).ListRec = Cnt&
    NEXT
  END IF
  CLOSE ListFile
  PayListCnt& = NumOfRecs&
RETURN
  
CheckPayList:
  EditFlag = False
  IF PayListCnt& > 0 THEN
    FOR Cnt = 1 TO PayListCnt&
      IF PayList(Cnt).CustRec = CUSTACCT& THEN
        PayListRec& = PayList(Cnt).ListRec
        EditFlag = True
        EXIT FOR
      END IF
    NEXT
  END IF
RETURN
  
CheckPaymentInfo:
  PaymentOKFlag = True
  
  TAmtRecv# = Value(Form$(AmtRecvFld, 0), ECode)
  TAmtPaid# = Value(Form$(AmtPaidFld, 0), ECode)
  
  ChangeAmt# = Value(Form$(ChangeFld, 0), ECode)
  
  IF TAmtPaid# = 0 THEN
    OK = MsgBox%("UB.QSL", "BADPYTOT")          'show bad scrn
    Action = 2
    PaymentOKFlag = False
    Frm(1).FldNo = Frm(1).PrevFld
    GOTO BadPayment
  END IF
  
  IF TAmtRecv# = Round#(TAmtPaid# + ChangeAmt#) AND TAmtRecv# > 0 AND ChangeAmt# >= 0 THEN
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

FUNCTION AskSavePayment% (DefaultFlag$)
  
  REDIM TempScrn(0)
  SaveScrn TempScrn()
  
  LibName$ = "UB"
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
  
  DisplayUBScrn ScrnName$
  
  ShowCursor
  
  Action = 1
  FirstTime = True
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      FirstTime = NOT FirstTime
      LSET Form$(1, 0) = DefaultFlag$
      Action = 1
    END IF
    
    SELECT CASE Frm(1).KeyCode
    CASE EscKey
      AskSavePayment% = False
      EXIT DO
    CASE F10Key
      Receipt$ = Form$(1, 0)
      IF Receipt$ = "Y" THEN
        AskSavePayment% = 1
      ELSE
        AskSavePayment% = True
      END IF
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

SUB AuthorizeOperator (oktoadd)
  
  LibName$ = "UB"
  ScrnName$ = "OPERAUTH"
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  Frm(1).StayOnField = True
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  BlockClear
  DisplayUBScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      IF Form$(1, 0) = "CASHMGMT" THEN
        oktoadd = 1
      ELSE
        oktoadd = 0
      END IF
      Done = True
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  
END SUB

SUB DeleteDepTrans (OPERNUM)
  
  PayFileName$ = "UBDEP" + QPTrim$(STR$(OPERNUM)) + ".DAT"
  TPayFileName$ = "UBDEP" + QPTrim$(STR$(OPERNUM)) + ".$$$"
  
  REDIM UBPaymentRec(1)   AS UBPaymentRecType
  
  UBPayRecLen = LEN(UBPaymentRec(1))
  
  NumPayRecs = FileSize&(PayFileName$) \ UBPayRecLen
  
  TotalRecs& = NumPayRecs
  
  REDIM TransList(1 TO TotalRecs&)    AS FLen2
  REDIM Picked(1 TO TotalRecs&) AS INTEGER
  
  WhatCnt& = 1
  
  FOpenS PayFileName$, PHandle
  NumOfRecs& = FLof(PHandle) \ UBPayRecLen
  IF NumOfRecs& > 0 THEN
    FOR Cnt& = 1 TO NumOfRecs&
      FGetRTA PHandle, UBPaymentRec(1), Cnt&, UBPayRecLen
      LSET TransList(WhatCnt&).V = UBPaymentRec(1).CUSTNAME
      MID$(TransList(WhatCnt&).V, 27) = "Deposit"
      MID$(TransList(WhatCnt&).V, 45) = FUsing$(STR$(UBPaymentRec(1).AMTPAID), "#####.##")
      'FUsing$ (Number$, Image$)
      WhatCnt& = WhatCnt& + 1
    NEXT
  END IF
  FClose PHandle
  PayListCnt& = NumOfRecs&
  
  MaxLen = 64
  Col = ((80 - MaxLen) \ 2) - 2
  
  BlockClear
  
  DelList TransList(), Picked(), NPicked
  
  IF NPicked > 0 THEN
    PFile = FREEFILE
    OPEN PayFileName$ FOR RANDOM SHARED AS PFile LEN = UBPayRecLen
    TFile = FREEFILE
    OPEN TPayFileName$ FOR RANDOM SHARED AS TFile LEN = UBPayRecLen
    FOR Cnt = 1 TO TotalRecs&
      GET PFile, Cnt, UBPaymentRec(1)
      SELECT CASE RIGHT$(TransList(Cnt).V, 1)
      CASE IS <> "û"
        PUT TFile, , UBPaymentRec(1)
      END SELECT
    NEXT
    CLOSE
    KillFile PayFileName$
    NAME TPayFileName$ AS PayFileName$
  END IF
  IF FileSize(PayFileName$) = 0 THEN
    KillFile PayFileName$
  END IF
END SUB

SUB DeletePayTrans (OPERNUM)
  
  PayFileName$ = "UBPAY" + QPTrim$(STR$(OPERNUM)) + ".DAT"
  TPayFileName$ = "UBPAY" + QPTrim$(STR$(OPERNUM)) + ".$$$"
  
  Fmt$ = "#######.##"
  
  REDIM UBPaymentRec(1)   AS UBPaymentRecType
  
  UBPayRecLen = LEN(UBPaymentRec(1))
  
  NumPayRecs = FileSize&(PayFileName$) \ UBPayRecLen
  
  TotalRecs& = NumPayRecs
  
  REDIM TransList(1 TO TotalRecs&)    AS FLen2
  REDIM Picked(1 TO TotalRecs&) AS INTEGER
  
  WhatCnt& = 1
  
  FOpenS PayFileName$, PHandle
  NumOfRecs& = FLof(PHandle) \ UBPayRecLen
  IF NumOfRecs& > 0 THEN
    FOR Cnt& = 1 TO NumOfRecs&
      FGetRTA PHandle, UBPaymentRec(1), Cnt&, UBPayRecLen
      LSET TransList(WhatCnt&).V = UBPaymentRec(1).CUSTNAME
      MID$(TransList(WhatCnt&).V, 28) = "Payment"
      MID$(TransList(WhatCnt&).V, 44) = FUsing$(STR$(UBPaymentRec(1).AMTPAID), Fmt$)
      WhatCnt& = WhatCnt& + 1
    NEXT
  END IF
  FClose PHandle
  PayListCnt& = NumOfRecs&
  
  MaxLen = 64
  Col = ((80 - MaxLen) \ 2) - 2
  
  BlockClear
  
  DelList TransList(), Picked(), NPicked
  IF NPicked > 0 THEN
    PFile = FREEFILE
    OPEN PayFileName$ FOR RANDOM SHARED AS PFile LEN = UBPayRecLen
    TFile = FREEFILE
    OPEN TPayFileName$ FOR RANDOM SHARED AS TFile LEN = UBPayRecLen
    FOR Cnt = 1 TO TotalRecs&
      GET PFile, Cnt, UBPaymentRec(1)
      SELECT CASE RIGHT$(TransList(Cnt).V, 1)
      CASE IS <> "û"
        PUT TFile, , UBPaymentRec(1)
      END SELECT
    NEXT
    CLOSE

    KillFile PayFileName$
    NAME TPayFileName$ AS PayFileName$
  END IF
  IF FileSize(PayFileName$) = 0 THEN
    KillFile PayFileName$
  END IF
  
END SUB

SUB DeleteTransMenu (OPERNUM)
  
  PayFileName$ = "UBPAY" + QPTrim$(STR$(OPERNUM)) + ".DAT"
  DepFileName$ = "UBDEP" + QPTrim$(STR$(OPERNUM)) + ".DAT"
  
  IF Exist(PayFileName$) AND FileSize(PayFileName$) > 0 THEN
    PayFlag = True
  END IF
  
  IF Exist(DepFileName$) AND FileSize(DepFileName$) THEN
    DepFlag = True
  END IF
  
  IF DepFlag AND PayFlag THEN
    DoMenu = True
  ELSEIF DepFlag THEN
    GOTO DoDelDeposit
  ELSEIF PayFlag THEN
    GOTO DoDelPayment
  END IF
  
  IF DoMenu THEN
    REDIM MChoice$(1 TO 2)
    
    MChoice$(1) = "Delete PAYMENT Transactions"
    MChoice$(2) = "Delete DEPOSIT Transactions"
    
    MaxLen = 0  'Set menu width to zero
    BoxBot = 17 'limit the box length to go no lower than line 20
    Action = 0  '0 means stay in the menu until they select something
    Choice = 1  'Pre-load choice to highlight
    
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
    
    BlockClear
    
    DO
      
      '--Set upper left corner of menu, turn off the cursor
      LOCATE Row, Col, 0
      BlockClear
      TitleBox 3, Col, MaxLen + 3, "Delete Transactions Menu ", Cnf
      TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
      ShowCursor
      
      VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      
      IF Ky$ = CHR$(27) THEN
        ExitFlag = True         'choice = 0
        EXIT DO
      END IF
      
      SELECT CASE Choice
      CASE 1
        DeletePayTrans OPERNUM
        ExitFlag = True
      CASE 2
        DeleteDepTrans OPERNUM
        ExitFlag = True
      END SELECT
    LOOP UNTIL ExitFlag
  ELSE
    BlockClear
    OK = MsgBox%("UB.QSL", "NOPAYDEP")          'None selected error
  END IF
  GOTO ExitDelTrans
  
DoDelPayment:
  DeletePayTrans OPERNUM
  GOTO ExitDelTrans
  
DoDelDeposit:
  DeleteDepTrans OPERNUM
  
ExitDelTrans:
END SUB

SUB DelList (Items() AS FLen2, Picked(), NPicked)               'STATIC
  
  Marker = 251  'Use ASCII code for - "û"
  MaxChoices = UBOUND(Picked)   'Find Max. number of items
  NPicked = 0   'Init. Number of choices
  MaxLen = 0    'Init. Longest Item
  
  Text$ = SPACE$(68)
  LSET Text$ = "  Customer               Transaction Type        Amount"
  MID$(Text$, 64) = "Taged"
  
  '----- Initialize the "Items$" array
  
  FOR N = 1 TO UBOUND(Items)    'Examine each item
    Lngth = LEN(Items(N))       'Get its length
    IF Lngth THEN               'Is it null?
      IF RIGHT$(Items(N).V, 3) <> "   " THEN    'Is it properly padded?
        Items(N).V = RTRIM$(Items(N).V) + "  "  'Pad it out
        Lngth = LEN(Items(N).V) 'Get new length
      ELSE      'Already padded,
        MID$(Items(N).V, Lngth - 1) = "³"       'Insert a bar character
      END IF
      IF Lngth > MaxLen THEN MaxLen = Lngth     'Update the max. length
      IF NPicked < MaxChoices THEN              'Are we at end of "Picked"?
        IF Picked(NPicked + 1) = N THEN         'Should this Item be marked?
          MID$(Items(N).V, Lngth) = CHR$(Marker)                'Mark it
          NPicked = NPicked + 1
        END IF
      END IF
    ELSE        'Bail out if null Item
      EXIT FOR
    END IF
  NEXT
  
  
  '----- Give menu of items and let User make choices
  Action = 1    'Set "Initialize Menu" Flag
  '  for multi-tasking mode
  BoxBot = 18   'No limit on bottom line
  ActMRow = 21
  
  KillemFlag = False
  
  LOCATE 5, 7, 0
  
  QPrintRC Text$, 4, 7, 112
  
  DisplayUBScrn "DelScrn"
  MPaintBox 5, 75, 5, 76, 8
  
  DO            'Poll the menu
    
    VertMenuT2 Items(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    ButtonPress 1, N, MooseButton, MCol, MRow   ' ----- Check Mouse activity
    
    IF N = 1 THEN
      MRow = (MRow \ 8) + 1     'Convert MCol, MRow to Row and Col.
      MCol = (MCol \ 8) + 1
      IF MRow = ActMRow THEN
        SELECT CASE MCol
        CASE 62 TO 72           'Cancel posting
          PressButton EscKey, ActMRow, 62, 72
        CASE 49 TO 60           'let it rip
          PressButton F10Key, ActMRow, 49, 60
        END SELECT
      END IF
    END IF
    
    IF Ky$ = CHR$(13) OR Ky$ = " " THEN         'Did User press Space or Enter?
      'Is Item not marked?
      Char = 32 'Set Space character code
      IF RIGHT$(Items(Choice).V, 1) = " " THEN
        IF NPicked < MaxChoices THEN            'Are we within the limit?
          Char = Marker         'Set marking character code
          NPicked = NPicked + 1 'Increment # of selections
          Picked(NPicked) = Choice
        END IF
      ELSE      'Un-select an item
        FOR i = 1 TO NPicked
          IF Picked(i) = Choice THEN            'Find the item in Picked() and
            DeleteT SEG Picked(i), -1, MaxChoices - i           ' delete it
            EXIT FOR            'All done,
          END IF
        NEXT
        NPicked = NPicked - 1   ' decrement # of selections
      END IF
      'Put mark character in Item$
      MID$(Items(Choice).V, LEN(Items(Choice).V)) = CHR$(Char)
      Action = 2
      CursorOff
      
    ELSEIF Ky$ = CHR$(0) + "D" THEN             'Did User press F10
      FOR CCnt = 1 TO MaxChoices
        IF Picked(CCnt) > 0 THEN
          FirstPick = CCnt
          OKFlag = True
          EXIT FOR
        END IF
      NEXT
      
      IF NOT OKFlag THEN
        OK = MsgBox%("UB.QSL", "NONEPDEL")      'None selected error
        CursorOff
      ELSE
        OK = MsgBox%("UB.QSL", "OKDELPAY")      'Ask if they are sure
        SELECT CASE OK
        CASE 1  'wipe them out
          KillemFlag = True
        CASE 2  'NOPE continue selecting
          OKFlag = False
        END SELECT
        CursorOff
      END IF
    ELSEIF Ky$ = CHR$(27) THEN
      NPicked = 0
      OKFlag = True
    END IF
    
  LOOP UNTIL OKFlag
  
  '----- Erase the menu from screen with Action of 5
  VertMenuT2 Items(), Choice, MaxLen, BoxBot, Ky$, 5, Cnf
  
END SUB

SUB DisplayOperEntry
  BlockClear
  DisplayUBScrn "UBCHOSOP"
END SUB

SUB OperatorEntry (OperOKFlag, OPERNUM, PostDate$)
  
  Action = 1
  added = 0
  
  LibName$ = "UB.QSL"
  ScrnName$ = "UBCHOSOP"
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  Frm(1).StayOnField = True
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  DisplayOperEntry
  
  FirstTime = True
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      LSET Form$(3, 0) = DATE$
      Action = 1
    END IF
    
    SELECT CASE Frm(1).KeyCode
    CASE F7KEY
      AddEditOperator
      Action = 1
      DisplayOperEntry
    CASE F10Key
      OPERNUM = VAL(Form$(1, 0))
      OperPassword$ = Form$(2, 0)
      PostDate$ = Form$(3, 0)
      GOSUB CheckPostDate
      IF BadDate THEN
        QPrintRC "<-- DATE NOT SET", 12, 49, 4
        BEEP
        SLEEP 1
        DisplayOperEntry
      END IF
      
      IF OPERNUM > 0 THEN
        ValidateOperator OPERNUM, OperPassword$, OperOKFlag
        IF OperOKFlag > 0 THEN
          Done = True
        ELSE
          UBLog "ERROR: INVALID OPERATOR LOGIN"
          UBLog " USED:" + STR$(OPERNUM) + "   PW: " + OperPassword$
          DisplayOperEntry
          Action = 1
          Frm(1).FldNo = 1
        END IF
      ELSE
        OperOKFlag = 0
        Done = True
      END IF
    CASE ESC
      OperOKFlag = -1
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  
  EXIT SUB
  
CheckPostDate:
  IF VAL(LEFT$(PostDate$, 2)) < 1 OR VAL(LEFT$(PostDate$, 2)) > 12 THEN
    BadDate = True
  ELSEIF VAL(MID$(PostDate$, 4, 2)) < 1 OR VAL(MID$(PostDate$, 4, 2)) > 31 THEN
    BadDate = True
  ELSE
    BadDate = False
  END IF
  RETURN
  
END SUB

SUB PostTransEntries (OPERNUM, PostDate$)
  
  CursorOff
  UBLog " IN: UB POST PAYMENTS,  OPER:" + STR$(OPERNUM)
  
  PayBillName$ = "UBPAY" + QPTrim$(STR$(OPERNUM)) + ".DAT"
  PayDepoName$ = "UBDEP" + QPTrim$(STR$(OPERNUM)) + ".DAT"
  
  IF FileSize&(PayBillName$) <= 0 AND FileSize&(PayDepoName$) <= 0 THEN
    BlockClear
    DisplayUBScrn "NOPAYPST"
    QPrintRC STR$(OPERNUM), 12, 31, 79
    WaitForAction
    GOTO ExitPayPost
  END IF
  
  BlockClear
  LibName$ = "UB"
  ScrnName$ = "PAYPOST"
  
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
  
  IF FileSize&("UBSNDEM.DAT") > 0 THEN
    FOR Cnt = 1 TO 3
      QPSound 1750, 2
      QPSound 1650, 2
    NEXT
  END IF
  
  DisplayUBScrn ScrnName$
  
  ShowCursor
  
  Action = 1
  FirstTime = True
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    SELECT CASE Frm(1).KeyCode
    CASE EscKey
      PostOKFlag = False
      EXIT DO
    CASE F10Key, 13
      PostOKFlag = True
      EXIT DO
    END SELECT
  LOOP
  
  IF NOT PostOKFlag THEN
    UBLog "POSTING ABORTED"
    BlockClear
    DisplayUBScrn "POSTCAN"
    WaitForAction
    GOTO ExitPayPost
  END IF
  
  UBLog "POSTING TRANSACTIONS START"
  
  REDIM UBTransRec(1)   AS UBTransRecType
  REDIM TUBTransRec(1)  AS UBTransRecType
  REDIM UBCustRec(1)    AS NewUBCustRecType
  REDIM UBPaymentRec(1) AS UBPaymentRecType
  
  UBCustRecLen = LEN(UBCustRec(1))
  UBPayRecLen = LEN(UBPaymentRec(1))
  UBTransRecLen = LEN(UBTransRec(1))
  
  TranFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS TranFile
  CLOSE TranFile
  
  
  UBLog "POSTING: DEPOSITS"
  IF FileSize&(PayDepoName$) > 0 THEN
    FOpenS "UBCUST.DAT", CHandle
    FOpenS "UBTRANS.DAT", Thandle
    FOpenS PayDepoName$, PHandle
    
    NumPayRecs& = FLof(PHandle) \ UBPayRecLen
    
    ShowProcessingScrn "Posting Deposit Transactions"
    FOR Cnt& = 1 TO NumPayRecs&
      LSET UBTransRec(1) = TUBTransRec(1)
      FGetRTA PHandle, UBPaymentRec(1), Cnt&, UBPayRecLen
      FGetRTA CHandle, UBCustRec(1), UBPaymentRec(1).CUSTACCT, UBCustRecLen
      UBTransRec(1).TransDate = UBPaymentRec(1).PAYDATE
      UBTransRec(1).TransType = TranDepositPayment
      
      '022098 Added
      IF LEN(QPTrim$(UBPaymentRec(1).DESC)) = 0 THEN
        UBTransRec(1).TransDesc = "DEPOSIT PAYMENT"
      ELSE
        UBTransRec(1).TransDesc = UBPaymentRec(1).DESC
      END IF
      '^^This holds the Payment Description
      
      'UBTransRec(1)CustLocation = UBPaymentRec(1).CUSTACCT
      UBTransRec(1).OperatorNumber = OPERNUM
      UBTransRec(1).CustAcctNo = UBPaymentRec(1).CUSTACCT
      UBTransRec(1).CustStatus = UBCustRec(1).Status
      UBTransRec(1).TransAmt = UBPaymentRec(1).AMTPAID
      UBTransRec(1).CheckAmount = UBPaymentRec(1).CHKAMT
      UBTransRec(1).CashAmount = UBPaymentRec(1).CASHAMT
      
      IF UBTransRec(1).CheckAmount > 0 AND UBTransRec(1).CashAmount > 0 THEN
        UBTransRec(1).PayTypeCode = 3
      ELSEIF UBTransRec(1).CashAmount > 0 THEN
        UBTransRec(1).PayTypeCode = 1
      ELSEIF UBTransRec(1).CheckAmount > 0 THEN
        UBTransRec(1).PayTypeCode = 2
      END IF
      
      FOR RevAmts = 1 TO MaxRevsCnt
        UBTransRec(1).RevAmt(RevAmts) = UBPaymentRec(1).PaidOwed(RevAmts).AMTPD1
      NEXT
      
      '05-05-97 FIX added run balance to deposit trans
      UBTransRec(1).RunBalance = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
      
      UBTransRec(1).PrevTrans = UBCustRec(1).LastTrans
      NextTransRec& = (FLof(Thandle) \ UBTransRecLen) + 1
      
      IF NextTransRec& <= 0 THEN
        NextTransRec& = 1
      END IF
      
      FPutRTA Thandle, UBTransRec(1), NextTransRec&, UBTransRecLen
      
      'UBCustRec(1).DepositAmt = UBTransRec(1).TransAmt
      
      '04-14-98 Testing
      UBCustRec(1).DepositAmt = Round#(UBCustRec(1).DepositAmt + UBTransRec(1).TransAmt)
      
      UBCustRec(1).LastTrans = NextTransRec&
      FPutRTA CHandle, UBCustRec(1), UBPaymentRec(1).CUSTACCT, UBCustRecLen
      ShowPctComp Cnt&, NumPayRecs&
    NEXT
    
    FClose CHandle
    FClose Thandle
    FClose PHandle
    
    KillFile PayDepoName$
  END IF
  UBLog "POSTED:" + STR$(NumPayRecs&)
  '**********
  UBLog "POSTING: PAYMENTS"
  IF FileSize&(PayBillName$) > 0 THEN
    FOpenS "UBCUST.DAT", CHandle
    FOpenS "UBTRANS.DAT", Thandle
    FOpenS PayBillName$, PHandle
    
    NumPayRecs& = FLof(PHandle) \ UBPayRecLen
    
    ShowProcessingScrn "Posting Payment Transactions"
    FOR Cnt& = 1 TO NumPayRecs&
      LSET UBTransRec(1) = TUBTransRec(1)
      FGetRTA PHandle, UBPaymentRec(1), Cnt&, UBPayRecLen
      FGetRTA CHandle, UBCustRec(1), UBPaymentRec(1).CUSTACCT, UBCustRecLen
      UBTransRec(1).TransDate = UBPaymentRec(1).PAYDATE
      UBTransRec(1).TransType = TranBillPayment
      
      '052698 Added tax exempt flag to trans rec. For payment summary report
      UBTransRec(1).TaxExempt = UBPaymentRec(1).TaxExempt
      
      '022098 Added
      IF LEN(QPTrim$(UBPaymentRec(1).DESC)) = 0 THEN
        UBTransRec(1).TransDesc = "BILLING-PAYMENT"
      ELSE
        UBTransRec(1).TransDesc = UBPaymentRec(1).DESC
      END IF
      '^^This holds the Payment Description
      'UBTransRec(1)CustLocation = UBPaymentRec(1).CUSTACCT
      UBTransRec(1).OperatorNumber = OPERNUM
      UBTransRec(1).CustAcctNo = UBPaymentRec(1).CUSTACCT
      UBTransRec(1).CustStatus = UBCustRec(1).Status
      UBTransRec(1).TransAmt = UBPaymentRec(1).AMTPAID
      UBTransRec(1).CheckAmount = UBPaymentRec(1).CHKAMT
      UBTransRec(1).CashAmount = UBPaymentRec(1).CASHAMT
      
      IF UBTransRec(1).CheckAmount > 0 AND UBTransRec(1).CashAmount > 0 THEN
        UBTransRec(1).PayTypeCode = 3
      ELSEIF UBTransRec(1).CashAmount > 0 THEN
        UBTransRec(1).PayTypeCode = 1
      ELSEIF UBTransRec(1).CheckAmount > 0 THEN
        UBTransRec(1).PayTypeCode = 2
      END IF
      
      'IF UBCustRec(1).PrevBalance > 0 THEN
      '050597 changed to zero if <> zero
      IF UBCustRec(1).PrevBalance <> 0 THEN
        IF UBTransRec(1).TransAmt >= UBCustRec(1).PrevBalance THEN
          UBCustRec(1).PrevBalance = 0
        ELSEIF UBTransRec(1).TransAmt < UBCustRec(1).PrevBalance THEN
          UBCustRec(1).PrevBalance = Round#(UBCustRec(1).PrevBalance - UBTransRec(1).TransAmt)
        END IF
      END IF
      
      FOR RevAmts = 1 TO MaxRevsCnt
        UBTransRec(1).RevAmt(RevAmts) = UBPaymentRec(1).PaidOwed(RevAmts).AMTPD1
        UBCustRec(1).CurrRevAmts(RevAmts) = Round#(UBCustRec(1).CurrRevAmts(RevAmts) - UBTransRec(1).RevAmt(RevAmts))
        'This is for previous bill distribution
        'UBCustRec(1).PrevRevAmts(RevAmts) = Round#(UBCustRec(1).PrevRevAmts(RevAmts) - UBTransRec(1).RevAmt(RevAmts))
        'IF UBCustRec(1).PrevRevAmts(RevAmts) < 0 THEN
        '  UBCustRec(1).PrevRevAmts(RevAmts) = 0
        'END IF
      NEXT
      TotalCustBalance# = 0
      FOR RevAmts = 1 TO MaxRevsCnt
        TotalCustBalance# = Round#(TotalCustBalance# + UBCustRec(1).CurrRevAmts(RevAmts))
      NEXT
      UBCustRec(1).CurrBalance = Round#(TotalCustBalance# - UBCustRec(1).PrevBalance)
      '02-26-97 Was not adding prev bal
      UBTransRec(1).RunBalance = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
      UBTransRec(1).PrevTrans = UBCustRec(1).LastTrans
      NextTransRec& = (FLof(Thandle) \ UBTransRecLen) + 1
      IF NextTransRec& <= 0 THEN
        NextTransRec& = 1
      END IF
      FPutRTA Thandle, UBTransRec(1), NextTransRec&, UBTransRecLen
      UBCustRec(1).LastTrans = NextTransRec&
      IF Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance) = 0 THEN
        IF UBCustRec(1).Status = "B" THEN
          UBCustRec(1).Status = "I"
          CustChCnt = CustChCnt + 1
          UBLog "POSTING: SET CUST STATUS to I. Acct:" + STR$(UBPaymentRec(1).CUSTACCT)
        END IF
      END IF
      FPutRTA CHandle, UBCustRec(1), UBPaymentRec(1).CUSTACCT, UBCustRecLen
      ShowPctComp Cnt&, NumPayRecs&
    NEXT
    
    FClose CHandle
    FClose Thandle
    FClose PHandle
    KillFile PayBillName$
  END IF
  UBLog "POSTED:" + STR$(NumPayRecs&)
  IF CustChCnt > 0 THEN
    UBLog "POSTING: CUST STATUS CHANGED:" + STR$(CustChCnt)
  END IF
  UBLog "POSTING TRANSACTIONS FINISH"
  
  BlockClear
  DisplayUBScrn "UPDATEOK"
  WaitForAction
  
  ERASE UBTransRec, TUBTransRec, UBCustRec, UBPaymentRec
  
ExitPayPost:
  UBLog "OUT: UB POST PAYMENTS,  OPER:" + STR$(OPERNUM) + CrLf$
END SUB

SUB PrintTransJournal (OPERNUM, PostDate$)

  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen
  IF INSTR(UBSetUpRec(1).UTILNAME, "HARRISBURG") > 0 THEN
    BlockClear
    OK = MsgBox("UBSETUP", "FINONLY")
    SELECT CASE OK
    CASE 1
      StatusFlag = False
    CASE ELSE
      StatusFlag = True
    END SELECT
  ELSE
    StatusFlag = False
  END IF


  REDIM RevText$(1 TO MaxRevsCnt)
  
  REDIM TaxRates(1 TO 15) AS SINGLE
  REDIM TaxAmt(1 TO 15) AS DOUBLE
  
  REDIM BookTotals(0 TO 99) AS BookTotalType
  
  REDIM UBCustRec(1)      AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBSetUp(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUp(), UBSetUpRecLen

  'IF StatusFlag THEN


  'IF INSTR(UBSetUp(1).UTILNAME, "AUTRY") > 0 THEN
  '  LptPort = 2
  'ELSE
    LPTPORT = 1
  'END IF

  LastRev = 15

  FOR Cnt = 1 TO MaxRevsCnt
    TempRev$ = QPTrim$(UBSetUp(1).Revenues(Cnt).RevName)
    IF LEN(TempRev$) = 0 THEN
      LastRev = Cnt - 1
      EXIT FOR
    ELSE
      RevText$(Cnt) = TempRev$
      TaxRates(Cnt) = UBSetUp(1).Revenues(Cnt).TAXRATE
    END IF
  NEXT
  
  REDIM RevAmts(1 TO LastRev) AS DOUBLE
  REDIM DepRevs(1 TO LastRev) AS DOUBLE
  
  CursorOff
  Operator$ = QPTrim$(STR$(OPERNUM))
  FF$ = CHR$(12)
  Page = 0
  LineCnt = 0
  MaxLines = 55
  Dash1$ = STRING$(107, "-")
  
  PayFileName$ = "UBPAY" + QPTrim$(STR$(OPERNUM)) + ".DAT"
  DepFileName$ = "UBDEP" + QPTrim$(STR$(OPERNUM)) + ".DAT"
  PayJourName$ = "UBPAY" + QPTrim$(STR$(OPERNUM)) + ".RPT"
  
  Header$ = "Utility Payment/Deposit Journal"
  
  REDIM CMOperRec(1) AS CMOperRecType
  CMOperRecLen = LEN(CMOperRec(1))
  
  REDIM UBPaymentRec(1) AS UBPaymentRecType
  UBPayRecLen = LEN(UBPaymentRec(1))
  
  CMFile = FREEFILE
  OPEN "CMOPER.DAT" FOR RANDOM SHARED AS CMFile LEN = CMOperRecLen
  NumRecs = LOF(CMFile) \ CMOperRecLen
  
  FOR Cnt = 1 TO NumRecs
    GET CMFile, Cnt, CMOperRec(1)
    IF CMOperRec(1).OperatorNumber = OPERNUM THEN
      Operator$ = QPTrim$(CMOperRec(1).OperatorName)
      EXIT FOR
    END IF
  NEXT
  CLOSE CMFile
  
  IF Exist(PayFileName$) AND FileSize(PayFileName$) > 0 THEN
    PayOKFlag = True
  END IF
  IF Exist(DepFileName$) AND FileSize(DepFileName$) > 0 THEN
    DepOKFlag = True
  END IF
  
  IF NOT DepOKFlag AND NOT PayOKFlag THEN
    BlockClear
    DisplayUBScrn "NOPAYJUR"
    QPrintRC STR$(OPERNUM), 12, 34, 79
    WaitForAction
    GOTO ExitJournal
  END IF
  
  CustFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CustFile LEN = UBCustRecLen
  
  TotalRecs& = (FileSize(DepFileName$) + FileSize(PayFileName$)) \ UBPayRecLen
  RptHandle = FREEFILE
  OPEN PayJourName$ FOR OUTPUT AS RptHandle
  ShowProcessingScrn Header$
  
  GOSUB PrintRptHeader
  IF PayOKFlag THEN
    FOpenS PayFileName$, PHandle
    NumOfRecs& = FLof(PHandle) \ UBPayRecLen
    FOR Cnt& = 1 TO NumOfRecs&
      FGetRTA PHandle, UBPaymentRec(1), Cnt&, UBPayRecLen
      GET CustFile, UBPaymentRec(1).CUSTACCT, UBCustRec(1)
      IF StatusFlag THEN
        IF UBPaymentRec(1).Status <> "F" THEN
          GOTO OnlyFinalSkip
        END IF
      END IF

      GOSUB GetCustBook
      DoneCnt = DoneCnt + 1
      IF UBPaymentRec(1).TaxExempt = "Y" THEN
        TaxExempt = True
      ELSE
        TaxExempt = False
      END IF
      IF LineCnt >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintRptHeader
      END IF
      IF UBPaymentRec(1).CASHAMT < 0 THEN UBPaymentRec(1).CASHAMT = 0
      IF UBPaymentRec(1).CHKAMT < 0 THEN UBPaymentRec(1).CHKAMT = 0
      
      PRINT #RptHandle, Num2Date(UBPaymentRec(1).PAYDATE);
      PRINT #RptHandle, TAB(13); USING "#####"; UBPaymentRec(1).CUSTACCT;
      
      IF OPERNUM = 99 THEN
        Pmnt$ = " DFT "
      ELSE
        Pmnt$ = " PMT "
      END IF

      'IF UBPaymentRec(1).Status = "F" THEN
      '  Pmnt$ = Pmnt$ + "*F"
      'END IF

      PRINT #RptHandle, TAB(20); UBPaymentRec(1).CUSTNAME; Pmnt$;
      PRINT #RptHandle, TAB(53); USING "#####,#.##"; UBPaymentRec(1).CASHAMT;
      PRINT #RptHandle, TAB(69); USING "#####,#.##"; UBPaymentRec(1).CHKAMT;
      PRINT #RptHandle, TAB(83); USING "#####,#.##"; Round#(Round#(UBPaymentRec(1).CHKAMT + UBPaymentRec(1).CASHAMT) - UBPaymentRec(1).CHANGE);
      PRINT #RptHandle, TAB(98); USING "#####,#.##"; UBPaymentRec(1).CHANGE
      BookTotals(CustBook).Check = Round#(BookTotals(CustBook).Check + UBPaymentRec(1).CHKAMT)
      BookTotals(CustBook).Cash = Round#(BookTotals(CustBook).Cash + UBPaymentRec(1).CASHAMT)
      
      TotalCash# = Round#(TotalCash# + UBPaymentRec(1).CASHAMT)
      TotalCheck# = Round#(TotalCheck# + UBPaymentRec(1).CHKAMT)
      TotalAmount# = Round#(TotalAmount# + UBPaymentRec(1).AMTPAID)
      TotalChange# = Round#(TotalChange# + UBPaymentRec(1).CHANGE)
      TotalReceipts = TotalReceipts + 1
      
      LineCnt = LineCnt + 1
      
      FOR RCnt = 1 TO LastRev
        IF NOT TaxExempt THEN
          IF TaxRates(RCnt) > 0 THEN
            Diff# = Round#(UBPaymentRec(1).PaidOwed(RCnt).AMTPD1 / (1 + TaxRates(RCnt)))
            Tax# = Round#(UBPaymentRec(1).PaidOwed(RCnt).AMTPD1 - Diff#)
            TaxAmt(RCnt) = Round#(TaxAmt(RCnt) + Tax#)
            RevAmts(RCnt) = Round#(RevAmts(RCnt) + (UBPaymentRec(1).PaidOwed(RCnt).AMTPD1 - Tax#))
          ELSE
            RevAmts(RCnt) = Round#(RevAmts(RCnt) + UBPaymentRec(1).PaidOwed(RCnt).AMTPD1)
            'IF UBPaymentRec(1).PaidOwed(RCnt).AMTPD1 < -10000 OR UBPaymentRec(1).PaidOwed(RCnt).AMTPD1 > 10000 THEN
            '  STOP
            'END IF
          END IF
        ELSE
          RevAmts(RCnt) = Round#(RevAmts(RCnt) + UBPaymentRec(1).PaidOwed(RCnt).AMTPD1)
        END IF
      NEXT
      '*********************
      ShowPctComp DoneCnt, TotalRecs&
OnlyFinalSkip:
    NEXT
    FClose PHandle
  END IF
  
  IF DepOKFlag THEN
    FOpenS DepFileName$, PHandle
    NumOfRecs& = FLof(PHandle) \ UBPayRecLen
    FOR Cnt& = 1 TO NumOfRecs&
      FGetRTA PHandle, UBPaymentRec(1), Cnt&, UBPayRecLen
      GET CustFile, UBPaymentRec(1).CUSTACCT, UBCustRec(1)
      GOSUB GetCustBook
      
      DoneCnt = DoneCnt + 1
      IF LineCnt >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintRptHeader
      END IF
      IF UBPaymentRec(1).CASHAMT < 0 THEN UBPaymentRec(1).CASHAMT = 0
      IF UBPaymentRec(1).CHKAMT < 0 THEN UBPaymentRec(1).CHKAMT = 0
      PRINT #RptHandle, Num2Date(UBPaymentRec(1).PAYDATE);
      PRINT #RptHandle, TAB(13); USING "#####"; UBPaymentRec(1).CUSTACCT;
      PRINT #RptHandle, TAB(20); UBPaymentRec(1).CUSTNAME; ; " DEP";
      PRINT #RptHandle, TAB(53); USING "#####,#.##"; UBPaymentRec(1).CASHAMT;
      PRINT #RptHandle, TAB(69); USING "#####,#.##"; UBPaymentRec(1).CHKAMT;
      PRINT #RptHandle, TAB(98); USING "#####,#.##"; UBPaymentRec(1).CHANGE
      BookTotals(CustBook).Check = Round#(BookTotals(CustBook).Check + UBPaymentRec(1).CHKAMT)
      BookTotals(CustBook).Cash = Round#(BookTotals(CustBook).Cash + UBPaymentRec(1).CASHAMT)
      
      TotalChange# = Round#(TotalChange# + UBPaymentRec(1).CHANGE)
      TotalCash# = Round#(TotalCash# + UBPaymentRec(1).CASHAMT)
      TotalCheck# = Round#(TotalCheck# + UBPaymentRec(1).CHKAMT)
      TotalReceipts = TotalReceipts + 1
      LineCnt = LineCnt + 1
      FOR RCnt = 1 TO LastRev
        DepositTot# = Round#(DepositTot# + UBPaymentRec(1).PaidOwed(RCnt).AMTPD1)
        DepRevs(RCnt) = Round#(DepRevs(RCnt) + UBPaymentRec(1).PaidOwed(RCnt).AMTPD1)
      NEXT
      ShowPctComp DoneCnt, TotalRecs&
    NEXT
    FClose PHandle
  END IF
  
  GOSUB PrintRptEnding
  
  CLOSE
  
  PrintRptFile Header$, PayJourName$, LPTPORT, -1, EntryPoint
  
  'KillFile PayJourName$
  
ExitJournal:
  EXIT SUB
  
GetCustBook:
  CustBook = Value#(UBCustRec(1).Book, ECode)
  BookTotals(CustBook).Count = BookTotals(CustBook).Count + 1
RETURN
  
PrintRptHeader:
  Page = Page + 1
  IF StatusFlag THEN
    PRINT #RptHandle, "Utility Payment Receipts Journal (Final Only)"
  ELSE
    PRINT #RptHandle, "Utility Payment/Deposit Receipts Journal"
  END IF
  PRINT #RptHandle, "Posting Date: "; PostDate$
  PRINT #RptHandle, "    Operator: "; Operator$; TAB(89); "Page #"; Page
  PRINT #RptHandle, ""
  PRINT #RptHandle, "       "; TAB(12); "        "; TAB(44); "           "; TAB(61); "            "; TAB(82); "Amount Paid"
  PRINT #RptHandle, " Date"; TAB(11); "Acct No      Customer"; TAB(52); "Cash Amount"; TAB(67); "Check Amount"; TAB(82); " on Account"; TAB(96); "Change Given"
  PRINT #RptHandle, Dash1$
  LineCnt = 6
  RETURN
  
PrintRptEnding:
  PRINT #RptHandle, Dash1$
  PRINT #RptHandle, "                  Totals: ";
  PRINT #RptHandle, TAB(53); USING "#####,#.##"; TotalCash#;
  PRINT #RptHandle, TAB(69); USING "#####,#.##"; TotalCheck#;
  PRINT #RptHandle, TAB(83); USING "#####,#.##"; TotalAmount#;
  PRINT #RptHandle, TAB(98); USING "#####,#.##"; TotalChange#
  PRINT #RptHandle, "Total Number of Receipts: "; USING "####,#"; TotalReceipts
  PRINT #RptHandle, FF$
  IF StatusFlag THEN
    PRINT #RptHandle, "Utility Payment Receipts Journal (Final Only)"
  ELSE
    PRINT #RptHandle, "Utility Payment/Deposit Receipts Journal"
  END IF
  PRINT #RptHandle, "Revenue Summary"
  PRINT #RptHandle, ""
  PRINT #RptHandle, "    Revenue"; TAB(34); "Payments         Deposits               Tax"
  PRINT #RptHandle, Dash1$
  GTotal# = 0
  FOR RCnt = 1 TO LastRev
    PRINT #RptHandle, TAB(5); RevText$(RCnt); TAB(30); USING "$$#####,#.##"; RevAmts(RCnt); TAB(47); DepRevs(RCnt); TAB(65); TaxAmt(RCnt)
    GTotal# = Round#(GTotal# + RevAmts(RCnt))
    TTax# = Round#(TTax# + TaxAmt(RCnt))
  NEXT
  PRINT #RptHandle, Dash1$
  PRINT #RptHandle, "Revenue Totals:"; TAB(29); USING "$$######,#.##"; GTotal#; TAB(46); DepositTot#; TAB(64); TTax#
  PRINT #RptHandle,             '"Deposit Total:"; TAB(35); USING "$$#####,#.##"; DepositTot#
  PRINT #RptHandle, "   Grand Total:"; TAB(29); USING "$$######,#.##"; Round#(GTotal# + DepositTot# + TTax#)
  PRINT #RptHandle, FF$
  IF StatusFlag THEN
    PRINT #RptHandle, "Utility Payment Receipts Journal (Final Only)"
  ELSE
    PRINT #RptHandle, "Utility Payment/Deposit Receipts Journal"
  END IF
  PRINT #RptHandle, "Posting Date: "; PostDate$; "    Operator: "; Operator$
  PRINT #RptHandle, "Books Summary"
  PRINT #RptHandle, "Book#     Count         Cash             Check             Total"
  PRINT #RptHandle, Dash1$
  FOR Cnt = 0 TO 99
    IF BookTotals(Cnt).Check <> 0 OR BookTotals(Cnt).Cash <> 0 THEN
      IF Cnt < 10 THEN
        PRINT #RptHandle, TAB(2); USING "0#"; Cnt;
      ELSE
        PRINT #RptHandle, TAB(2); USING "##"; Cnt;
      END IF
      PRINT #RptHandle, TAB(12); USING "####"; BookTotals(Cnt).Count;
      PRINT #RptHandle, TAB(19); USING "#######.##"; BookTotals(Cnt).Cash; TAB(37); BookTotals(Cnt).Check; TAB(55); Round#(BookTotals(Cnt).Cash + BookTotals(Cnt).Check)
    END IF
  NEXT
  PRINT #RptHandle, Dash1$
  PRINT #RptHandle, "Totals:"; TAB(12); USING "####"; TotalReceipts;
  PRINT #RptHandle, TAB(19); USING "#######.##"; TotalCash#;
  PRINT #RptHandle, TAB(37); USING "#######.##"; TotalCheck#;
  PRINT #RptHandle, TAB(55); USING "#######.##"; Round#(GTotal# + DepositTot# + TTax#)
  PRINT #RptHandle, FF$
  RETURN
  
END SUB

SUB ValidateOperator (OPERNUM, OperPassword$, OperOKFlag)
  
  REDIM CMOperRec(1) AS CMOperRecType
  OperFileLength = LEN(CMOperRec(1))
  
  FileHandle = FREEFILE
  OPEN "CMOPER.DAT" FOR RANDOM SHARED AS #FileHandle LEN = OperFileLength
  
  NumOperRecs = LOF(FileHandle) \ OperFileLength
  
  IF NumOperRecs = 0 THEN
    OperOKFlag = 0
    GOTO ExitValidate
  END IF
  
  FOR Cnt = 1 TO NumOperRecs
    GET FileHandle, Cnt, CMOperRec(1)
    IF OPERNUM = CMOperRec(1).OperatorNumber AND OperPassword$ = CMOperRec(1).OperatorPassword THEN
      OperOKFlag = Cnt
      EXIT FOR
    END IF
  NEXT
  
ExitValidate:
  CLOSE
  IF OperOKFlag = 0 THEN
    BEEP
    OK = MsgBox("UB", "BADOPER")
  END IF
  
END SUB

