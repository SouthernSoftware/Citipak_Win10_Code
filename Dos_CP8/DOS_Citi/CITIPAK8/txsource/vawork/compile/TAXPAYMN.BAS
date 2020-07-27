DEFINT A-Z
DECLARE SUB ShowCustHistory (CustRec&, TaxType%)
DECLARE SUB DMVTest ()
DECLARE SUB DelList (Items() AS ANY, Picked%(), NPicked%)
DECLARE SUB DeletePayments (TaxType%)
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB DeleteT (SEG Element, ElSize, NumEls)
DECLARE SUB PostTaxPayments ()
DECLARE SUB PrintPaymentJournal ()
DECLARE FUNCTION AskSavePayment% (DefaultFlag$)
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
DECLARE SUB EnterPayments (TaxType%)
DECLARE SUB FClose (Handle%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB HideCursor ()
DECLARE SUB KillFile (FileName$)
DECLARE SUB LookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, ActiveOnly%, TaxType%)
DECLARE SUB OperatorEntry ()
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB ButtonPress (WhatBtn, N, MooseButton, MCol, MRow)
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
  
  
  TYPE FLen2
    V AS STRING * 64
  END TYPE
  
  TYPE VACustPayListType
    CustAcct     AS LONG
    LastPayRec  AS LONG
    NumPayRec   AS LONG
    BillType AS STRING * 1      'R  or P
  END TYPE
  
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TaxCONST.BI'
  '$INCLUDE: 'PROPAbst.BI'
  '$INCLUDE: 'VApaymnt.BI'
  '$INCLUDE: 'cmfiles.BI'
  
  CONST False = 0, True = NOT False
  
  STACK 5000
  
  CrLf$ = CHR$(13) + CHR$(10)
  
  DIM SHARED OperNum AS INTEGER, OperOKFlag AS INTEGER
  DIM SHARED PostDate AS STRING, OperPassword AS STRING
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 8)
  
  MChoice$(1) = "Enter/Edit Real Estate Tax Payments"
  MChoice$(2) = "Enter/Edit Personal Tax Payments"
  MChoice$(3) = "Delete a Tax Payment"
  MChoice$(4) = "Print Transaction Edit List"
  MChoice$(5) = "Post Transactions"
  MChoice$(6) = "Exit to OS"
  
  IF WEnvTest AND INSTR(COMMAND$, "TEST") = 0 THEN
    OperNum = 1
    OperPassword = "ME"
    PostDate = DATE$

    'PostDate = "10-31-1998"
  ELSE
    DO
      OperatorEntry
    LOOP WHILE OperOKFlag = 0
  END IF
  
  IF OperNum = 0 THEN
    GOTO TaxPaymentExit
  END IF
  
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
    
    TitleBox 3, Col, MaxLen + 3, "Tax Billing Payment Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      EnterPayments 1
    CASE 2
      EnterPayments 2
    CASE 3
      DeletePayments 1
    CASE 4
      PrintPaymentJournal
    CASE 5
      PostTaxPayments
    CASE 6
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP
  
TaxPaymentExit:
  
  RUN "TaxMenu"
  
  END

SUB AddEditOperator
  
  AuthorizeOperator oktoadd
  
  IF oktoadd = 0 THEN EXIT SUB
  
  LibName$ = "TAX"
  ScrnName$ = "TXNEWOP"
  BlockClear
  DisplayTaxScrn ScrnName$
  
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
      
      Done = True
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  
  CLOSE CMFile
END SUB

FUNCTION AskSavePayment% (DefaultFlag$)
  
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
  
  LibName$ = "TAX"
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
  DisplayTaxScrn ScrnName$
  
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

SUB DeletePayments (TaxType%)
  
  Oper$ = QPTrim$(STR$(OperNum))
  
  TaxCPRFileName$ = "TAXCPR" + Oper$ + ".DAT"   'Customers Payment Record file
  TaxLOPFileName$ = "TAXLOP" + Oper$ + ".DAT"   'List Of Payments customers
  TPayFileName$ = "TAXCPR" + Oper$ + ".$$$"
  
  Fmt$ = "#######.##"
  FmtA$ = "######"
  
  
  REDIM TaxPaymentRec(1)   AS TaxPaymentRecType
  REDIM PayList(1 TO 2)      AS PayListType
  TaxPayRecLen = LEN(TaxPaymentRec(1))
  PayListLen = LEN(PayList(1))
  
  WhatCnt& = 1
  
  PayRecFile = FREEFILE
  OPEN TaxCPRFileName$ FOR RANDOM SHARED AS PayRecFile LEN = TaxPayRecLen
  NumOfRecs& = LOF(PayRecFile) \ TaxPayRecLen
  
  IF NumOfRecs& > 0 THEN
    REDIM TransList(1 TO NumOfRecs&)    AS FLen2
    REDIM Picked(1 TO NumOfRecs&) AS INTEGER
    FOR Cnt& = 1 TO NumOfRecs&
      GET #PayRecFile, Cnt&, TaxPaymentRec(1)
      LSET TransList(WhatCnt&).V = FUsing$(STR$(TaxPaymentRec(1).CustAcct), FmtA$)
      MID$(TransList(WhatCnt&).V, 8) = TaxPaymentRec(1).BillType
      MID$(TransList(WhatCnt&).V, 10) = LTRIM$(TaxPaymentRec(1).CustName)
      MID$(TransList(WhatCnt&).V, 44) = FUsing$(STR$(TaxPaymentRec(1).AmtPaid), Fmt$)
      WhatCnt& = WhatCnt& + 1
    NEXT
  ELSE
    CLOSE
    EXIT SUB
  END IF
  CLOSE
  'PayListCnt& = NumOfRecs&
  
  MaxLen = 64
  Col = ((80 - MaxLen) \ 2) - 2
  
  BlockClear
  
  DelList TransList(), Picked(), NPicked
  
  'STOP
  
  IF NPicked > 0 THEN
    PFile = FREEFILE
    OPEN TaxCPRFileName$ FOR RANDOM SHARED AS PFile LEN = TaxPayRecLen
    TFile = FREEFILE
    OPEN TPayFileName$ FOR RANDOM SHARED AS TFile LEN = TaxPayRecLen
    PayListFile = FREEFILE
    OPEN TaxLOPFileName$ FOR RANDOM AS PayListFile LEN = PayListLen
    FOR Cnt = 1 TO NumOfRecs&
      GET PFile, Cnt, TaxPaymentRec(1)
      SELECT CASE RIGHT$(TransList(Cnt).V, 1)
      CASE "û"
        ThisPayRec& = TaxPaymentRec(1).LastPayRec
        DO WHILE ThisPayRec& > 0
          GET #PayListFile, ThisPayRec&, PayList(1)             'get paylist rec
          PUT #PayListFile, ThisPayRec&, PayList(2)             'put blank rec
          ThisPayRec& = PayList(1).PrevListRec
        LOOP
      CASE ELSE
        PUT TFile, , TaxPaymentRec(1)
      END SELECT
    NEXT
    CLOSE
    KillFile TaxCPRFileName$
    NAME TPayFileName$ AS TaxCPRFileName$
  END IF
  IF FileSize(TaxCPRFileName$) = 0 THEN
    KillFile TaxCPRFileName$
  END IF
  
  'add check for none and deleted
  
END SUB

SUB DelList (Items() AS FLen2, Picked(), NPicked)               'STATIC
  
  Marker = 251  'Use ASCII code for - "û"
  MaxChoices = UBOUND(Picked)   'Find Max. number of items
  NPicked = 0   'Init. Number of choices
  MaxLen = 0    'Init. Longest Item
  
  Text$ = SPACE$(68)
  LSET Text$ = "  Acct# Type   Customer                           Amount"
  MID$(Text$, 63) = "Tagged"
  
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
  BoxBot = 17   'No limit on bottom line
  ActMRow = 21
  
  KillemFlag = False
  
  LOCATE 5, 7, 0
  
  QPrintRC Text$, 4, 7, 112
  
  DisplayTaxScrn "DelScrn"
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
        OK = MsgBox%("TAX.QSL", "NONEPDEL")     'None selected error
        CursorOff
      ELSE
        OK = MsgBox%("TAX.QSL", "OKDELPAY")     'Ask if they are sure
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
  DisplayTaxScrn "TXCHOSOP"
END SUB

SUB EnterPayments (TaxType%)
  
  SHARED Choice$()
  
  REDIM TempScrn(0)
  REDIM TaxPaymentRec(1)     AS TaxPaymentRecType
  REDIM PayList(1 TO 1)      AS PayListType
  REDIM TaxCustRec(1)        AS TaxCustType
  REDIM AmtFlds(1 TO 7)      AS INTEGER
  
  TaxPayRecLen = LEN(TaxPaymentRec(1))
  PayListLen = LEN(PayList(1))
  TaxCustRecLen = LEN(TaxCustRec(1))
  
  REDIM TaxSetUp(1) AS TaxMasterType
  TaxSetupLen = LEN(TaxSetUp(1))
  FGetAH "TAXSETUP.DAT", TaxSetUp(1), TaxSetupLen, 1            'load it
  
  RcptPort = TaxSetUp(1).RcptPort
  RecptDef$ = TaxSetUp(1).RecptDef

  IF RcptPort < 1 THEN
    RcptPort = 1
  ELSEIF RcptPort > 2 THEN
    RcptPort = 2
  END IF
  GOSUB LoadCustPayList
  
  IF RecpPort < 1 OR RecpPort > 2 THEN
    RecpPort = 1
  END IF
  
  TOWNNAME$ = TaxSetUp(1).NAME
  
  LibName$ = "TAX"
  IF TaxType% = 1 THEN
    ScrnName$ = "VAREPAY"
  ELSE
    ScrnName$ = "VAPPPAY"
  END IF
  
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

  IF TaxType% = 2 THEN
    FOR zz = 1 TO 7
      AmtFlds(zz) = FldNum%("AMTOWE" + QPTrim$(STR$(zz)), Fld())
    NEXT zz
  ELSE
    FOR zz = 1 TO 3
      AmtFlds(zz) = FldNum%("AMTOWE" + QPTrim$(STR$(zz)), Fld())
    NEXT zz
  END IF
  
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

    CASE F4Key
      IF CustAcct& > 0 THEN
        ShowCustHistory CustAcct&, TaxType%
      END IF

    CASE F7KEY  'Lookup Customer
      IF Frm(1).FldNo = 3 THEN  'if user is on the Customer field
        SaveScrn TempScrn()     'and F7key then do lookup routine
        MPaintBox 4, 5, 22, 75, 8
        LastCust& = CustAcct&
        LookUp CustAcct&, "Payment", 0, False, False, TaxType%
        RestScrn TempScrn()
        IF CustAcct& > 0 THEN   'if this is a valid customer
          GOSUB ChkCustList
          GOSUB GetCustInfo     'go get customer info
          Frm(1).FldNo = 4
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
      TempAmtRecv# = Value#(Form$(AmtRecvFld, 0), ECode)
      IF TempAmtRecv# > 0 THEN
        GOSUB AutoDistribute
      END IF
      
    CASE F10Key 'Save
      GOSUB CheckPaymentInfo
      IF PaymentOKFlag THEN
        SELECT CASE AskSavePayment(RecptDef$)
        CASE 1  'Save trans print receipt
          GOSUB SaveTransaction
          GOSUB PrintReceipt
          GOSUB ClearForm
          Frm(1).FldNo = 1
          Action = 1
          EditFlag = False
          GOSUB LoadCustPayList
        CASE True               'Save trans no receipt
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
  IF TaxType% = 2 THEN
    NumOfCustRecs& = FileSize("PPTXCUST.DAT") \ TaxCustRecLen
  ELSE
    NumOfCustRecs& = FileSize("RETXCUST.DAT") \ TaxCustRecLen
  END IF
  IF CustAcct& > NumOfCustRecs& OR CustAcct& = 0 THEN
    CustAcct& = 0
    OK = MsgBox%("TAX.QSL", "BADACCTN")
    Frm(1).FldNo = 1
    GOSUB SetOperInfo
    GOTO SkipCustInfo
  ELSEIF IsCustDeleted(CustAcct&, TaxType%) THEN
    CustAcct& = 0
    OK = MsgBox%("TAX.QSL", "DELACCTN")
    Frm(1).FldNo = 1
    GOSUB SetOperInfo
    GOTO SkipCustInfo
  END IF
  
  CustFile = FREEFILE
  IF TaxType% = 2 THEN
    OPEN "PPTXCUST.DAT" FOR RANDOM SHARED AS CustFile LEN = TaxCustRecLen
  ELSE
    OPEN "RETXCUST.DAT" FOR RANDOM SHARED AS CustFile LEN = TaxCustRecLen
  END IF
  GET CustFile, CustAcct&, TaxCustRec(1)
  CLOSE CustFile
  
  IF NOT EditFlag THEN
    IF NOT DoesCustOwe%(TaxCustRec(1)) THEN
      CustAcct& = 0
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
  TPrinciple# = 0
  TPrinciple2# = 0
  TPrinciple3# = 0
  TPrinciple4# = 0
  TPrinciple5# = 0
  Interest# = 0: Penalty# = 0
  
  BeenEditedFlag = False        'clear the edited flag
RETURN
  
SaveTransaction:
  SaveScrn TempScrn()
  DisplayTaxScrn "UPDATDSK"
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(TaxPaymentRec(1)), VARPTR(TaxPaymentRec(1)), LEN(Form$(0, 0)), 0
  IF TaxType% = 1 THEN LastField = 3 ELSE LastField = 7
  FOR Cnt = 1 TO LastField      'make sure there are no blanks
    IF TaxPaymentRec(1).PaidOwed(Cnt).AmtPaid < -100000# THEN
      TaxPaymentRec(1).PaidOwed(Cnt).AmtPaid = 0
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
  TaxPaymentRec(1).AmtOwed = TAmtOwed#
  TaxPaymentRec(1).AmtPaid = TAmtPaid#
  IF TaxType% = 2 THEN
    TaxPaymentRec(1).BillType = "P"
  ELSE
    TaxPaymentRec(1).BillType = "R"
  END IF
  PUT #PayRecFile, CustPayRec&, TaxPaymentRec(1)
  EditFlag = False
  CLOSE PayRecFile
  NPicked = 0
  LastCust& = 0
  CustAcct& = 0
  RestScrn TempScrn()
  FirstTime = True
RETURN
  
UPDateListOfPayments:
  
  IF EditFlag THEN
    REDIM TPayList(1 TO 2) AS PayListType
    TempName$ = Unique$("")
    NAME TaxLOPFileName$ AS TempName$
    NPayListFile = FREEFILE
    OPEN TaxLOPFileName$ FOR RANDOM AS NPayListFile LEN = PayListLen
    PayListFile = FREEFILE
    OPEN TempName$ FOR RANDOM AS PayListFile LEN = PayListLen
    NumListRec& = (LOF(PayListFile) \ PayListLen)
    FOR Cnt& = 1 TO NumListRec&
      GET #PayListFile, Cnt&, TPayList(1)
      IF TPayList(1).CustRec <> CustAcct& THEN  'if this rec isn't this cust
        PUT #NPayListFile, , TPayList(1)        'write it to disk
      ELSE
        PUT #NPayListFile, , TPayList(2)        'write it to disk
      END IF
    NEXT
    CLOSE PayListFile, NPayListFile
    KILL TempName$              'erase the old temp file
    'Paylist file is now clean of cust old records.
    'Common code below will handle putting new ones back
  END IF
  
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
    REDIM CustList(1 TO NumOfRecs&)    AS VACustPayListType
    FOR Cnt& = 1 TO NumOfRecs&
      GET #PayRecFile, Cnt&, TaxPaymentRec(1)
      CustList(Cnt&).CustAcct = TaxPaymentRec(1).CustAcct
      CustList(Cnt&).LastPayRec = TaxPaymentRec(1).LastPayRec
      CustList(Cnt&).NumPayRec = TaxPaymentRec(1).NumPayRec
      CustList(Cnt&).BillType = TaxPaymentRec(1).BillType
    NEXT
  END IF
  CLOSE PayRecFile
  CustListCnt& = NumOfRecs&
  RETURN
  
ChkCustList:
  IF TaxType% = 1 THEN BillType$ = "R" ELSE BillType$ = "P"
  EditFlag = False
  IF CustListCnt& > 0 THEN
    FOR Cnt = 1 TO CustListCnt&
      IF CustList(Cnt).CustAcct = CustAcct& AND CustList(Cnt).BillType = BillType$ THEN
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
  TPenalty# = 0
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
    PayList(LCnt).Interest = TPayList(1).Interest
    TInterest# = Round#(TInterest# + TPayList(1).Interest)
    PayList(LCnt).Penalty = TPayList(1).Penalty
    TPenalty# = Round#(TPenalty# + TPayList(1).Penalty)
    ThisPayRec& = TPayList(1).PrevListRec
  LOOP
  NPicked = LCnt
  CLOSE PayListFile
  RETURN
  
CheckPaymentInfo:
  'Parse and move data to Paylist records here
  PaymentOKFlag = True
  IF TaxType% = 2 THEN
    PrinceOW# = 0: PrincePD# = 0
    PrinceOW# = Value#(Form$(AmtFlds(1), 0), ECode)
    PrincePD# = Value#(Form$(AmtFlds(1) + 1, 0), ECode)
    PrinceOW# = PrinceOW# + Value#(Form$(AmtFlds(2), 0), ECode)
    PrincePD# = PrincePD# + Value#(Form$(AmtFlds(2) + 1, 0), ECode)
    PrinceOW# = PrinceOW# + Value#(Form$(AmtFlds(3), 0), ECode)
    PrincePD# = PrincePD# + Value#(Form$(AmtFlds(3) + 1, 0), ECode)
    PrinceOW# = PrinceOW# + Value#(Form$(AmtFlds(4), 0), ECode)
    PrincePD# = PrincePD# + Value#(Form$(AmtFlds(4) + 1, 0), ECode)
    PrinceOW# = PrinceOW# + Value#(Form$(AmtFlds(5), 0), ECode)
    PrincePD# = PrincePD# + Value#(Form$(AmtFlds(5) + 1, 0), ECode)
    InterestOw# = Value#(Form$(AmtFlds(6), 0), ECode)
    InterestPd# = Value#(Form$(AmtFlds(6) + 1, 0), ECode)
    PenaltyOw# = Value#(Form$(AmtFlds(7), 0), ECode)
    PenaltyPd# = Value#(Form$(AmtFlds(7) + 1, 0), ECode)
  ELSE
    PrinceOW# = 0: PrincePD# = 0
    PrinceOW# = Value#(Form$(AmtFlds(1), 0), ECode)
    PrincePD# = Value#(Form$(AmtFlds(1) + 1, 0), ECode)
    InterestOw# = Value#(Form$(AmtFlds(2), 0), ECode)
    InterestPd# = Value#(Form$(AmtFlds(2) + 1, 0), ECode)
    PenaltyOw# = Value#(Form$(AmtFlds(3), 0), ECode)
    PenaltyPd# = Value#(Form$(AmtFlds(3) + 1, 0), ECode)
  END IF
  
  TAmtOwed# = Value#(Form$(AmtOwedFld, 0), ECode)
  TAmtRecv# = Value#(Form$(AmtRecvFld, 0), ECode)
  TAmtPaid# = Value#(Form$(AmtPaidFld, 0), ECode)
  ChangeAmt# = Value#(Form$(ChangeFld, 0), ECode)
  
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
  
  IF (PrincePD# > PrinceOW#) OR (InterestPd# > InterestOw#) OR (PenaltyPd# > PenaltyOw#) THEN
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
    IF TaxType% = 1 THEN
      PPrinciple# = PayList(Cnt).Principle1
    ELSE
      PPrinciple# = PayList(Cnt).Principle1 + PayList(Cnt).Principle2 + PayList(Cnt).Principle3 + PayList(Cnt).Principle4 + PayList(Cnt).Principle5
    END IF
    
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
    PInterest# = PayList(Cnt).Interest
    IF (InterestPd# >= PInterest#) AND (InterestPd# > 0) THEN
      InterestPd# = Round#(InterestPd# - PInterest#)
    ELSE
      IF InterestPd# > 0 THEN
        PayList(Cnt).Interest = InterestPd#
        InterestPd# = 0
      ELSE
        PayList(Cnt).Interest = 0
      END IF
    END IF
    PPenalty# = PayList(Cnt).Penalty
    IF (PenaltyPd# >= PPenalty#) AND (PenaltyPd# > 0) THEN
      PenaltyPd# = Round#(PenaltyPd# - PPenalty#)
    ELSE
      IF PenaltyPd# > 0 THEN
        PayList(Cnt).Penalty = PenaltyPd#
        PenaltyPd# = 0
      ELSE
        PayList(Cnt).Penalty = 0
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
    DO WHILE TransRecord& > 0
      GET TransFile, TransRecord&, TaxTrans(1)
      IF TaxTrans(1).TranType = 1 THEN
        Balance# = Round#(TaxTrans(1).Revenue.Principle1 + TaxTrans(1).Revenue.Principle2 + TaxTrans(1).Revenue.Principle3 + TaxTrans(1).Revenue.Principle4 + TaxTrans(1).Revenue.Principle5)
        Balance# = Round#(Balance# + TaxTrans(1).Revenue.Interest + TaxTrans(1).Revenue.Penalty + TaxTrans(1).Revenue.Collection)
        Balance# = Round#(Balance# - (TaxTrans(1).Revenue.Principle1Pd + TaxTrans(1).Revenue.Principle2Pd + TaxTrans(1).Revenue.Principle3Pd + TaxTrans(1).Revenue.Principle4Pd + TaxTrans(1).Revenue.Principle5Pd))
        Balance# = Round#(Balance# - (TaxTrans(1).Revenue.InterestPd + TaxTrans(1).Revenue.PenaltyPd + TaxTrans(1).Revenue.CollectionPD))
        IF Balance# > 0 THEN
          BillCnt = BillCnt + 1
          REDIM PRESERVE Items(1 TO BillCnt) AS FLen2
          LSET Items(BillCnt).V = Num2Date(TaxTrans(1).TransDate)
          MID$(Items(BillCnt).V, 20) = FUsing$(STR$(TaxTrans(1).TaxYear), "####")
          MID$(Items(BillCnt).V, 30) = FUsing$(STR$(TransRecord&), "######")
          MID$(Items(BillCnt).V, 40) = FUsing$(STR$(TaxTrans(1).Amount), "######.##")
          MID$(Items(BillCnt).V, 50) = FUsing$(STR$(Balance#), "######.##")
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
    TPenalty# = 0
    TCollect# = 0
    REDIM PayList(1 TO NPicked)      AS PayListType
    TransFile = FREEFILE
    OPEN "TaxTrans.dat" FOR RANDOM SHARED AS TransFile LEN = LEN(TaxTrans(1))
    FOR zz = 1 TO NPicked
      TNum$ = MID$(Items(Picked(zz)).V, 30, 7)
      TransNum& = QPValL&(TNum$)
      GET #TransFile, TransNum&, TaxTrans(1)
      PayList(zz).BillRec = TransNum&
      PayList(zz).CustRec = CustAcct&
      PayList(zz).Principle1 = Round#(TaxTrans(1).Revenue.Principle1 - TaxTrans(1).Revenue.Principle1Pd)
      TPrinciple# = Round#(TPrinciple# + PayList(zz).Principle1)
      PayList(zz).Principle2 = Round#(TaxTrans(1).Revenue.Principle2 - TaxTrans(1).Revenue.Principle2Pd)
      TPrinciple2# = Round#(TPrinciple2# + PayList(zz).Principle2)
      PayList(zz).Principle3 = Round#(TaxTrans(1).Revenue.Principle3 - TaxTrans(1).Revenue.Principle3Pd)
      TPrinciple3# = Round#(TPrinciple3# + PayList(zz).Principle3)
      PayList(zz).Principle4 = Round#(TaxTrans(1).Revenue.Principle4 - TaxTrans(1).Revenue.Principle4Pd)
      TPrinciple4# = Round#(TPrinciple4# + PayList(zz).Principle4)
      PayList(zz).Principle5 = Round#(TaxTrans(1).Revenue.Principle5 - TaxTrans(1).Revenue.Principle5Pd)
      TPrinciple5# = Round#(TPrinciple5# + PayList(zz).Principle5)
      
      PayList(zz).Interest = Round#(TaxTrans(1).Revenue.Interest - TaxTrans(1).Revenue.InterestPd)
      TInterest# = Round#(TInterest# + PayList(zz).Interest)

      PayList(zz).Penalty = Round#(TaxTrans(1).Revenue.Penalty - TaxTrans(1).Revenue.PenaltyPd)
      TPenalty# = Round#(TPenalty# + PayList(zz).Penalty)

      Frm(1).FldNo = TenderFld
    NEXT
    CLOSE
    
    IF TaxType% = 2 THEN
      LSET Form$(AmtFlds(1), 0) = STR$(TPrinciple#)
      LSET Form$(AmtFlds(1) + 1, 0) = STR$(0)
      LSET Form$(AmtFlds(2), 0) = STR$(TPrinciple2#)
      LSET Form$(AmtFlds(2) + 1, 0) = STR$(0)
      LSET Form$(AmtFlds(3), 0) = STR$(TPrinciple3#)
      LSET Form$(AmtFlds(3) + 1, 0) = STR$(0)
      LSET Form$(AmtFlds(4), 0) = STR$(TPrinciple4#)
      LSET Form$(AmtFlds(4) + 1, 0) = STR$(0)
      LSET Form$(AmtFlds(5), 0) = STR$(TPrinciple5#)
      LSET Form$(AmtFlds(5) + 1, 0) = STR$(0)
      
      LSET Form$(AmtFlds(6), 0) = STR$(TInterest#)
      LSET Form$(AmtFlds(6) + 1, 0) = STR$(0)
      LSET Form$(AmtFlds(7), 0) = STR$(TPenalty#)
      LSET Form$(AmtFlds(7) + 1, 0) = STR$(0)
      TAmtOwed# = Round#(TPrinciple# + TPrinciple2# + TPrinciple3# + TPrinciple4# + TPrinciple5# + TInterest# + TPenalty#)
      LSET Form$(AmtOwedFld, 0) = STR$(TAmtOwed#)
      
      LSET Form$(ChkAmtFld, 0) = STR$(0)
      LSET Form$(CashAmtFld, 0) = STR$(0)
      LSET Form$(TotalFld, 0) = STR$(0)
      LSET Form$(AmtRecvFld, 0) = STR$(0)
      LSET Form$(ChangeFld, 0) = STR$(0)
      LSET Form$(AmtPaidFld, 0) = STR$(0)
      
      FOR zz = 1 TO 7
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
    ELSE
      LSET Form$(AmtFlds(1), 0) = STR$(TPrinciple#)
      LSET Form$(AmtFlds(1) + 1, 0) = STR$(0)
      
      LSET Form$(AmtFlds(2), 0) = STR$(TInterest#)
      LSET Form$(AmtFlds(2) + 1, 0) = STR$(0)
      LSET Form$(AmtFlds(3), 0) = STR$(TPenalty#)
      LSET Form$(AmtFlds(3) + 1, 0) = STR$(0)
      
      
      TAmtOwed# = Round#(TPrinciple# + TInterest# + TPenalty#)
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
    
  END IF
  
  Action = 1
  RestScrn TempScrn()
  
  ERASE Items
  
  RETURN
  
AutoDistribute:
  IF TaxType% = 2 THEN
    FOR Cnt = 7 TO 1 STEP -1
      ThisAmt# = Value(Form$(AmtFlds(Cnt), 0), ECode)
      IF ThisAmt# < 0 THEN
        TempAmtRecv# = Round#(TempAmtRecv# - ThisAmt#)
      END IF
    NEXT
    
    FOR Cnt = 7 TO 1 STEP -1
      ThisAmt# = Value(Form$(AmtFlds(Cnt), 0), ECode)
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
  ELSE
    FOR Cnt = 3 TO 1 STEP -1
      ThisAmt# = Value(Form$(AmtFlds(Cnt), 0), ECode)
      IF ThisAmt# < 0 THEN
        TempAmtRecv# = Round#(TempAmtRecv# - ThisAmt#)
      END IF
    NEXT
    
    FOR Cnt = 3 TO 1 STEP -1
      ThisAmt# = Value(Form$(AmtFlds(Cnt), 0), ECode)
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
    
  END IF
  
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
  RETURN
  
PrintReceipt:
  SaveScrn TempScrn()
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
  PRINT #RHandle, TaxPaymentRec(1).DESC
  PRINT #RHandle, "Acct. No. "; TaxPaymentRec(1).CustAcct
  PRINT #RHandle,
  PRINT #RHandle, "Total Owed: "; USING "$$####,#.##"; TaxPaymentRec(1).AmtOwed
  PRINT #RHandle, "Total Paid: "; USING "$$####,#.##"; TaxPaymentRec(1).AmtPaid
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
  
  KillFile PayRecpName$
  RestScrn TempScrn()
  RETURN
  
  
END SUB

SUB OperatorEntry
  
  Action = 1
  added = 0
  
  LibName$ = "TAX.QSL"
  ScrnName$ = "TXCHOSOP"
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
      OperNum = VAL(Form$(1, 0))
      OperPassword$ = Form$(2, 0)
      PostDate$ = Form$(3, 0)
      GOSUB CheckPostDate
      IF BadDate THEN
        QPrintRC "<-- DATE NOT SET", 12, 49, 4
        BEEP
        SLEEP 1
        DisplayOperEntry
      END IF
      
      IF OperNum > 0 THEN
        ValidateOperator
        IF OperOKFlag > 0 THEN
          Done = True
        ELSE
          'UBLog "ERROR: INVALID OPERATOR LOGIN"
          'UBLog " USED:" + STR$(OPERNUM) + "   PW: " + OperPassword$
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

SUB PostTaxPayments
  '  EXIT SUB
  
  CursorOff
  
  Oper$ = QPTrim$(STR$(OperNum))
  
  TaxCPRFileName$ = "TAXCPR" + Oper$ + ".DAT"   'Customers Payment Record file
  TaxLOPFileName$ = "TAXLOP" + Oper$ + ".DAT"   'List Of Payments customers
  
  IF FileSize&(TaxCPRFileName$) <= 0 THEN
    BlockClear
    DisplayTaxScrn "NOPAYPST"
    QPrintRC STR$(OperNum), 12, 31, 79
    WaitForAction
    GOTO ExitPayPost
  END IF
  
  BlockClear
  LibName$ = "TAX"
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
  
  DisplayTaxScrn ScrnName$
  
  ShowCursor
  
  Action = 1
  FirstTime = True
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    SELECT CASE Frm(1).KeyCode
    CASE EscKey
      CursorOff
      PostOKFlag = False
      EXIT DO
    CASE F10Key, 13
      PostOKFlag = True
      EXIT DO
    END SELECT
  LOOP
  
  IF NOT PostOKFlag THEN
    BlockClear
    DisplayTaxScrn "POSTCAN"
    WaitForAction
    GOTO ExitPayPost
  END IF
  
  REDIM TaxPaymentRec(1)     AS TaxPaymentRecType
  REDIM PayListRec(1 TO 1)   AS PayListType
  REDIM TaxCustRec(1)        AS TaxCustType
  REDIM TaxTranRec(1)        AS TaxTransactionType
  REDIM PayTranRec(1)         AS TaxTransactionType
  
  TaxPayRecLen = LEN(TaxPaymentRec(1))
  PayListLen = LEN(PayListRec(1))
  TaxCustRecLen = LEN(TaxCustRec(1))
  TaxTranRecLen = LEN(TaxTranRec(1))
  
  BlockClear
  ShowProcessingScrn "Posting Payment Transactions"
  
  PayRecFile = FREEFILE
  OPEN TaxCPRFileName$ FOR RANDOM SHARED AS PayRecFile LEN = TaxPayRecLen
  PayListFile = FREEFILE
  OPEN TaxLOPFileName$ FOR RANDOM AS PayListFile LEN = PayListLen
  TranFile = FREEFILE
  OPEN "TAXTRANS.DAT" FOR RANDOM SHARED AS TranFile LEN = TaxTranRecLen
  
  NumOfRec& = LOF(PayRecFile) \ TaxPayRecLen
  ShowPctComp 1, NumOfRec&
  
  FOR Cnt& = 1 TO NumOfRec&
    GET #PayRecFile, Cnt&, TaxPaymentRec(1)
    
    ThisListRec& = TaxPaymentRec(1).LastPayRec
    DO WHILE ThisListRec& > 0
      GET #PayListFile, ThisListRec&, PayListRec(1)  'get paylist rec

      CLOSE CustFile
      CustFile = FREEFILE
      IF TaxPaymentRec(1).BillType = "R" THEN
        OPEN "RETXCUST.DAT" FOR RANDOM SHARED AS CustFile LEN = TaxCustRecLen
      ELSE
        OPEN "PPTXCUST.DAT" FOR RANDOM SHARED AS CustFile LEN = TaxCustRecLen
      END IF
      GET #CustFile, TaxPaymentRec(1).CustAcct, TaxCustRec(1) 'get cust rec
      GET #TranFile, PayListRec(1).BillRec, TaxTranRec(1) 'get bill trans
      REDIM PayTranRec(1)  AS TaxTransactionType
      'make a new clean payment trans
      '090799 Found bug in multi bill payment. Only by pure luck.
      '      TotalPaid# = Round#(TaxPaymentRec(1).AmtPaid)
      TotalPaid# = Round#(PayListRec(1).Principle1 + PayListRec(1).Principle2 + PayListRec(1).Principle3)
      TotalPaid# = Round#(TotalPaid# + PayListRec(1).Principle4 + PayListRec(1).Principle5)
      TotalPaid# = Round#(TotalPaid# + PayListRec(1).Interest + PayListRec(1).Penalty)
      IF TotalPaid# = 0 THEN
        GOTO SkipThisRec
      END IF

      PayTranRec(1).TransDate = TaxPaymentRec(1).PayDate
      PayTranRec(1).TranType = 2
      PayTranRec(1).Revenue.Principle1Pd = PayListRec(1).Principle1
      'TaxType% = 1 is Real, 2 is Personal
      PayTranRec(1).Revenue.Principle2Pd = PayListRec(1).Principle2
      PayTranRec(1).Revenue.Principle3Pd = PayListRec(1).Principle3
      PayTranRec(1).Revenue.Principle4Pd = PayListRec(1).Principle4
      PayTranRec(1).Revenue.Principle5Pd = PayListRec(1).Principle5
      PayTranRec(1).Revenue.InterestPd = PayListRec(1).Interest
      PayTranRec(1).Revenue.PenaltyPd = PayListRec(1).Penalty
      PayTranRec(1).Amount = TotalPaid#
      PayTranRec(1).Description = TaxPaymentRec(1).DESC
      PayTranRec(1).CustomerRec = TaxPaymentRec(1).CustAcct
      PayTranRec(1).LastTrans = TaxCustRec(1).LastTrans
      PayTranRec(1).BelongTo = PayListRec(1).BillRec
      
      IF TaxPaymentRec(1).BillType = "R" THEN
'*********************************************************
'01-08-2001 Found bug in posting that was using the total
'payment record (TaxPaymentRec) instead of the pay list
'record for each payment (PayListRec)
'This fixes the (multiple bills) single payment bug!!
'*********************************************************
        TaxTranRec(1).Revenue.Principle1Pd = Round#(TaxTranRec(1).Revenue.Principle1Pd + PayTranRec(1).Revenue.Principle1Pd)
        TaxTranRec(1).Revenue.InterestPd = Round#(TaxTranRec(1).Revenue.InterestPd + PayTranRec(1).Revenue.InterestPd)
        TaxTranRec(1).Revenue.PenaltyPd = Round#(TaxTranRec(1).Revenue.PenaltyPd + PayTranRec(1).Revenue.PenaltyPd)
      ELSE
        TaxTranRec(1).Revenue.Principle1Pd = Round#(TaxTranRec(1).Revenue.Principle1Pd + PayTranRec(1).Revenue.Principle1Pd)
        TaxTranRec(1).Revenue.Principle2Pd = Round#(TaxTranRec(1).Revenue.Principle2Pd + PayTranRec(1).Revenue.Principle2Pd)
        TaxTranRec(1).Revenue.Principle3Pd = Round#(TaxTranRec(1).Revenue.Principle3Pd + PayTranRec(1).Revenue.Principle3Pd)
        TaxTranRec(1).Revenue.Principle4Pd = Round#(TaxTranRec(1).Revenue.Principle4Pd + PayTranRec(1).Revenue.Principle4Pd)
        TaxTranRec(1).Revenue.Principle5Pd = Round#(TaxTranRec(1).Revenue.Principle5Pd + PayTranRec(1).Revenue.Principle5Pd)
        TaxTranRec(1).Revenue.InterestPd = Round#(TaxTranRec(1).Revenue.InterestPd + PayTranRec(1).Revenue.InterestPd)
        TaxTranRec(1).Revenue.PenaltyPd = Round#(TaxTranRec(1).Revenue.PenaltyPd + PayTranRec(1).Revenue.PenaltyPd)
      END IF
      PUT #TranFile, PayListRec(1).BillRec, TaxTranRec(1)
      
      NextTransRec& = (LOF(TranFile) \ TaxTranRecLen) + 1
      
      PUT #TranFile, NextTransRec&, PayTranRec(1)
      
      TaxCustRec(1).LastTrans = NextTransRec&
      PUT #CustFile, TaxPaymentRec(1).CustAcct, TaxCustRec(1)
      
SkipThisRec:
      ThisListRec& = PayListRec(1).PrevListRec
      ShowPctComp Cnt&, NumOfRec&
    LOOP
    
    ShowPctComp Cnt&, NumOfRec&
  NEXT
  
  CLOSE
  
  KillFile TaxCPRFileName$
  KillFile TaxLOPFileName$
  
  BlockClear
  
  DisplayTaxScrn "UPDATEOK"
  WaitForAction
  
ExitPayPost:
END SUB

SUB PrintPaymentJournal
  
  REDIM TaxSetUp(1) AS TaxMasterType
  TaxSetupLen = LEN(TaxSetUp(1))
  FGetAH "TAXSETUP.DAT", TaxSetUp(1), TaxSetupLen, 1            'load it
  
  REDIM RERevAmts(1 TO 3) AS DOUBLE
  REDIM PPRevAmts(1 TO 7) AS DOUBLE
  
  REDIM RERevText$(1 TO 3)
  RERevText$(1) = "Tax Principle"
  RERevText$(2) = "Interest"
  RERevText$(3) = "Penalty"
  
  REDIM PPRevText$(1 TO 7)
  PPRevText$(1) = "Pers Prop Tax"
  PPRevText$(2) = "Mach/Tool Tax"
  PPRevText$(3) = "Merch Cap Tax"
  PPRevText$(4) = "Farm Equip Tax"
  PPRevText$(5) = "Mobile Hme Tax"
  PPRevText$(6) = "Interest"
  PPRevText$(7) = "Penalty"
  
  CursorOff
  Oper$ = QPTrim$(STR$(OperNum))
  FF$ = CHR$(12)
  Page = 0
  LineCnt = 0
  MaxLines = 55
  Dash1$ = STRING$(80, "-")
  
  TaxCPRFileName$ = "TAXCPR" + Oper$ + ".DAT"   'Customers Payment Record file
  
  PayJourName$ = "TAXPAY" + QPTrim$(STR$(OperNum)) + ".RPT"
  Header$ = "Utility Payment/Deposit Journal"
  
  REDIM CMOperRec(1) AS CMOperRecType
  CMOperRecLen = LEN(CMOperRec(1))
  
  REDIM TaxPaymentRec(1) AS TaxPaymentRecType
  TaxPayRecLen = LEN(TaxPaymentRec(1))
  
  CMFile = FREEFILE
  OPEN "CMOPER.DAT" FOR RANDOM SHARED AS CMFile LEN = CMOperRecLen
  NumRecs = LOF(CMFile) \ CMOperRecLen
  
  FOR Cnt = 1 TO NumRecs
    GET CMFile, Cnt, CMOperRec(1)
    IF CMOperRec(1).OperatorNumber = OperNum THEN
      Operator$ = QPTrim$(CMOperRec(1).OperatorName)
      EXIT FOR
    END IF
  NEXT
  CLOSE CMFile
  
  IF Exist(TaxCPRFileName$) AND FileSize(TaxCPRFileName$) > 0 THEN
    PayOKFlag = True
  END IF
  
  IF NOT PayOKFlag THEN
    BlockClear
    DisplayTaxScrn "NOPAYJUR"
    QPrintRC STR$(OperNum), 12, 34, 79
    WaitForAction
    GOTO ExitJournal
  END IF
  
  TotalRecs& = FileSize(TaxCPRFileName$) \ TaxPayRecLen
  BlockClear
  
  RptHandle = FREEFILE
  OPEN PayJourName$ FOR OUTPUT AS RptHandle
  ShowProcessingScrn Header$
  
  GOSUB PrintRptHeader
  PayRecFile = FREEFILE
  OPEN TaxCPRFileName$ FOR RANDOM SHARED AS PayRecFile LEN = TaxPayRecLen
  NumOfRecs& = LOF(PayRecFile) \ TaxPayRecLen
  FOR Cnt& = 1 TO NumOfRecs&
    GET #PayRecFile, Cnt&, TaxPaymentRec(1)
    DoneCnt = DoneCnt + 1
'IF TaxPaymentRec(1).BillType = "P" THEN
'  GOTO SkipEm
'END IF
    IF LineCnt >= MaxLines THEN
      PRINT #RptHandle, FF$
      GOSUB PrintRptHeader
    END IF
    IF TaxPaymentRec(1).CashAmt < 0 THEN TaxPaymentRec(1).CashAmt = 0
    IF TaxPaymentRec(1).ChkAmt < 0 THEN TaxPaymentRec(1).ChkAmt = 0
    
    PRINT #RptHandle, Num2Date(TaxPaymentRec(1).PayDate);
    PRINT #RptHandle, TAB(12); USING "#####"; TaxPaymentRec(1).CustAcct;
    PRINT #RptHandle, TaxPaymentRec(1).BillType;
    PRINT #RptHandle, TAB(19); LTRIM$(TaxPaymentRec(1).CustName);
    PRINT #RptHandle, TAB(44); USING "######.##"; TaxPaymentRec(1).CashAmt;
    PRINT #RptHandle, TAB(54); USING "######.##"; TaxPaymentRec(1).ChkAmt;
    PRINT #RptHandle, TAB(64); USING "######.##"; Round#(Round#(TaxPaymentRec(1).ChkAmt + TaxPaymentRec(1).CashAmt) - TaxPaymentRec(1).Change);
    PRINT #RptHandle, TAB(74); USING "####.##"; TaxPaymentRec(1).Change
    
    TotalCash# = Round#(TotalCash# + TaxPaymentRec(1).CashAmt)
    TotalCheck# = Round#(TotalCheck# + TaxPaymentRec(1).ChkAmt)
    TotalAmount# = Round#(TotalAmount# + TaxPaymentRec(1).AmtPaid)
    TotalChange# = Round#(TotalChange# + TaxPaymentRec(1).Change)
    TotalReceipts = TotalReceipts + 1
    LineCnt = LineCnt + 1
    IF TaxPaymentRec(1).BillType = "R" THEN
      FOR RCnt = 1 TO 3
        RERevAmts(RCnt) = Round#(RERevAmts(RCnt) + TaxPaymentRec(1).PaidOwed(RCnt).AmtPaid)
      NEXT
    ELSE
      FOR RCnt = 1 TO 7
        PPRevAmts(RCnt) = Round#(PPRevAmts(RCnt) + TaxPaymentRec(1).PaidOwed(RCnt).AmtPaid)
      NEXT
    END IF
SkipEm:
    ShowPctComp DoneCnt, TotalRecs&
  NEXT
  
  GOSUB PrintRptEnding
  
  CLOSE
  
  LPTPORT = 1
  PrintRptFile Header$, PayJourName$, LPTPORT%, RetCode%, EntryPoint
  
  KillFile PayJourName$
  
ExitJournal:
  EXIT SUB
  
PrintRptHeader:
  Page = Page + 1
  PRINT #RptHandle, "Tax Payment Receipts Journal"
  PRINT #RptHandle, "Posting Date: "; PostDate$
  PRINT #RptHandle, "    Operator: "; Operator$; TAB(70); "Page #"; Page
  PRINT #RptHandle, ""
  PRINT #RptHandle, "                                                                 Paid on"
  PRINT #RptHandle, " Date       Acct       Customer                 Cash     Check   Account  Change"
  PRINT #RptHandle, Dash1$
  LineCnt = 6
  RETURN
  
PrintRptEnding:
  PRINT #RptHandle, Dash1$
  PRINT #RptHandle, "                  Totals: ";
  PRINT #RptHandle, TAB(44); USING "######.##"; TotalCash#;
  PRINT #RptHandle, TAB(54); USING "######.##"; TotalCheck#;
  PRINT #RptHandle, TAB(64); USING "######.##"; TotalAmount#;
  PRINT #RptHandle, TAB(74); USING "####.##"; TotalChange#
  
  PRINT #RptHandle, "Total Number of Receipts: "; USING "####,#"; TotalReceipts
  PRINT #RptHandle, FF$
  PRINT #RptHandle, "Tax Payment Receipts Journal"
  PRINT #RptHandle, "Source Summary"
  PRINT #RptHandle, ""
  PRINT #RptHandle, "Real Estate Receipts"
  PRINT #RptHandle, ""
  PRINT #RptHandle, "     Source"; TAB(34); "Payments"
  PRINT #RptHandle, Dash1$
  REGTotal# = 0
  FOR RCnt = 1 TO 3
    PRINT #RptHandle, TAB(5); RERevText$(RCnt); TAB(30); USING "$$#####,#.##"; RERevAmts(RCnt)
    REGTotal# = Round#(REGTotal# + RERevAmts(RCnt))
    GTotal# = Round#(GTotal# + RERevAmts(RCnt))
  NEXT
  
  PRINT #RptHandle, Dash1$
  PRINT #RptHandle, " Real Estate Grand Total:"; TAB(39); USING "$$######,#.##"; REGTotal#
  PRINT #RptHandle, ""
  PRINT #RptHandle, "Personal Property Receipts"
  PRINT #RptHandle, ""
  PRINT #RptHandle, "     Source"; TAB(34); "Payments"
  PRINT #RptHandle, Dash1$
  REGTotal# = 0
  FOR RCnt = 1 TO 7
    PRINT #RptHandle, TAB(5); PPRevText$(RCnt); TAB(30); USING "$$#####,#.##"; PPRevAmts(RCnt)
    PPGTotal# = Round#(PPGTotal# + PPRevAmts(RCnt))
    GTotal# = Round#(GTotal# + PPRevAmts(RCnt))
  NEXT
  PRINT #RptHandle, Dash1$
  PRINT #RptHandle, " Personal Property Grand Total:"; TAB(39); USING "$$######,#.##"; PPGTotal#
  PRINT #RptHandle, ""
  PRINT #RptHandle, "     Grand Total: All Receipts:"; TAB(39); USING "$$######,#.##"; GTotal#
  PRINT #RptHandle, FF$
  RETURN
  
  
END SUB

SUB ValidateOperator
  
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
    IF OperNum = CMOperRec(1).OperatorNumber AND OperPassword$ = CMOperRec(1).OperatorPassword THEN
      OperOKFlag = Cnt
      EXIT FOR
    END IF
  NEXT
  
ExitValidate:
  CLOSE
  IF OperOKFlag = 0 THEN
    BEEP
    OK = MsgBox("TAX", "BADOPER")
  END IF
  
END SUB

