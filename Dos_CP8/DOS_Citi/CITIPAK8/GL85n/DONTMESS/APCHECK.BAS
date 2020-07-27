'AP check for writing Detail to glif

DEFINT A-Z

'DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB MnuTitle (Title$, Row, Col, MaxLen, Colr)
DECLARE SUB VoidCheckTrans (VoidVendorRecNum&, VoidTransRecNum&, VoidChkNum&)
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB PrintCheckListing ()
DECLARE SUB VoidPrintedCheck ()
DECLARE SUB VoidAPCheck ()
DECLARE SUB OpenPayableReport ()
DECLARE SUB PreAuditRpt ()
DECLARE SUB PostCHKDAT (ChkFileName$)
DECLARE SUB PrintChecks2 (FirstBadNum&, CheckNum&, CheckDate%, LPTPort%, RestartFlag%, BankCode%)
'DECLARE SUB PrintChecks (FirstBadNum&, CheckNum&, CheckDate%, LPTPort%, RestartFlag%, BankCode%)
DECLARE SUB PostAPChecks (RPTOnlyFlag)
DECLARE SUB GetCheckInfo (BYVAL RestartFlag%)
DECLARE SUB CheckPREAudit ()
DECLARE SUB PickList (Items() AS ANY, Picked%(), NPicked%, Cnf AS ANY)
DECLARE SUB SelectOpenPayables ()
DECLARE FUNCTION FindVendorRec% (VendorCode$)
DECLARE SUB PostInvTrans (APType%)
DECLARE SUB StuffBuf (Ky$)
DECLARE SUB OpenAPLedgerFile (APLedgerFile%, NumTrans%, RecLen%)
DECLARE SUB OpenAPDistFile (APDistFile%, NumDistRecs&, RecLen%)
DECLARE SUB OpenAPVendorFile (APVendorFile%, NumVRecs%, VendorRecLen%)
DECLARE SUB OpenVendorFile (VendorFile%, NumVRecs%)
DECLARE SUB OpenVendorIdx (VendorIdxFile%, NumActiveVendors%)
DECLARE FUNCTION ChkForVendor% (ChkVendor$)
DECLARE SUB VertMenuT (Items() AS ANY, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB Prn2Scrn2 (TitleText$(), FileName$, TopRow%, BotRow%, Clr%, PrnFlag%)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB BLPrint (LPTPort%, Text$, ErrCode%)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION Round# (DoubleNum#)
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE SUB FileView (FileName$, Ky, Action, FVI AS ANY, SEG Array)
DECLARE FUNCTION FUsing$ (Number$, Mask$)
DECLARE SUB HideCursor ()
DECLARE FUNCTION MGetKey% (Row%, Col%)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION QPValL& (LongNumber$)
DECLARE FUNCTION QPValI (IntNumber$)
DECLARE FUNCTION ConvDateStr1$ (D$)
DECLARE SUB MScrnSave (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MScrnRest (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MPaintBox (ULRow%, ULCol%, LRRow%, LRCol%, Colr%)
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION Peek1% (Segment%, Address%)
DECLARE SUB QPrint (X$, Colr%, Page%)
DECLARE SUB QPrintRC (t$, r%, c%, Clr%)
DECLARE FUNCTION QPTrim$ (X$)
DECLARE SUB ShowCursor ()
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE SUB KillFile (FileName$)
DECLARE FUNCTION WaitKey% ()
DECLARE FUNCTION FileSize& (FileName$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE FUNCTION SpellNumber$ (Number$)
DECLARE SUB GetFundList (FundList$(), NumFunds)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB ClrBG ()
DECLARE SUB GetPostDates (LPDate, HPDate)
DECLARE FUNCTION GetBankGLAcct$ (BankNum)
DECLARE SUB GetBankList (BankList$(), NumBanks%)
  
  '$INCLUDE: 'DefCnf.bi'                    'monitor and color info
  '$INCLUDE: 'FieldInf.bi'                  'field information
  '$INCLUDE: 'FormEdit.bi'                  'form editing information
  '$INCLUDE: 'pageinfo.BI'                  'Form Page info
  '$INCLUDE: 'QSCR.BI'                      'QuickScreen Routines
  '$INCLUDE: 'GLAUX.BI'                     'Common GL Routines
  '$INCLUDE: 'GL.BI'                        'GL File Types
  '$INCLUDE: 'AP.BI'
  '$INCLUDE: 'SCRAUX.bi'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'PROSCHK.bi'
  
  DIM SHARED User$, FundLen, DetLen, CashAcct$, APAcct$, CDActive$, CDCash$, CDDue$
  
  '$INCLUDE: 'init.BI'
  STACK 8000
  
  TYPE TPayListType
    LedgerRecNum AS LONG
    VendorRecNum AS INTEGER
  END TYPE
  
  TYPE CheckInfoType2
    Ledger1st     AS LONG
    LedgerLst     AS LONG
    StartChk      AS LONG
    LastChk       AS LONG
    ChkAmt        AS DOUBLE
    ChkDate       AS INTEGER
    VendorRecNum  AS INTEGER
  END TYPE
  
  TYPE oCheckInfoType
    Ledger1st     AS LONG
    LedgerLst     AS LONG
    ChkNum        AS LONG
    ChkAmt        AS DOUBLE
    ChkDate       AS INTEGER
    VendorRecNum  AS INTEGER
  END TYPE
  
  TYPE CheckInfoType3
    ListFirst     AS INTEGER
    ListLast      AS INTEGER
    StartChk      AS LONG
    LastChk       AS LONG
    ChkAmt        AS DOUBLE
    ChkDate       AS INTEGER
    VendorRecNum  AS INTEGER
    VoidFlag      AS INTEGER
    BankCode      AS INTEGER 'To track o/s checks by bank
  END TYPE
  
  TYPE LedgerInfoType
    LedDate       AS STRING * 17
    LedInvNum     AS STRING * 25
    INVAMT        AS STRING * 17
  END TYPE
  
  TYPE DistInfoType
    Fill1         AS STRING * 19
    DistAcct      AS STRING * 16
    DistAmt       AS STRING * 23
    'InvAmt        AS STRING * 16
  END TYPE
  
  TYPE CheckRegType
    ChkNum     AS STRING * 10
    ChkDate    AS STRING * 13
    VENDNAME   AS STRING * 30
    ChkAmt     AS STRING * 14
  END TYPE
  
  TYPE FLen
    V AS STRING * 42
  END TYPE
  
  TYPE FLen2
    V AS STRING * 64
  END TYPE
  
  TYPE LedgerInfoType2
    INVDATE       AS STRING * 12
    DUEDATE       AS STRING * 12
    INVNUM        AS STRING * 27
    PONum         AS STRING * 14
    Amt           AS STRING * 12
    DistAcct      AS STRING * 16
    DistAmt       AS STRING * 12
  END TYPE
  
  'TYPE ChkBankCodeType
  '  BankCode AS INTEGER
  'END TYPE


  DIM SHARED VENDOR AS VendorRecType
  
  DIM SHARED VendorIdx AS VendorIdxRecType
  
  CONST True = -1, False = NOT True
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 10)
  
  MChoice$(1) = "Open Payables Report"
  MChoice$(2) = "Select Invoices for Payment"
  MChoice$(3) = "Check Pre-Audit Report"
  MChoice$(4) = "Print A/P Checks"
  MChoice$(5) = "Restart Aborted Check Printing"
  MChoice$(6) = "Void a Printed A/P Check"
  MChoice$(7) = "Print Checks Register"
  MChoice$(8) = "Post A/P Checks"
  MChoice$(9) = "Void a Posted A/P Check"
  MChoice$(10) = "Quit"
  
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
  Help$ = "Check Processing Menu"
  
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    LibFile2Scrn "GL.QSL", "MENUBAK", MonoCode, -1, ErrorCode

    'TitleBox 3, Col, MaxLen + 3, "Check Processing Menu ", Cnf
    'TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

    Title$ = "Check Processing Menu"
    MnuTitle Title$, Row, Col, MaxLen, 15

    PrintTitle User$
    PrintHelp Help$
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      ClrBG
      OpenPayableReport
    CASE 2
      ClrBG
      SelectOpenPayables
    CASE 3
      ClrBG
      PreAuditRpt
    CASE 4
      ClrBG
      GetCheckInfo 0
    CASE 5
      ClrBG
      GetCheckInfo -1
    CASE 6
      ClrBG
      VoidPrintedCheck
    CASE 7
      ClrBG
      PrintCheckListing
    CASE 8
      ClrBG
      PostAPChecks False
    CASE 9
      ClrBG
      VoidAPCheck
    CASE 10
      CALL HideCursor
      CLS
      END
    END SELECT
  LOOP
  
  CALL HideCursor
  ClrBG
  RUN "apMenu"
  END

SUB ClearLedgerPrintCodes
  
  REDIM APLedgerRec(1) AS APLedger81RecType
  'REDIM APLedgerRec(1) AS APLedgerRecType
  LdRecLen = LEN(APLedgerRec(1))
  OpenAPLedgerFile APLedgerFile, NumTrans, LdRecLen
  FOR Cnt = 1 TO NumTrans
    GET APLedgerFile, Cnt, APLedgerRec(1)
    
    APLedgerRec(1).PrintCode = 0
    PUT APLedgerFile, Cnt, APLedgerRec(1)
  NEXT
  CLOSE
  
END SUB

SUB GetCheckInfo (BYVAL RestartFlag%)

  IF NOT Exist("TPAYLIST.LST") THEN EXIT SUB

  GetPostDates LPDate, HPDate
   
  '--Under Construction
  REDIM BankList$(1)
  GetBankList BankList$(), NumBanks
  IF NumBanks > 2 THEN
    MaxChoice = NumBanks
  ELSE
    MaxChoice = 2
  END IF

  SHARED PrnDef$()
  SHARED Choice$()
  
 ' DIM ChkBankCode AS ChkBankCodeType
  
  REDIM Choice$(0 TO MaxChoice, 0 TO 1)
  Choice$(0, 0) = "5"
  Choice$(1, 0) = "Prn#1"
  Choice$(2, 0) = "Prn#2"

'--Under Construction
  IF NumBanks > 0 THEN
    Choice$(0, 1) = "2"
    FOR Cnt = 1 TO NumBanks
      Choice$(Cnt, 1) = BankList$(Cnt)
    NEXT
  END IF

  REDIM Frm(1) AS FormInfo
  
  TNumInvoices = FileSize("TPAYLIST.LST") \ 6
  
  IF TNumInvoices = 0 THEN EXIT SUB
  
  FormName$ = "CHKINF"
  
  IF RestartFlag THEN
    FormName$ = FormName$ + "2"
  ELSE
    FormName$ = FormName$ + "1"
  END IF
  
  FChkNumFld = 2
  LChkNumFld = 3
  ChkDateFld = 4
  PPortFld = 5
  
  NumFlds = LibNumberOfFields("AP", FormName$)
  REDIM Form$(NumFlds, 2)    'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo            'DIM the field information array
  StartEl = 0   'Load first form at array start
  LibGetFldDef "AP", FormName$, StartEl, Fld(), Form$(), ErrCode
  
  '----- Setup TYPE for setting and reading form editing information.
  Frm(1).FldNo = 1           'Start editing on field #1
  Frm(1).InsStat = False     'Set the insert state (-1 = Insert on)
  Frm(1).StartEl = 0         'Set form starting element to 0 and
  
  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  
  CheckDate = Date2Num(DATE$)
  
  Action = 1
  LSET Form$(ChkDateFld, 0) = Num2Date$(CheckDate)
  LSET Form$(PPortFld, 0) = "1"
  
  FirstTime = True
  
  DO            'Pole the editing procedure
    
    RetryFlag = False
    LibFile2Scrn "AP", FormName$, MonoCode, -1, ErrCode
    DO
      EditForm Form$(), Fld(), Frm(1), Cnf, Action
      IF FirstTime THEN
        FirstTime = NOT FirstTime
        LSET Form$(ChkDateFld, 0) = Num2Date(CheckDate)
        LSET Form$(PPortFld, 0) = "1"
        SaveField 1, Form$(), Fld(), BadFld
        Action = 1
        Frm(1).FldNo = 1
        PrintHelp " Ready to Print Checks"
      END IF
      
      IF Frm(1).Presses AND Frm(1).MRow = 25 THEN
        SELECT CASE Frm(1).MCol
        CASE 2 TO 13            'F1
          'PressButton F1Key, 25, 2, 13
        END SELECT
      END IF
      
      SELECT CASE Frm(1).KeyCode
      CASE F10Key
        'add logic to verify user entry
        CheckDate = Date2Num(Form$(ChkDateFld, 0))
        BankCode = VAL(LEFT$(Form$(2, 0), 2))
        CheckNum& = VAL(QPTrim$(Form$(LChkNumFld, 0)))

        BadFld = 0
        IF CheckDate < LPDate OR CheckDate > HPDate THEN
          OK = MsgBox("GL", "BADDATE")
          BadFld = ChkDateFld
        END IF
        
        IF BankCode < 1 THEN
          BadFld = 2
        END IF

        IF CheckNum& < 1 THEN
          BadFld = LChkNumFld
        END IF
        
        IF RestartFlag THEN
          FirstBadNum& = QPValL(Form$(FChkNumFld, 0))
        END IF
        
        IF BadFld = False THEN Ok2Print = True
        IF Ok2Print THEN
          BankCode = VAL(Form$(2, 0))
          CheckDate = Date2Num(Form$(ChkDateFld, 0))
          CheckNum& = VAL(QPTrim$(Form$(LChkNumFld, 0)))
          SELECT CASE Form$(PPortFld, 0)
          CASE "Prn#1"
            LPTPort = 1
          CASE "Prn#2"
            LPTPort = 2
          CASE ELSE
            LPTPort = 1
          END SELECT
          
          PrintChecks2 FirstBadNum&, CheckNum&, CheckDate, LPTPort, RestartFlag, BankCode
          ExitFlag = True
        ELSE
          BEEP
          Frm(1).FldNo = BadFld
          Action = 1
        END IF
        
      CASE F5KEY
        
        SELECT CASE Form$(PPortFld, 0)
        CASE "Prn#1"
          LPTPort = 1
        CASE "Prn#2"
          LPTPort = 2
        CASE ELSE
          LPTPort = 1
        END SELECT
        
        'CursorOff
        'BlockClear
        'LPTPort = QPValI(Form$(PPortFld, 0))
        'REDIM PRNSet(1 TO 16) AS INTEGER
        'FGetAH PrinterSetUPFile, PRNSet(1), 2, 16
        'RPTPitch = PRNSet(15)
        'ERASE PRNSet
        
        'SELECT CASE RPTPitch
        'CASE 10
        '  ToPrint$ = PrnDef$(2)
        'CASE 12
        '  ToPrint$ = PrnDef$(3)
        'CASE 17
        '  ToPrint$ = PrnDef$(4)
        'CASE ELSE
        'END SELECT
        
        'BLPrint LPTPort, PrnDef$(1), ErrCount
        'BLPrint LPTPort, ToPrint$, ErrCount
        PrintRptFile "", "APCHECK.MSK", LPTPort, RetCode, 5
        'BLPrint LPTPort, PrnDef$(1), ErrCount
        
        CheckNum& = QPValL(Form$(LChkNumFld, 0))
        'CheckNum& = VAL(Form$(LChkNumFld, 0))
        'CheckDate = Date2Num(Form$(ChkDateFld, 0))
        
        IF RetCode = 0 AND CheckNum& > 0 THEN
          CheckNum& = CheckNum& + 1
          'CheckNum& = VAL(QPTrim$(Form$(LChkNumFld, 0)))
          LSET Form$(LChkNumFld, 0) = QPTrim$(STR$(CheckNum&))
          SaveField LChkNumFld, Form$(), Fld(), BadFld
          UnPackBuffer 0, 0, Form$(), Fld()
        END IF
        RetryFlag = True
        Action = 2
        'FirstTime = True
      CASE EscKey
        ExitFlag = True
      END SELECT
      
    LOOP UNTIL ExitFlag OR RetryFlag
    
  LOOP UNTIL ExitFlag           '
  
  'CursorOff
  
  ERASE Form$, Fld, Frm                ', TempScrn
  
END SUB

SUB OpenPayableReport
  
  '  SHARED VENDOR AS VendorRecType
  
  SHARED Choice$()
  REDIM Choice$(0 TO 3, 1)
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn#1"
  Choice$(3, 0) = "Prn#2"
  
  REDIM Frm(1) AS FormInfo
  
  FormName$ = "PRNSPEC"
  
  NumFlds = LibNumberOfFields("GL", FormName$)
  REDIM Form$(NumFlds, 2)    'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo            'DIM the field information array
  StartEl = 0   'Load first form at array start
  LibGetFldDef "GL", FormName$, StartEl, Fld(), Form$(), ErrCode
  
  '----- Setup TYPE for setting and reading form editing information.
  Frm(1).FldNo = 1           'Start editing on field #1
  Frm(1).InsStat = False     'Set the insert state (-1 = Insert on)
  Frm(1).StartEl = 0         'Set form starting element to 0 and
  
  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1
  
  FirstTime = True
  
  LibFile2Scrn "GL", FormName$, MonoCode, -1, ErrCode
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF Frm(1).KeyCode = EscKey THEN
      ExitFlag = True
    ELSEIF Frm(1).KeyCode = F10Key THEN
      OkFlag = True
    END IF
  LOOP UNTIL ExitFlag OR OkFlag
  IF ExitFlag THEN GOTO oExitPreAudit
  
  LPTPort = VAL(RIGHT$(QPTrim$(Form$(1, 0)), 1))
  
  'IF LPTPort = 0 THEN LPTPort = 1
  
  FF$ = CHR$(12)
  MaxLines = 50
  Dash2$ = STRING$(78, "=")
  Dash$ = Dash2$
  MID$(Dash2$, 1, 7) = SPACE$(7)
  PageNum = 0
  
  REDIM FundList$(1)
  GetFundList FundList$(), NumFunds
  IF NumFunds = 0 THEN EXIT SUB
  
  REDIM FundAmts(1 TO NumFunds) AS DOUBLE
  
  REDIM APLedgerRec(1) AS APLedger81RecType
  'REDIM APLedgerRec(1) AS APLedgerRecType
  REDIM APDistRec(1) AS APDistRecType
  REDIM LedInfo(1) AS LedgerInfoType2
  REDIM DistInfo(1) AS DistInfoType
  REDIM ChkRegInfo(1) AS CheckRegType
  
  NumVendRecs = (FileSize("apvendor.idx") \ 12)
  REDIM VIndex(1 TO NumVendRecs) AS VendorIdxRecType
  FGetAH "apvendor.idx", VIndex(1), 12, NumVendRecs
  
  DistInfo(1).Fill1 = ""
  APDistRecLen = LEN(APDistRec(1))
  RecLen = LEN(APLedgerRec(1))
  
  PrintFile = FREEFILE
  
  OPEN "OPENPAYB.PRN" FOR OUTPUT AS PrintFile
  
  OpenVendorFile VendorFile, NumVRecs
  OpenAPLedgerFile APLedgerFile, NumTrans, RecLen
  OpenAPDistFile APDistFile, NumDistRecs&, APDistRecLen
  
  GOSUB PrintOpenPayRptHeader
  
  FOR VCnt = 1 TO NumVendRecs
    DoneVHeader = 0
    GET VendorFile, VIndex(VCnt).RecNum, VENDOR

    'IF VENDOR.DelFlag <> 0 THEN STOP

    NextTrans& = VENDOR.FrstTran
    DO UNTIL NextTrans& = 0
      GET APLedgerFile, NextTrans&, APLedgerRec(1)
      IF APLedgerRec(1).TrCode = 1 AND APLedgerRec(1).PAYCODE = 1 THEN
        IF NOT DoneVHeader THEN
          IF LineCnt > MaxLines THEN
            PRINT #PrintFile, FF$
            LineCnt = 0
          END IF
          GOSUB oPrintVendHeader
          GOSUB oPAInvHeader
        END IF
        GOSUB oPrintDist
        IF LineCnt > MaxLines THEN
          PRINT #PrintFile, FF$
          GOSUB PrintOpenPayRptHeader
          GOSUB oPrintVendHeader
          GOSUB oPAInvHeader
        END IF

      END IF
      NextTrans& = APLedgerRec(1).NextTrans
    LOOP

    IF DoneVHeader THEN
      GOSUB oFinishVendor
    END IF
    IF LineCnt > MaxLines THEN
      PRINT #PrintFile, FF$
      GOSUB PrintOpenPayRptHeader
    END IF
  NEXT

  PRINT #PrintFile, 'JB
  PRINT #PrintFile, Dash$
  IF LineCnt > MaxLines THEN
    PRINT #PrintFile, FF$
  END IF
  GOSUB FinishOpenReport
  PRINT #PrintFile, FF$
  
  CLOSE
  
  ERASE FundList$, FundAmts, APLedgerRec, APDistRec
  ERASE LedInfo, DistInfo, ChkRegInfo           ', ChkInfo
  ERASE Form$, Fld, Fld
  
  IF LPTPort > 0 THEN
    EntryPoint = 5
  ELSE
    EntryPoint = 2
  END IF
  Title$ = "Open Payables Report"
  PrintRptFile Title$, "OPENPAYB.PRN", LPTPort%, RetCode%, EntryPoint
  
oExitPreAudit:
  
  
  EXIT SUB
  
FinishOpenReport:
  'PageNum = PageNum + 1
  'Page$ = FUsing(STR$(PageNum), "###")
  PRINT #PrintFile, "Report Totals:"
  PRINT #PrintFile, "Vendors with Open Invoices: "; FUsing(STR$(ChkCnt), ",#############")
  PRINT #PrintFile, "                  Totaling: "; FUsing(STR$(TotalChkAmt#), "$$########,.##")
  'PRINT #PrintFile,
  'LineCnt = 7 who cares now?
  RETURN

PrintOpenPayRptHeader:
  PageNum = PageNum + 1
  Page$ = FUsing(STR$(PageNum), "###")
  PRINT #PrintFile, "A/P Open Payables Report                                              Page:" + Page$
  PRINT #PrintFile, "Run Date: " + DATE$
  LineCnt = 2
RETURN
  
oPAInvHeader:
  PRINT #PrintFile,             'Dash2$
  PRINT #PrintFile, "Inv Date    Due Date    Inv Num                    PO                  Amount"
  PRINT #PrintFile, "----------  ----------  -------------------------  ----------    ------------"
  LineCnt = LineCnt + 3
RETURN
  
oPrintDist:
  LSET ChkRegInfo(1).VENDNAME = VENDOR.VNAME
  LSET LedInfo(1).INVDATE = Num2Date(APLedgerRec(1).TRDate)
  LSET LedInfo(1).DUEDATE = Num2Date(APLedgerRec(1).DUEDATE)
  LSET LedInfo(1).INVNUM = LEFT$(APLedgerRec(1).DOCNum, 25)
  LSET LedInfo(1).PONum = LEFT$(APLedgerRec(1).PONum, 10)
  RSET LedInfo(1).Amt = FUsing(STR$(APLedgerRec(1).Amt), ",########.##")
  VendTotal# = Round(VendTotal# + APLedgerRec(1).Amt)
  TotalChkAmt# = Round(TotalChkAmt# + APLedgerRec(1).Amt)
  PRINT #PrintFile, LedInfo(1).INVDATE; LedInfo(1).DUEDATE; LedInfo(1).INVNUM; LedInfo(1).PONum; LedInfo(1).Amt
  LineCnt = LineCnt + 1

  'NextDist& = APLedgerRec(1).FrstDist
  'DO UNTIL NextDist& = 0
  '  GET APDistFile, NextDist&, APDistRec(1)
  '  IF LineCnt > MaxLines THEN
  '    PRINT #PrintFile, FF$
  '    GOSUB PrintOpenPayRptHeader
  '    'GOSUB PADistHeader
  '  END IF
  '  LSET LedInfo(1).InvDate = ""
  '  LSET LedInfo(1).DueDate = ""
  '  LSET LedInfo(1).InvNum = ""
  '  LSET LedInfo(1).PONum = ""
  '  RSET LedInfo(1).Amt = ""

  '  LSET LedInfo(1).DistAcct = APDistRec(1).DistAcctNum
  '  RSET LedInfo(1).DistAmt = FUsing(STR$(APDistRec(1).DistAmt), ",########.##")
  'PRINT #PrintFile, DistInfo(1).Fill1; DistInfo(1).DistAcct; DistInfo(1).DistAmt
  '  PRINT #PrintFile, LedInfo(1).InvDate; LedInfo(1).DueDate; LedInfo(1).InvNum; LedInfo(1).PONum; LedInfo(1).Amt; "  "; LedInfo(1).DistAcct; LedInfo(1).DistAmt

  '  LineCnt = LineCnt + 1
  '  ThisFund$ = LEFT$(APDistRec(1).DistAcctNum, FundLen)
  '  FOR FundCnt = 1 TO NumFunds
  '    IF ThisFund$ = FundList$(FundCnt) THEN
  '      FundAmts(FundCnt) = Round(FundAmts(FundCnt) + APDistRec(1).DistAmt)
  '      EXIT FOR
  '    END IF
  '  NEXT
  '  NextDist& = APDistRec(1).NextDist
  'LOOP
RETURN
  
oPrintVendHeader:
  PRINT #PrintFile,
  PRINT #PrintFile, Dash$
  PRINT #PrintFile, VENDOR.VNum; VENDOR.VNAME
  DoneVHeader = -1
  LineCnt = LineCnt + 3
  RETURN
  
oFinishVendor:
  PRINT #PrintFile, TAB(66); "------------"
  PRINT #PrintFile, TAB(50); "Vendor Total: "; TAB(66); FUsing(STR$(VendTotal#), ",########.##")
  VendTotal# = 0
  ChkCnt = ChkCnt + 1
  LineCnt = LineCnt + 2
  VActive = 0
  RETURN
  
END SUB

SUB PostAPChecks (RPTOnlyFlag)
  
  '--if there are no checks to post then get outta here
  IF NOT Exist("APCHKINF.DAT") THEN EXIT SUB

  SHARED Choice$()
  REDIM Choice$(0 TO 3, 1)
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Prn#1"
  Choice$(2, 0) = "Prn#2"
  Choice$(3, 0) = ""

  LibName$ = "GL"
  FormName$ = "PRNSPEC"
  NumFlds = LibNumberOfFields(LibName$, FormName$)
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)           '--DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo   '--DIM the field information array
  StartEl = 0                       '--Load first form at array start
  LibGetFldDef LibName$, FormName$, StartEl, Fld(), Form$(), ErrCode
  Frm(1).FldNo = 1                  '--Start editing on field #1
  Frm(1).InsStat = False            '--Set the insert state (-1 = Insert on)
  Frm(1).StartEl = 0                '--Set form starting element to 0 and
  Action = 1                        '--Init form edit
  
  LibFile2Scrn LibName$, FormName$, MonoCode, -1, ErrCode

  FirstTime = True

  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF Frm(1).KeyCode = EscKey THEN
      ExitFlag = True
    ELSEIF Frm(1).KeyCode = F10Key THEN
      OkFlag = True
    END IF
  LOOP UNTIL ExitFlag OR OkFlag
  IF ExitFlag THEN GOTO ExitPost
  LPTPort = VAL(RIGHT$(QPTrim$(Form$(1, 0)), 1))
  
  FF$ = CHR$(12)
  MaxLines = 53
  ToPrint$ = SPACE$(78)
  Dash$ = STRING$(75, "-")
  Dash2$ = STRING$(61, "-")
  MID$(Dash2$, 1, 7) = SPACE$(7)
  PageNum = 0
  
  '--For Central Depository - number of characters to pad detail code with
  IF CDActive$ = "Y" THEN
    PadChars = DetLen - FundLen
    IF PadChars > 0 THEN
      DetPad$ = STRING$(PadChars, "0")
    END IF
  END IF

  REDIM FundList$(1)
  GetFundList FundList$(), NumFunds
  IF NumFunds = 0 THEN EXIT SUB
  
  REDIM FundAmts(1 TO NumFunds) AS DOUBLE
  REDIM APLedgerRec(1) AS APLedger81RecType
  REDIM APLedgerRec2(1) AS APLedger81RecType
  REDIM APDistRec(1) AS APDistRecType
  REDIM LedInfo(1) AS LedgerInfoType
  REDIM DistInfo(1) AS DistInfoType
  REDIM ChkRegInfo(1) AS CheckRegType
  
  DistInfo(1).Fill1 = ""
  APDistRecLen = LEN(APDistRec(1))
  RecLen = LEN(APLedgerRec(1))
  
  TPayCnt = (FileSize("TPAYLIST.LST") \ 6)
  IF TPayCnt = 0 THEN EXIT SUB
  REDIM TPayArray(1 TO TPayCnt) AS TPayListType
  FGetAH "TPAYLIST.LST", TPayArray(1), 6, TPayCnt
  
  REDIM ChkInfo(1 TO 1) AS CheckInfoType3
  ChkInfoRecLen = LEN(ChkInfo(1))
  ChkCnt = (FileSize("APCHKINF.DAT") \ ChkInfoRecLen)
  
  IF ChkCnt = 0 THEN EXIT SUB
  
  REDIM ChkInfo(1 TO ChkCnt) AS CheckInfoType3
  FGetAH "APCHKINF.DAT", ChkInfo(1), ChkInfoRecLen, ChkCnt
  CheckDate$ = Num2Date(ChkInfo(1).ChkDate)

  '--If we are not using a central depository, then use the
  '--cash account$ assigned with the record (glbank.dat)
  IF CDActive$ <> "Y" THEN
    CashAcct$ = GetBankGLAcct$(ChkInfo(1).BankCode)
  END IF
  
  PrintFile = FREEFILE
  PrintFileName$ = "APCHKREG.PRN"
  OPEN PrintFileName$ FOR OUTPUT AS PrintFile

  APIFFileName$ = "APCHKIF.DAT"
  KillFile APIFFileName$
  REDIM GLifRec(1) AS GLTransRecType
  GLIFRecLen = LEN(GLifRec(1))
  GLIFFile = FREEFILE
  OPEN APIFFileName$ FOR RANDOM AS GLIFFile LEN = GLIFRecLen
  
  OpenVendorFile VendorFile, NumVRecs
  OpenAPLedgerFile APLedgerFile, NumTrans, RecLen
  OpenAPDistFile APDistFile, NumDistRecs&, APDistRecLen
  
  GOSUB DoChkRegHeader
  
  '--for each check
  FOR Cnt = 1 TO ChkCnt

    '--if the check was'nt voided
    IF ChkInfo(Cnt).VoidFlag = False THEN

      '--Create a list to hold the check's distribution amounts by fund
      REDIM ThisChkFunds(1 TO NumFunds) AS DOUBLE

      '--update vendor's pointer to their last transaction
      GET VendorFile, ChkInfo(Cnt).VendorRecNum, VENDOR
      NumTrans = NumTrans + 1
      IF VENDOR.LastTran > 0 THEN
        GET APLedgerFile, VENDOR.LastTran, APLedgerRec(1)
        APLedgerRec(1).NextTrans = NumTrans
        PUT APLedgerFile, VENDOR.LastTran, APLedgerRec(1)
      END IF
      VENDOR.LastTran = NumTrans
      PUT VendorFile, ChkInfo(Cnt).VendorRecNum, VENDOR
      
      '---Post the check to the apledger file
      APLedgerRec2(1).VIN = VENDOR.VIN
      APLedgerRec2(1).VendorCode = VENDOR.VNum
      APLedgerRec2(1).VRecNum = ChkInfo(Cnt).VendorRecNum
      APLedgerRec2(1).TRDate = ChkInfo(Cnt).ChkDate
      APLedgerRec2(1).DOCNum = STR$(ChkInfo(Cnt).LastChk)
      APLedgerRec2(1).TrCode = 3
      APLedgerRec2(1).PAYCODE = 3
      APLedgerRec2(1).Amt = ChkInfo(Cnt).ChkAmt
      APLedgerRec2(1).NextTrans = 0
      PUT APLedgerFile, NumTrans, APLedgerRec2(1)

      'write out cent dep cash as a single amount
      '--IF Cent Dep write out check here
      IF CDActive$ = "Y" THEN
        '--Credit Central Depository Cash
        TRec& = TRec& + 1
        GLifRec(1).TRDate = ChkInfo(1).ChkDate
        GLifRec(1).AcctNum = CDCash$
        GLifRec(1).Desc = VENDOR.VNAME
        GLifRec(1).DrAmt = 0
        GLifRec(1).CrAmt = ChkInfo(Cnt).ChkAmt
        GLifRec(1).Src = "CK" + ConvDateStr$(DATE$)
        GLifRec(1).Ref = STR$(ChkInfo(Cnt).LastChk)
        PUT GLIFFile, TRec&, GLifRec(1)
      END IF
     
      '--check lines printed to see if we need a page break
      IF LineCnt > MaxLines THEN
        PRINT #PrintFile, FF$
        GOSUB DoChkRegHeader
      END IF
      
      '--print this check to the report
      RSET ChkRegInfo(1).ChkNum = STR$(ChkInfo(Cnt).LastChk)
      RSET ChkRegInfo(1).ChkDate = Num2Date(ChkInfo(Cnt).ChkDate)
      RSET ChkRegInfo(1).ChkAmt = FUsing(STR$(ChkInfo(Cnt).ChkAmt), ",########.##")
      LSET ChkRegInfo(1).VENDNAME = VENDOR.VNAME
      PRINT #PrintFile, Dash$
      PRINT #PrintFile, ChkRegInfo(1).ChkNum; ChkRegInfo(1).ChkDate; "   ";
      PRINT #PrintFile, ChkRegInfo(1).VENDNAME; "    "; ChkRegInfo(1).ChkAmt
      LineCnt = LineCnt + 2
      
      '--For each invoice being paid by this check...
      FOR LedCnt& = ChkInfo(Cnt).ListFirst TO ChkInfo(Cnt).ListLast

        '--update the invoice record with paid check num and date
        GET APLedgerFile, TPayArray(LedCnt&).LedgerRecNum, APLedgerRec(1)
        APLedgerRec(1).PDCheckDate = ChkInfo(Cnt).ChkDate
        APLedgerRec(1).PDCheckNum = ChkInfo(Cnt).LastChk
        APLedgerRec(1).PAYCODE = 3
        PUT APLedgerFile, TPayArray(LedCnt&).LedgerRecNum, APLedgerRec(1)

        '--print the invoice header
        GOSUB InvHeader

        '--check for a page break
        IF LineCnt > MaxLines THEN
          PRINT #PrintFile, FF$
          GOSUB DoChkRegHeader
          GOSUB InvHeader
        END IF

        '--Print the invoice to the report
        RSET LedInfo(1).LedDate = Num2Date(APLedgerRec(1).TRDate)
        LSET LedInfo(1).LedInvNum = APLedgerRec(1).DOCNum
        RSET LedInfo(1).INVAMT = FUsing(STR$(APLedgerRec(1).Amt), ",########.##")
        TotalChkAmt# = Round(TotalChkAmt# + APLedgerRec(1).Amt)
        PRINT #PrintFile, LedInfo(1).LedDate; "  "; LedInfo(1).LedInvNum; LedInfo(1).INVAMT
        LineCnt = LineCnt + 2

        '--Print the distribution header
        GOSUB DistHeader
        
        '--print the distributions of the invoice
        NextDist& = APLedgerRec(1).FrstDist
        DO UNTIL NextDist& = 0
          GET APDistFile, NextDist&, APDistRec(1)
          IF LineCnt > MaxLines THEN
            PRINT #PrintFile, FF$
            GOSUB DoChkRegHeader
            GOSUB DistHeader
          END IF
          RSET DistInfo(1).DistAcct = APDistRec(1).DistAcctNum
          RSET DistInfo(1).DistAmt = FUsing(STR$(APDistRec(1).DistAmt), ",########.##")
          PRINT #PrintFile, DistInfo(1).Fill1; DistInfo(1).DistAcct; "   "; DistInfo(1).DistAmt
          LineCnt = LineCnt + 1

          '--Summarize the distributions by fund
          ThisFund$ = LEFT$(APDistRec(1).DistAcctNum, FundLen)
          FOR FundCnt = 1 TO NumFunds
            IF ThisFund$ = FundList$(FundCnt) THEN
              '--Update grand total by fund
              FundAmts(FundCnt) = Round(FundAmts(FundCnt) + APDistRec(1).DistAmt)
              '--Update check totals by fund
              ThisChkFunds(FundCnt) = Round(ThisChkFunds(FundCnt) + APDistRec(1).DistAmt)
              EXIT FOR
            END IF
          NEXT

          NextDist& = APDistRec(1).NextDist
        LOOP '--distribution printing
      NEXT  '--Invoice being paid
      
      PRINT #PrintFile,
      
      '--get the record number for a new distribution record
      NextDistRec& = (LOF(APDistFile) \ APDistRecLen) + 1
      FirstDist& = NextDistRec&
      
      '--see how many distributions we need to write
      Num2Write = 0
      FOR ChkDist = 1 TO NumFunds
        IF ThisChkFunds(ChkDist) > 0 THEN  'what if we have negative entries?
          Num2Write = Num2Write + 1
        END IF
      NEXT
       
      '--Write out check distribution by fund to the apdist.dat file
      REDIM NewDistRec(1)  AS APDistRecType
      FOR ChkFDist = 1 TO NumFunds
        IF ThisChkFunds(ChkFDist) > 0 THEN
          Dist = Dist + 1
          IF Dist = Num2Write THEN
            NewDistRec(1).APLedgerRec = NumTrans
            NewDistRec(1).DistAcctNum = FundList$(ChkFDist) + APAcct$
            NewDistRec(1).DistAmt = ThisChkFunds(ChkFDist)
            NewDistRec(1).NextDist = 0
            PUT APDistFile, NextDistRec&, NewDistRec(1)
            Dist = 0
            EXIT FOR
          ELSE
            NewDistRec(1).APLedgerRec = NumTrans
            NewDistRec(1).DistAcctNum = FundList$(ChkFDist) + APAcct$
            NewDistRec(1).DistAmt = ThisChkFunds(ChkFDist)
            NewDistRec(1).NextDist = NextDistRec& + 1
            PUT APDistFile, NextDistRec&, NewDistRec(1)
            NextDistRec& = NextDistRec& + 1
          END IF
        END IF
      NEXT

      '--update ledger rec pointers to first and last dist here
      APLedgerRec2(1).FrstDist = FirstDist&
      APLedgerRec2(1).LastDist = (FirstDist& + Num2Write) - 1
      PUT APLedgerFile, NumTrans, APLedgerRec2(1)
      
      '--Write out the GL entries to update cash and a/p
      '--creates one entry per check per fund
      FOR ChkFDist = 1 TO NumFunds

        IF ThisChkFunds(ChkFDist) > 0 THEN
          '--Credit Cash or Due to Central Depository
          TRec& = TRec& + 1
          GLifRec(1).Src = "CK" + ConvDateStr$(DATE$)
          GLifRec(1).AcctNum = FundList$(ChkFDist) + CashAcct$
          GLifRec(1).TRDate = ChkInfo(1).ChkDate
          GLifRec(1).Desc = VENDOR.VNAME
          GLifRec(1).Ref = STR$(APLedgerRec(1).PDCheckNum)
          GLifRec(1).CrAmt = ThisChkFunds(ChkFDist)
          GLifRec(1).DrAmt = 0
          PUT GLIFFile, TRec&, GLifRec(1)
          '--Debit A/P
          TRec& = TRec& + 1
          GLifRec(1).AcctNum = FundList$(ChkFDist) + APAcct$
          GLifRec(1).CrAmt = 0
          GLifRec(1).DrAmt = ThisChkFunds(ChkFDist)
          PUT GLIFFile, TRec&, GLifRec(1)
          '--update Central Depository
          IF CDActive$ = "Y" THEN
            '--Debit Central Depository Due From
            TRec& = TRec& + 1
            GLifRec(1).AcctNum = CDDue$ + FundList$(ChkFDist) + DetPad$
            GLifRec(1).DrAmt = ThisChkFunds(ChkFDist)
            GLifRec(1).CrAmt = 0
            PUT GLIFFile, TRec&, GLifRec(1)
          END IF
        END IF

      NEXT

    ELSE
      UVoidCnt = UVoidCnt + 1
    END IF

  NEXT 'check
  
  '--Finish Check Detail Report
  PRINT #PrintFile, Dash$
  PRINT #PrintFile, FF$
  LineCnt = 1
  PageNum = 0
  GOSUB DoFundHeader
  FOR Cnt = 1 TO NumFunds
    IF FundAmts(Cnt) > 0 THEN
      PRINT #PrintFile, "      Fund "; FUsing(STR$(VAL(FundList$(Cnt))), ",####"); "  Amt  "; FUsing(STR$(FundAmts(Cnt)), ",########.##")
      LineCnt = LineCnt + 1
    END IF
  NEXT
  PRINT #PrintFile, FF$
  CLOSE

  '--Create the A/P Check Journal Entry report
  GLIFFile = FREEFILE
  OPEN APIFFileName$ FOR RANDOM AS GLIFFile LEN = GLIFRecLen
  NumIFRecs = LOF(GLIFFile) \ GLIFRecLen

  REDIM SortTrans(NumIFRecs) AS GLTransRecType
  FOR i = 1 TO NumIFRecs
    GET GLIFFile, i, GLifRec(1)
    SortTrans(i) = GLifRec(1)
  NEXT
  CLOSE

  SortT SortTrans(1), NumIFRecs, 0, 96, 2, 14

  '--print the journal entry report
  RptFile = FREEFILE
  RptFileName$ = "APCHKJE.PRN"
  OPEN RptFileName$ FOR OUTPUT AS RptFile
  JE$ = SPACE$(80)
  GOSUB JEHeader

  FOR i = 1 TO NumIFRecs
    '--Print Entry
    LSET JE$ = ""
    MID$(JE$, 1) = SortTrans(i).AcctNum
    MID$(JE$, 20) = LEFT$(SortTrans(i).Desc, 19)
    MID$(JE$, 40) = SortTrans(i).Ref
    MID$(JE$, 49) = FUsing(STR$(SortTrans(i).DrAmt), "#######,.##")
    MID$(JE$, 63) = FUsing(STR$(SortTrans(i).CrAmt), "#######,.##")
    PRINT #RptFile, JE$

    JEDr# = JEDr# + SortTrans(i).DrAmt
    JECr# = JECr# + SortTrans(i).CrAmt

    JELineCnt = JELineCnt + 1
    IF JELineCnt > 55 THEN
       PRINT #RptFile, FF$
       GOSUB JEHeader
    END IF
  NEXT

  PRINT #RptFile,
  PRINT #RptFile, "Journal Entry Totals";
  PRINT #RptFile, TAB(49); FUsing(STR$(JEDr#), "#######,.##");
  PRINT #RptFile, TAB(63); FUsing(STR$(JECr#), "#######,.##")
  PRINT #RptFile, FF$
  CLOSE

  '--Post and clean up
  Post2GL APIFFileName$, BadAcct
  KillFile APIFFileName$
  PostCHKDAT "APCHKINF.DAT"
  KillFile "APCHKINF.DAT"
  KillFile "TPAYLIST.LST"
  
  ERASE FundList$, FundAmts, APLedgerRec, APLedgerRec2, APDistRec
  ERASE LedInfo, DistInfo, ChkRegInfo, ChkInfo
  ERASE Form$, Fld, GLifRec
  
  EntryPoint = 0
  Title$ = "A/P Check Register"
  PrintRptFile Title$, PrintFileName$, LPTPort%, RetCode%, EntryPoint

  EntryPoint = 0
  Title$ = "A/P Check Journal Entries"
  PrintRptFile Title$, RptFileName$, LPTPort%, RetCode%, EntryPoint

ExitPost:
  
EXIT SUB


JEHeader:
  PRINT #RptFile, "A/P Check Journal Entries"
  PRINT #RptFile, "Posting Date: " + CheckDate$
  PRINT #RptFile,
  MID$(JE$, 1) = "Account Number"
  MID$(JE$, 20) = "Description"
  MID$(JE$, 40) = "Reference"
  MID$(JE$, 49) = "     Debit"
  MID$(JE$, 63) = "    Credit"
  PRINT #RptFile, JE$
  LSET JE$ = ""
  JELineCnt = 4
RETURN

  
DoFundHeader:
  PageNum = PageNum + 1
  Page$ = FUsing(STR$(PageNum), "###")
  
  Text1$ = "A/P Check Register Detail                                         Page:" + Page$
  Text2$ = "Total Checks Issued        "
  PRINT #PrintFile, Text1$
  PRINT #PrintFile, Dash$
  PRINT #PrintFile,
  PRINT #PrintFile, Text2$; FUsing(STR$(ChkCnt - UVoidCnt), ",####")
  PRINT #PrintFile, "           Totaling  "; FUsing(STR$(TotalChkAmt#), "$$########,.##")
  PRINT #PrintFile,
  PRINT #PrintFile, " By Fund:"
  LineCnt = 7
  RETURN
  
DoChkRegHeader:
  PageNum = PageNum + 1
  Page$ = FUsing(STR$(PageNum), "###")
  Text1$ = "A/P Check Register Summary                                        Page:" + Page$
  
  PRINT #PrintFile, Text1$
  PRINT #PrintFile,
  PRINT #PrintFile, " Check No.   Check Date   Vendor Name                         Check Amount"
  LineCnt = 3
  RETURN
  
InvHeader:
  PRINT #PrintFile, Dash2$
  PRINT #PrintFile, "       Inv Date    Inv Num                            Inv Amt"
  LineCnt = LineCnt + 2
  RETURN
  
DistHeader:
  PRINT #PrintFile, "                   Dist Acct                         Dist Amt"
  LineCnt = LineCnt + 1
  RETURN
  
END SUB

SUB PostCHKDAT (ChkFileName$)

  REDIM OSIFRec(1)    AS OSChkRecType
  OSChkRecLen = LEN(OSIFRec(1))
  
  REDIM CheckInfo(1) AS CheckInfoType3
  ChkRecLen = LEN(CheckInfo(1))
  ChkNumRec = FileSize(ChkFileName$) \ ChkRecLen
  REDIM CheckInfo(1 TO ChkNumRec) AS CheckInfoType3
  FGetAH ChkFileName$, CheckInfo(1), ChkRecLen, ChkNumRec
  
  OSChkFile$ = "CRCHK.DAT"
  OSChkHandle = FREEFILE
  OPEN OSChkFile$ FOR RANDOM AS #OSChkHandle LEN = OSChkRecLen
  OSChkNumRec = (LOF(OSChkHandle) \ OSChkRecLen) + 1
  OpenVendorFile VendorFile, NumVRecs
  
  FOR Cnt = 1 TO ChkNumRec
    GET VendorFile, CheckInfo(Cnt).VendorRecNum, VENDOR

    GLIFDate$ = Num2Date(CheckInfo(Cnt).ChkDate)
    OSIFRec(1).ChkDate = ConvDateStr1$(GLIFDate$)
    OSIFRec(1).ChkNum = CheckInfo(Cnt).LastChk
    OSIFRec(1).Desc = VENDOR.VNAME
    OSIFRec(1).Amt = CheckInfo(Cnt).ChkAmt
    OSIFRec(1).Src = 0
    '--Track Checks by Bank.  For Gate City, VA
    OSIFRec(1).BankCode = CheckInfo(Cnt).BankCode

    'OSIFRec(1).TrType="D"


    PUT OSChkHandle, OSChkNumRec, OSIFRec(1)
    OSChkNumRec = OSChkNumRec + 1

  NEXT

  CLOSE
  
END SUB

SUB PreAuditRpt
  
  '  SHARED VENDOR AS VendorRecType

  TPayCnt = (FileSize("TPAYLIST.LST") \ 6)
  IF TPayCnt = 0 THEN
    EXIT SUB
  END IF
  
  SHARED Choice$()
  
  REDIM Choice$(0 TO 3, 1)
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn#1"
  Choice$(3, 0) = "Prn#2"
  
  REDIM Frm(1) AS FormInfo
  
  FormName$ = "PRNSPEC"
  
  NumFlds = LibNumberOfFields("GL", FormName$)
  REDIM Form$(NumFlds, 2)    'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo            'DIM the field information array
  StartEl = 0   'Load first form at array start
  LibGetFldDef "GL", FormName$, StartEl, Fld(), Form$(), ErrCode
  
  '----- Setup TYPE for setting and reading form editing information.
  Frm(1).FldNo = 1           'Start editing on field #1
  Frm(1).InsStat = False     'Set the insert state (-1 = Insert on)
  Frm(1).StartEl = 0         'Set form starting element to 0 and
  
  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1
  
  FirstTime = True
  
  LibFile2Scrn "GL", FormName$, MonoCode, -1, ErrCode
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF Frm(1).KeyCode = EscKey THEN
      ExitFlag = True
    ELSEIF Frm(1).KeyCode = F10Key THEN
      OkFlag = True
    END IF
  LOOP UNTIL ExitFlag OR OkFlag
  IF ExitFlag THEN GOTO ExitPreAudit
  
  LPTPort = VAL(RIGHT$(QPTrim$(Form$(1, 0)), 1))
  
  'IF LPTPort = 0 THEN LPTPort = 1
  
  FF$ = CHR$(12)
  MaxLines = 50
  Dash$ = STRING$(78, "-")
  Dash2$ = STRING$(61, "-")
  MID$(Dash2$, 1, 7) = SPACE$(7)
  PageNum = 0
  
  REDIM FundList$(1)
  GetFundList FundList$(), NumFunds
  IF NumFunds = 0 THEN EXIT SUB
  
  REDIM FundAmts(1 TO NumFunds) AS DOUBLE
  
  REDIM APLedgerRec(1) AS APLedger81RecType
  'REDIM APLedgerRec(1) AS APLedgerRecType

  REDIM APDistRec(1) AS APDistRecType
  REDIM LedInfo(1) AS LedgerInfoType2
  REDIM DistInfo(1) AS DistInfoType
  REDIM ChkRegInfo(1) AS CheckRegType
  
  DistInfo(1).Fill1 = ""
  APDistRecLen = LEN(APDistRec(1))
  RecLen = LEN(APLedgerRec(1))
  
  'TPayCnt = (FileSize("TPAYLIST.LST") \ 6)
  'IF TPayCnt = 0 THEN
  '  EXIT SUB
  'END IF
  REDIM TPayArray(1 TO TPayCnt) AS TPayListType
  
  FGetAH "TPAYLIST.LST", TPayArray(1), 6, TPayCnt
  
  PrintFile = FREEFILE
  
  OPEN "APCHKREG.PRN" FOR OUTPUT AS PrintFile
  
  OpenVendorFile VendorFile, NumVRecs
  OpenAPLedgerFile APLedgerFile, NumTrans, RecLen
  OpenAPDistFile APDistFile, NumDistRecs&, APDistRecLen
  
  GOSUB PADoChkRegHeader
  GET VendorFile, TPayArray(1).VendorRecNum, VENDOR
  GET APLedgerFile, TPayArray(1).LedgerRecNum, APLedgerRec(1)
  GOSUB PrintVendHeader
  GOSUB PAInvHeader
  GOSUB PrintDist
  VRecNum& = TPayArray(1).VendorRecNum
  ChkCnt = 1
  FOR Cnt = 2 TO TPayCnt
    IF VRecNum& <> TPayArray(Cnt).VendorRecNum THEN
      GOSUB FinishVendor
      GET VendorFile, TPayArray(Cnt).VendorRecNum, VENDOR
      GOSUB PrintVendHeader
      GOSUB PAInvHeader
      VRecNum& = TPayArray(Cnt).VendorRecNum
      ChkCnt = ChkCnt + 1
    END IF
    GET APLedgerFile, TPayArray(Cnt).LedgerRecNum, APLedgerRec(1)
    GOSUB PrintDist
    IF LineCnt > MaxLines THEN
      PRINT #PrintFile, FF$
      GOSUB PADoChkRegHeader
    END IF
  NEXT

  GOSUB FinishVendor

  PRINT #PrintFile, FF$
  
  'print fund stuff
  LineCnt = 1
  GOSUB PADoFundHeader
  FOR Cnt = 1 TO NumFunds
    IF FundAmts(Cnt) > 0 THEN
      IF LineCnt > MaxLines THEN
        PRINT #PrintFile, FF$
        GOSUB PADoFundHeader
      END IF
      PRINT #PrintFile, "      Fund "; FUsing(STR$(VAL(FundList$(Cnt))), ",####"); "  Amt  "; FUsing(STR$(FundAmts(Cnt)), ",########.##")
      LineCnt = LineCnt + 1
    END IF
  NEXT

  PRINT #PrintFile, FF$
  CLOSE
  
  ERASE FundList$, FundAmts, APLedgerRec, APDistRec
  ERASE LedInfo, DistInfo, ChkRegInfo           ', ChkInfo
  ERASE Form$, Fld, Fld
  
  
  IF LPTPort > 0 THEN
    EntryPoint = 5
  ELSE
    EntryPoint = 2
  END IF
  Title$ = "Check Pre-Audit Report"
  
  PrintRptFile Title$, "APCHKREG.PRN", LPTPort%, RetCode%, EntryPoint
  
ExitPreAudit:
  
  
  EXIT SUB
  
PADoFundHeader:
  PageNum = PageNum + 1
  Page$ = FUsing(STR$(PageNum), "###")
  PRINT #PrintFile, "Check Pre-Audit Report Summary                                    Page:" + Page$
  PRINT #PrintFile, Dash$
  PRINT #PrintFile,
  PRINT #PrintFile, "    Checks to Print        "; FUsing(STR$(ChkCnt), ",####")
  PRINT #PrintFile, "           Totaling  "; FUsing(STR$(TotalChkAmt#), "$$########,.##")
  PRINT #PrintFile,
  PRINT #PrintFile, " By Fund:"
  LineCnt = 7
  RETURN
  
  
PADoChkRegHeader:
  PageNum = PageNum + 1
  Page$ = FUsing(STR$(PageNum), "###")
  PRINT #PrintFile, "A/P Check Pre-Audit Report                                            Page:" + Page$
  PRINT #PrintFile, "Run Date: " + DATE$
  PRINT #PrintFile,
  PRINT #PrintFile, Dash$
  LineCnt = 4
RETURN

  
PAInvHeader:
  PRINT #PrintFile,
  PRINT #PrintFile, "Inv Date    Due Date    Inv Num                    PO                  Amount"
  PRINT #PrintFile, "----------  ----------  -------------------------  ----------    ------------"
  LineCnt = LineCnt + 2
RETURN


PrintDist:
  LSET ChkRegInfo(1).VENDNAME = VENDOR.VNAME
  LSET LedInfo(1).INVDATE = Num2Date(APLedgerRec(1).TRDate)
  LSET LedInfo(1).DUEDATE = Num2Date(APLedgerRec(1).DUEDATE)
  LSET LedInfo(1).INVNUM = LEFT$(APLedgerRec(1).DOCNum, 25)
  LSET LedInfo(1).PONum = LEFT$(APLedgerRec(1).PONum, 10)
  RSET LedInfo(1).Amt = FUsing(STR$(APLedgerRec(1).Amt), ",########.##")
  VendTotal# = Round(VendTotal# + APLedgerRec(1).Amt)
  TotalChkAmt# = Round(TotalChkAmt# + APLedgerRec(1).Amt)
  PRINT #PrintFile, LedInfo(1).INVDATE; LedInfo(1).DUEDATE; LedInfo(1).INVNUM; LedInfo(1).PONum; LedInfo(1).Amt
  LineCnt = LineCnt + 2
  NextDist& = APLedgerRec(1).FrstDist
  
  PRINT #PrintFile, TAB(50); "Dist Acct        Dist Amount"
  LineCnt = LineCnt + 1
  
  DO UNTIL NextDist& = 0
    GET APDistFile, NextDist&, APDistRec(1)
    IF LineCnt > MaxLines THEN
      PRINT #PrintFile, FF$
      GOSUB PADoChkRegHeader
      'GOSUB PADistHeader
    END IF
    LSET LedInfo(1).INVDATE = ""
    LSET LedInfo(1).DUEDATE = ""
    LSET LedInfo(1).INVNUM = ""
    LSET LedInfo(1).PONum = ""
    RSET LedInfo(1).Amt = ""
    
    LSET LedInfo(1).DistAcct = APDistRec(1).DistAcctNum
    LSET LedInfo(1).DistAmt = FUsing(STR$(APDistRec(1).DistAmt), ",########.##")
    'PRINT #PrintFile, DistInfo(1).Fill1; DistInfo(1).DistAcct; DistInfo(1).DistAmt
    'PRINT #PrintFile, LedInfo(1).InvDate; LedInfo(1).DueDate; LedInfo(1).InvNum; LedInfo(1).PONum; LedInfo(1).Amt; "  "; LedInfo(1).DistAcct; LedInfo(1).DistAmt
    PRINT #PrintFile, TAB(50); LedInfo(1).DistAcct; LedInfo(1).DistAmt
    
    LineCnt = LineCnt + 1
    ThisFund$ = LEFT$(APDistRec(1).DistAcctNum, FundLen)
    FOR FundCnt = 1 TO NumFunds
      IF ThisFund$ = FundList$(FundCnt) THEN
        FundAmts(FundCnt) = Round(FundAmts(FundCnt) + APDistRec(1).DistAmt)
        EXIT FOR
      END IF
    NEXT
    NextDist& = APDistRec(1).NextDist
  LOOP
  PRINT #PrintFile, ""
  LineCnt = LineCnt + 1
RETURN
  
  
PrintVendHeader:
  'PRINT #PrintFile, Dash$
  PRINT #PrintFile, VENDOR.VNum; VENDOR.VNAME
  LineCnt = LineCnt + 1
RETURN


FinishVendor:
  PRINT #PrintFile, TAB(50); "Vendor Total: "; TAB(66); FUsing(STR$(VendTotal#), ",########.##")
  PRINT #PrintFile, Dash$
  VendTotal# = 0
  LineCnt = LineCnt + 2
RETURN

  
END SUB

SUB PrintCheckListing

  IF NOT Exist("APCHKINF.DAT") THEN EXIT SUB
  Dash$ = STRING$(78, "-")
  FF$ = CHR$(12)

  REDIM ChkInfo(1 TO 1) AS CheckInfoType3
  ChkInfoRecLen = LEN(ChkInfo(1))
  VCnt = (FileSize("APCHKINF.DAT") \ ChkInfoRecLen)
  REDIM ChkInfo(1 TO VCnt)  AS CheckInfoType3
  FGetAH "APCHKINF.DAT", ChkInfo(1), ChkInfoRecLen, VCnt

  OpenVendorFile VendorFile, NumVRecs
  PrintFile = FREEFILE
  OPEN "APCHKLST.PRN" FOR OUTPUT AS #PrintFile

  PRINT #PrintFile, " A/P Check Listing:  " + Num2Date(ChkInfo(1).ChkDate)
  PRINT #PrintFile,
  PRINT #PrintFile, " Check No.         Description                              Amount"
  PRINT #PrintFile, Dash$
  FOR Cnt = 1 TO VCnt
    GET VendorFile, ChkInfo(Cnt).VendorRecNum, VENDOR
    FOR Cnt2& = ChkInfo(Cnt).StartChk TO ChkInfo(Cnt).LastChk
      PRINT #PrintFile, FUsing(STR$(Cnt2&), "#######");
      IF ChkInfo(Cnt).VoidFlag THEN
        PRINT #PrintFile, TAB(20); VENDOR.VNAME; TAB(52); "  VOIDED BY USER"
      ELSEIF Cnt2& < ChkInfo(Cnt).LastChk THEN
        PRINT #PrintFile, TAB(20); VENDOR.VNAME; TAB(52); "            VOID"
      ELSE
        TCheckAmt# = Round(TCheckAmt# + ChkInfo(Cnt).ChkAmt)
        PRINT #PrintFile, TAB(20); VENDOR.VNAME; TAB(54); FUsing(STR$(ChkInfo(Cnt).ChkAmt), "$$########,.##")
      END IF
    NEXT
  NEXT

  PRINT #PrintFile, Dash$

  PRINT #PrintFile, "Total"; TAB(54); FUsing(STR$(TCheckAmt#), "$$########,.##")
  PRINT #PrintFile, FF$
  CLOSE

  'TOP
  SHARED Choice$()

  REDIM Choice$(0 TO 3, 1)
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn#1"
  Choice$(3, 0) = "Prn#2"

  REDIM Frm(1) AS FormInfo

  FormName$ = "PRNSPEC"

  NumFlds = LibNumberOfFields("GL", FormName$)
  REDIM Form$(NumFlds, 2)    'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo            'DIM the field information array
  StartEl = 0   'Load first form at array start
  LibGetFldDef "GL", FormName$, StartEl, Fld(), Form$(), ErrCode

  '----- Setup TYPE for setting and reading form editing information.
  Frm(1).FldNo = 1           'Start editing on field #1
  Frm(1).InsStat = False     'Set the insert state (-1 = Insert on)
  Frm(1).StartEl = 0         'Set form starting element to 0 and

  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1

  FirstTime = True

  LibFile2Scrn "GL", FormName$, MonoCode, -1, ErrCode
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF Frm(1).KeyCode = EscKey THEN
      ExitFlag = True
    ELSEIF Frm(1).KeyCode = F10Key THEN
      OkFlag = True
    END IF
  LOOP UNTIL ExitFlag OR OkFlag

  IF ExitFlag THEN GOTO ExitCheckListing:

  SELECT CASE LEFT$(Form$(1, 0), 1)
  CASE "S"
    EntryPoint = 2
  CASE ELSE
    EntryPoint = 5
    LPTPort = VAL(RIGHT$(QPTrim$(Form$(1, 0)), 1))
  END SELECT
  Title$ = "A/P Check Listing"

  PrintRptFile Title$, "APCHKLST.PRN", LPTPort, RetCode%, EntryPoint

ExitCheckListing:

END SUB

SUB PrintChecks2 (FirstBadNum&, CheckNum&, CheckDate, LPTPort, RestartFlag, BankCode)
  
  REDIM APLedgerRec(1) AS APLedger81RecType
  'REDIM APLedgerRec(1) AS APLedgerRecType
   REDIM APDistRec(1) AS APDistRecType

   APDistRecLen = LEN(APDistRec(1))

  RecLen = LEN(APLedgerRec(1))
  
  TCheckNum& = CheckNum&
  
  REDIM PayListRec(1) AS TPayListType
  REDIM ChkInfo(1 TO 1) AS CheckInfoType3
  ChkInfoRecLen = LEN(ChkInfo(1))
  
  PayListRecLen = LEN(PayListRec(1))
  
  TNumInvoices = (FileSize("TPAYLIST.LST") \ PayListRecLen)
  REDIM InvList(1 TO TNumInvoices) AS TPayListType
  FGetAH "TPAYLIST.LST", InvList(1), PayListRecLen, TNumInvoices
  
  IF NOT RestartFlag THEN
    VCnt = 1
    LCnt = 1
    VRecNum1& = InvList(LCnt).VendorRecNum
    DO UNTIL LCnt >= TNumInvoices
      LCnt = LCnt + 1
      VRecNum2& = InvList(LCnt).VendorRecNum
      IF VRecNum1& <> VRecNum2& THEN
        VCnt = VCnt + 1
        VRecNum1& = VRecNum2&
      END IF
    LOOP
    
    REDIM ChkInfo(1 TO VCnt) AS CheckInfoType3
    
    ChkCnt = 1
    VRecNum1& = InvList(ChkCnt).VendorRecNum
    ChkInfo(ChkCnt).ListFirst = 1
    ChkInfo(ChkCnt).VendorRecNum = VRecNum1&
    
    FOR Cnt = 2 TO TNumInvoices
      VRecNum2& = InvList(Cnt).VendorRecNum
      IF VRecNum1& <> VRecNum2& THEN
        ChkInfo(ChkCnt).ListLast = Cnt - 1
        ChkInfo(ChkCnt).VendorRecNum = InvList(Cnt - 1).VendorRecNum
        ChkInfo(ChkCnt).BankCode = BankCode
        ChkCnt = ChkCnt + 1
        ChkInfo(ChkCnt).ListFirst = Cnt
        VRecNum1& = VRecNum2&
      END IF
    NEXT
    
    ChkInfo(ChkCnt).ListLast = TNumInvoices
    ChkInfo(ChkCnt).VendorRecNum = InvList(TNumInvoices).VendorRecNum
    ChkInfo(ChkCnt).BankCode = BankCode

    KillFile "APCHKINF.DAT"
    FPutAH "APCHKINF.DAT", ChkInfo(1), ChkInfoRecLen, ChkCnt
    StartCnt = 1
  ELSE          'We are restarting printing
    VCnt = (FileSize("APCHKINF.DAT") \ ChkInfoRecLen)
    REDIM ChkInfo(1 TO VCnt) AS CheckInfoType3
    FGetAH "APCHKINF.DAT", ChkInfo(1), ChkInfoRecLen, VCnt
    FOR TCnt = 1 TO VCnt
      IF ChkInfo(TCnt).LastChk = FirstBadNum& THEN
        StartCnt = TCnt
        EXIT FOR
      END IF
    NEXT
  END IF
  
'**********check for check less than 0
'****************
'Check Formats **************************************************************
''$INCLUDE: 'ATMORE.CAP'     'Atmore
''$INCLUDE: 'wingate.CAP'     'Wingate
''$INCLUDE: 'NEWSTD.CAP'     'Form 9013
''$INCLUDE: 'NEWSTD42.CAP'   'Form 9013 with 3 extra top line (42 total)
''$INCLUDE: 'OLDSTD.CAP'     'Form 9028
''$INCLUDE: 'ROLESVIL.CAP'    'ROLESVILLE, NC  Form #9016
''$INCLUDE: 'HOINWALD.CAP'   'laser check
''$INCLUDE: 'LASER.CAP'      'Flat Rock
''$INCLUDE: 'TROY.CAP'       'bottom stub
''$INCLUDE: 'ALGY.CAP'       'Alleghany Co. (Yancey's with addr line up 2)
''$INCLUDE: 'APPVA.CAP'      'for Appalachia, VA & Appalachian Dist Health Dept.
''$INCLUDE: 'CARTHAGE.CAP'   '05/28/96 - 42 line blank top stub
''$INCLUDE: 'CAROBCH.CAP'
'$INCLUDE: 'FAISON.CAP'
''$INCLUDE: 'DUMFRIES.CAP'
''$INCLUDE: 'drexel.CAP'
''$INCLUDE: 'GRANITEQ.CAP'
''$INCLUDE: 'grover.CAP'
''$INCLUDE: 'LONGVIEW.CAP'
''$INCLUDE: 'ORANGEAL.CAP'    'Orange Alamance (Mebane)
''$INCLUDE: 'SYLVA.CAP'      'Last update 5/1/96
''$INCLUDE: 'SDS.CAP'        'RICHLAND.CAP'   'SDS, Troutman
''$INCLUDE: 'truman.CAP'
''$INCLUDE: 'wadesbor.cap'    'Wadesboro Format
''$INCLUDE: 'whitaker.CAP'
''$INCLUDE: 'YANCEY.CAP'     'also for Spruce Pine, Region D COG
''$INCLUDE: 'WRGHTBCH.CAP'     'Wrightsville Beech
'****************************************************************************
ExitCheckPrinting:    'DO NOT REMOVE THIS LABEL
                      'IT IS USED TO EXIT INCLUDE FILES
END SUB

SUB SelectOpenPayables
  
  SHARED Choice$()
  
  REDIM APLedgerRec(1) AS APLedger81RecType
  'REDIM APLedgerRec(1) AS APLedgerRecType
  REDIM PayListRec(1) AS TPayListType
  PayListRecLen = LEN(PayListRec(1))
  
  PickHeader$ = "  Vendor                    Inv Num       Inv Date       Inv Amt    "
  LdRecLen = LEN(APLedgerRec(1))
  
  NumFlds = LibNumberOfFields("AP", "SELOPEN")
  REDIM Frm(1)    AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  StartEl = 0   'Load first form at array start
  LibGetFldDef "AP", "SELOPEN", StartEl, Fld(), Form$(), ErrCode
  LibFile2Scrn "AP", "SELOPEN", MonoCode, -1, ErrCode           'display screen

  PrintTitle User$
  PrintHelp "Select Open Payables"
  Help$ = "(*)=All"
  QPrintRC Help$, 25, 73, -1
  Action = 1
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF Frm(1).KeyCode = 27 OR Frm(1).KeyCode = -68 THEN
      SELECT CASE Frm(1).KeyCode
      CASE EscKey
        ExitFlag = True
      CASE F10Key
        'add other stuff for pre exit form here.
      END SELECT
      EXIT DO
    END IF
    IF Frm(1).KeyCode = ASC("*") THEN
      GOSUB Get1stLastVendor
    END IF
    IF Frm(1).PrevFld <> Frm(1).FldNo THEN      'User is moving around the form
      SELECT CASE Frm(1).PrevFld
      CASE 1, 2
        ChkVendor$ = QPTrim$(Form$(Frm(1).PrevFld, 0))
        IF LEN(ChkVendor$) > 0 THEN
          GOSUB Check4Vendor
        END IF
      CASE ELSE
      END SELECT
    END IF
  LOOP
  
  IF ExitFlag GOTO ExitSelPayables
  
  REDIM TempVList(1 TO 2) AS FLen
  RSET TempVList(1).V = Form$(1, 0)
  RSET TempVList(2).V = Form$(2, 0)
  ' (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
  SortT TempVList(1), 2, 0, 32, 0, 32
  
  Vend1St$ = QPTrim$(TempVList(1).V)
  VendLst$ = QPTrim$(TempVList(2).V)
  DUEDATE = Date2Num(Form$(3, 0))
  
  NumVendRecs = (FileSize("apvendor.idx") \ 12)
  REDIM VIndex(1 TO NumVendRecs) AS VendorIdxRecType
  FGetAH "apvendor.idx", VIndex(1), 12, NumVendRecs
  
  KillFile "TPAYLIST.LST"
  TPayListFile = FREEFILE
  OPEN "TPAYLIST.LST" FOR RANDOM AS TPayListFile LEN = PayListRecLen
  
  OpenVendorFile VendorFile, NumVRecs
  OpenAPLedgerFile APLedgerFile, NumTrans, LdRecLen
  
  FOR Cnt = 1 TO NumVendRecs
    VCode$ = QPTrim$(VIndex(Cnt).VendorCode)
    IF VCode$ >= Vend1St$ AND VCode$ <= VendLst$ THEN
      
      GET VendorFile, VIndex(Cnt).RecNum, VENDOR
      NextTrans& = VENDOR.FrstTran
      
      DO UNTIL NextTrans& = 0
        GET APLedgerFile, NextTrans&, APLedgerRec(1)
        'IF APLedgerRec(1).PayCode = 0 THEN
        '  APLedgerRec(1).PayCode = 1
        '  PUT APLedgerFile, NextTrans&, APLedgerRec(1)
        'END IF
        
        IF APLedgerRec(1).TrCode = 1 AND APLedgerRec(1).PAYCODE = 1 THEN
          IF APLedgerRec(1).DUEDATE <= DUEDATE THEN
            '  LRec& = NextTrans&
            PayListRec(1).LedgerRecNum = NextTrans&
            PayListRec(1).VendorRecNum = VIndex(Cnt).RecNum
            PUT TPayListFile, , PayListRec(1)
          END IF
        END IF
        NextTrans& = APLedgerRec(1).NextTrans
      LOOP
    END IF
  NEXT
  CLOSE
  
  ERASE VIndex
  
  ListCnt = (FileSize("TPAYLIST.LST") \ PayListRecLen)
  
  IF ListCnt = 0 THEN EXIT SUB
  'add a trap here to display an error scrn if no matching ledeger recs
  'to pay
  
  REDIM ListArray(1 TO ListCnt) AS TPayListType
  FGetAH "TPAYLIST.LST", ListArray(1), 6, ListCnt
  REDIM Items(1 TO ListCnt) AS FLen2
  REDIM Picked(1 TO ListCnt)
  OpenAPLedgerFile APLedgerFile, NumTrans, LdRecLen
  OpenVendorFile VendorFile, NumVRecs

  FOR Cnt = 1 TO ListCnt
    GET APLedgerFile, ListArray(Cnt).LedgerRecNum, APLedgerRec(1)
    GET VendorFile, APLedgerRec(1).VRecNum, VENDOR
    Items(Cnt).V = LEFT$(VENDOR.VNAME, 24)
    MID$(Items(Cnt).V, 27) = APLedgerRec(1).DOCNum
    MID$(Items(Cnt).V, 41) = Num2Date(APLedgerRec(1).DUEDATE)
    MID$(Items(Cnt).V, 52) = FUsing$(STR$(APLedgerRec(1).Amt), "#######.##")
  NEXT
  
  LibFile2Scrn "AP.QSL", "BG", MonoCode, -1, ErrorCode
  QPrintRC PickHeader$, 4, 7, 11
  QPrintRC "F10 To Continue, SpaceBar To Toggle, Alt-M Marks All, Alt-C Clear All Marks", 25, 4, -1
  
  MPaintBox 5, 74, 5, 76, 8
  StuffBuf (CHR$(0) + "2")
  PickList Items(), Picked(), NPicked, Cnf
  
  'GOTO ExitSelPayables:
  'look at the picked()

  '--process the pick list into a file of only the selected choices
  KillFile "TPAYLIST.LST"
  TPayListFile = FREEFILE
  OPEN "TPAYLIST.LST" FOR RANDOM AS TPayListFile LEN = PayListRecLen
  FOR NCnt = 1 TO NPicked
    ThisRec& = Picked(NCnt)
    PUT #TPayListFile, NCnt, ListArray(ThisRec&)
  NEXT
  CLOSE
  EXIT SUB
  REDIM PayListRec(1) AS TPayListType
  REDIM ChkInfo(1 TO 1) AS CheckInfoType3
  
  ChkInfoRecLen = LEN(ChkInfo(1))
  PayListRecLen = LEN(PayListRec(1))
  
  TNumInvoices = (FileSize("TPAYLIST.LST") \ PayListRecLen)
  IF TNumInvoices = 0 THEN EXIT SUB
  REDIM InvList(1 TO TNumInvoices) AS TPayListType
  FGetAH "TPAYLIST.LST", InvList(1), PayListRecLen, TNumInvoices
  
  VCnt = 1
  VRecNum1& = InvList(1).VendorRecNum
  ChkInfo(VCnt).ListFirst = InvList(1).LedgerRecNum
  FOR TCnt = 2 TO TNumInvoices
    VRecNum2& = InvList(TCnt).VendorRecNum
    IF VRecNum1& <> VRecNum2& THEN
      ChkInfo(VCnt).ListLast = InvList(TCnt - 1).LedgerRecNum
      ChkInfo(VCnt).VendorRecNum = InvList(TCnt - 1).VendorRecNum
      
      VCnt = VCnt + 1
      REDIM PRESERVE ChkInfo(1 TO VCnt) AS CheckInfoType3
      '      IF TCnt = TNumInvoices THEN
      ChkInfo(VCnt).ListFirst = InvList(TCnt).LedgerRecNum
      '      END IF
      
      ' ChkInfo(VCnt).Ledger1St = TCnt
      VRecNum1& = InvList(TCnt).VendorRecNum
    END IF
  NEXT
  'ChkInfo(VCnt).LedgerLst = TNumInvoices
  ChkInfo(VCnt).ListLast = InvList(TCnt - 1).LedgerRecNum
  ChkInfo(VCnt).VendorRecNum = InvList(TCnt - 1).VendorRecNum
  StartCnt = 1
  FPutAH "APCHKINF.DAT", ChkInfo(1), ChkInfoRecLen, VCnt
  
  '-------------------------
ExitSelPayables:
  
  
  EXIT SUB
  
Check4Vendor:
  MatchingVendors = 0
  List$ = SPACE$(65)
  
  VLookup$ = QPTrim$(UCASE$(ChkVendor$))
  VLookuplen = LEN(VLookup$)
  
  REDIM ScrBuf(2000)
  
  MScrnSave 1, 1, 25, 80, ScrBuf(0)             'Save the existing screen
  
  OpenVendorIdx VendorIdxFile, NumActiveVendors
  
  IF NumActiveVendors = 0 THEN
    CLOSE
    ListVendors = 0
    EXIT SUB
  END IF
  
  OpenVendorFile VendorFile, NumVRecs
  
'  REDIM VendorList(1 TO NumVRecs) AS FLen       'Dim array to hold trans list
  REDIM VendorList(1 TO NumActiveVendors) AS FLen       'Dim array to hold trans list
  
  FOR Cnt = 1 TO NumActiveVendors               'Read the file
    GET VendorIdxFile, Cnt, VendorIdx
    IF LEFT$(VendorIdx.VendorCode, VLookuplen) = VLookup$ THEN
      MatchingVendors = MatchingVendors + 1
      GET VendorFile, VendorIdx.RecNum, VENDOR
      LSET VendorList(MatchingVendors).V = List$
      VendorList(MatchingVendors).V = VendorIdx.VendorCode
      MID$(VendorList(MatchingVendors).V, 11) = VENDOR.VNAME
      MID$(VendorList(MatchingVendors).V, 41) = MKI$(VendorIdx.RecNum)
    END IF
  NEXT
  
  CLOSE
  
  IF MatchingVendors = 0 THEN
    OK = MsgBox("AP", "NOVENDOR")
    VendorPickList = 0
    LSET Form$(3, 0) = ""
    Frm(1).FldNo = 1
    Action = 1
    GOTO ChkReturn
  END IF
  
  IF MatchingVendors = 1 THEN
    LSET Form$(Frm(1).PrevFld, 0) = VendorList(1).V
    SaveField Frm(1).PrevFld, Form$(), Fld(), BadFld
    Action = 1
    GOTO ChkReturn
  END IF
  
  IF MatchingVendors > 1 THEN
    REDIM PRESERVE VendorList(1 TO MatchingVendors) AS FLen
    Row = 5: Col = 15
    LOCATE Row, Col
    Action = 0
    MaxWidth = 39
    BoxBot = 18
    ListDescTxt$ = "VendorCode  Name"
    ListDesc$ = SPACE$(MaxWidth + 4)
    MID$(ListDesc$, 3) = ListDescTxt$
    QPrintRC ListDesc$, Row - 1, Col, 112       'Cnf.ActivCh
    Choice = 1
    DO
      'call vertmenu to display matching recs to choose from
      VertMenuT VendorList(), Choice, MaxWidth, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        CLOSE
        Abort = True
        Frm(1).FldNo = 1
        Action = 1
        GOTO ChkReturn
        '        EXIT DO
      END IF
      IF Choice > 0 THEN EXIT DO
    LOOP
    '  MScrnRest 1, 1, 25, 80, ScrBuf(0)           'Restore Screen
  ELSE
  END IF
  
  IF Choice > 0 THEN
    LSET Form$(Frm(1).PrevFld, 0) = VendorList(Choice).V
    SaveField Frm(1).PrevFld, Form$(), Fld(), BadFld
    Choice = 0
  END IF
  
ChkReturn:
  MScrnRest 1, 1, 25, 80, ScrBuf(0)             'Restore Screen
  QPrintRC TName$, 3, 27, -1
  ERASE VendorList
  
  RETURN
  
  
Get1stLastVendor:
  OpenVendorIdx VendorIdxFile, NumActiveVendors
  IF NumActiveVendors > 0 THEN
    GET VendorIdxFile, 1, VendorIdx
    Form$(1, 0) = VendorIdx.VendorCode
    GET VendorIdxFile, NumActiveVendors, VendorIdx
    Form$(2, 0) = VendorIdx.VendorCode
    Form$(3, 0) = DATE$
    SaveField 1, Form$(), Fld(), BadFld
    Frm(1).FldNo = 3
    Action = 2
  END IF
  CLOSE VendorIdxFile
  RETURN
END SUB

SUB VoidAPCheck

  REDIM VoidAPLedgerRec(1) AS APLedger81RecType

  LdRecLen = LEN(VoidAPLedgerRec(1))

  REDIM Frm(1) AS FormInfo

  FormName$ = "VOIDCHK"

  NumFlds = LibNumberOfFields("AP", FormName$)
  REDIM Form$(NumFlds, 2)    'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo            'DIM the field information array
  StartEl = 0   'Load first form at array start
  LibGetFldDef "AP", FormName$, StartEl, Fld(), Form$(), ErrCode

  '----- Setup TYPE for setting and reading form editing information.
  Frm(1).FldNo = 1           'Start editing on field #1
  Frm(1).InsStat = False     'Set the insert state (-1 = Insert on)
  Frm(1).StartEl = 0         'Set form starting element to 0 and

  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1

  FirstTime = True

  LibFile2Scrn "AP", FormName$, MonoCode, -1, ErrCode
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF Frm(1).FldNo > 1 AND Frm(1).PrevFld = 1 THEN   'User is moving around the form
      ChkVendor$ = QPTrim$(Form$(1, 0))
      GOSUB VoidCheck4Vendor
      IF NOT Abort THEN
        GOSUB VoidCheck4ChkNum
      ELSE
        Abort = False
      END IF
    END IF

    IF Frm(1).KeyCode = EscKey THEN
      ExitFlag = True
    ELSEIF Frm(1).KeyCode = F10Key THEN
      OkFlag = True
    END IF
  LOOP UNTIL ExitFlag OR OkFlag

  IF ExitFlag THEN EXIT SUB


  IF OkFlag THEN
    Temp$ = InputBox$("AP", "OK2VOID", Cancel%)
  END IF

  IF Cancel THEN
    EXIT SUB
  ELSE
    VoidChkNum& = QPValL(Form$(2, 0))
    VoidCheckTrans VoidVendorRecNum&, VoidTransRecNum&, VoidChkNum&
  END IF

  EXIT SUB


VoidCheck4ChkNum:
  REDIM ScrBuf(2000)
  MScrnSave 1, 1, 25, 80, ScrBuf(0)             'Save the existing screen
  CheckCnt = 0
  List$ = SPACE$(45)
  OpenVendorFile VendorFile, NumVRecs
  OpenAPLedgerFile APLedgerFile, NumTrans, LdRecLen
  GET VendorFile, VoidVendorRecNum&, VENDOR
  NextTrans& = VENDOR.FrstTran
  DO UNTIL NextTrans& = 0
    GET APLedgerFile, NextTrans&, VoidAPLedgerRec(1)
    IF VoidAPLedgerRec(1).TrCode = 3 THEN
       CheckCnt = CheckCnt + 1
    END IF
    NextTrans& = VoidAPLedgerRec(1).NextTrans
  LOOP

  IF CheckCnt = 0 THEN
    OK = MsgBox("AP", "NOCHECKS")
    Frm(1).FldNo = 1
    Action = 1
    GOTO VoidCheck4ChkNumExit
  END IF

  REDIM CheckList(1 TO CheckCnt) AS FLen       'Dim array to hold trans list

  MatchCnt = 0
  NextTrans& = VENDOR.FrstTran

  DO UNTIL NextTrans& = 0
    GET APLedgerFile, NextTrans&, VoidAPLedgerRec(1)
      IF VoidAPLedgerRec(1).TrCode = 3 THEN
        MatchCnt = MatchCnt + 1
        LSET CheckList(MatchCnt).V = List$
        CheckList(MatchCnt).V = QPTrim$(VoidAPLedgerRec(1).DOCNum)
        MID$(CheckList(MatchCnt).V, 11) = Num2Date(VoidAPLedgerRec(1).TRDate)
        MID$(CheckList(MatchCnt).V, 22) = FUsing(STR$(VoidAPLedgerRec(1).Amt), "$$########,.##")
        MID$(CheckList(MatchCnt).V, 39) = MKL$(NextTrans&)
      END IF
      NextTrans& = VoidAPLedgerRec(1).NextTrans
    LOOP

  CLOSE

  Row = 12
  Col = 15
  LOCATE Row, Col
  Action = 0
  MaxWidth = 37
  BoxBot = 18
  ListDescTxt$ = "Check Num     Date          Amount"
  ListDesc$ = SPACE$(MaxWidth + 4)
  MID$(ListDesc$, 3) = ListDescTxt$
  QPrintRC ListDesc$, Row - 1, Col, 112       'Cnf.ActivCh
  Choice = 1
  DO
    'call vertmenu to display matching recs to choose from
    VertMenuT CheckList(), Choice, MaxWidth, BoxBot, Ky$, Action, Cnf
    IF Ky$ = CHR$(27) THEN
      Abort = True
      Frm(1).FldNo = 1
      Action = 1
      GOTO VoidCheck4ChkNumExit
    END IF
    IF Choice > 0 THEN
      LSET Form$(2, 0) = QPTrim$(LEFT$(CheckList(Choice).V, 10))
      SaveField 2, Form$(), Fld(), BadFld
      VoidTransRecNum& = CVL(RIGHT$(CheckList(Choice).V, 4))
      Frm(1).FldNo = 3
      EXIT DO
    END IF
  LOOP

VoidCheck4ChkNumExit:
  MScrnRest 1, 1, 25, 80, ScrBuf(0)           'Restore Screen

RETURN


VoidCheck4Vendor:
  MatchingVendors = 0
  List$ = SPACE$(65)

  VLookup$ = QPTrim$(UCASE$(ChkVendor$))
  VLookuplen = LEN(VLookup$)

  REDIM ScrBuf(2000)

  MScrnSave 1, 1, 25, 80, ScrBuf(0)             'Save the existing screen

  OpenVendorIdx VendorIdxFile, NumActiveVendors

  IF NumActiveVendors = 0 THEN
    CLOSE
    ListVendors = 0
    EXIT SUB
  END IF

  OpenVendorFile VendorFile, NumVRecs

  REDIM VendorList(1 TO NumActiveVendors) AS FLen       'Dim array to hold trans list

  FOR Cnt = 1 TO NumActiveVendors               'Read the file
    GET VendorIdxFile, Cnt, VendorIdx
    IF LEFT$(VendorIdx.VendorCode, VLookuplen) = VLookup$ THEN
      GET VendorFile, VendorIdx.RecNum, VENDOR
      IF VENDOR.DELFLAG = 0 THEN
        MatchingVendors = MatchingVendors + 1
        LSET VendorList(MatchingVendors).V = List$
        VendorList(MatchingVendors).V = VendorIdx.VendorCode
        MID$(VendorList(MatchingVendors).V, 12) = VENDOR.VNAME
        MID$(VendorList(MatchingVendors).V, 41) = MKI$(VendorIdx.RecNum)
      END IF
    END IF
  NEXT

  CLOSE

  IF MatchingVendors = 0 THEN
    OK = MsgBox("AP", "NOVENDOR")
    VendPickList = 0
    Frm(1).FldNo = 1
    Action = 1
    GOTO VoidChkReturn
  END IF

  IF MatchingVendors = 1 THEN
    LSET Form$(Frm(1).PrevFld, 0) = VendorList(1).V
    SaveField Frm(1).PrevFld, Form$(), Fld(), BadFld
    Action = 1
    VoidVendorRecNum& = CVI(RIGHT$(VendorList(1).V, 2))
    GOTO VoidChkReturn
  END IF

  IF MatchingVendors > 1 THEN
    REDIM PRESERVE VendorList(1 TO MatchingVendors) AS FLen
    Row = 5: Col = 15
    LOCATE Row, Col
    Action = 0
    MaxWidth = 39
    BoxBot = 18
    ListDescTxt$ = "VendorCode  Name"
    ListDesc$ = SPACE$(MaxWidth + 4)
    MID$(ListDesc$, 3) = ListDescTxt$
    QPrintRC ListDesc$, Row - 1, Col, 112       'Cnf.ActivCh
    Choice = 1
    DO
      'call vertmenu to display matching recs to choose from
      VertMenuT VendorList(), Choice, MaxWidth, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        CLOSE
        Abort = True
        Frm(1).FldNo = 1
        Action = 1
        GOTO VoidChkReturn
      END IF
      IF Choice > 0 THEN
        VoidVendorRecNum& = CVI(RIGHT$(VendorList(Choice).V, 2))
        EXIT DO
      END IF
    LOOP
    '  MScrnRest 1, 1, 25, 80, ScrBuf(0)           'Restore Screen
  ELSE
  END IF

  IF Choice > 0 THEN
    LSET Form$(Frm(1).PrevFld, 0) = VendorList(Choice).V
    SaveField Frm(1).PrevFld, Form$(), Fld(), BadFld
    Choice = 0
  END IF

VoidChkReturn:
  MScrnRest 1, 1, 25, 80, ScrBuf(0)             'Restore Screen
  QPrintRC TName$, 3, 27, -1
  ERASE VendorList

  RETURN

  
END SUB

SUB VoidCheckTrans (VoidVendorRecNum&, VoidTransRecNum&, VoidChkNum&)

  InterfaceFile$ = "APVOIDIF.DAT"
  IF Exist(InterfaceFile$) THEN
    STOP
  END IF

  PrnFileName$ = "VOIDCHK.PRN"
  PrnTitle$ = "Voided Check Report"
  Fmt$ = "######,.##"
  LPTNo = 1
  
  REDIM VoidAPLedgerRec(1) AS APLedger81RecType
  REDIM DistRec(1) AS APDistRecType
  REDIM GLifRec(1) AS GLTransRecType

  GLIFRecLen = LEN(GLifRec(1))
  LdRecLen = LEN(VoidAPLedgerRec(1))
  DistRecLen = LEN(DistRec(1))
  
  OpenAPLedgerFile APLedgerFile, NumTrans, LdRecLen
  OpenVendorFile VendorFile, NumVRecs

  PrnFile = FREEFILE
  OPEN PrnFileName$ FOR OUTPUT AS PrnFile
  GOSUB PrintReportHeader

  GET VendorFile, VoidVendorRecNum&, VENDOR
  
  NextTrans& = VENDOR.FrstTran
  DO UNTIL NextTrans& = 0
    GET APLedgerFile, NextTrans&, VoidAPLedgerRec(1)
    IF VoidAPLedgerRec(1).TrCode = 1 AND VoidAPLedgerRec(1).PAYCODE = 3 AND VoidAPLedgerRec(1).PDCheckNum = VoidChkNum& THEN
      VoidAPLedgerRec(1).PAYCODE = 1
      VoidAPLedgerRec(1).PDCheckNum = 0
      VoidAPLedgerRec(1).PDCheckDate = 0
      PUT APLedgerFile, NextTrans&, VoidAPLedgerRec(1)
    END IF
    NextTrans& = VoidAPLedgerRec(1).NextTrans
  LOOP
  
  GET APLedgerFile, VoidTransRecNum&, VoidAPLedgerRec(1)
  VoidAPLedgerRec(1).TrCode = -3
  PUT APLedgerFile, VoidTransRecNum&, VoidAPLedgerRec(1)
  CLOSE APLedgerFile

  GLIFFile = FREEFILE
  OPEN InterfaceFile$ FOR RANDOM AS GLIFFile LEN = GLIFRecLen

  OpenAPDistFile APDistFile, NumDistRecs&, DistRecLen

  NextDist& = VoidAPLedgerRec(1).FrstDist
  DistCnt = 0
  DO UNTIL NextDist& = 0
    GET APDistFile, NextDist&, DistRec(1)

    '--Set common values
    GLifRec(1).Src = "VC" + ConvDateStr$(DATE$)
    GLifRec(1).TRDate = Date2Num(DATE$)
    GLifRec(1).Desc = "TO VOID CHK" + STR$(VoidChkNum&)
    GLifRec(1).Ref = STR$(VoidChkNum&)

    '--make a debit
    GLifRec(1).AcctNum = LEFT$(DistRec(1).DistAcctNum, FundLen) + CashAcct$
    GLifRec(1).CrAmt = 0
    GLifRec(1).DrAmt = DistRec(1).DistAmt
    PUT GLIFFile, , GLifRec(1)
    GTDebit# = GTDebit# + GLifRec(1).DrAmt
    GOSUB PrintGLIF
    
    '--Make a credit
    GLifRec(1).AcctNum = LEFT$(DistRec(1).DistAcctNum, FundLen) + APAcct$
    GLifRec(1).CrAmt = DistRec(1).DistAmt
    GLifRec(1).DrAmt = 0
    PUT GLIFFile, , GLifRec(1)
    GTCredit# = GTCredit# + GLifRec(1).CrAmt
    GOSUB PrintGLIF
    
    NextDist& = DistRec(1).NextDist

  LOOP

  PRINT #PrnFile,
  PRINT #PrnFile, "Grand Total";
  PRINT #PrnFile, TAB(58); USING Fmt$; GTDebit#;
  PRINT #PrnFile, TAB(70); USING Fmt$; GTCredit#

  CLOSE

  Post2GL InterfaceFile$, BadAcct
  KILL InterfaceFile$

  IF BadAcct THEN
    CLS
    QPrintRC "Posting Error.  Review GL Utility Log for details." + STR$(BadAcct), 2, 1, -1
    QPrintRC "Uposted Transactions: " + STR$(BadAcct), 2, 1, -1
    QPrintRC "Press any key for posting report", 3, 1, -1
    KeyPress$ = INPUT$(1)
  END IF

  PrintRptFile PrnTitle$, PrnFileName$, LPTNo, RetCode%, EntryPoint

  OK = MsgBox("AP", "VDCHKMSG")

EXIT SUB
  
PrintReportHeader:
  PRINT #PrnFile, PrnTitle$
  PRINT #PrnFile, "Run Date: " + DATE$
  PRINT #PrnFile,
  PRINT #PrnFile, "Acct No         Date        Description        Ref            Debit      Credit"
  PRINT #PrnFile, "==============================================================================="
  RETURN

PrintGLIF:
  PRINT #PrnFile, GLifRec(1).AcctNum;
  PRINT #PrnFile, TAB(17); Num2Date(GLifRec(1).TRDate);
  PRINT #PrnFile, TAB(29); LEFT$(GLifRec(1).Desc, 17);
  PRINT #PrnFile, TAB(48); GLifRec(1).Ref;
  PRINT #PrnFile, TAB(58); USING Fmt$; GLifRec(1).DrAmt;
  PRINT #PrnFile, TAB(70); USING Fmt$; GLifRec(1).CrAmt
  RETURN
  
END SUB

SUB VoidPrintedCheck
  
  IF NOT Exist("APCHKINF.DAT") THEN EXIT SUB

  REDIM ChkInfo(1 TO 1) AS CheckInfoType3
  ChkInfoRecLen = LEN(ChkInfo(1))
  VCnt = (FileSize("APCHKINF.DAT") \ ChkInfoRecLen)
  REDIM ChkInfo(1 TO VCnt)  AS CheckInfoType3
  FGetAH "APCHKINF.DAT", ChkInfo(1), ChkInfoRecLen, VCnt
  
  REDIM Items(1 TO VCnt) AS FLen2
  REDIM Picked(1 TO VCnt)
  OpenVendorFile APVendorFile, NumVRecs
  FOR Cnt = 1 TO VCnt
    GET APVendorFile, ChkInfo(Cnt).VendorRecNum, VENDOR
  '  IF ChkInfo(Cnt).VoidFlag = True THEN
  '      STOP
  '    PickOffset = PickOffset + 1
  '    Picked(PickOffset) = Cnt
  '  END IF
    'Items(Cnt).V = Vendor.VendorCode
    Items(Cnt).V = STR$(ChkInfo(Cnt).LastChk)
    MID$(Items(Cnt).V, 11) = VENDOR.VNAME
    'MID$(Items(Cnt).V, 37) = Num2Date(APLedgerRec(1).DueDate)
    MID$(Items(Cnt).V, 49) = FUsing$(STR$(ChkInfo(Cnt).ChkAmt), "#######.##")
  NEXT
  
  LibFile2Scrn "AP.QSL", "BG", MonoCode, -1, ErrorCode
  QPrintRC PickHeader$, 4, 7, 11
  QPrintRC "Press SpaceBar To Toggle, Alt-M Marks All, Alt-C Clear All, F10 to Continue.", 25, 3, -1
  
  MPaintBox 5, 74, 5, 76, 8
  PickList Items(), Picked(), NPicked, Cnf
  IF NPicked > 0 THEN
    FOR Cnt = 1 TO NPicked
      ChkInfo(Picked(Cnt)).VoidFlag = True
    NEXT
    FPutAH "APCHKINF.DAT", ChkInfo(1), ChkInfoRecLen, VCnt
  ELSE
    FOR Cnt = 1 TO VCnt
      ChkInfo(Cnt).VoidFlag = False
    NEXT
    FPutAH "APCHKINF.DAT", ChkInfo(1), ChkInfoRecLen, VCnt
  END IF

'GOTO VoidPrintedExit
VoidPrintedExit:

END SUB

