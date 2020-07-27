
'APRPT.BAS (c) Southern Software, Inc.
'by Jim Brown and Dale Watson
DEFINT A-Z
TYPE FLen
  V AS STRING * 64
END TYPE
DECLARE SUB OpenPayablebyDate ()
DECLARE SUB SalesTaxRptE2 ()
DECLARE SUB SalesTaxRichland ()
DECLARE SUB test2 ()
DECLARE SUB test ()
DECLARE SUB SalesTaxReport ()
DECLARE SUB PaidSupplyList ()
DECLARE SUB OpenVendorFile (VendorFile%, NumVRecs%)
DECLARE SUB OpenVendorIdx (VendorIdxFile%, NumActiveVendors%)
DECLARE SUB OpenAPDistFile (APDistFile%, NumDistRecs&, RecLen%)
DECLARE SUB OpenAPLedgerFile (APLedgerFile%, NumTrans%, RecLen%)
DECLARE SUB ClrBG ()
DECLARE SUB VendorHistory ()
DECLARE SUB PrintOpenPO ()
DECLARE SUB POAcctHist ()
DECLARE SUB OpenPayable ()
DECLARE SUB PrintChkList ()
DECLARE SUB VertMenuT2 (Items() AS FLen, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB HideCursor ()
DECLARE SUB MScrnSave (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MScrnRest (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MPaintBox (ULRow%, ULCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB QPrint (x$, Colr%, Page%)
DECLARE SUB QPrintRC (T$, R%, C%, Clr%)
DECLARE SUB ShowCursor ()
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TextCursor (FG%, BG%)
'DECLARE SUB GetFundCodes (FirstFund$, LastFund$)
DECLARE SUB OpenDeptFile (DeptFile%, NumDepts%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB MnuTitle (Title$, Row, Col, MaxLen, Colr)
DECLARE SUB GetFundList (FundList$(), NumFunds%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB KillFile (FileName$)
DECLARE SUB StuffBuf (Ky$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE FUNCTION FUsing$ (Number$, Mask$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION QPTrim$ (x$)
DECLARE FUNCTION WaitKey% ()
DECLARE FUNCTION GetDeptTitle$ (DeptRecNum%)
DECLARE FUNCTION FindDept% (DeptNum$)
DECLARE FUNCTION Num2Month% (Dt%)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION Peek1% (Segment%, Address%)


CONST False = 0, True = NOT False
  
'$INCLUDE: 'DefCnf.bi'                    'monitor and color info
'$INCLUDE: 'FieldInf.bi'                  'field information
'$INCLUDE: 'FormEdit.bi'                  'form editing information
'$INCLUDE: 'GLAUX.BI'                     'Common GL Routines
'$INCLUDE: 'GL.bi'                        'GL File Types
'$INCLUDE: 'AP.bi'                        'AP File Types
'$INCLUDE: 'SCRAUX.BI'
'$INCLUDE: 'QSCR.BI'          'QuickScreen Declarations

DIM SHARED TotLen%, FundLen%, AcctLen%, DetLen%
DIM SHARED Acct AS GLAcctRecType
DIM SHARED AcctIdx AS GLAcctIndexType
DIM SHARED Trans AS GLTransRecType
  
'REDIM SHARED Choice$(0, 0)    'Used to hold multiple choices
  
  '$INCLUDE: 'SetCnf.bi'        'Set monitor and color information
  '$INCLUDE: 'init.bi'
  
  LOCATE 1, 1, 0
  TextCursor -2, -2             'Define the Mouse cursor
  ShowCursor    'Show it
  
  REDIM MChoice$(1 TO 10)
  MChoice$(1) = "Vendor History"
  MChoice$(2) = "Open Payables"
  MChoice$(3) = "Open Payables by Date"
  MChoice$(4) = "Open Purchase Orders"
  MChoice$(5) = "P/O History by G/L Acct"
  MChoice$(6) = "A/P Check Listing"
  MChoice$(7) = "Paid Supply Listing"
  MChoice$(8) = "1099 Forms"
  MChoice$(9) = "Sales Tax Report"
  MChoice$(10) = "Quit"
  
  MaxLen = 0  'Set menu width to zero
  BoxBot = 17   'limit the box length to go no lower than line 20
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  FOR Cnt = 1 TO UBOUND(MChoice$)               'Find max menu width
    TLen = LEN(MChoice$(Cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT
  
  'Center Menu within Screen
  Row = ((25 - (UBOUND(MChoice$))) \ 2) - 1
  Col = ((80 - MaxLen) \ 2) - 2
  help$ = "Accounts Payable Reports Menu"

  DO
    
    LibFile2Scrn "GL.QSL", "MENUBAK", MonoCode, Attribute, ErrorCode
    'TitleBox 5, Col, MaxLen + 3, "Reports Menu ", Cnf
    'TitleBox 18, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    Title$ = "Reports Menu  "
    MnuTitle Title$, Row, Col, MaxLen, 15
    
    PrintTitle User$
    PrintHelp help$
    
    LOCATE Row, Col, 0
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO
    
    SELECT CASE Choice
    CASE 1
      VendorHistory
    CASE 2
      OpenPayable
    CASE 3
      OpenPayablebyDate
    CASE 4
      PrintOpenPO
    CASE 5
      POAcctHist
    CASE 6
      PrintChkList
    CASE 7
      PaidSupplyList
    CASE 8
      'ClrBG
      'OK = MsgBox("GL", "NA")
      RUN "AP1099"
    CASE 9
      ClrBG
      'OK = MsgBox("GL", "NA")
      SalesTaxReport
      'SalesTaxRichland
      'SalesTaxRptE2
      'test
      'test2
    CASE 10
      HideCursor                'Be nice and turn the mouse cursor Off
      CLS
      END
    END SELECT
    
  LOOP
  
  HideCursor    'Be nice and turn the mouse cursor Off
  LibFile2Scrn "GL.QSL", "MENUBAK", MonoCode, Attribute, ErrorCode
  DispLoad
  RUN "CPRDMENU"

FUNCTION Num2Month% (Dt%)
    
    M$ = LEFT$(Num2Date(Dt%), 2)
    Num2Month% = VAL(M$)
    
END FUNCTION

SUB OpenPayable

  SHARED User$
  SHARED Choice$()
  REDIM Choice$(0 TO 3, 1)
  Choice$(0, 0) = "2"
  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn#1"
  Choice$(3, 0) = "Prn#2"
  User$ = QPTrim$(User$)
  Page = 0

  REDIM Frm(1) AS FormInfo
  FormName$ = "PRNOPAY"
  LibName$ = "AP"
  NumFlds = LibNumberOfFields(LibName$, FormName$)
  REDIM Form$(NumFlds, 2)    'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo            'DIM the field information array
  StartEl = 0   'Load first form at array start
  LibGetFldDef LibName$, FormName$, StartEl, Fld(), Form$(), ErrCode
  Frm(1).FldNo = 1           'Start editing on field #1
  Frm(1).InsStat = False     'Set the insert state (-1 = Insert on)
  Frm(1).StartEl = 0         'Set form starting element to 0 and

  Action = 1
  FirstTime = True
  Form$(1, 0) = "N"
  Form$(2, 0) = Choice$(1, 0)

  LibFile2Scrn LibName$, FormName$, MonoCode, -1, ErrCode
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF Frm(1).KeyCode = EscKey THEN
      ExitFlag = True
    ELSEIF Frm(1).KeyCode = F10Key THEN
      OkFlag = True
    END IF
  LOOP UNTIL ExitFlag OR OkFlag

  IF ExitFlag THEN EXIT SUB 'GOTO oExitPreAudit

  LPTPort = VAL(RIGHT$(QPTrim$(Form$(1, 0)), 1))

  IF Form$(1, 0) = "Y" THEN
    ShowDist = True
  ELSE
    ShowDist = False
  END IF

  FF$ = CHR$(12)
  MaxLines = 50
  DblDash$ = STRING$(78, "=")
  PageNum = 0
  Header$ = "Open Payables Report"

  REDIM FundList$(1)
  GetFundList FundList$(), NumFunds
  IF NumFunds = 0 THEN EXIT SUB

  REDIM FundAmts(1 TO NumFunds) AS DOUBLE

  'REDIM APLedgerRec(1) AS APLedgerRecType
  REDIM APLedgerRec(1) AS APLedger81RecType
  RecLen = LEN(APLedgerRec(1))

  REDIM APDistRec(1) AS APDistRecType
  APDistRecLen = LEN(APDistRec(1))

  DIM Vendor AS VendorRecType

  NumVendRecs = (FileSize("apvendor.idx") \ 12)
  REDIM VIndex(1 TO NumVendRecs) AS VendorIdxRecType
  FGetAH "apvendor.idx", VIndex(1), 12, NumVendRecs

  PrintFile = FREEFILE
  PrnFile$ = "PAYABLES.PRN"
  OPEN PrnFile$ FOR OUTPUT AS PrintFile

  OpenVendorFile VendorFile, NumVRecs
  OpenAPLedgerFile APLedgerFile, NumTrans, RecLen
  OpenAPDistFile APDistFile, NumDistRecs&, APDistRecLen

  GOSUB PrintOpenPayRptHeader

  QPrintRC "Processing.. Please wait.", 25, 2, -1

  FOR VCnt = 1 TO NumVendRecs
    
    DoneVHeader = 0
    NumItems = 0
    GET VendorFile, VIndex(VCnt).RecNum, Vendor
    NextTrans& = Vendor.FrstTran
    DO UNTIL NextTrans& = 0
      GET APLedgerFile, NextTrans&, APLedgerRec(1)
      IF APLedgerRec(1).TrCode = 1 AND APLedgerRec(1).PayCode = 1 THEN

        NumItems = NumItems + 1

        IF MaxLines - LineCnt < 5 THEN
          PRINT #PrintFile, FF$
          GOSUB PrintOpenPayRptHeader
          GOSUB PrintVendHeader
          'GOSUB InvHeader
        END IF
        
        '--Print the Vendor Header First time thru
        IF NOT DoneVHeader THEN
          IF LineCnt > MaxLines THEN
            PRINT #PrintFile, FF$
            LineCnt = 0
          END IF
          GOSUB PrintVendHeader
          'IF ShowDist = False THEN
          '  GOSUB InvHeader
          'END IF
        END IF

        IF LineCnt > MaxLines THEN
          PRINT #PrintFile, FF$
          GOSUB PrintOpenPayRptHeader
          GOSUB PrintVendHeader
          'GOSUB InvHeader
        END IF
        
        GOSUB PrintItem
        'IF LineCnt > MaxLines THEN
        '  PRINT #PrintFile, FF$
        '  GOSUB PrintOpenPayRptHeader
        '  GOSUB PrintVendHeader
        '  'GOSUB InvHeader
        'END IF

      END IF
      NextTrans& = APLedgerRec(1).NextTrans
    LOOP

    IF DoneVHeader THEN
      GOSUB FinishVendor
    END IF
    IF LineCnt > MaxLines THEN
      PRINT #PrintFile, FF$
      GOSUB PrintOpenPayRptHeader
    END IF
  NEXT

  PRINT #PrintFile, 'JB
  'PRINT #PrintFile, DblDash$
  IF LineCnt > MaxLines THEN
    PRINT #PrintFile, FF$
  END IF
  GOSUB FinishOpenReport
  PRINT #PrintFile, FF$

  CLOSE

  ERASE FundList$, FundAmts, APLedgerRec, APDistRec, Form$, Fld

  IF LPTPort > 0 THEN
    EntryPoint = 5
  ELSE
    EntryPoint = 2
  END IF

  Title$ = "Open Payables Report"
  PrintRptFile Title$, PrnFile$, LPTPort%, RetCode%, EntryPoint

EXIT SUB

FinishOpenReport:
  PRINT #PrintFile, "Report Totals:"
  PRINT #PrintFile, "Vendors with Open Invoices: "; FUsing(STR$(ChkCnt), ",#############")
  PRINT #PrintFile, "                  Totaling: "; FUsing(STR$(TotalChkAmt#), "$$########,.##")
  IF ShowDist THEN
    PRINT #PrintFile, "Total Open By Fund:"
    FOR FundCnt = 1 TO NumFunds
      IF FundAmts(FundCnt) > 0 THEN
         PRINT #PrintFile, FundList$(FundCnt); TAB(27); FUsing(STR$(FundAmts(FundCnt)), "$$########,.##")
      END IF
    NEXT
  END IF
RETURN

PrintOpenPayRptHeader:
Page = Page + 1
  PRINT #PrintFile, TAB(40 - (INT(LEN(User$) / 2))); User$
  PRINT #PrintFile, TAB(40 - (INT(LEN(Header$) / 2))); Header$
  PRINT #PrintFile,
  PRINT #PrintFile, "Report Date: "; DATE$; TAB(67); "Page #"; Page
  PRINT #PrintFile,
  PRINT #PrintFile, "Inv Num                   Inv Date    Due Date    PO                  Amount"
  PRINT #PrintFile, STRING$(78, "=")
  LineCnt = 6
  
RETURN


InvHeader:

  'PRINT #PrintFile,
  'PRINT #PrintFile, "Inv Num                   Inv Date    Due Date    PO                  Amount"
  'PRINT #PrintFile, STRING$(78, "-")
  'LineCnt = LineCnt + 2
  '"----------  ----------  -------------------------  ----------    ------------"
RETURN

PrintVendHeader:
  'PRINT #PrintFile,
  'PRINT #PrintFile, DblDash$
  PRINT #PrintFile, Vendor.VNum; TAB(15); Vendor.Vname
  DoneVHeader = -1
  LineCnt = LineCnt + 1

RETURN

PrintItem:
  'IF ShowDist THEN
  '  GOSUB InvHeader
  'END IF

  ToPrint$ = SPACE$(80)
  MID$(ToPrint$, 1) = LEFT$(APLedgerRec(1).DOCNum, 25)
  MID$(ToPrint$, 27) = Num2Date(APLedgerRec(1).TrDate)
  MID$(ToPrint$, 39) = Num2Date(APLedgerRec(1).DUEDATE)
  MID$(ToPrint$, 51) = LEFT$(APLedgerRec(1).PONUM, 10)
  MID$(ToPrint$, 66) = FUsing(STR$(APLedgerRec(1).Amt), ",########.##")
  PRINT #PrintFile, ToPrint$
  LineCnt = LineCnt + 1

  VendTotal# = Round(VendTotal# + APLedgerRec(1).Amt)
  TotalChkAmt# = Round(TotalChkAmt# + APLedgerRec(1).Amt)
  NextDist& = APLedgerRec(1).FrstDist

  IF ShowDist THEN

    PRINT #PrintFile,

    ToPrint$ = SPACE$(80)
    MID$(ToPrint$, 5) = "Accounting Distribution"
    PRINT #PrintFile, ToPrint$
    LineCnt = LineCnt + 1

    ToPrint$ = SPACE$(80)
    MID$(ToPrint$, 10) = "Acct Number"
    MID$(ToPrint$, 28) = "Description"
    MID$(ToPrint$, 56) = "      Amount"
    PRINT #PrintFile, ToPrint$
    LineCnt = LineCnt + 1

    ToPrint$ = SPACE$(80)
    DO UNTIL NextDist& = 0
      GET APDistFile, NextDist&, APDistRec(1)
      InvTotal# = InvTotal# + APDistRec(1).DistAmt

      DistAcctRec = FindAcct(APDistRec(1).DistAcctNum)
      AcctName$ = GetAcctTitle(DistAcctRec)

      IF LineCnt > MaxLines THEN
        PRINT #PrintFile, FF$
        GOSUB PrintOpenPayRptHeader
      END IF

      MID$(ToPrint$, 10) = APDistRec(1).DistAcctNum
      MID$(ToPrint$, 28) = AcctName$
      MID$(ToPrint$, 56) = FUsing(STR$(APDistRec(1).DistAmt), ",########.##")
      PRINT #PrintFile, ToPrint$
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

    ToPrint$ = SPACE$(80)
    MID$(ToPrint$, 56) = STRING$(13, "-")
    PRINT #PrintFile, ToPrint$

    ToPrint$ = SPACE$(80)
    MID$(ToPrint$, 28) = "Total Distributed"
    MID$(ToPrint$, 56) = FUsing(STR$(InvTotal#), ",########.##")
    PRINT #PrintFile, ToPrint$

    LineCnt = LineCnt + 2
    InvTotal# = 0

  END IF

  IF ShowDist THEN
    PRINT #PrintFile,
    PRINT #PrintFile, STRING$(78, "-")
    LineCnt = LineCnt + 2
  END IF
  
RETURN


FinishVendor:
  IF ShowDist THEN
    IF NumItems > 1 THEN
      PRINT #PrintFile, QPTrim$(Vendor.Vname); " Total: "; TAB(66); FUsing(STR$(VendTotal#), ",########.##")
      PRINT #PrintFile, STRING$(78, "-")
      LineCnt = LineCnt + 2
    'ELSE
    '  PRINT #PrintFile, DblDash$
    '  LineCnt = LineCnt + 1
    END IF
  ELSE
    IF NumItems > 1 THEN
      PRINT #PrintFile, TAB(66); "------------"
      PRINT #PrintFile, QPTrim$(Vendor.Vname); " Total: "; TAB(66); FUsing(STR$(VendTotal#), ",########.##")
      PRINT #PrintFile, DblDash$
      LineCnt = LineCnt + 3 '2
    ELSE
      PRINT #PrintFile, DblDash$
      LineCnt = LineCnt + 1
    END IF
  END IF
  VendTotal# = 0
  ChkCnt = ChkCnt + 1
  VActive = 0

RETURN

  
END SUB

SUB OpenPayablebyDate

  SHARED User$
  SHARED Choice$()
  REDIM Choice$(0 TO 3, 1)
  Choice$(0, 0) = "3"
  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn#1"
  Choice$(3, 0) = "Prn#2"
  User$ = QPTrim$(User$)
  Page = 0
  Header$ = "Open Payables Report"
  REDIM Frm(1) AS FormInfo
  FormName$ = "PRNOPYD"
  LibName$ = "AP"
  NumFlds = LibNumberOfFields(LibName$, FormName$)
  REDIM Form$(NumFlds, 2)    'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo            'DIM the field information array
  StartEl = 0   'Load first form at array start
  LibGetFldDef LibName$, FormName$, StartEl, Fld(), Form$(), ErrCode
  Frm(1).FldNo = 1           'Start editing on field #1
  Frm(1).InsStat = False     'Set the insert state (-1 = Insert on)
  Frm(1).StartEl = 0         'Set form starting element to 0 and

  Action = 1
  FirstTime = True
  Form$(1, 0) = "N"
  Form$(3, 0) = Choice$(1, 0)

  LibFile2Scrn LibName$, FormName$, MonoCode, -1, ErrCode
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF Frm(1).KeyCode = EscKey THEN
      ExitFlag = True
    ELSEIF Frm(1).KeyCode = F10Key THEN
      OkFlag = True
    END IF
  LOOP UNTIL ExitFlag OR OkFlag

  IF ExitFlag THEN EXIT SUB 'GOTO oExitPreAudit

  LPTPort = VAL(RIGHT$(QPTrim$(Form$(1, 0)), 1))

  IF Form$(1, 0) = "Y" THEN
    ShowDist = True
  ELSE
    ShowDist = False
  END IF

  OpenDate = Date2Num(Form$(2, 0))

  FF$ = CHR$(12)
  MaxLines = 50
  DblDash$ = STRING$(78, "=")
  PageNum = 0

  REDIM FundList$(1)
  GetFundList FundList$(), NumFunds
  IF NumFunds = 0 THEN EXIT SUB

  REDIM FundAmts(1 TO NumFunds) AS DOUBLE

  'REDIM APLedgerRec(1) AS APLedgerRecType
  REDIM APLedgerRec(1) AS APLedger81RecType
  RecLen = LEN(APLedgerRec(1))

  REDIM APDistRec(1) AS APDistRecType
  APDistRecLen = LEN(APDistRec(1))

  DIM Vendor AS VendorRecType

  NumVendRecs = (FileSize("apvendor.idx") \ 12)
  REDIM VIndex(1 TO NumVendRecs) AS VendorIdxRecType
  FGetAH "apvendor.idx", VIndex(1), 12, NumVendRecs

  PrintFile = FREEFILE
  PrnFile$ = "PAYABLES.PRN"
  OPEN PrnFile$ FOR OUTPUT AS PrintFile

  OpenVendorFile VendorFile, NumVRecs
  OpenAPLedgerFile APLedgerFile, NumTrans, RecLen
  OpenAPDistFile APDistFile, NumDistRecs&, APDistRecLen

  GOSUB PrintRptHeader

  QPrintRC "Processing.. Please wait.", 25, 2, -1

  FOR VCnt = 1 TO NumVendRecs
    DoneVHeader = 0
    NumItems = 0
    GET VendorFile, VIndex(VCnt).RecNum, Vendor
    NextTrans& = Vendor.FrstTran
    DO UNTIL NextTrans& = 0
      GET APLedgerFile, NextTrans&, APLedgerRec(1)

      'IF APLedgerRec(1).TrCode = 1 AND APLedgerRec(1).PAYCODE = 1 THEN

      '--Filter Invoices dated earlier than search date
      IF APLedgerRec(1).TrCode = 1 AND APLedgerRec(1).GLDistDate <= OpenDate THEN

        '--Then filter those which are unpaid or paid after the search date
        IF APLedgerRec(1).PayCode = 1 OR APLedgerRec(1).PdCheckDate > OpenDate THEN
              
              NumItems = NumItems + 1

              IF MaxLines - LineCnt < 5 THEN
                PRINT #PrintFile, FF$
                GOSUB PrintRptHeader
                GOSUB PrintVendHeader1
                'GOSUB InvHeader
              END IF

              '--Print the Vendor Header First time thru
              IF NOT DoneVHeader THEN
                IF LineCnt > MaxLines THEN
                  PRINT #PrintFile, FF$
                  LineCnt = 0
                END IF
                GOSUB PrintVendHeader1
                'IF ShowDist = False THEN
                '  GOSUB InvHeader
                'END IF
              END IF

              IF LineCnt > MaxLines THEN
                PRINT #PrintFile, FF$
                GOSUB PrintRptHeader
                GOSUB PrintVendHeader1
                'GOSUB InvHeader
              END IF

              GOSUB PrintItem1
              'IF LineCnt > MaxLines THEN
              '  PRINT #PrintFile, FF$
              '  GOSUB PrintOpenPayRptHeader
              '  GOSUB PrintVendHeader
              '  'GOSUB InvHeader
              'END IF
        
        END IF
      END IF
      NextTrans& = APLedgerRec(1).NextTrans
    LOOP

    IF DoneVHeader THEN
      GOSUB FinishVendor1
    END IF
    IF LineCnt > MaxLines THEN
      PRINT #PrintFile, FF$
      GOSUB PrintRptHeader
    END IF
  NEXT

  PRINT #PrintFile, 'JB
  'PRINT #PrintFile, DblDash$
  IF LineCnt > MaxLines THEN
    PRINT #PrintFile, FF$
  END IF
  GOSUB FinishOpenReport1
  PRINT #PrintFile, FF$

  CLOSE

  ERASE FundList$, FundAmts, APLedgerRec, APDistRec, Form$, Fld

  IF LPTPort > 0 THEN
    EntryPoint = 5
  ELSE
    EntryPoint = 2
  END IF

  Title$ = "Open Payables Report"
  PrintRptFile Title$, PrnFile$, LPTPort%, RetCode%, EntryPoint

EXIT SUB

FinishOpenReport1:
  PRINT #PrintFile, "Report Totals:"
  PRINT #PrintFile, "Vendors with Open Invoices: "; FUsing(STR$(ChkCnt), ",#############")
  PRINT #PrintFile, "                  Totaling: "; FUsing(STR$(TotalChkAmt#), "$$########,.##")
  IF ShowDist THEN
    PRINT #PrintFile, "Total Open By Fund:"
    FOR FundCnt = 1 TO NumFunds
      IF FundAmts(FundCnt) > 0 THEN
         PRINT #PrintFile, FundList$(FundCnt); TAB(27); FUsing(STR$(FundAmts(FundCnt)), "$$########,.##")
      END IF
    NEXT
  END IF
RETURN

PrintRptHeader:
Page = Page + 1
  PRINT #PrintFile, TAB(40 - (INT(LEN(User$) / 2))); User$
  PRINT #PrintFile, TAB(40 - (INT(LEN(Header$) / 2))); Header$
  PRINT #PrintFile,
  PRINT #PrintFile, "Report Date: "; DATE$; TAB(67); "Page #"; Page
  PRINT #PrintFile, "Inv Num                   Dist Date   Date Paid    PO                  Amount"
  PRINT #PrintFile, STRING$(78, "=")
  LineCnt = 6

RETURN


InvHeader1:

  'PRINT #PrintFile,
  'PRINT #PrintFile, "Inv Num                   Inv Date   Paid Date    PO                  Amount"
  'PRINT #PrintFile, STRING$(78, "-")
  'LineCnt = LineCnt + 2
  '"----------  ----------  -------------------------  ----------    ------------"
RETURN

PrintVendHeader1:
  'PRINT #PrintFile,
  'PRINT #PrintFile, DblDash$
  PRINT #PrintFile, Vendor.VNum; TAB(15); Vendor.Vname
  DoneVHeader = -1
  LineCnt = LineCnt + 1

RETURN

PrintItem1:
  'IF ShowDist THEN
  '  GOSUB InvHeader
  'END IF

  ToPrint$ = SPACE$(80)
  MID$(ToPrint$, 1) = LEFT$(APLedgerRec(1).DOCNum, 25)
  'MID$(ToPrint$, 27) = Num2Date(APLedgerRec(1).TRDate)
  MID$(ToPrint$, 27) = Num2Date(APLedgerRec(1).GLDistDate)
  IF APLedgerRec(1).PdCheckDate = 0 THEN
     MID$(ToPrint$, 39) = "Unpaid"
  ELSE
    MID$(ToPrint$, 39) = Num2Date(APLedgerRec(1).PdCheckDate)
  END IF
  MID$(ToPrint$, 51) = LEFT$(APLedgerRec(1).PONUM, 10)
  MID$(ToPrint$, 66) = FUsing(STR$(APLedgerRec(1).Amt), ",########.##")
  PRINT #PrintFile, ToPrint$
  LineCnt = LineCnt + 1

  VendTotal# = Round(VendTotal# + APLedgerRec(1).Amt)
  TotalChkAmt# = Round(TotalChkAmt# + APLedgerRec(1).Amt)
  NextDist& = APLedgerRec(1).FrstDist

  IF ShowDist THEN

    PRINT #PrintFile,

    ToPrint$ = SPACE$(80)
    MID$(ToPrint$, 5) = "Accounting Distribution"
    PRINT #PrintFile, ToPrint$
    LineCnt = LineCnt + 1

    ToPrint$ = SPACE$(80)
    MID$(ToPrint$, 10) = "Acct Number"
    MID$(ToPrint$, 28) = "Description"
    MID$(ToPrint$, 56) = "      Amount"
    PRINT #PrintFile, ToPrint$
    LineCnt = LineCnt + 1

    ToPrint$ = SPACE$(80)
    DO UNTIL NextDist& = 0
      GET APDistFile, NextDist&, APDistRec(1)
      InvTotal# = InvTotal# + APDistRec(1).DistAmt

      DistAcctRec = FindAcct(APDistRec(1).DistAcctNum)
      
      AcctName$ = GetAcctTitle(DistAcctRec)

      IF LineCnt > MaxLines THEN
        PRINT #PrintFile, FF$
        GOSUB PrintRptHeader
      END IF

      MID$(ToPrint$, 10) = APDistRec(1).DistAcctNum
      MID$(ToPrint$, 28) = AcctName$
      MID$(ToPrint$, 56) = FUsing(STR$(APDistRec(1).DistAmt), ",########.##")
      PRINT #PrintFile, ToPrint$
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

    ToPrint$ = SPACE$(80)
    MID$(ToPrint$, 56) = STRING$(13, "-")
    PRINT #PrintFile, ToPrint$

    ToPrint$ = SPACE$(80)
    MID$(ToPrint$, 28) = "Total Distributed"
    MID$(ToPrint$, 56) = FUsing(STR$(InvTotal#), ",########.##")
    PRINT #PrintFile, ToPrint$

    LineCnt = LineCnt + 2
    InvTotal# = 0

  END IF

  IF ShowDist THEN
    PRINT #PrintFile,
    PRINT #PrintFile, STRING$(78, "-")
    LineCnt = LineCnt + 2
  END IF

RETURN


FinishVendor1:
  IF ShowDist THEN
    IF NumItems > 1 THEN
      PRINT #PrintFile, QPTrim$(Vendor.Vname); " Total: "; TAB(66); FUsing(STR$(VendTotal#), ",########.##")
      PRINT #PrintFile, STRING$(78, "-")
      LineCnt = LineCnt + 2
    'ELSE
    '  PRINT #PrintFile, DblDash$
    '  LineCnt = LineCnt + 1
    END IF
  ELSE
    IF NumItems > 1 THEN
      PRINT #PrintFile, TAB(66); "------------"
      PRINT #PrintFile, QPTrim$(Vendor.Vname); " Total: "; TAB(66); FUsing(STR$(VendTotal#), ",########.##")
      PRINT #PrintFile, DblDash$
      LineCnt = LineCnt + 3 '2
    ELSE
      PRINT #PrintFile, DblDash$
      LineCnt = LineCnt + 1
    END IF
  END IF
  VendTotal# = 0
  ChkCnt = ChkCnt + 1
  VActive = 0

RETURN


END SUB

SUB PaidSupplyList

  LOCATE 1, 1, 0

  SHARED FundLen
  SHARED Choice$()
  SHARED User$

  REDIM Choice$(0 TO 3, 1)
  Choice$(0, 0) = "4"
  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn#1"
  Choice$(3, 0) = "Prn#2"
  User$ = QPTrim$(User$)
  Page = 0
  REDIM Form$(0, 0)
  REDIM Fld(0) AS FieldInfo
  DIM Frm AS FormInfo

  LibName$ = "AP.QSL"
  FormName$ = "PRNPSL"
  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()

  '--Assign Default Values
  'Form$(3, 0) = Choice$(1, 1)
  'Form$(4, 0) = "N"
  'Form$(5, 0) = Choice$(1, 0)

  Action = 1
  CancelFlag = False

  DO
    EditForm Form$(), Fld(), Frm, Cnf, Action
    SELECT CASE Frm.KeyCode
      CASE F10Key
        EXIT DO
      CASE ESC
        CancelFlag = True
        EXIT DO
    END SELECT
  LOOP

  IF CancelFlag THEN EXIT SUB

  BegDate = Date2Num(Form$(1, 0))
  EndDate = Date2Num(Form$(2, 0))
  Dev$ = LEFT$(Form$(4, 0), 1)
  IF Dev$ = "P" THEN LPTNo = VAL(RIGHT$(Form$(4, 0), 1))
  IF Form$(3, 0) = "S" THEN ShowSelected = True

  'ChkNum$ = SPACE$(14)
  CommaFmt$ = "########,.##"
  FF$ = CHR$(12)
  MaxLines = 55
  
  DIM APLedger AS APLedger81RecType
  APLRecLen = LEN(APLedger)
  OpenAPLedgerFile APLedgerFile, NumTrans, APLRecLen

  REDIM ChkList(1 TO 1) AS GLAcctIndexType  '--borrowing this type
  ChkNum$ = SPACE$(14)
  
  DIM APDist AS APDistRecType
  APDRecLen = LEN(APDist)
  OpenAPDistFile APDistFile, NumDistRecs&, APDRecLen

  RptFile = FREEFILE
  RptFileName$ = "PSL.PRN"
  OPEN RptFileName$ FOR OUTPUT AS RptFile

  DIM Vendor AS VendorRecType
  VRecLen = LEN(Vendor)
  OpenVendorFile VendorFile, NumVRecs

  GOSUB PSLPageHdr
  
  'Get a list of checks
  FOR Cnt = 1 TO NumTrans
    Pct$ = STR$(INT((Cnt / NumTrans) * 100))
    QPrintRC "Searching for checks..." + Pct$ + "% Completed", 25, 2, -1

    GET APLedgerFile, Cnt, APLedger
    GET VendorFile, APLedger.VRecNum, Vendor

    IF APLedger.TrCode = 3 THEN
      IF APLedger.TrDate >= BegDate AND APLedger.TrDate <= EndDate THEN
        NumChks = NumChks + 1
        REDIM PRESERVE ChkList(1 TO NumChks) AS GLAcctIndexType
        ChkList(NumChks).RecNum = Cnt
        RSET ChkNum$ = QPTrim$(APLedger.DOCNum)
        ChkList(NumChks).AcctNum = ChkNum$
     END IF
    END IF
  NEXT

  IF NumChks > 0 THEN
    GOSUB SortPSL
    GOSUB SearchforPSLTrans
    PRINT #RptFile, USING "Total Checks ########,.##"; RunTotal#;
    PRINT #RptFile, FF$
  ELSE
    PRINT #RptFile, "  No Checks on file."
  END IF

  CLOSE

  SELECT CASE Dev$
   CASE "S"
     EntryPoint = 2
   CASE "P"
     EntryPoint = 5
  END SELECT

  PrintRptFile Header$, RptFileName$, LPTNo, RetCode%, EntryPoint

EXIT SUB


SortPSL:
  SortT ChkList(1), NumChks, 0, 16, 0, 14
RETURN

SearchforPSLTrans:
  QPrintRC SPACE$(40), 25, 2, -1

  '--search thru check list
  FOR Cnt = 1 TO NumChks
    Pct$ = STR$(INT((Cnt / NumChks) * 100))
    QPrintRC "Searching for PSL Transactions..." + Pct$ + "% Complete.", 25, 2, -1
    ChkNum$ = SPACE$(14)

    '--Get the pointers we need
    GET APLedgerFile, ChkList(Cnt).RecNum, APLedger
    GET VendorFile, APLedger.VRecNum, Vendor
    NextTrans& = Vendor.FrstTran

    '--Look for this check Number in the pd check num field
    CheckNumber& = VAL(ChkList(Cnt).AcctNum)
    PrintedCheck = False
    
    '--Search thru the ledger file to see if trans is flagged for PSL
    DO UNTIL NextTrans& = 0
      GET APLedgerFile, NextTrans&, APLedger
      IF APLedger.PDCheckNum = CheckNumber& THEN
        IF ShowSelected THEN
          IF APLedger.PSLFlag = "Y" THEN
            '--print check info only one time - check may be for multiple invoices
            IF NOT PrintedCheck THEN
              GOSUB PrintCheckInfo
            END IF
            NextDist& = APLedger.FrstDist
            DO
              GET APDistFile, NextDist&, APDist
              DistAmt# = DistAmt# + APDist.DistAmt
              NextDist& = APDist.NextDist
              GOSUB PrintDist
            LOOP UNTIL NextDist& = 0
          END IF
        ELSE
          IF NOT PrintedCheck THEN
            GOSUB PrintCheckInfo
          END IF
          NextDist& = APLedger.FrstDist
          DO
            GET APDistFile, NextDist&, APDist
            DistAmt# = DistAmt# + APDist.DistAmt
            NextDist& = APDist.NextDist
            GOSUB PrintDist
          LOOP UNTIL NextDist& = 0
        END IF
      END IF
      NextTrans& = APLedger.NextTrans
    LOOP
    IF PrintedCheck THEN GOSUB PrintChkTotals
  NEXT

RETURN

PSLPageHdr:
  Page = Page + 1
  PRINT #RptFile, TAB(40 - (LEN(User$) / 2)); User$
  PRINT #RptFile, TAB(30); " Paid Supply List                                                  Page:"; Page
  PRINT #RptFile, TAB(27); Form$(1, 0) + " thru " + Form$(2, 0)
  PRINT #RptFile,
  PRINT #RptFile, "Chk#     Vendor                     Description"
  PRINT #RptFile, "                                    Department                          Amount"
  PRINT #RptFile, STRING$(78, "=")
  LineCnt = 7
RETURN

PrintCheckInfo:
  PrintedCheck = True
  RunTotal# = RunTotal# + APLedger.Amt
  ToPrint$ = SPACE$(80)
  MID$(ToPrint$, 2) = STR$(CheckNumber&)
  MID$(ToPrint$, 10) = Vendor.Vname
  MID$(ToPrint$, 37) = APLedger.Comment
  PRINT #RptFile, ToPrint$
  LineCnt = LineCnt + 1
  IF LineCnt > MaxLines THEN
     PRINT #RptFile, FF$
     GOSUB PSLPageHdr
  END IF
RETURN

PrintDist:
  DeptCode$ = MID$(APDist.DistAcctNum, FundLen + 2, AcctLen)
  DeptRecNum = FindDept%(DeptCode$)
  IF DeptRecNum > 0 THEN
    DeptName$ = GetDeptTitle$(DeptRecNum)
  ELSE
    DeptName$ = "Undefined Dept" + DeptCode$
  END IF
  ToPrint$ = SPACE$(80)
  MID$(ToPrint$, 37) = DeptName$
  MID$(ToPrint$, 67) = FUsing(STR$(APDist.DistAmt), CommaFmt$)
  PRINT #RptFile, ToPrint$
  LineCnt = LineCnt + 1
  IF LineCnt > MaxLines THEN
     PRINT #RptFile, FF$
     GOSUB PSLPageHdr
  END IF
RETURN

PrintChkTotals:
  ToPrint$ = SPACE$(80)
  MID$(ToPrint$, 67) = "------------"
  PRINT #RptFile, ToPrint$
  ToPrint$ = SPACE$(80)
  MID$(ToPrint$, 67) = FUsing(STR$(DistAmt#), CommaFmt$)
  PRINT #RptFile, ToPrint$
  ToPrint$ = SPACE$(80)
  PRINT #RptFile, ToPrint$
  LineCnt = LineCnt + 3
  IF LineCnt > MaxLines THEN
     PRINT #RptFile, FF$
     GOSUB PSLPageHdr
  END IF
  DistAmt# = 0
RETURN


END SUB

SUB POAcctHist

  REDIM Desc$(1)

  LibName$ = "AP.QSL"           'screen library
  FormName$ = "PRNPOHST"        'form to edit
  DevFld = 5    'Set to output device field number

  'SHARED Cnf AS Config
  SHARED Choice$()

  REDIM Form$(0, 0)
  REDIM Fld(0) AS FieldInfo
  REDIM Choice$(0 TO 3, 1)
  DIM Frm AS FormInfo

  Choice$(0, 0) = STR$(DevFld)  'Choices for Output Device
  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn#1"
  Choice$(3, 0) = "Prn#2"

  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()

  Action = 1

  DO            'Allow user to input report params

    EditForm Form$(), Fld(), Frm, Cnf, Action

    IF Frm.PrevFld <> Frm.FldNo THEN            'User is moving around in form

      IF Frm.FldNo = 1 OR Frm.FldNo = 2 THEN
        Form$(Frm.FldNo, 0) = ""
        Action = 1
      END IF

      '--strip out dashes for editing
      IF Frm.FldNo = 1 OR Frm.FldNo = 2 THEN
        IF LEN(Form$(Frm.FldNo, 0)) > 0 THEN
          Form$(Frm.FldNo, 0) = StripAcct$(Form$(Frm.FldNo, 0))
          SaveField Frm.FldNo, Form$(), Fld(), 0
          Action = 1
        END IF
      END IF

      SELECT CASE Frm.PrevFld
      CASE 1, 2
        Form$(Frm.PrevFld, 0) = FmtAcct$(Form$(Frm.PrevFld, 0), FundLen%, AcctLen%, DetLen%)
      CASE ELSE
      END SELECT
      Action = 1

    END IF

    SELECT CASE Frm.KeyCode     'Check the user's key presses

    CASE 42
      Form$(1, 0) = "First Account"
      Form$(2, 0) = "Last Account"
      Form$(3, 0) = Num2Date(GetBegDate)
      Form$(4, 0) = Num2Date(GetEndDate)
      Frm.FldNo = DevFld
      Action = 1

    CASE F10Key 'SaveButton
      'First Acct
      IF UCASE$(LEFT$(QPTrim$(Form$(1, 0)), 1)) = "F" THEN
        FirstAcct$ = "000000000000"
      ELSE
        FirstAcct$ = QPTrim$(Form$(1, 0))
      END IF

      'Last Acct
      IF UCASE$(LEFT$(QPTrim$(Form$(2, 0)), 1)) = "L" THEN
        LastAcct$ = "999999999999"
      ELSE
        LastAcct$ = QPTrim$(Form$(2, 0))
      END IF

      BegDate = Date2Num(Form$(3, 0))
      EndDate = Date2Num(Form$(4, 0))

      IF LEN(LTRIM$(RTRIM$(Form$(DevFld, 0)))) = 0 THEN
        Dev$ = "S"
        LPTNo = 1
      ELSE
        Dev$ = LEFT$(Form$(DevFld, 0), 1)
        LPTNo = VAL(RIGHT$(RTRIM$(Form$(DevFld, 0)), 1))
      END IF

      IF BegDate > EndDate OR Form$(1, 0) > Form$(2, 0) THEN
        OK = MsgBox("GL", "BADPARMS")
        Ok2Go = 0
        Frm.FldNo = 5
        Action = 1
      ELSE
        Ok2Go = -1
      END IF

    CASE ESC
      EXIT SUB

    END SELECT

  LOOP UNTIL Ok2Go

  LOCATE , , 0

  IF Dev$ = "F" THEN Dev$ = "S"
  
  'End of Input
  '=====================================================
  'Start Report Processing

  ReportFile$ = "POHIST.PRN"   'Report File Name
  CommaFmt$ = "########,.##"    'format takes 13 chars
  TotalFmt$ = "#########,.##"   'format takes 14 chars
  SumLine$ = STRING$(13, "-")   'column summary line
  DivLine$ = STRING$(77, "-")   'dashed line
  DivLine2$ = STRING$(77, "=")  'Double Line
  CrLf$ = CHR$(13) + CHR$(10)

  Header$ = "Purchase Order History"
  Desc$(1) = "Date       Description             Reference       Debit        Credit  Post Ref"

  TotDr# = 0
  TotCr# = 0
  OpenDate$ = ConvDateStr1(Num2Date$(GetBegDate))
  OpenDesc$ = "Opening Balance"

  PrnFile = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #PrnFile

  PrintHelp "   Processing:"

  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum, NumGLAcctRecs

  SHARED Trans AS GLTransRecType
  TransLen = LEN(Trans)
  TransFile = FREEFILE
  OPEN "potrans.dat" FOR RANDOM ACCESS READ WRITE SHARED AS TransFile LEN = TransLen
  NumTrans = LOF(TransFile) \ TransLen

  IF NumTrans = 0 THEN
    CLOSE
    EXIT SUB
  END IF

  DIM TrSort(1 TO NumTrans)  AS TrSortType

  GrTotDr# = 0
  GrTotCr# = 0

  FOR Cnt = 1 TO NumGLAccts

    DrFwd# = 0
    CrFwd# = 0
    TotAcctDr# = 0
    TotAcctCr# = 0

    GET AcctIdxFileNum, Cnt, AcctIdx
    GET AcctFileNum, AcctIdx.RecNum, Acct

    AcctNum$ = QPTrim$(Acct.Num)
    PrintHelp "Processing account: " + AcctNum$

    IF AcctNum$ >= FirstAcct$ AND AcctNum$ <= LastAcct$ THEN
      IF Acct.Typ = "E" THEN

        NextTr = Acct.FrstPTran 'get the first trans for this acct
        AcctNumber$ = "Account " + AcctNum$ + " - " + QPTrim$(Acct.Title)

        ToPrint$ = SPACE$(80)
        LSET ToPrint$ = AcctNumber$
        PRINT #PrnFile, ToPrint$

        DO UNTIL NextTr = 0     'keep going 'til we run out of trans

          GET TransFile, NextTr, Trans

          IF Trans.TrDate >= BegDate AND Trans.TrDate <= EndDate THEN
            '--within range - assign to array for sorting
            NumTrans = NumTrans + 1
            TrSort(NumTrans).TrDate = Trans.TrDate
            TrSort(NumTrans).Record = NextTr

          ELSE
            '--check the transaction to see if we need to carry it in
            '  the balance fwd
            IF Trans.TrDate < BegDate THEN
              DrFwd# = DrFwd# + Trans.DrAmt
              CrFwd# = CrFwd# + Trans.CrAmt
              FwdFlag = -1
            END IF
          END IF

          NextTr = Trans.NextTran            'Get the next transaction

        LOOP

        IF FwdFlag THEN
          '--
          'FwdFlag = 0
          ToPrint$ = SPACE$(80)
          MID$(ToPrint$, 1) = "Balance Forward"

          BalFwd# = DrFwd# - CrFwd#
          IF BalFwd# >= 0 THEN
            Debit$ = FUsing$(STR$(BalFwd#), CommaFmt$)
            Credit$ = ""
          ELSE
            Credit$ = FUsing$(STR$(ABS(BalFwd#)), CommaFmt$)
            Debit$ = ""
          END IF
          
          MID$(ToPrint$, 45) = Debit$
          MID$(ToPrint$, 59) = Credit$
          PRINT #PrnFile, ToPrint$

        END IF

        IF NumTrans > 0 OR BalFwd# <> 0 THEN

          PrintHelp "Sorting Transactions..."

          SortT TrSort(1), NumAcctTrans, 0, 6, 0, -1

          PrintHelp "Writing to report file..."
          FOR Trn = 1 TO NumTrans
            GET TransFile, TrSort(Trn).Record, Trans

            ToPrint$ = SPACE$(80)

            MID$(ToPrint$, 1) = ConvDateStr1(Num2Date(Trans.TrDate))
            MID$(ToPrint$, 12) = Trans.Desc
            MID$(ToPrint$, 36) = Trans.Ref

            IF Trans.DrAmt <> 0 THEN
              MID$(ToPrint$, 45) = FUsing$(STR$(Trans.DrAmt), CommaFmt$)
            END IF

            IF Trans.CrAmt <> 0 THEN
              MID$(ToPrint$, 59) = FUsing$(STR$(Trans.CrAmt), CommaFmt$)
            END IF

            MID$(ToPrint$, 74) = LEFT$(Trans.Src, 6)

            PRINT #PrnFile, ToPrint$

            TotAcctDr# = TotAcctDr# + Trans.DrAmt
            TotAcctCr# = TotAcctCr# + Trans.CrAmt

            GrTotDr# = GrTotDr# + Trans.DrAmt
            GrTotCr# = GrTotCr# + Trans.CrAmt

          NEXT

          '--Print summary lines
          ToPrint$ = SPACE$(80)
          MID$(ToPrint$, 44) = SumLine$
          MID$(ToPrint$, 58) = SumLine$
          PRINT #PrnFile, ToPrint$

          '--Print transaction totals
          IF NumAcctTrans > 0 THEN
            ToPrint$ = SPACE$(80)
            MID$(ToPrint$, 1) = "Transaction Totals"
            MID$(ToPrint$, 44) = FUsing$(STR$(TotAcctDr#), TotalFmt$)
            MID$(ToPrint$, 58) = FUsing$(STR$(TotAcctCr#), TotalFmt$)
            PRINT #PrnFile, ToPrint$
          END IF

          '--Print ending balance
          ToPrint$ = SPACE$(80)
          MID$(ToPrint$, 1) = "Encumbered Balance"
          SELECT CASE Acct.Typ
          CASE "E"
            AcctBal# = BalFwd# + TotAcctDr# - TotAcctCr#
            IF AcctBal# >= 0 THEN
              Debit$ = FUsing$(STR$(AcctBal#), TotalFmt$)
              Credit$ = ""
            ELSE
              Credit$ = FUsing$(STR$(ABS(AcctBal#)), TotalFmt$)
              Debit$ = ""
            END IF

          CASE "R"
            AcctBal# = BalFwd# + TotAcctCr# - TotAcctDr#
            IF AcctBal# >= 0 THEN
              Credit$ = FUsing$(STR$(AcctBal#), TotalFmt$)
              Debit$ = ""
            ELSE
              Debit$ = FUsing$(STR$(ABS(AcctBal#)), TotalFmt$)
              Credit$ = ""
            END IF

          END SELECT

          MID$(ToPrint$, 44) = Debit$
          MID$(ToPrint$, 58) = Credit$
          PRINT #PrnFile, ToPrint$

          ToPrint$ = STRING$(80, "*")
          PRINT #PrnFile, ToPrint$

        ELSE
          ToPrint$ = SPACE$(80)
          MID$(ToPrint$, 5) = " -- No Activity --"
          PRINT #PrnFile, ToPrint$

          ToPrint$ = STRING$(80, "*")
          PRINT #PrnFile, ToPrint$

        END IF

      END IF
    END IF      'Account is not of this fund

    NumTrans = 0             'reset for next account
    BalFwd# = 0
    FwdFlag = 0
    TotAcctDr# = 0
    TotAcctCr# = 0
  NEXT


  'ToPrint$ = SPACE$(80)
  'LSET ToPrint$ = "Grand Total Debits"
  'MID$(ToPrint$, 25) = FUsing$(STR$(GrTotDr#), "##########,.##")
  'PRINT #PrnFile, ToPrint$

  'ToPrint$ = SPACE$(80)
  'LSET ToPrint$ = "Grand Total Credits"
  'MID$(ToPrint$, 25) = FUsing$(STR$(GrTotCr#), "##########,.##")
  'PRINT #PrnFile, ToPrint$

  CLOSE

  SELECT CASE Dev$
    CASE "S"
      EntryPoint = 2
    CASE "P"
      EntryPoint = 5
  END SELECT
  PrintRptFile Header$, ReportFile$, LPTNo, RetCode%, EntryPoint

  KILL ReportFile$


END SUB

SUB PrintChkList

  LOCATE 1, 1, 0
  SHARED User$
  SHARED FundLen
  SHARED Choice$()

  REDIM Choice$(0 TO 3, 0 TO 1)
  Choice$(0, 0) = "5"
  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn#1"
  Choice$(3, 0) = "Prn#2"
  Choice$(0, 1) = "3"
  Choice$(1, 1) = "Check Number"
  Choice$(2, 1) = "Vendor"

  REDIM Form$(0, 0)
  REDIM Fld(0) AS FieldInfo
  DIM Frm AS FormInfo

  LibName$ = "AP.QSL"
  FormName$ = "PRNCHK"
  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()

  '--Assign Default Values
  Form$(3, 0) = Choice$(1, 1)
  Form$(4, 0) = "N"
  Form$(5, 0) = Choice$(1, 0)
  Action = 1
  CancelFlag = False

  DO
    EditForm Form$(), Fld(), Frm, Cnf, Action
    SELECT CASE Frm.KeyCode
      CASE F10Key
        EXIT DO
      CASE ESC
        CancelFlag = True
        EXIT DO
    END SELECT
  LOOP

  IF CancelFlag THEN EXIT SUB

  BegDate = Date2Num(Form$(1, 0))
  EndDate = Date2Num(Form$(2, 0))
  SortSpec$ = LEFT$(Form$(3, 0), 1)
  ShowFundDist = False
  IF Form$(4, 0) = "Y" THEN ShowFundDist = True 'unrem for new release
  Dev$ = LEFT$(Form$(5, 0), 1)
  IF Dev$ = "P" THEN LPTNo = VAL(RIGHT$(Form$(5, 0), 1))

  SELECT CASE SortSpec$
     CASE "V"
       ColTitle$ = " Vendor                           Chk Num         Date              Amt"
       Header$ = "A/P Checks by Vendor"
     CASE "C"
       ColTitle$ = " Chk Num  Vendor                                  Date              Amt"
       Header$ = "A/P Check Listing"
  END SELECT

  A$ = SPACE$(14)
  CommaFmt$ = "########,.##"
  FF$ = CHR$(12)
  MaxLines = 55
  User$ = QPTrim$(User$)
  Page = 0
  '--Get a list of active funds
  REDIM FundList$(1)
  GetFundList FundList$(), NumFunds
  REDIM FundGrdTot#(1 TO NumFunds)

  DIM APLedger AS APLedger81RecType
  APLRecLen = LEN(APLedger)
  OpenAPLedgerFile APLedgerFile, NumTrans, APLRecLen
   
  REDIM ChkList(1 TO 1) AS GLAcctIndexType  '--borrowing this type

  DIM APDist AS APDistRecType
  APDRecLen = LEN(APDist)
  OpenAPDistFile APDistFile, NumDistRecs&, APDRecLen

  RptFile = FREEFILE
  RptFileName$ = "apchks.prn"
  OPEN RptFileName$ FOR OUTPUT AS RptFile

  DIM Vendor AS VendorRecType
  VRecLen = LEN(Vendor)
  OpenVendorFile VendorFile, NumVRecs
  
  GOSUB OpenChkPageHdr

  OhShoot = False
  FOR Cnt = 1 TO NumTrans
    Pct$ = STR$(INT((Cnt / NumTrans) * 100))
    QPrintRC "Reading..." + Pct$ + "%", 25, 2, -1

    GET APLedgerFile, Cnt, APLedger
    IF APLedger.VRecNum > 0 THEN

    GET VendorFile, APLedger.VRecNum, Vendor
                                 
    IF APLedger.TrCode = 3 OR APLedger.TrCode = -3 THEN
      IF APLedger.TrDate >= BegDate AND APLedger.TrDate <= EndDate THEN

        NumChks = NumChks + 1
        REDIM PRESERVE ChkList(1 TO NumChks) AS GLAcctIndexType
        ChkList(NumChks).RecNum = Cnt
        SELECT CASE SortSpec$
          CASE "C"
            RSET A$ = QPTrim$(APLedger.DOCNum)
            ChkList(NumChks).AcctNum = A$
          CASE "V"
            ChkList(NumChks).AcctNum = Vendor.Vname
        END SELECT

      END IF
    END IF
    END IF
  NEXT

  IF NumChks > 0 THEN
    GOSUB SortChk
    GOSUB PrintChkList
  ELSE
    PRINT #RptFile, "  No Checks on file."
  END IF

  CLOSE

  SELECT CASE Dev$
   CASE "S"
     EntryPoint = 2
   CASE "P"
     EntryPoint = 5
  END SELECT

  PrintRptFile Header$, RptFileName$, LPTNo, RetCode%, EntryPoint

EXIT SUB


SortChk:
'  REDIM PRESERVE ChkList(1 TO NumChks) AS GLAcctIndexType
  SortT ChkList(1), NumChks, 0, 16, 0, 14
RETURN


OpenChkPageHdr:
  Page = Page + 1
  PRINT #RptFile, TAB(40 - (INT(LEN(User$) / 2))); User$
  PRINT #RptFile, TAB(40 - (INT(LEN(Header$) / 2))); Header$
  PRINT #RptFile,
  PRINT #RptFile, "Report Date: "; DATE$; TAB(67); "Page #"; Page
  PRINT #RptFile, ColTitle$
  PRINT #RptFile, STRING$(78, "=")
  LineCnt = 5
RETURN


PrintChkList:
  QPrintRC SPACE$(20), 25, 2, -1

  FOR Cnt = 1 TO NumChks
    Pct$ = STR$(INT((Cnt / NumChks) * 100))
    QPrintRC "Writing..." + Pct$ + "%", 25, 2, -1
    A$ = SPACE$(14)

    Rec = ChkList(Cnt).RecNum
    GET APLedgerFile, Rec, APLedger
    GET VendorFile, APLedger.VRecNum, Vendor

    RunTotal# = RunTotal# + APLedger.Amt

    ToPrint$ = SPACE$(80)
    SELECT CASE SortSpec$
      CASE "V"
        MID$(ToPrint$, 2) = Vendor.Vname
        MID$(ToPrint$, 35) = LEFT$(APLedger.DOCNum, 15)
      CASE "C"
        MID$(ToPrint$, 2) = LEFT$(APLedger.DOCNum, 15)
        MID$(ToPrint$, 10) = Vendor.Vname
    END SELECT
    MID$(ToPrint$, 51) = Num2Date$(APLedger.TrDate)
    MID$(ToPrint$, 62) = FUsing(STR$(APLedger.Amt), CommaFmt$)

    IF APLedger.TrCode = -3 THEN
      MID$(ToPrint$, 76) = "Void"
    END IF

    PRINT #RptFile, ToPrint$

    LineCnt = LineCnt + 1
    IF LineCnt > MaxLines THEN
       PRINT #RptFile, FF$
       GOSUB OpenChkPageHdr
    END IF

    IF ShowFundDist THEN
      NextDist& = APLedger.FrstDist
      DistAmt# = 0

      IF NextDist& > 0 THEN
         PRINT #RptFile,
         PRINT #RptFile, TAB(50); "Fund Distribution:"
         LineCnt = LineCnt + 2

         DO
            GET APDistFile, NextDist&, APDist
            DistAmt# = DistAmt# + Round#(APDist.DistAmt)

            ToPrint$ = SPACE$(80)
            MID$(ToPrint$, 50) = LEFT$(APDist.DistAcctNum, FundLen)
            MID$(ToPrint$, 62) = FUsing(STR$(APDist.DistAmt), CommaFmt$)
            PRINT #RptFile, ToPrint$
            LineCnt = LineCnt + 1
            IF LineCnt > MaxLines THEN
               PRINT #RptFile, FF$
               GOSUB OpenChkPageHdr
            END IF
            
            '--summarize by fund
            Found = False
            FOR Fund = 1 TO NumFunds
              FundNum$ = LEFT$(APDist.DistAcctNum, FundLen)
              IF FundNum$ = FundList$(Fund) THEN
                Found = True
                FundGrdTot#(Fund) = FundGrdTot#(Fund) + Round#(APDist.DistAmt)
                EXIT FOR
              END IF
            NEXT

            NextDist& = APDist.NextDist
         LOOP UNTIL NextDist& = 0
          
         PRINT #RptFile, STRING$(78, "-")
      END IF  '--showing Distribution
    END IF
  NEXT

  PRINT #RptFile,
  PRINT #RptFile, "Total Checks Listed: " + FUsing(STR$(RunTotal#), CommaFmt$)
  PRINT #RptFile,

  IF ShowFundDist THEN
    FOR FCnt = 1 TO NumFunds
      IF FundGrdTot#(FCnt) > 0 THEN
        PRINT #RptFile, "Fund: "; FundList$(FCnt); " " + FUsing(STR$(FundGrdTot#(FCnt)), CommaFmt$)
      END IF
    NEXT
  END IF

  PRINT #RptFile, FF$

RETURN


END SUB

SUB PrintOpenPO

  LOCATE 1, 1, 0
  SHARED Choice$()
  SHARED User$
  SHARED Vendor AS VendorRecType

  LibName$ = "AP.QSL"           'screen library
  FormName$ = "PRNPO"        'form to edit
  DevFld = 5    'Set to output device field number
  
  REDIM Choice$(0 TO 3, 0 TO 1)
  Choice$(0, 0) = "4"
  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn#1"
  Choice$(3, 0) = "Prn#2"
  Choice$(0, 1) = "1"
  Choice$(1, 1) = "P O Number"
  Choice$(2, 1) = "Vendor"
  
  REDIM Form$(0, 0)
  REDIM Fld(0) AS FieldInfo
  DIM Frm AS FormInfo
ResetForm:
  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()

  Form$(1, 0) = Choice$(1, 1)
  Form$(2, 0) = "N"
  Form$(4, 0) = Choice$(1, 0)
  Frm.FldNo = 3
  Action = 1
  CancelFlag = False

  DO
    EditForm Form$(), Fld(), Frm, Cnf, Action
    SELECT CASE Frm.KeyCode
      CASE F10Key
        Dept$ = Form$(3, 0)
        Dept$ = QPTrim$(Dept$)
        IF LEN(Dept$) = 0 THEN
         help$ = "You Must Enter a Valid Dept Number"
         PrintHelp help$
         SLEEP 2
         GOTO ResetForm
        END IF

        EXIT DO
      CASE ESC
        CancelFlag = True
        EXIT DO
    END SELECT
  LOOP

  IF CancelFlag THEN EXIT SUB

  SortSpec$ = LEFT$(Form$(1, 0), 1)
  ShowEnc = False
  IF Form$(2, 0) = "Y" THEN ShowEnc = True
  Dev$ = LEFT$(Form$(4, 0), 1)
  IF Dev$ = "P" THEN LPTNo = VAL(RIGHT$(Form$(4, 0), 1))

  SELECT CASE SortSpec$
     CASE "V"
       ColTitle$ = " Vendor                           PO Num          Date              Amt"
       Header$ = "Open Purchase Orders by Vendor"
     CASE "P"
       ColTitle$ = " PO Num           Vendor                          Date              Amt"
       Header$ = "Open Purchase Orders"
  END SELECT
  CommaFmt$ = "######,.##"
  MaxLines = 55
  User$ = QPTrim$(User$)
  Page = 0
  FF$ = CHR$(12)
  MaxPO = 400
  REDIM POList(1 TO MaxPO) AS GLAcctIndexType  '--borrowing this type
  
  '--Get a list of active funds
  REDIM FundList$(1)
  GetFundList FundList$(), NumFunds
  REDIM FundGrdTot#(1 TO NumFunds)
  
  DIM APLedger AS APLedger81RecType
  APLRecLen = LEN(APLedger)
  OpenAPLedgerFile APLedgerFile, NumTrans, APLRecLen

  DIM APDist AS APDistRecType
  APDRecLen = LEN(APDist)
  OpenAPDistFile APDistFile, NumDistRecs&, APDRecLen

  OpenVendorFile VendorFile, NumVRecs
  
  RptFile = FREEFILE
  RptFileName$ = "openpo.prn"
  OPEN RptFileName$ FOR OUTPUT AS RptFile

  GOSUB OpenPOPageHdr

  OhShoot = False
  FOR Cnt = 1 TO NumTrans
    Pct$ = STR$(INT((Cnt / NumTrans) * 100))
    QPrintRC "Reading..." + Pct$ + "%", 25, 2, -1

    GET APLedgerFile, Cnt, APLedger
 
    IF APLedger.TrCode = 4 THEN
     IF LEFT$(Dept$, 3) = "ALL" OR APLedger.DeptNumb = VAL(Dept$) THEN
      NumPO = NumPO + 1
      IF NumPO = MaxPO THEN
        OhShoot = True
        EXIT FOR
      ELSE
        POList(NumPO).RecNum = Cnt
        SELECT CASE SortSpec$
          CASE "P"
            POList(NumPO).AcctNum = LEFT$(APLedger.PONUM, 14)
          CASE "V"
            POList(NumPO).AcctNum = APLedger.VendorCode
        END SELECT
      END IF
     END IF
    END IF
  NEXT
   
  IF OhShoot = True THEN
    CLOSE
    CLS
    QPrintRC "Error: Available elements exceed needs. Unable to run report.", 10, 1, 12
    QPrintRC "Press any key to continue.", 11, 1, 12
    K$ = INPUT$(1)
    EXIT SUB
  END IF
   
  IF NumPO > 0 THEN
    GOSUB SortPO
    IF ShowEnc THEN
     GOSUB ClearEnc
    END IF
    GOSUB PrintPOList
  ELSE
    PRINT #RptFile, "No Open Purchase Orders"
  END IF

  CLOSE
  
  SELECT CASE Dev$
   CASE "S"
     EntryPoint = 2
   CASE "P"
     EntryPoint = 5
  END SELECT

  PrintRptFile Header$, RptFileName$, LPTNo, RetCode%, EntryPoint

EXIT SUB


SortPO:
  REDIM PRESERVE POList(1 TO NumPO) AS GLAcctIndexType
  SortT POList(1), NumPO, 0, 16, 0, 14 'sort the $#&+ out of it!
RETURN


OpenPOPageHdr:
  Page = Page + 1
  PRINT #RptFile, TAB(40 - (INT(LEN(User$) / 2))); User$
  PRINT #RptFile, TAB(40 - (INT(LEN(Header$) / 2))); Header$
  PRINT #RptFile,
  PRINT #RptFile, "Dept Number: "; Dept$
  PRINT #RptFile, "Report Date: "; DATE$; TAB(67); "Page #"; Page
  PRINT #RptFile, ColTitle$
  PRINT #RptFile, STRING$(78, "=")
  LineCnt = 4
RETURN


PrintPOList:
  FOR Cnt = 1 TO NumPO
    Pct$ = STR$(INT((Cnt / NumPO) * 100))
    QPrintRC "Writing..." + Pct$ + "%", 25, 2, -1
    
    Rec = POList(Cnt).RecNum
    GET APLedgerFile, Rec, APLedger
    GET VendorFile, APLedger.VRecNum, Vendor
    RunTotal# = RunTotal# + APLedger.Amt

    ToPrint$ = SPACE$(80)
    SELECT CASE SortSpec$
      CASE "V"
        MID$(ToPrint$, 2) = Vendor.Vname 'APLedger.VendorCode
        MID$(ToPrint$, 35) = LEFT$(APLedger.PONUM, 15)
      CASE "P"
        MID$(ToPrint$, 2) = LEFT$(APLedger.PONUM, 15)
        MID$(ToPrint$, 19) = Vendor.Vname 'APLedger.VendorCode
    END SELECT
    MID$(ToPrint$, 51) = Num2Date$(APLedger.TrDate)
    MID$(ToPrint$, 64) = FUsing(STR$(APLedger.Amt), CommaFmt$)
    PRINT #RptFile, ToPrint$
    LineCnt = LineCnt + 1
    IF LineCnt > MaxLines THEN
       PRINT #RptFile, FF$
       GOSUB OpenPOPageHdr
    END IF

    IF ShowEnc THEN
      '--Now print the distribution
      NextDist& = APLedger.FrstDist
      DistAmt# = 0

      IF NextDist& > 0 THEN
         PRINT #RptFile,
         PRINT #RptFile, TAB(50); "Encumbered Accounts:"
         LineCnt = LineCnt + 2

         DO
            GET APDistFile, NextDist&, APDist

            GOSUB UpdateGLAcct          'Reseting the Encumbered Amt

            DistAmt# = DistAmt# + APDist.DistAmt
            NextDist& = APDist.NextDist

            ToPrint$ = SPACE$(80)
            MID$(ToPrint$, 50) = APDist.DistAcctNum
            MID$(ToPrint$, 64) = FUsing(STR$(APDist.DistAmt), CommaFmt$)
            PRINT #RptFile, ToPrint$

            LineCnt = LineCnt + 1
            IF LineCnt > MaxLines THEN
               PRINT #RptFile, FF$
               GOSUB OpenPOPageHdr
            END IF

            '--summarize by fund
            Found = False
            FOR Fund = 1 TO NumFunds
              FundNum$ = LEFT$(APDist.DistAcctNum, FundLen)
              IF FundNum$ = FundList$(Fund) THEN
                Found = True
                FundGrdTot#(Fund) = FundGrdTot#(Fund) + Round#(APDist.DistAmt)
                EXIT FOR
              END IF
            NEXT
            
         LOOP UNTIL NextDist& = 0

         PRINT #RptFile, TAB(64); "----------"
         PRINT #RptFile, TAB(50); "Total:"; TAB(64); FUsing(STR$(DistAmt#), CommaFmt$)
         LineCnt = LineCnt + 2
         PRINT #RptFile, STRING$(78, "-")
      '--showing encumbrances
      END IF
    END IF
  NEXT

  PRINT #RptFile,
  PRINT #RptFile, "Total Open PO's: " + FUsing(STR$(RunTotal#), CommaFmt$)
  IF ShowEnc THEN
    PRINT #RptFile, "By Fund:"
    FOR FCnt = 1 TO NumFunds
      IF FundGrdTot#(FCnt) > 0 THEN
        PRINT #RptFile, "Fund: "; FundList$(FCnt); " " + FUsing(STR$(FundGrdTot#(FCnt)), CommaFmt$)
      END IF
    NEXT
  END IF
  
  PRINT #RptFile, FF$
  
RETURN
UpdateGLAcct:          'Reseting the Encumbered Amt
   Amt# = APDist.DistAmt
   DistAcctRec = FindAcct(APDist.DistAcctNum)
   IF DistAcctRec > 0 THEN
    OpenAcctFile AcctFileNum, NumGLAcctRecs
    GET AcctFileNum, DistAcctRec, Acct
    Acct.Encumb = Acct.Encumb + Amt#
    PUT AcctFileNum, DistAcctRec, Acct
    CLOSE AcctFileNum
   END IF
   RETURN


ClearEnc:
  OpenAcctFile AcctFileNum, NumGLAcctRecs
  FOR Cnt1! = 1 TO NumGLAcctRecs
  GET AcctFileNum, Cnt1!, Acct
  Acct.Encumb = 0
  PUT AcctFileNum, Cnt1!, Acct
  NEXT Cnt1!
  CLOSE AcctFileNum
  RETURN

END SUB

SUB SalesTaxReport

  LOCATE 1, 1, 0

  LibName$ = "AP.QSL"           'screen library
  FormName$ = "SALESTX"        'form to edit
  DevFld = 7    'Set to output device field number
  REDIM Form$(0, 0)
  REDIM Fld(0) AS FieldInfo
  DIM Frm AS FormInfo
  SHARED Choice$()
  REDIM Choice$(0 TO 3, 1)
  Choice$(0, 0) = STR$(DevFld)  'Choices for Output Device
  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn#1"
  Choice$(3, 0) = "Prn#2"

  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
  Action = 1
  DO
    EditForm Form$(), Fld(), Frm, Cnf, Action
    IF Frm.PrevFld <> Frm.FldNo THEN
      'IF Form$(1, 0) = "Y" THEN
      '  Fld(3).Protected = True
      'ELSE
      '  Fld(4).Protected = True
      'END IF
      IF Frm.FldNo = 2 OR Frm.FldNo = 3 THEN
        Form$(Frm.FldNo, 0) = ""
        Action = 1
      END IF
      '--strip out dashes for editing
      IF Frm.FldNo = 2 OR Frm.FldNo = 3 THEN
        IF LEN(Form$(Frm.FldNo, 0)) > 0 THEN
          Form$(Frm.FldNo, 0) = StripAcct$(Form$(Frm.FldNo, 0))
          SaveField Frm.FldNo, Form$(), Fld(), 0
          Action = 1
        END IF
      END IF
      SELECT CASE Frm.PrevFld
      CASE 2, 3
        Form$(Frm.PrevFld, 0) = FmtAcct$(Form$(Frm.PrevFld, 0), FundLen%, AcctLen%, DetLen%)
      CASE ELSE
      END SELECT
      Action = 1
    END IF
    SELECT CASE Frm.KeyCode
    CASE F10Key 'SaveButton
      IF Form$(1, 0) = "Y" THEN Combined = True
      StateTaxRecAcct$ = QPTrim$(Form$(2, 0))
      CountyTaxRecAcct$ = QPTrim$(Form$(3, 0))
      StateFactor# = VAL(Form$(4, 0))
      BegDate = Date2Num(Form$(5, 0))
      EndDate = Date2Num(Form$(6, 0))
    
      IF LEN(LTRIM$(RTRIM$(Form$(DevFld, 0)))) = 0 THEN
        Dev$ = "S"
        LPTNo = 1
      ELSE
        Dev$ = LEFT$(Form$(DevFld, 0), 1)
        LPTNo = VAL(RIGHT$(RTRIM$(Form$(DevFld, 0)), 1))
      END IF
      IF BegDate > EndDate THEN 'OR Form$(1, 0) > Form$(2, 0) THEN
        OK = MsgBox("GL", "BADPARMS")
        Ok2Go = 0
        Frm.FldNo = 5
        Action = 1
      ELSE
        Ok2Go = -1
      END IF
    CASE ESC
      EXIT SUB
    END SELECT
  LOOP UNTIL Ok2Go

  
  

  

  DIM StSalesTaxPaid#(0 TO 999)
  IF NOT Combined THEN
    DIM CoSalesTaxPaid#(0 TO 999)
  END IF

  FOR state = 0 TO 99
  PrintFlag = 0: StTotTax# = 0: TotState# = 0: TotCounty# = 0: TCnt = 0
  CoTotTax# = 0
  FOR ll = 0 TO 999
   StSalesTaxPaid#(ll) = 0: CoSalesTaxPaid#(ll) = 0
  NEXT ll

  DIM Vendor AS VendorRecType
  CLOSE VendorFile
  OpenVendorFile VendorFile, NumVRecs

  DIM APLedger AS APLedger81RecType
  LdRecLen = LEN(APLedger)
   CLOSE APLedgerFile
  OpenAPLedgerFile APLedgerFile, NumTrans, LdRecLen
  
  DIM APDist AS APDistRecType
  APDRecLen = LEN(APDist)
  CLOSE APDistFile
  OpenAPDistFile APDistFile, NumDistRecs&, APDRecLen

  

  FOR Cnt = 1 TO NumVRecs
    GET VendorFile, Cnt, Vendor
    NextTran = Vendor.FrstTran
    IF NextTran > 0 THEN
      DO
        GET APLedgerFile, NextTran, APLedger
        IF APLedger.TrCode = 1 THEN
         '--check this out
         IF APLedger.GLDistDate >= BegDate AND APLedger.GLDistDate <= EndDate THEN
          TCnt = TCnt + 1
          NextDist& = APLedger.FrstDist
          DO
            GET APDistFile, NextDist&, APDist
            
            IF QPTrim$(APDist.DistAcctNum) = StateTaxRecAcct$ THEN
               GET VendorFile, APLedger.VRecNum, Vendor
               offset = VAL(Vendor.StCode)
               COffSet = VAL(Vendor.CoCode)
               
               IF state = offset THEN
                SCnt = SCnt + 1
                StTax# = StTax# + APDist.DistAmt
                PrintFlag = 1
                offset = VAL(Vendor.StCode)
                IF offset < 0 OR offset > 999 THEN offset = 0
                StSalesTaxPaid#(COffSet) = StSalesTaxPaid#(COffSet) + APDist.DistAmt
                'LPRINT APDist.DistAmt
               END IF
             END IF
             IF NOT Combined THEN
               IF QPTrim$(APDist.DistAcctNum) = CountyTaxRecAcct$ THEN
                 GET VendorFile, APLedger.VRecNum, Vendor
                 IF state = VAL(Vendor.StCode) THEN
                 CCnt = CCnt + 1
                 CoTax# = CoTax# + APDist.DistAmt
                 PrintFlag = 1
                 offset = VAL(Vendor.CoCode)
                 IF offset < 0 OR offset > 999 THEN offset = 0
               '  LPRINT Vendor.Vname, APDist.DistAmt
                 CoSalesTaxPaid#(offset) = CoSalesTaxPaid#(offset) + APDist.DistAmt
                 END IF
               END IF
             END IF
             NextDist& = APDist.NextDist
          LOOP UNTIL NextDist& = 0

         END IF

        END IF
        NextTran = APLedger.NextTrans
      LOOP UNTIL NextTran = 0
    END IF

  NEXT Cnt
  
  IF PrintFlag = 1 THEN
   GOSUB PrintState
  END IF
  CLS
  COLOR 11
  LOCATE 3, 1: PRINT "Sales Tax Report : Now Processing"
  LOCATE 5, 1: PRINT "State Code Being Searched: "; : COLOR 15: PRINT state
  NEXT state
  CLOSE
  EXIT SUB

  
PrintState:
  PrnFile = FREEFILE
  ReportFile$ = "SALESTAX.PRN"
  OPEN ReportFile$ FOR OUTPUT AS #PrnFile

  PRINT #PrnFile, "Sales Tax Report"
  PRINT #PrnFile, "State Code: "; state
  PRINT #PrnFile,
  PRINT #PrnFile, "County Code          State          County"
  PRINT #PrnFile, "--------------------------------------------"

  IF Combined THEN
    FOR Cnt = 0 TO 999
      IF StSalesTaxPaid#(Cnt) > 0 THEN
        StateAmt# = (StSalesTaxPaid#(Cnt) * StateFactor#)
        CountyAmt# = (StSalesTaxPaid#(Cnt) - StateAmt#)
        StTotTax# = StTotTax# + StSalesTaxPaid#(Cnt)
        TotState# = TotState# + StateAmt#
        TotCounty# = TotCounty# + CountyAmt#
        PRINT #PrnFile, Cnt; USING "######.##"; TAB(20); StateAmt#; TAB(35); CountyAmt#
      END IF
    NEXT
    PRINT #PrnFile, "--------------------------------------------"
    PRINT #PrnFile, "Totals"; USING "######.##"; TAB(20); TotState#; TAB(35); TotCounty#
  ELSE
    FOR Cnt = 0 TO 999
      IF StSalesTaxPaid#(Cnt) > 0 OR CoSalesTaxPaid#(Cnt) > 0 THEN
        PRINT #PrnFile, Cnt; USING "#####.##"; TAB(20); StSalesTaxPaid#(Cnt); TAB(35); CoSalesTaxPaid#(Cnt)
        StTotTax# = StTotTax# + StSalesTaxPaid#(Cnt)
        CoTotTax# = CoTotTax# + CoSalesTaxPaid#(Cnt)
      END IF
    NEXT
    PRINT #PrnFile, "--------------------------------------------"
    PRINT #PrnFile, "Totals"; USING "######.##"; TAB(19); StTotTax#; TAB(34); CoTotTax#
  END IF

  PRINT #PrnFile, CHR$(12)

  CLOSE

  SELECT CASE Dev$
   CASE "S"
     EntryPoint = 2
   CASE "P"
     EntryPoint = 5
  END SELECT

  Header$ = "Sales Tax Report"
  PrintRptFile Header$, ReportFile$, LPTNo, RetCode%, EntryPoint
  RETURN


END SUB

SUB SalesTaxRichland

  LOCATE 1, 1, 0

  SHARED User$
  LibName$ = "AP.QSL"           'screen library
  FormName$ = "SALESTX"        'form to edit
  DevFld = 7    'Set to output device field number
  REDIM Form$(0, 0)
  REDIM Fld(0) AS FieldInfo
  DIM Frm AS FormInfo
  SHARED Choice$()
  REDIM Choice$(0 TO 3, 1)
  Choice$(0, 0) = STR$(DevFld)  'Choices for Output Device
  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn#1"
  Choice$(3, 0) = "Prn#2"

  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
  Action = 1
  DO
    EditForm Form$(), Fld(), Frm, Cnf, Action
    IF Frm.PrevFld <> Frm.FldNo THEN
      IF Frm.FldNo = 2 OR Frm.FldNo = 3 THEN
        Form$(Frm.FldNo, 0) = ""
        Action = 1
      END IF
      '--strip out dashes for editing
      IF Frm.FldNo = 2 OR Frm.FldNo = 3 THEN
        IF LEN(Form$(Frm.FldNo, 0)) > 0 THEN
          Form$(Frm.FldNo, 0) = StripAcct$(Form$(Frm.FldNo, 0))
          SaveField Frm.FldNo, Form$(), Fld(), 0
          Action = 1
        END IF
      END IF
      SELECT CASE Frm.PrevFld
      CASE 2, 3
        Form$(Frm.PrevFld, 0) = FmtAcct$(Form$(Frm.PrevFld, 0), FundLen%, AcctLen%, DetLen%)
      CASE ELSE
      END SELECT
      Action = 1
    END IF
    SELECT CASE Frm.KeyCode
    CASE F10Key 'SaveButton
      IF Form$(1, 0) = "Y" THEN Combined = True
      StateTaxExpAcct$ = "3700"      'QPTrim$(Form$(2, 0))
      CountyTaxExpAcct$ = "4000"     'QPTrim$(Form$(3, 0))
      StateFactor# = VAL(Form$(4, 0))
      BegDate = Date2Num(Form$(5, 0))
      EndDate = Date2Num(Form$(6, 0))
      IF LEN(LTRIM$(RTRIM$(Form$(DevFld, 0)))) = 0 THEN
        Dev$ = "S"
        LPTNo = 1
      ELSE
        Dev$ = LEFT$(Form$(DevFld, 0), 1)
        LPTNo = VAL(RIGHT$(RTRIM$(Form$(DevFld, 0)), 1))
      END IF
      IF BegDate > EndDate THEN 'OR Form$(1, 0) > Form$(2, 0) THEN
        OK = MsgBox("GL", "BADPARMS")
        Ok2Go = 0
        Frm.FldNo = 5
        Action = 1
      ELSE
        Ok2Go = -1
      END IF
    CASE ESC
      EXIT SUB
    END SELECT
  LOOP UNTIL Ok2Go

  SHARED Vendor AS VendorRecType
  OpenVendorFile VendorFile, NumVRecs

  'DIM APLedger AS APLedgerRecType
  DIM APLedger AS APLedger81RecType
  LdRecLen = LEN(APLedger)
  OpenAPLedgerFile APLedgerFile, NumTrans, LdRecLen

  DIM APDist AS APDistRecType
  APDRecLen = LEN(APDist)
  OpenAPDistFile APDistFile, NumDistRecs&, APDRecLen

  DIM StSalesTaxPaid#(0 TO 100)
  DIM CoSalesTaxPaid#(0 TO 100)

  Fmt$ = "#####,.##"

  FOR Cnt = 1 TO NumVRecs
    GET VendorFile, Cnt, Vendor
    NextTran = Vendor.FrstTran
    IF NextTran > 0 THEN
      DO
        GET APLedgerFile, NextTran, APLedger
        IF APLedger.TrCode = 1 THEN
          TCnt = TCnt + 1
          NextDist& = APLedger.FrstDist
          DO
             GET APDistFile, NextDist&, APDist
             ObjCode$ = MID$(APDist.DistAcctNum, 8, 4)
             IF ObjCode$ = StateTaxExpAcct$ THEN
               GET VendorFile, APLedger.VRecNum, Vendor
               'PRINT #PrnFile, Vendor.VName, Vendor.StCode, Vendor.CoCode
               SCnt = SCnt + 1
               StTax# = StTax# + APDist.DistAmt
               offset = VAL(Vendor.CoCode)
               IF offset < 0 OR offset > 100 THEN offset = 0
               StSalesTaxPaid#(offset) = StSalesTaxPaid#(offset) + APDist.DistAmt
             END IF
             
             IF ObjCode$ = CountyTaxExpAcct$ THEN
               GET VendorFile, APLedger.VRecNum, Vendor
               'PRINT #PrnFile, Vendor.VName, Vendor.StCode, Vendor.CoCode
               CCnt = CCnt + 1
               CoTax# = CoTax# + APDist.DistAmt
               offset = VAL(Vendor.CoCode)
               IF offset < 0 OR offset > 100 THEN offset = 0
               CoSalesTaxPaid#(offset) = CoSalesTaxPaid#(offset) + APDist.DistAmt
             END IF
             
             NextDist& = APDist.NextDist
          LOOP UNTIL NextDist& = 0
        END IF
        NextTran = APLedger.NextTrans
      LOOP UNTIL NextTran = 0
    END IF
  NEXT

  CLOSE

  PrnFile = FREEFILE
  ReportFile$ = "SALESTAX.PRN"
  OPEN ReportFile$ FOR OUTPUT AS #PrnFile

  GOSUB PrintReptHdr
  
  FOR Cnt = 0 TO 100
    IF StSalesTaxPaid#(Cnt) > 0 OR CoSalesTaxPaid#(Cnt) > 0 THEN
      PRINT #PrnFile, Cnt; TAB(20); USING Fmt$; StSalesTaxPaid#(Cnt);
      PRINT #PrnFile, TAB(40); USING Fmt$; CoSalesTaxPaid#(Cnt)
      StTotTax# = StTotTax# + StSalesTaxPaid#(Cnt)
      CoTotTax# = CoTotTax# + CoSalesTaxPaid#(Cnt)
    END IF
  NEXT
  PRINT #PrnFile, "------------------------------------------------"
  PRINT #PrnFile, "Totals";
  PRINT #PrnFile, TAB(20); USING Fmt$; StTotTax#;
  PRINT #PrnFile, TAB(40); USING Fmt$; CoTotTax#
  
  CLOSE

  SELECT CASE Dev$
   CASE "S"
     EntryPoint = 2
   CASE "P"
     EntryPoint = 5
  END SELECT

  Header$ = "Sales Tax Report"
  PrintRptFile Header$, ReportFile$, LPTNo, RetCode%, EntryPoint

EXIT SUB

PrintReptHdr:
PRINT #PrnFile, User$
PRINT #PrnFile, "Sales Tax Report"
PRINT #PrnFile, Form$(5, 0) + " through " + Form$(6, 0)
PRINT #PrnFile,
PRINT #PrnFile, "County Code    State Tax Pd        County Tax Pd"
PRINT #PrnFile, "================================================"
RETURN


END SUB

SUB SalesTaxRptE2

  LOCATE 1, 1, 0

  SHARED User$
  LibName$ = "AP.QSL"           'screen library
  FormName$ = "STAXE2"        'form to edit
  
  REDIM Form$(0, 0)
  REDIM Fld(0) AS FieldInfo
  DIM Frm AS FormInfo
  SHARED Choice$()
  REDIM Choice$(0 TO 3, 1)
  Choice$(0, 0) = "5"
  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn#1"
  Choice$(3, 0) = "Prn#2"

  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
  Action = 1
  DO
    EditForm Form$(), Fld(), Frm, Cnf, Action

    SELECT CASE Frm.KeyCode
      CASE F10Key 'SaveButton
        StateTaxExpAcct$ = QPTrim$(Form$(1, 0))
        Factor# = VAL(Form$(2, 0))
        BegDate = Date2Num(Form$(3, 0))
        EndDate = Date2Num(Form$(4, 0))
        IF LEN(LTRIM$(RTRIM$(Form$(DevFld, 0)))) = 0 THEN
          Dev$ = "S"
          LPTNo = 1
        ELSE
          Dev$ = LEFT$(Form$(DevFld, 0), 1)
          LPTNo = VAL(RIGHT$(RTRIM$(Form$(DevFld, 0)), 1))
        END IF
        IF BegDate > EndDate THEN 'OR Form$(1, 0) > Form$(2, 0) THEN
          OK = MsgBox("GL", "BADPARMS")
          Ok2Go = 0
          Frm.FldNo = 5
          Action = 1
        ELSE
          Ok2Go = -1
        END IF
      CASE ESC
        EXIT SUB
    END SELECT

  LOOP UNTIL Ok2Go

  SHARED Vendor AS VendorRecType
  OpenVendorFile VendorFile, NumVRecs

  'DIM APLedger AS APLedgerRecType
  DIM APLedger AS APLedger81RecType
  LdRecLen = LEN(APLedger)
  OpenAPLedgerFile APLedgerFile, NumTrans, LdRecLen

  DIM APDist AS APDistRecType
  APDRecLen = LEN(APDist)
  OpenAPDistFile APDistFile, NumDistRecs&, APDRecLen

  DIM StSalesTaxPaid#(0 TO 100)
  
  Fmt$ = "#####,.##"

  FOR Cnt = 1 TO NumVRecs
    GET VendorFile, Cnt, Vendor
    NextTran = Vendor.FrstTran
    IF NextTran > 0 THEN
      DO
        GET APLedgerFile, NextTran, APLedger
        IF APLedger.TrCode = 1 THEN
          TCnt = TCnt + 1
          NextDist& = APLedger.FrstDist
          DO
             GET APDistFile, NextDist&, APDist
             ObjCode$ = MID$(APDist.DistAcctNum, FundLen + AcctLen + 3, DetLen)
             IF ObjCode$ = StateTaxExpAcct$ THEN
               GET VendorFile, APLedger.VRecNum, Vendor
               'PRINT #PrnFile, Vendor.VName, Vendor.StCode, Vendor.CoCode
               SCnt = SCnt + 1
               StTax# = StTax# + APDist.DistAmt
               offset = VAL(Vendor.CoCode)
               IF offset < 0 OR offset > 100 THEN offset = 0
               StSalesTaxPaid#(offset) = StSalesTaxPaid#(offset) + APDist.DistAmt
             END IF

             'IF ObjCode$ = CountyTaxExpAcct$ THEN
             '  GET VendorFile, APLedger.VRecNum, Vendor
             '  'PRINT #PrnFile, Vendor.VName, Vendor.StCode, Vendor.CoCode
             '  CCnt = CCnt + 1
             '  CoTax# = CoTax# + APDist.DistAmt
             '  Offset = VAL(Vendor.CoCode)
             '  IF Offset < 0 OR Offset > 100 THEN Offset = 0
             '  CoSalesTaxPaid#(Offset) = CoSalesTaxPaid#(Offset) + APDist.DistAmt
             'END IF

             NextDist& = APDist.NextDist
          LOOP UNTIL NextDist& = 0
        END IF
        NextTran = APLedger.NextTrans
      LOOP UNTIL NextTran = 0
    END IF
  NEXT

  CLOSE

  PrnFile = FREEFILE
  ReportFile$ = "SALESTAX.PRN"
  OPEN ReportFile$ FOR OUTPUT AS #PrnFile

  GOSUB PrintE2ReptHdr
  
  FOR Cnt = 0 TO 100
    'IF StSalesTaxPaid#(Cnt) > 0 OR CoSalesTaxPaid#(Cnt) > 0 THEN
    '  PRINT #PrnFile, Cnt; TAB(20); USING Fmt$; StSalesTaxPaid#(Cnt);
    '  PRINT #PrnFile, TAB(40); USING Fmt$; CoSalesTaxPaid#(Cnt)
    '  StTotTax# = StTotTax# + StSalesTaxPaid#(Cnt)
    '  CoTotTax# = CoTotTax# + CoSalesTaxPaid#(Cnt)
    'END IF

    IF StSalesTaxPaid#(Cnt) > 0 THEN
      StAmt# = StSalesTaxPaid#(Cnt) * Factor#
      CoAmt# = StSalesTaxPaid#(Cnt) - StAmt#
      PRINT #PrnFile, Cnt; TAB(20); USING Fmt$; StAmt#;
      PRINT #PrnFile, TAB(40); USING Fmt$; CoAmt#
      TotStTax# = TotStAmt# + StAmt#
      TotCoTax# = TotCoAmt# + CoAmt#
    END IF

  NEXT
  PRINT #PrnFile, "------------------------------------------------"
  PRINT #PrnFile, "Totals";
  PRINT #PrnFile, TAB(20); USING Fmt$; TotStTax#;
  PRINT #PrnFile, TAB(40); USING Fmt$; TotCoTax#

  CLOSE

  SELECT CASE Dev$
   CASE "S"
     EntryPoint = 2
   CASE "P"
     EntryPoint = 5
  END SELECT

  Header$ = "Sales Tax Report"
  PrintRptFile Header$, ReportFile$, LPTNo, RetCode%, EntryPoint

EXIT SUB

PrintE2ReptHdr:
PRINT #PrnFile, User$
PRINT #PrnFile, "Sales Tax Report"
PRINT #PrnFile, Form$(3, 0) + " through " + Form$(4, 0)
PRINT #PrnFile,
PRINT #PrnFile, "County Code    State Tax Pd        County Tax Pd"
PRINT #PrnFile, "================================================"
RETURN

END SUB

SUB test
                
  CLS
  
  DIM APDist AS APDistRecType
  APDRecLen = LEN(APDist)
  OpenAPDistFile APDistFile, NumDistRecs&, APDRecLen
  
  'DIM APLedger AS APLedgerRecType
  DIM APLedger AS APLedger81RecType
  LdRecLen = LEN(APLedger)
  OpenAPLedgerFile APLedgerFile, NumTrans, LdRecLen
  
  SHARED Vendor AS VendorRecType
  OpenVendorFile VendorFile, NumVRecs
  
  PrnFile = FREEFILE
  ReportFile$ = "SALESTAX.PRN"
  OPEN ReportFile$ FOR OUTPUT AS #PrnFile
  
  FOR Cnt = 1 TO NumVRecs
    GET VendorFile, Cnt, Vendor
    NextTran = Vendor.FrstTran
    IF NextTran > 0 THEN
      DO
        GET APLedgerFile, NextTran, APLedger
        IF APLedger.TrCode = 1 THEN
          TCnt = TCnt + 1
          NextDist& = APLedger.FrstDist
          DO
             GET APDistFile, NextDist&, APDist
             IF QPTrim$(APDist.DistAcctNum) = "10-130-00" THEN
               GET VendorFile, APLedger.VRecNum, Vendor
               PRINT #PrnFile, Vendor.Vname, Vendor.StCode, Vendor.CoCode
               SCnt = SCnt + 1
               Tax# = Tax# + APDist.DistAmt
             END IF
             NextDist& = APDist.NextDist
          LOOP UNTIL NextDist& = 0
        END IF
        NextTran = APLedger.NextTrans
      LOOP UNTIL NextTran = 0
    END IF

  NEXT

  CLOSE
  PRINT TCnt, SCnt, Tax#

  K$ = INPUT$(1)

  EXIT SUB

'  FOR Cnt = 1 TO NumDistRecs&
'    GET APDistFile, Cnt, APDist
'    IF QPTrim$(APDist.DistAcctNum) = "10-130-00" THEN
'      Hits = Hits + 1
'      Total# = Total# + APDist.DistAmt
'      GET APLedgerFile, APDist.APLedgerRec, APLedger
'      GET VendorFile, APLedger.VRecNum, Vendor
'      PRINT #PrnFile, Vendor.VName; "TrCode: "; APLedger.TrCode
'      'PRINT "Press any key to continue.."
'      'K$ = INPUT$(1)
'    END IF
'  NEXT

'  PRINT #PrnFile, Hits; Total#
'  CLOSE

  'SELECT CASE Dev$
  ' CASE "S"
     EntryPoint = 2
  ' CASE "P"
  '   EntryPoint = 5
  'END SELECT
  
  Header$ = "Sales Tax Report"
  PrintRptFile Header$, ReportFile$, LPTNo, RetCode%, EntryPoint
 
  'PRINT Hits; Total#
  'PRINT "Press any key to continue.."
  'K$ = INPUT$(1)
  
  'LOCATE 1, 1, 0
  '
  'BegDate = Date2Num("07-01-1995")
  'EndDate = Date2Num("06-30-1996")
  'StateTaxRecAcct$ = "10-130-00"
  'CountyTaxRecAcct$ = "10-133-00"
  '
  'PrnFile = FREEFILE
  'ReportFile$ = "SALESTAX.PRN"
  'OPEN ReportFile$ FOR OUTPUT AS #PrnFile
  '
  'DIM APLedger AS APLedgerRecType
  ''DIM APLedger AS APLedger81RecType
  'LdRecLen = LEN(APLedger)
  'OpenAPLedgerFile APLedgerFile, NumTrans, LdRecLen
  '
  'DIM APDist AS APDistRecType
  'APDRecLen = LEN(APDist)
  'OpenAPDistFile APDistFile, NumDistRecs&, APDRecLen
  '
  'SHARED Vendor AS VendorRecType
  'OpenVendorFile VendorFile, NumVRecs
  '
  'DIM SalesTaxPaid#(0 TO 100)
  '
  'FOR Cnt = 1 TO NumTrans
  '  GET APLedgerFile, Cnt, APLedger
  '  IF APLedger.TrCode = 1 THEN
  '    NextDist& = APLedger.FrstDist
  '    DO UNTIL NextDist& = 0
  '      GET APDistFile, NextDist&, APDist
  '      DistRecsRead = DistRecsRead + 1
  '      IF QPTrim$(APDist.DistAcctNum) = StateTaxRecAcct$ THEN
  '        Hits = Hits + 1
  '        GET VendorFile, APLedger.VRecNum, Vendor
  '        x = VAL(Vendor.CoCode)
  '        IF x < 0 OR x > 100 THEN x = 0
  '        SalesTaxPaid#(x) = APDist.DistAmt
  '      END IF
  '      NextDist& = APDist.NextDist
  '    LOOP
  '    NextTrans& = APLedger.NextTrans
  '  END IF
  'NEXT
  '
  ''PRINT Hits; Total#
  'PRINT "Press any key to continue.."
  '''K$ = INPUT$(1)

  'PRINT #PrnFile, "Read: "; DistRecsRead; "Hits: "; Hits
  '
  'FOR Cnt = 0 TO 100
  '  IF SalesTaxPaid#(Cnt) > 0 THEN
  '    PRINT #PrnFile, Cnt; SalesTaxPaid#(Cnt)
  '    TotalTax# = TotalTax# + SalesTaxPaid#(Cnt)
  '  END IF
  'NEXT
  '
  'PRINT #PrnFile, TotalTax#
  '
  'CLOSE
  '
  ''SELECT CASE Dev$
  '' CASE "S"
  '   EntryPoint = 2
  '' CASE "P"
  ''   EntryPoint = 5
  ''END SELECT
  '
  'Header$ = "Sales Tax Report"
  'PrintRptFile Header$, ReportFile$, LPTNo, RetCode%, EntryPoint
  
  
END SUB

SUB test2

  'DIM APLedger AS APLedgerRecType
  'DIM APLedger AS APLedger81RecType
  'LdRecLen = LEN(APLedger)
  'OpenAPLedgerFile APledgerFile, NumTrans, LdRecLen
  '
  'FOR Cnt = 1 TO NumTrans
  '  GET APledgerFile, Cnt, APLedger
  '  IF APLedger.TrCode = 1 THEN
  '    InvCnt = InvCnt + 1
  '  END IF
  'NEXT
  'CLOSE
  '
  'PRINT "Inv Cnt:"; InvCnt
  '
  'PRINT "Press any key to continue"
  'K$ = INPUT$(1)

  DIM APDist AS APDistRecType
  APDRecLen = LEN(APDist)
  OpenAPDistFile APDistFile, NumDistRecs&, APDRecLen

  FOR R& = 1 TO NumDistRecs&
    GET APDistFile, R&, APDist
    IF QPTrim$(APDist.DistAcctNum) = "10-133-0000" THEN
      DAmt# = DAmt# + APDist.DistAmt
      TCnt = TCnt + 1
    END IF

  NEXT
  PRINT TCnt

END SUB

SUB VendorHistory

  SHARED User$
  SHARED Choice$()
  Dash$ = STRING$(78, "-")

  REDIM Choice$(0 TO 3, 1)
  Choice$(0, 0) = "3"

  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn#1"
  Choice$(3, 0) = "Prn#2"

  FirstTime = True

  FF$ = CHR$(12)

  SHARED VendorIdx AS VendorIdxRecType
  VendIdxRecLen = LEN(VendorIdx)

  SHARED Vendor AS VendorRecType

  REDIM APLedgerRec(1) AS APLedger81RecType
  LdRecLen = LEN(APLedgerRec(1))
  
  NumFlds = LibNumberOfFields("AP", "VENDHIST")
  Header$ = "Vendor History"
  MaxLines = 63
  LineCnt = 0
  User$ = QPTrim$(User$)
  Page = 0

  REDIM Frm(1)    AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  'for each page, get first and last fields

  StartEl = 0   'Load first form at array start
  LibGetFldDef "AP", "VENDHIST", StartEl, Fld(), Form$(), ErrCode
  LibFile2Scrn "AP", "VENDHIST", MonoCode, -1, ErrCode          'display screen
  'PrintTitle User$
  PrintHelp "Vendor History Report"
  help$ = "(*)=All"
  QPrintRC help$, 25, 73, -1
  Action = 1

  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    'IF FirstTime THEN
    '  FirstTime = False
    '  GOSUB Get1stLastVendor
    'END IF
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
  OUTWhere$ = QPTrim$(Form$(3, 0))

  SELECT CASE OUTWhere$
  CASE "Screen"
    EntryPoint = 2
    LPTPort = 1
    ToScrnOUT = True
  CASE ""
    EntryPoint = 2
    LPTPort = 1
    ToScrnOUT = True
  CASE "Prn#1"
    EntryPoint = 5
    LPTPort = 1
    ToScrnOUT = False
  CASE "Prn#2"
    EntryPoint = 5
    LPTPort = 2
    ToScrnOUT = False
  END SELECT

  ' (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
  SortT TempVList(1), 2, 0, 32, 0, 32

  Vend1St$ = QPTrim$(TempVList(1).V)
  VendLst$ = QPTrim$(TempVList(2).V)

  NumVendRecs = (FileSize("apvendor.idx") \ VendIdxRecLen)

  REDIM VIndex(1 TO NumVendRecs) AS VendorIdxRecType
  FGetAH "apvendor.idx", VIndex(1), 12, NumVendRecs

  PrnFile = FREEFILE

  ReportFile$ = "vendhist.PRN"

  OPEN ReportFile$ FOR OUTPUT AS #PrnFile

  OpenVendorFile VendorFile, NumVRecs
  OpenAPLedgerFile APLedgerFile, NumTrans, LdRecLen
  GOSUB VendHistHeader

  FOR Cnt = 1 TO NumVendRecs
    VCode$ = QPTrim$(VIndex(Cnt).VendorCode)
    IF VCode$ >= Vend1St$ AND VCode$ <= VendLst$ THEN
      TransCnt = 0
      VTBal# = 0
      VTDebit# = 0
      VTCredit# = 0
      POCredit# = 0
      PODebit# = 0
      POTBal# = 0
      GET VendorFile, VIndex(Cnt).RecNum, Vendor
      IF Vendor.DELFLAG = 0 THEN

        IF LineCnt >= MaxLines THEN
          PRINT #PrnFile, FF$
          GOSUB VendHistHeader
        END IF

        PRINT #PrnFile, "Vendor: "; Vendor.Vname
        LineCnt = LineCnt + 1

        NextTrans& = Vendor.FrstTran
        DO UNTIL NextTrans& = 0

          GET APLedgerFile, NextTrans&, APLedgerRec(1)
          TransCnt = TransCnt + 1
          IF LineCnt >= MaxLines THEN
            PRINT #PrnFile, FF$
            GOSUB VendHistHeader
            PRINT #PrnFile, "Vendor: "; Vendor.Vname
          END IF
          PRINT #PrnFile, Num2Date(APLedgerRec(1).TrDate); "  ";
          LineCnt = LineCnt + 1

          IF APLedgerRec(1).TrCode = 1 THEN
            VTCredit# = Round(VTCredit# + APLedgerRec(1).Amt)
            IF APLedgerRec(1).PdCheckDate > 0 THEN
              'PRINT #PRNFile, "Inv "; LEFT$(APLedgerRec(1).DOCNum, 20); APLedgerRec(1).PdCheckNum; "  "; Num2Date(APLedgerRec(1).PdCheckDate); TAB(62); FUsing(STR$(APLedgerRec(1).Amt), "$$,######.##");"  Pd "
              PRINT #PrnFile, "Inv "; LEFT$(APLedgerRec(1).DOCNum, 20); TAB(54); FUsing(STR$(APLedgerRec(1).Amt), "$$,######.##"); "  Pd"; APLedgerRec(1).PDCheckNum
            ELSE
              PRINT #PrnFile, "Inv "; LEFT$(APLedgerRec(1).DOCNum, 20); TAB(54); FUsing(STR$(APLedgerRec(1).Amt), "$$,######.##"); "  Open"
            END IF

          '--02/27/97 to handle voided invoices
          ELSEIF APLedgerRec(1).TrCode = -1 THEN
             PRINT #PrnFile, "Inv "; LEFT$(APLedgerRec(1).DOCNum, 20); TAB(54); FUsing(STR$(APLedgerRec(1).Amt), "$$,######.##"); "  Void"
          '--02/27/97

          ELSEIF APLedgerRec(1).TrCode = 4 THEN
            POCredit# = Round(POCredit# + APLedgerRec(1).Amt)
            PRINT #PrnFile, "PO  "; LEFT$(APLedgerRec(1).DOCNum, 20); TAB(54); FUsing(STR$(APLedgerRec(1).Amt), "$$,######.##"); "  Open"
          ELSEIF APLedgerRec(1).TrCode = -4 THEN  'for cleared PO's
            PODebit# = Round(PODebit# + APLedgerRec(1).Amt)
            PRINT #PrnFile, "PO  "; LEFT$(APLedgerRec(1).DOCNum, 20); TAB(54); FUsing(STR$(APLedgerRec(1).Amt), "$$,######.##"); "  Closed"
          ELSEIF APLedgerRec(1).TrCode = 3 THEN
            VTDebit# = Round(VTDebit# + APLedgerRec(1).Amt)
            PRINT #PrnFile, "Chk "; LEFT$(APLedgerRec(1).DOCNum, 20); TAB(42); FUsing(STR$(APLedgerRec(1).Amt), "$$,######.##")
          ELSEIF APLedgerRec(1).TrCode = -3 THEN
            PRINT #PrnFile, "Chk "; LEFT$(APLedgerRec(1).DOCNum, 10); " VOID"; TAB(42); FUsing(STR$(APLedgerRec(1).Amt), "$$,######.##")
          ELSE
            PRINT #PrnFile, "Code is:"; APLedgerRec(1).TrCode
          END IF
          LineCnt = LineCnt + 1

         ' IF LEN(QPTrim$(APLedgerRec(1).Get1099)) > 0 THEN
         '   PRINT #PrnFile, "!!"; APLedgerRec(1).Get1099; "!!"
         ' END IF

          NextTrans& = APLedgerRec(1).NextTrans

        LOOP

        IF TransCnt = 0 THEN
          PRINT #PrnFile, "  NO TRANSACTIONS."
          PRINT #PrnFile, Dash$
          LineCnt = LineCnt + 2
        ELSE
          VTBal# = Round(VTCredit# - VTDebit#)
          PRINT #PrnFile, "Balance :"; FUsing(STR$(VTBal#), "$$,######.##")
          PRINT #PrnFile, "On Order:"; FUsing(STR$(POCredit#), "$$,######.##")

          PRINT #PrnFile, Dash$
          LineCnt = LineCnt + 2
        END IF
      END IF
    END IF
  NEXT
  PRINT #PrnFile, FF$
  CLOSE
  ERASE VIndex

  PrintRptFile "Vendor History Report", ReportFile$, LPTPort%, RetCode%, EntryPoint

  IF ListCnt = 0 THEN EXIT SUB
  'add a trap here to display an error scrn if no matching ledeger recs
  'to pay

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

  REDIM VendorList(1 TO NumVRecs) AS FLen       'Dim array to hold trans list
  'REDIM VendorList(1 TO 10) AS FLen       'Dim array to hold trans list

  FOR Cnt = 1 TO NumActiveVendors               'Read the file
    GET VendorIdxFile, Cnt, VendorIdx
    IF LEFT$(VendorIdx.VendorCode, VLookuplen) = VLookup$ THEN
      GET VendorFile, VendorIdx.RecNum, Vendor
      IF Vendor.DELFLAG = 0 THEN
        MatchingVendors = MatchingVendors + 1
        LSET VendorList(MatchingVendors).V = List$
        VendorList(MatchingVendors).V = VendorIdx.VendorCode
        MID$(VendorList(MatchingVendors).V, 13) = Vendor.Vname
        MID$(VendorList(MatchingVendors).V, 31) = MKI$(VendorIdx.RecNum)
      END IF
    END IF
  NEXT

  CLOSE

  IF MatchingVendors = 0 THEN
    'Ok = MsgBox("AP", "NOVENDOR")
    VendPickList = 0
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
    MaxWidth = 29
    BoxBot = 18
    ListDescTxt$ = "VendorCode  Name"
    ListDesc$ = SPACE$(MaxWidth + 4)
    MID$(ListDesc$, 3) = ListDescTxt$
    QPrintRC ListDesc$, Row - 1, Col, 112       'Cnf.ActivCh
    Choice = 1
    DO
      'call vertmenu to display matching recs to choose from
      VertMenuT2 VendorList(), Choice, MaxWidth, BoxBot, Ky$, Action, Cnf
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
    'Form$(3, 0) = DATE$
    SaveField 1, Form$(), Fld(), BadFld
    Frm(1).FldNo = 3
    Action = 2
  END IF
  CLOSE VendorIdxFile
  RETURN

VendHistHeader:
Page = Page + 1
  PRINT #PrnFile, TAB(40 - (INT(LEN(User$) / 2))); User$
  PRINT #PrnFile, TAB(40 - (INT(LEN(Header$) / 2))); Header$
  PRINT #PrnFile,
  PRINT #PrnFile, "Report Date: "; DATE$; TAB(67); "Page #"; Page
  PRINT #PrnFile, "Inv Date    TrCode  Desc                        Debit      Credit  Status"
  PRINT #PrnFile, "=============================================================================="
  LineCnt = 6
RETURN



END SUB

SUB xSalesTaxReport

  LOCATE 1, 1, 0

  LibName$ = "AP.QSL"           'screen library
  FormName$ = "SALESTX"        'form to edit
  DevFld = 7    'Set to output device field number
  REDIM Form$(0, 0)
  REDIM Fld(0) AS FieldInfo
  DIM Frm AS FormInfo
  SHARED Choice$()
  REDIM Choice$(0 TO 3, 1)
  Choice$(0, 0) = STR$(DevFld)  'Choices for Output Device
  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn#1"
  Choice$(3, 0) = "Prn#2"

  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
  Action = 1
  DO
    EditForm Form$(), Fld(), Frm, Cnf, Action
    IF Frm.PrevFld <> Frm.FldNo THEN
      'IF Form$(1, 0) = "Y" THEN
      '  Fld(3).Protected = True
      'ELSE
      '  Fld(4).Protected = True
      'END IF
      IF Frm.FldNo = 2 OR Frm.FldNo = 3 THEN
        Form$(Frm.FldNo, 0) = ""
        Action = 1
      END IF
      '--strip out dashes for editing
      IF Frm.FldNo = 2 OR Frm.FldNo = 3 THEN
        IF LEN(Form$(Frm.FldNo, 0)) > 0 THEN
          Form$(Frm.FldNo, 0) = StripAcct$(Form$(Frm.FldNo, 0))
          SaveField Frm.FldNo, Form$(), Fld(), 0
          Action = 1
        END IF
      END IF
      SELECT CASE Frm.PrevFld
      CASE 2, 3
        Form$(Frm.PrevFld, 0) = FmtAcct$(Form$(Frm.PrevFld, 0), FundLen%, AcctLen%, DetLen%)
      CASE ELSE
      END SELECT
      Action = 1
    END IF
    SELECT CASE Frm.KeyCode
    CASE F10Key 'SaveButton
      IF Form$(1, 0) = "Y" THEN Combined = True
      StateTaxRecAcct$ = QPTrim$(Form$(2, 0))
      CountyTaxRecAcct$ = QPTrim$(Form$(3, 0))
      StateFactor# = VAL(Form$(4, 0))
      BegDate = Date2Num(Form$(5, 0))
      EndDate = Date2Num(Form$(6, 0))
      IF LEN(LTRIM$(RTRIM$(Form$(DevFld, 0)))) = 0 THEN
        Dev$ = "S"
        LPTNo = 1
      ELSE
        Dev$ = LEFT$(Form$(DevFld, 0), 1)
        LPTNo = VAL(RIGHT$(RTRIM$(Form$(DevFld, 0)), 1))
      END IF
      IF BegDate > EndDate THEN 'OR Form$(1, 0) > Form$(2, 0) THEN
        OK = MsgBox("GL", "BADPARMS")
        Ok2Go = 0
        Frm.FldNo = 5
        Action = 1
      ELSE
        Ok2Go = -1
      END IF
    CASE ESC
      EXIT SUB
    END SELECT
  LOOP UNTIL Ok2Go
  
  SHARED Vendor AS VendorRecType
  OpenVendorFile VendorFile, NumVRecs

  'DIM APLedger AS APLedgerRecType
  DIM APLedger AS APLedger81RecType
  LdRecLen = LEN(APLedger)
  OpenAPLedgerFile APLedgerFile, NumTrans, LdRecLen

  DIM APDist AS APDistRecType
  APDRecLen = LEN(APDist)
  OpenAPDistFile APDistFile, NumDistRecs&, APDRecLen
  
  DIM StSalesTaxPaid#(0 TO 100)
  IF NOT Combined THEN
    DIM CoSalesTaxPaid#(0 TO 100)
  END IF

  FOR Cnt = 1 TO NumVRecs
    GET VendorFile, Cnt, Vendor
    NextTran = Vendor.FrstTran
    IF NextTran > 0 THEN
      DO
        GET APLedgerFile, NextTran, APLedger
        IF APLedger.TrCode = 1 THEN
          TCnt = TCnt + 1
          NextDist& = APLedger.FrstDist
          DO
             GET APDistFile, NextDist&, APDist
             IF QPTrim$(APDist.DistAcctNum) = StateTaxRecAcct$ THEN
               GET VendorFile, APLedger.VRecNum, Vendor
               'PRINT #PrnFile, Vendor.VName, Vendor.StCode, Vendor.CoCode
               SCnt = SCnt + 1
               StTax# = StTax# + APDist.DistAmt
               IF Combined THEN
                  offset = VAL(Vendor.CoCode)
               ELSE
                  offset = VAL(Vendor.StCode)
               END IF
               IF offset < 0 OR offset > 100 THEN offset = 0
               StSalesTaxPaid#(offset) = StSalesTaxPaid#(offset) + APDist.DistAmt
             END IF
             IF NOT Combined THEN
               IF QPTrim$(APDist.DistAcctNum) = CountyTaxRecAcct$ THEN
                 GET VendorFile, APLedger.VRecNum, Vendor
                 'PRINT #PrnFile, Vendor.VName, Vendor.StCode, Vendor.CoCode
                 CCnt = CCnt + 1
                 CoTax# = CoTax# + APDist.DistAmt
                 offset = VAL(Vendor.CoCode)
                 IF offset < 0 OR offset > 100 THEN offset = 0
                 CoSalesTaxPaid#(offset) = CoSalesTaxPaid#(offset) + APDist.DistAmt
               END IF
             END IF
             NextDist& = APDist.NextDist
          LOOP UNTIL NextDist& = 0
        END IF
        NextTran = APLedger.NextTrans
      LOOP UNTIL NextTran = 0
    END IF
  NEXT
  
  CLOSE

  PrnFile = FREEFILE
  ReportFile$ = "SALESTAX.PRN"
  OPEN ReportFile$ FOR OUTPUT AS #PrnFile
  IF Combined THEN
    FOR Cnt = 0 TO 100
      IF StSalesTaxPaid#(Cnt) > 0 THEN
        StateAmt# = (StSalesTaxPaid#(Cnt) * StateFactor#)
        CountyAmt# = (StSalesTaxPaid#(Cnt) - StateAmt#)
        StTotTax# = StTotTax# + StSalesTaxPaid#(Cnt)
        TotState# = TotState# + StateAmt#
        TotCounty# = TotCounty# + CountyAmt#
        PRINT #PrnFile, Cnt; StateAmt#; CountyAmt#
      END IF
    NEXT
    PRINT #PrnFile, "State Total: "; TotState#; "County Total: "; TotCounty#
    PRINT #PrnFile, "Grand Total: "; TotState# + TotCounty#
    PRINT #PrnFile, StTax#
  ELSE
    FOR Cnt = 0 TO 100
      IF StSalesTaxPaid#(Cnt) > 0 OR CoSalesTaxPaid#(Cnt) > 0 THEN
        PRINT #PrnFile, Cnt; StSalesTaxPaid#(Cnt); CoSalesTaxPaid#(Cnt)
        StTotTax# = StTotTax# + StSalesTaxPaid#(Cnt)
        CoTotTax# = CoTotTax# + CoSalesTaxPaid#(Cnt)
      END IF
    NEXT
    PRINT #PrnFile, "State Total:  "; StTotTax#
    PRINT #PrnFile, "County Total: "; CoTotTax#
  END IF

  CLOSE
  
  SELECT CASE Dev$
   CASE "S"
     EntryPoint = 2
   CASE "P"
     EntryPoint = 5
  END SELECT

  Header$ = "Sales Tax Report"
  PrintRptFile Header$, ReportFile$, LPTNo, RetCode%, EntryPoint
          
END SUB

