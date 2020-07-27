DEFINT A-Z
DECLARE SUB ClearBox ()
DECLARE FUNCTION ConvDateStr$ (D$)
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE FUNCTION EDCheck4Acct% (Form$(), AcctFld%)
DECLARE FUNCTION Exist (FileName$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FLof& (FileHandle%)
DECLARE FUNCTION FUsing$ (number$, Mask$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION FindAcct% (AcctNum$)
DECLARE FUNCTION FindDept% (DeptNum$)
DECLARE FUNCTION FindFund% (FundNum$)
DECLARE FUNCTION FmtAcct$ (AcctNum$, FundLen%, AcctLen%, DetLen%)
DECLARE FUNCTION FmtDate$ (D$)
DECLARE FUNCTION GetAcctTitle$ (AcctRecNum)
DECLARE FUNCTION GetAcctType$ (AcctRecNum)
DECLARE FUNCTION GetBegDate% ()
DECLARE FUNCTION GetDeptTitle$ (DeptRecNum%)
DECLARE FUNCTION GetEndDate% ()
DECLARE FUNCTION GetFundTitle$ (FundRecNum%)
DECLARE FUNCTION GetPct$ (N1#, N2#)
DECLARE FUNCTION MGetKey% (Row%, Col%)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION Peek1% (Segment%, Address%)
DECLARE FUNCTION QPStrI$ (number%)
DECLARE FUNCTION QPStrL$ (number&)
DECLARE FUNCTION QPTrim$ (X$)
DECLARE FUNCTION Round# (N#)
DECLARE FUNCTION StripAcct$ (AcctNum$)
DECLARE FUNCTION WaitKey% ()
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB BLPrint (LPTPORT%, Text$, ErrCode%)
DECLARE SUB CalcAcctBal (StartFund$, EndFund$, FundLen, EndDate)
DECLARE SUB ClrBG ()
DECLARE SUB DisplayTotals ()
DECLARE SUB EdIFEntry ()
DECLARE SUB ExtractBL (ThruDate%)
DECLARE SUB ExtractCM (ThruDate%)
DECLARE SUB ExtractDC (ThruDate%)
DECLARE SUB ExtractTX (ThruDate%)
DECLARE SUB ExtractUB (ThruDate%)
DECLARE SUB FClose (Handle%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetRTA (Handle, UBSetUpRec(), A&, UBSetUpLen%)
DECLARE SUB FOpenAll (FileName$, AccessMode%, SharedMode%, Handle%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FileView (FileName$, Ky, Action, FVI AS ANY, SEG Array)
DECLARE SUB GetFundCodes (FirstFund$, LastFund$)
DECLARE SUB GetFundList (FundList$(), NumFunds%)
DECLARE SUB GetPostDates (LPDate, HPEndDate)
DECLARE SUB HideCursor ()
DECLARE SUB IFInit ()
DECLARE SUB LoadUBSetUpFile (UBSetUpFileNum%, UBSetUpRecLen%)
DECLARE SUB MPaintBox (ULRow%, ULCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB MScrnRest (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MScrnSave (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MnuTitle (Title$, Row, Col, Wdth, TxtColor)
DECLARE SUB OpenAcctFile (AcctFileNum%, NumAccts%)
DECLARE SUB OpenAcctIdx (AcctIdxFileNum%, NumIdxRecs%)
DECLARE SUB OpenDeptFile (DeptFile%, NumDepts%)
DECLARE SUB OpenDeptIdx (DeptIdxFile%, NumActiveDepts%)
DECLARE SUB OpenFundFile (FundFileNum%, NumFunds%)
DECLARE SUB OpenFundIdx (FundIdxFileNum%, NumFIdxRecs%)
DECLARE SUB OpenIFEditFile (IFEditFileNum%, NumIFTrans%)
DECLARE SUB OpenTransFile (TransFileNum%, NumTrans&)
DECLARE SUB Post2GL (FileName$, BadTrans%)
DECLARE SUB PostTrans ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPORT%, RetCode%, EntryPoint%)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PrnEditList ()
DECLARE SUB QPrint (X$, Colr%, PAGE%)
DECLARE SUB QPrintRC (T$, R%, c%, Clr%)
DECLARE SUB ShowCursor ()
DECLARE SUB SortAcctIndex ()
DECLARE SUB SortFundIndex ()
DECLARE SUB SortStr2 (BYVAL Address%, Size%, Direction%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TRGetAcctStruct (GLFundLen%, GLAcctLen%, GLDetLen%)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE SUB Trxfer2GJ ()
  'DECLARE SUB VertMenuT (Items() AS ANY, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
  
  '$INCLUDE: 'DefCnf.bi'                    'monitor and color info
  '$INCLUDE: 'FieldInf.bi'                  'field information
  '$INCLUDE: 'FormEdit.bi'                  'form editing information
  '$INCLUDE: 'QSCR.BI'                      'QuickScreen Routines
  '$INCLUDE: 'GLAUX.BI'                     'Common GL Routines
  '$INCLUDE: 'GL.BI'                        'GL File Types
  '$INCLUDE: 'cmfiles.bi'
  '$INCLUDE: 'ARv85.bi'                        'A/R FILE LAYOUTS
  '$INCLUDE: 'DC.bi'                        'Decal File Layouts
  '$INCLUDE: 'SCRAUX.BI'
  '$INCLUDE: 'SetCnf.bi'        'Set monitor and color information
  '$INCLUDE: 'init.bi'
  
  CONST False = 0, True = NOT False
  
  TYPE FLen
    v AS STRING * 42
  END TYPE
  
  TYPE TranRecInfoType
    TranDate  AS INTEGER
    TranRecNo AS LONG
  END TYPE
  
  DIM SHARED MCodeRec(1) AS MiscCodeRecType
  DIM SHARED GLSetup(1) AS GLSetupRecType
  DIM SHARED GJRec(1) AS TrEditRecType
  DIM SHARED GJEdit AS TrEditRecType
  DIM SHARED IFEdit AS TrEditRecType
  
  'changedate
  'setfieldvalue
  
  LOCATE 1, 1, 0
  TextCursor -2, -2             'Define the Mouse cursor
  ShowCursor    'Show it
  
  'GOTO ungrab
  
  '***Master Menu
  
  REDIM mChoice$(1 TO 6)
  
  mChoice$(1) = "Grab Transactions"
  mChoice$(2) = "Print Journal Register"
  mChoice$(3) = "Post Journal Entries"
  mChoice$(4) = "Transfer to General Journal"
  mChoice$(5) = "Initialize Transactions for I/F"
  mChoice$(6) = "Quit"
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 20   'limit the box to line 20
  Action = 0    '0 means stay in the menu
  Choice = 1    'Pre-load choice to highlight
  
  FOR Cnt = 1 TO UBOUND(mChoice$)               'Find max menu width
    TLen = LEN(mChoice$(Cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT
  
  'Center Menu within Screen
  Row = ((25 - (UBOUND(mChoice$))) \ 2) - 1
  Col = ((80 - MaxLen) \ 2) - 2
  
  ShowCursor
  
  DO
    LibFile2Scrn "GL.QSL", "MENUBAK", MonoCode, Attribute, ErrorCode
    'TitleBox 3, Col, MaxLen + 3, "Get Distribution Menu", Cnf
    'TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    PrintTitle User$
    Title$ = "Get Distribution Menu  "
    MnuTitle Title$, Row, Col, MaxLen, 15
    PrintHelp "General Journal Processing"
    
    LOCATE Row, Col, 0
    VertMenu mChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO
    
    SELECT CASE Choice
    CASE 1
      ClrBG
      EdIFEntry
    CASE 2
      ClrBG
      PrnEditList
    CASE 3
      ClrBG
      PostTrans
    CASE 4
      ClrBG
      Trxfer2GJ
    CASE 5
      ClrBG
      IFInit
    CASE 6
      CLS : END
    CASE ELSE
    END SELECT
    
  LOOP
  
  ClrBG
  HideCursor
  RUN "glmenu"
  
  'ungrab:
  '
  '  ADate = Date2Num("03-16-2000")
  '  REDIM CMTransRec(1)  AS CMTransRecType
  '  CMTransRecLen = LEN(CMTransRec(1))
  '
  '  OPEN "CMTRANS.DAT" FOR RANDOM AS #1 LEN = CMTransRecLen
  '  NumTrans& = LOF(1) \ CMTransRecLen
  '  FOR Cnt& = 1 TO NumTrans&
  '    GET #1, Cnt&, CMTransRec(1)
  '    IF CMTransRec(1).TransDate >= ADate THEN
  '      IF CMTransRec(1).Trans2GL = "Y" THEN
  '        CMTransRec(1).Trans2GL = ""
  '        PUT #1, Cnt&, CMTransRec(1)
  '        Found = Found + 1
  '      END IF
  '    END IF
  '  NEXT
  '  CLOSE
  '  PRINT "Found:"; Found
  '  PRINT "NFound:"; NFound
  '
  '  END

SUB EdIFEntry
  
  GetPostDates LPDate, HPDate
  
  SHARED Choice$()              'Multiple choice array
  REDIM Form$(0, 0)             'Holds data from form
  REDIM Fld(1) AS FieldInfo     'Field editing array
  
  REDIM Choice$(0 TO 2, 1)      'Choices for entry type field
  Choice$(0, 0) = "5"           'Field Number
  Choice$(1, 0) = "Debit "
  Choice$(2, 0) = "Credit"
  
  
  OpenIFEditFile IFEditFileNum, NumIFTrans      'Opens the GJ Edit File
  
  
  LibName$ = "GL"
  ScrnName$ = "CPIF"
  Help$ = "Cash Management Operator Entry"
  LOCATE 1, 1, 0
  
  Help$ = "General Ledger Transaction Entry"
  ShowCursor
  
  
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Rem Set Choice
  REDIM Choice$(5, 0)
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "CM-Misc Cash Receipts"
  Choice$(2, 0) = "UB-Utility Billing"
  Choice$(3, 0) = "BL-Business Licence"
  Choice$(4, 0) = "TX-Tax Billing"
  Choice$(5, 0) = "DC-Va. Vehicle Decals"
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  ' Get operator name and display
  
  
  
  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%
  PrintHelp Help$
  
  
  PrintTitle User$
  
  'Set Processing Thru Date equal to last date in posting range
  Form$(2, 0) = Num2Date$(HPDate)
  
  Action = 1
  
  DIM Frm AS FormInfo           'Dim the form info array
  
  CLOSE         'Close file while doing input
  
  DO            'Edit the form
    
    EditForm Form$(), Fld(), Frm, Cnf, Action
    
    SELECT CASE Frm.KeyCode     'Respond to the key presses
    CASE F10Key 'Save
      
      GOSUB CheckFlds
      IF BadType = 1 THEN
        Frm.FldNo = 1
        PRINT CHR$(7);
        Help$ = "** Must Select Type **"
        PrintHelp Help$
        Action = 1
      END IF
      IF BadDate = 1 THEN
        Form$(2, 0) = Num2Date$(HPDate)
        Frm.FldNo = 2
        PRINT CHR$(7);
        Help$ = "** Must Enter Date In Posting Range **"
        PrintHelp Help$
        Action = 1
      END IF
      
      IF Good = 1 THEN
        IF Form$(1, 0) = "CM" THEN ExtractCM (Date2Num%(Form$(2, 0))): EXIT SUB
        IF Form$(1, 0) = "UB" THEN ExtractUB (Date2Num%(Form$(2, 0))): EXIT SUB
        IF Form$(1, 0) = "TX" THEN ExtractTX (Date2Num%(Form$(2, 0))): EXIT SUB
        IF Form$(1, 0) = "BL" THEN ExtractBL (Date2Num%(Form$(2, 0))): EXIT SUB
        IF Form$(1, 0) = "DC" THEN ExtractDC (Date2Num%(Form$(2, 0))): EXIT SUB
      END IF
    CASE EscKey
      CLOSE
      EXIT SUB
    END SELECT
  LOOP
CheckFlds:
  'Set Default Flags
  BadDate = 0: BadType = 0: Good = 0
  IF LEN(RTRIM$(Form$(1, 0))) = 0 THEN Good = 0: BadType = 1: RETURN
  IF Date2Num%(Form$(2, 0)) > HPDate THEN Good = 0: BadDate = 1: RETURN
  Good = 1
  
  RETURN
  
  
  EXIT SUB
  
END SUB

SUB ExtractBL (ThruDate)
  
  Today$ = DATE$
  Ref$ = "BL" + LEFT$(Today$, 2) + MID$(Today$, 4, 2) + RIGHT$(Today$, 2)
        
  REDIM TranInfo(1) AS TranRecInfoType
  DIM Rec#(500), RevAmt#(500), ttype%(500), Fund$(100), FundAmt#(100)
  
  Dash80$ = STRING$(80, "-")
  P2S$ = SPACE$(4)
  
  TRGetAcctStruct FundLen%, AcctLen%, DetLen%
  GJRecLen = LEN(GJRec(1))
  
  RPTFile = FREEFILE
  OPEN "GLCMTRX.RPT" FOR OUTPUT AS RPTFile
  
  ClearBox

  QPrintRC "Searching Cash Transactions.", 11, 26, 126
  QPrintRC "New Transactions:", 13, 29, Cnf.HiLite
  
  REDIM ARTransRec(1) AS ARTransRecType
  ARTransRecLen = LEN(ARTransRec(1))
  ARTransFile = FREEFILE
  OPEN "ARTrans.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARTransFile LEN = ARTransRecLen
  NumOfTRecs& = LOF(ARTransFile) \ ARTransRecLen
  LOCK #ARTransFile
  
  FOR TCnt& = NumOfTRecs& TO 1 STEP -1
    GET #ARTransFile, TCnt&, ARTransRec(1)
    IF ARTransRec(1).TransDate <= ThruDate THEN
      IF LEN(QPTrim$(ARTransRec(1).Posted2GL)) = 0 OR QPTrim$(ARTransRec(1).Posted2GL) = "N" THEN
        'Store trans rec numbers and dates in array
        FoundCnt = FoundCnt + 1
        REDIM PRESERVE TranInfo(FoundCnt) AS TranRecInfoType
        TranInfo(FoundCnt).TranDate = ARTransRec(1).TransDate
        TranInfo(FoundCnt).TranRecNo = TCnt&
     ' ELSE
     '   NGCnt = NGCnt + 1
      END IF
    END IF
    RSET P2S$ = QPStrI$(FoundCnt)
    QPrintRC P2S$, 13, 47, Cnf.HiLite
    'SmallPause
    'Allow 500 Bad Entries Before Exiting
    'IF NGCnt >= 500 THEN EXIT FOR
  NEXT
  '
  IF FoundCnt = 0 THEN
    CLOSE
    ClearBox
    PRINT CHR$(7);
    QPrintRC "No Transactions Found To InterFace", 12, 24, 126
    SLEEP 4
    GOTO BLSendExit
  END IF
  
  'Array (1), NumElem, Dir, StructSize, MemOff, MemSize
  SortT TranInfo(1), FoundCnt, 0, 6, 0, -1      'sort'em by date. oldest first
  
  
  'Open GL InterFace File
  GJFile = FREEFILE
  OPEN "GLTRXED.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS GJFile LEN = GJRecLen
  NumEdTrans = LOF(GJFile) \ GJRecLen
  LOCK #GJFile
  
  'OPEN Business License CODES Here
  
  REDIM ARCatCodeRec(1) AS ARNewCatCodeRecType

  ARCatCodeRecLen = LEN(ARCatCodeRec(1))
  ARCatFile = FREEFILE
  OPEN "ARCODE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARCatFile LEN = ARCatCodeRecLen
  NumOFARCatRecs = LOF(ARCatFile) \ ARCatCodeRecLen

  IF NumOFARCatRecs > 1 THEN
    MiddleRec = NumOFARCatRecs \ 2
  END IF

  FirstTran = 1
  ThisDate = TranInfo(1).TranDate
  WorkDate = ThisDate
  
  FOR Cnt = 1 TO FoundCnt
    IF ThisDate <> TranInfo(Cnt).TranDate THEN
      ThisDate = TranInfo(Cnt).TranDate
      GOSUB BLProcessThisBunch
      DayCount = 0
      WorkDate = ThisDate
    END IF
    
    GET #ARTransFile, TranInfo(Cnt).TranRecNo, ARTransRec(1)
    
    'Find Catagory Record Number So We Can Pull G/L Revenue Account
    CatCodeRecord = 0
    CatCodeRecord = ARTransRec(1).CatCodeRec

    IF CatCodeRecord = 0 THEN
      CatCodeRecord = MiddleRec
    END IF

    IF CatCodeRecord > 0 THEN
      GET ARCatFile, CatCodeRecord, ARCatCodeRec(1)
      IF DayCount = 0 THEN
        RevAmt# = ARTransRec(1).TransAmount
        RevAmt# = Round#(RevAmt#)
        IF RevAmt# <> 0 THEN
          'If There Is an Amount get catagory code record
          DayCount = DayCount + 1
          Rec#(DayCount) = CatCodeRecord
          RevAmt#(DayCount) = RevAmt#
          ttype%(DayCount) = ARTransRec(1).TransType
        END IF
      ELSE
        RevAmt# = ARTransRec(1).TransAmount
        RevAmt# = Round#(RevAmt#)
        DO WHILE RevAmt# <> 0
          FOR FindCount = 1 TO DayCount
            IF Rec#(FindCount) = CatCodeRecord THEN
              RevAmt#(FindCount) = RevAmt#(FindCount) + RevAmt#
              RevAmt# = 0
              EXIT DO
            END IF
          NEXT FindCount
          DayCount = DayCount + 1
          Rec#(DayCount) = CatCodeRecord
          RevAmt#(DayCount) = RevAmt#
          ttype%(DayCount) = ARTransRec(1).TransType
          RevAmt# = 0
        LOOP
        
      END IF
      
    END IF
    
  NEXT Cnt
  
  GOSUB BLProcessThisBunch
  
  'Mark Transactions as interfaced
  FOR Cnt = 1 TO FoundCnt
    GET #ARTransFile, TranInfo(Cnt).TranRecNo, ARTransRec(1)
    ARTransRec(1).Posted2GL = "Y"
    PUT #ARTransFile, TranInfo(Cnt).TranRecNo, ARTransRec(1)
  NEXT
  CLOSE
  
BLSendExit:
  EXIT SUB
  
  
BLProcessThisBunch:
  ' Must Combine By Date and Then Do Cash Debit Entry For Total by Fund
  
  IF DayCount <= 0 THEN RETURN
  
  FundCnt = 0   ' Set Funds Used to Zero
  
  'Process Payments First Type=2
  
  FOR Process = 1 TO DayCount
    IF ttype%(Process) = 2 THEN
      GET #ARCatFile, Rec#(Process), ARCatCodeRec(1)
      
      IF LEN(QPTrim$(ARCatCodeRec(1).ARGLACCT)) > 0 THEN
        acct$ = ARCatCodeRec(1).ARGLACCT
        acct$ = StripAcct$(acct$)
        acct$ = FmtAcct$(acct$, FundLen%, AcctLen%, DetLen%)
      ELSE
        acct$ = ARCatCodeRec(1).REVGLNUM
        acct$ = StripAcct$(acct$)
        acct$ = FmtAcct$(acct$, FundLen%, AcctLen%, DetLen%)
      END IF
      
      GOSUB BLGetAcctName
      
      GJRec(1).AcctRec = 0
      GJRec(1).AcctNum = QPTrim$(acct$)
      GJRec(1).AcctName = AcctName$
      GJRec(1).TrDate = WorkDate
      GJRec(1).Ref = Ref$
      GJRec(1).CrAmt = RevAmt#(Process)
      GJRec(1).DrAmt = 0
      GJRec(1).EType = "C"
      GJRec(1).Desc = "FROM CASH MGMT"
      GJRec(1).Src = "CR"
      PUT #GJFile, , GJRec(1)
      
      
      'Now Make Matching Debit Entries to Cash Account
      SetupFile = FREEFILE
      SetupRecLen = LEN(GLSetup(1))
      OPEN "GLSETUP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS SetupFile LEN = SetupRecLen
      GET SetupFile, 1, GLSetup(1)
      CASHACCT$ = GLSetup(1).CRCashAcct
      CLOSE SetupFile
      acct$ = LEFT$(acct$, FundLen%) + CASHACCT$
      acct$ = QPTrim$(acct$)
      acct$ = StripAcct$(acct$)
      acct$ = FmtAcct$(acct$, FundLen%, AcctLen%, DetLen%)
      
      GOSUB BLGetAcctName
      
      GJRec(1).AcctRec = 0
      GJRec(1).AcctNum = acct$
      GJRec(1).AcctName = AcctName$
      GJRec(1).TrDate = WorkDate
      GJRec(1).Ref = Ref$
      GJRec(1).DrAmt = RevAmt#(Process)
      GJRec(1).CrAmt = 0
      GJRec(1).EType = "D"
      GJRec(1).Desc = "FROM CASH MGMT"
      GJRec(1).Src = "CR"
      PUT #GJFile, , GJRec(1)
      
    END IF
  NEXT Process
  
  
  'Process Charges if Needed Type=1 For Accrual Only
  
  FundCnt = 0
  
  FOR Process = 1 TO DayCount
    IF ttype%(Process) = 1 THEN
      GET #ARCatFile, Rec#(Process), ARCatCodeRec(1)
      IF LEN(QPTrim$(ARCatCodeRec(1).ARGLACCT)) > 0 THEN
        acct$ = ARCatCodeRec(1).REVGLNUM
        acct$ = StripAcct$(acct$)
        acct$ = FmtAcct$(acct$, FundLen%, AcctLen%, DetLen%)
        
        GOSUB BLGetAcctName
        
        GJRec(1).AcctRec = 0
        GJRec(1).AcctNum = QPTrim$(acct$)
        GJRec(1).AcctName = AcctName$
        GJRec(1).TrDate = WorkDate
        GJRec(1).Ref = Ref$
        GJRec(1).CrAmt = RevAmt#(Process)
        GJRec(1).DrAmt = 0
        GJRec(1).EType = "C"
        GJRec(1).Desc = "FROM CASH MGMT"
        GJRec(1).Src = "CR"
        PUT #GJFile, , GJRec(1)
        
        
        
        'Now Make Matching Debit Entries to Cash Account
        
        
        acct$ = ARCatCodeRec(1).ARGLACCT
        acct$ = StripAcct$(acct$)
        acct$ = FmtAcct$(acct$, FundLen%, AcctLen%, DetLen%)
        
        GOSUB BLGetAcctName
        
        GJRec(1).AcctRec = 0
        GJRec(1).AcctNum = acct$
        GJRec(1).AcctName = AcctName$
        GJRec(1).TrDate = WorkDate
        GJRec(1).Ref = Ref$
        GJRec(1).DrAmt = RevAmt#(Process)
        GJRec(1).CrAmt = 0
        GJRec(1).EType = "D"
        GJRec(1).Desc = "FROM CASH MGMT"
        GJRec(1).Src = "CR"
        PUT #GJFile, , GJRec(1)
      END IF
    END IF
  NEXT Process
  
  
  
  
BLBunchReturn:
  RETURN
  
BLGetAcctName:
  AcctName$ = "UnDefined"
  
  
  OPEN "glacct.idx" FOR RANDOM ACCESS READ WRITE SHARED AS #13 LEN = 16: FIELD 13, 14 AS AcctNumber$, 2 AS RecordNumber$
  FOR Cnt! = 1 TO LOF(13) / 16: GET 13, Cnt!
    IF RTRIM$(AcctNumber$) = RTRIM$(acct$) THEN
      Record = CVI(RecordNumber$): ng = 1:
      OPEN "glacct.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #12 LEN = 178
      FIELD 12, 16 AS nu$, 30 AS AName$
      GET 12, Record
      AcctName$ = AName$
      CLOSE 12: CLOSE 13
      RETURN
    END IF
  NEXT Cnt!
  ng = 0
  CLOSE 12: CLOSE 13
  RETURN
  
  
END SUB

SUB ExtractCM (ThruDate%)
  
  Today$ = DATE$
  Ref$ = "CM" + LEFT$(Today$, 2) + MID$(Today$, 4, 2) + RIGHT$(Today$, 2)
  
  TRGetAcctStruct FundLen%, AcctLen%, DetLen%
  
  REDIM TranInfo(1) AS TranRecInfoType
  DIM MiscRec#(500), MiscAmt#(500), Fund$(100), FundAmt#(100)
  
  Dash80$ = STRING$(80, "-")
  P2S$ = SPACE$(4)
  
  GJRecLen = LEN(GJRec(1))
  
  RPTFile = FREEFILE
  OPEN "GLCMTRX.RPT" FOR OUTPUT AS RPTFile
  
  ClearBox

  QPrintRC "Searching Cash Transactions.", 11, 26, 126
  QPrintRC "New Transactions:", 13, 29, Cnf.HiLite
  
  REDIM CMTransRec(1)  AS CMTransRecType
  CMTransRecLen = LEN(CMTransRec(1))
  
  CMTran = FREEFILE
  OPEN "CMTRANS.DAT" FOR RANDOM SHARED AS CMTran LEN = CMTransRecLen
  NumOfTRecs& = LOF(CMTran) \ CMTransRecLen
  LOCK #CMTran
  
  FOR TCnt& = NumOfTRecs& TO 1 STEP -1
    GET #CMTran, TCnt&, CMTransRec(1)
    IF LEN(QPTrim$(CMTransRec(1).Trans2GL)) = 0 OR QPTrim$(CMTransRec(1).Trans2GL) = "N" THEN
      'Store trans rec numbers and dates in array
      IF CMTransRec(1).TransDate <= ThruDate% THEN
        FoundCnt = FoundCnt + 1
        REDIM PRESERVE TranInfo(FoundCnt) AS TranRecInfoType
        TranInfo(FoundCnt).TranDate = CMTransRec(1).TransDate
        TranInfo(FoundCnt).TranRecNo = TCnt&
      END IF
    ELSE
      NGCnt = NGCnt + 1
    END IF
    RSET P2S$ = QPStrI$(FoundCnt)
    QPrintRC P2S$, 13, 47, Cnf.HiLite
    'SmallPause
    'Allow 250 Bad Dates Before Exiting
    IF NGCnt >= 250 THEN EXIT FOR
  NEXT
  
  IF FoundCnt = 0 THEN
    CLOSE
    ClearBox
    PRINT CHR$(7);
    QPrintRC "No Transactions Found To InterFace", 12, 24, 126
    SLEEP 4
    GOTO SendExit
  END IF
  
  SortT TranInfo(1), FoundCnt, 0, 6, 0, -1      'sort'em by date. oldest first
  'Array (1), NumElem, Dir, StructSize, MemOff, MemSize
  
  GJFile = FREEFILE
  OPEN "GLTRXED.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS GJFile LEN = GJRecLen
  NumEdTrans = LOF(GJFile) \ GJRecLen
  
  MCFile = FREEFILE
  OPEN "CMMISCCD.DAT" FOR RANDOM SHARED AS MCFile LEN = LEN(MCodeRec(1))
  
  FirstTran = 1
  ThisDate = TranInfo(1).TranDate
  WorkDate = ThisDate
  
  FOR Cnt = 1 TO FoundCnt
    IF ThisDate <> TranInfo(Cnt).TranDate THEN
      ThisDate = TranInfo(Cnt).TranDate
      GOSUB ProcessThisBunch
      DayCount = 0
      WorkDate = ThisDate
    END IF
    
    GET #CMTran, TranInfo(Cnt).TranRecNo, CMTransRec(1)
    
    IF CMTransRec(1).TransSource = 1 THEN
      
      IF DayCount = 0 THEN
        
        FOR MCnt = 1 TO 5
          MiscRevAmt# = (CMTransRec(1).TransRevAmt(MCnt))
          MiscRevAmt# = Round#(MiscRevAmt#)
          IF MiscRevAmt# <> 0 THEN
            'If There Is an Amount in Misc Rev 1-5 then get code record number from 6-10
            IF CMTransRec(1).TransRevAmt(MCnt + 5) >= 1 THEN
              DayCount = DayCount + 1
              MiscRec#(DayCount) = CMTransRec(1).TransRevAmt(MCnt + 5)
              MiscAmt#(DayCount) = MiscRevAmt#
            END IF
          END IF
          
        NEXT MCnt
        
      ELSE
        FOR MCnt = 1 TO 5
          MiscRevAmt# = (CMTransRec(1).TransRevAmt(MCnt))
          MiscRevAmt# = Round#(MiscRevAmt#)
          DO WHILE MiscRevAmt# <> 0
            FOR FindCount = 1 TO DayCount
              IF MiscRec#(FindCount) = CMTransRec(1).TransRevAmt(MCnt + 5) THEN
                MiscAmt#(FindCount) = MiscAmt#(FindCount) + MiscRevAmt#
                MiscRevAmt# = 0
                EXIT DO
              END IF
            NEXT FindCount
            DayCount = DayCount + 1
            MiscRec#(DayCount) = CMTransRec(1).TransRevAmt(MCnt + 5)
            MiscAmt#(DayCount) = MiscRevAmt#
            MiscRevAmt# = 0
          LOOP
        NEXT MCnt
      END IF
    END IF
  NEXT Cnt
  GOSUB ProcessThisBunch
  
  'transactions as interfaced
  
  FOR Cnt = 1 TO FoundCnt
    GET #CMTran, TranInfo(Cnt).TranRecNo, CMTransRec(1)
    CMTransRec(1).Trans2GL = "Y"
    PUT #CMTran, TranInfo(Cnt).TranRecNo, CMTransRec(1)
  NEXT
  CLOSE
SendExit:
  EXIT SUB
  
  
ProcessThisBunch:
  ' Must Combine By Date and Then Do Cash Debit Entry For Total by Fund
  
  IF DayCount <= 0 THEN RETURN
  
  FundCnt = 0   ' Set Funds Used to Zero
  
  FOR Process = 1 TO DayCount
    IF MiscRec#(Process) >= 1 THEN
      GET #MCFile, MiscRec#(Process), MCodeRec(1)
      acct$ = MCodeRec(1).GlAcctNumb
      acct$ = StripAcct$(acct$)
      acct$ = FmtAcct$(acct$, FundLen%, AcctLen%, DetLen%)
      
      GOSUB GetAcctName
      GJRec(1).AcctRec = 0
      GJRec(1).AcctNum = QPTrim$(acct$)
      GJRec(1).AcctName = AcctName$
      GJRec(1).TrDate = WorkDate
      GJRec(1).Ref = Ref$
      GJRec(1).CrAmt = MiscAmt#(Process)
      GJRec(1).DrAmt = 0
      GJRec(1).EType = "C"
      GJRec(1).Desc = "FROM CASH MGMT"
      GJRec(1).Src = "CR"
      PUT #GJFile, , GJRec(1)
      
      'Add Up Fund Total Here for Cash Credit Entry
      IF FundCnt = 0 THEN
        FundCnt = 1
        Fund$(FundCnt) = LEFT$(acct$, FundLen%)
        FundAmt#(FundCnt) = MiscAmt#(Process)
      ELSE
        FoundFund = 0
        FOR FCnt = 1 TO FundCnt
          IF Fund$(FCnt) = LEFT$(acct$, FundLen%) THEN
            FoundFund = 1
            FundAmt#(FCnt) = FundAmt#(FCnt) + MiscAmt#(Process)
          END IF
        NEXT FCnt
        IF FoundFund = 0 THEN
          FundCnt = FundCnt + 1
          Fund$(FundCnt) = LEFT$(acct$, FundLen%)
          FundAmt#(FundCnt) = MiscAmt#(Process)
        END IF
      END IF
    END IF
  NEXT Process
  
  'Now Make Matching Debit Entries to Cash Account
  'Open GLSETUP File Here
  SetupFile = FREEFILE
  SetupRecLen = LEN(GLSetup(1))
  OPEN "GLSETUP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS SetupFile LEN = SetupRecLen
  GET SetupFile, 1, GLSetup(1)
  CASHACCT$ = GLSetup(1).CRCashAcct
  CLOSE SetupFile
  
  
  FOR Cash = 1 TO FundCnt
    acct$ = Fund$(Cash) + CASHACCT$
    acct$ = QPTrim$(acct$)
    acct$ = StripAcct$(acct$)
    acct$ = FmtAcct$(acct$, FundLen%, AcctLen%, DetLen%)
    GOSUB GetAcctName
    
    GJRec(1).AcctRec = 0
    GJRec(1).AcctNum = acct$
    GJRec(1).AcctName = AcctName$
    GJRec(1).TrDate = WorkDate
    GJRec(1).Ref = Ref$
    GJRec(1).DrAmt = FundAmt#(Cash)
    GJRec(1).CrAmt = 0
    GJRec(1).EType = "D"
    GJRec(1).Desc = "FROM CASH MGMT"
    GJRec(1).Src = "CR"
    PUT #GJFile, , GJRec(1)
  NEXT Cash
BunchReturn:
  RETURN
  
  
  
GetAcctName:
  OPEN "glacct.idx" FOR RANDOM ACCESS READ WRITE SHARED AS #13 LEN = 16
  FIELD 13, 14 AS AcctNumber$, 2 AS RecordNumber$
  AcctName$ = "UnDefined"
  FOR Cnt! = 1 TO LOF(13) / 16
    GET 13, Cnt!
    IF RTRIM$(AcctNumber$) = RTRIM$(acct$) THEN
      Record = CVI(RecordNumber$)
      ng = 1
      OPEN "glacct.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #12 LEN = 178
      FIELD 12, 16 AS nu$, 30 AS AName$
      GET 12, Record
      AcctName$ = AName$
      CLOSE 12
      CLOSE 13
      RETURN
    END IF
  NEXT Cnt!
  ng = 0
  CLOSE 12
  CLOSE 13
  RETURN
  
END SUB

SUB ExtractDC (ThruDate)
  Today$ = DATE$
  Ref$ = "IF" + LEFT$(Today$, 2) + MID$(Today$, 4, 2) + RIGHT$(Today$, 2)
  
  REDIM TranInfo(1) AS TranRecInfoType
  DIM Rec#(500), RevAmt#(500), Fund$(100), FundAmt#(100)
  
  Dash80$ = STRING$(80, "-")
  P2S$ = SPACE$(4)
  
  TRGetAcctStruct FundLen%, AcctLen%, DetLen%
  
  GJRecLen = LEN(GJRec(1))
  
  
  RPTFile = FREEFILE
  OPEN "GLCMTRX.RPT" FOR OUTPUT AS RPTFile
  
  ClearBox

  QPrintRC "Searching Cash Transactions.", 11, 26, 126
  QPrintRC "New Transactions:", 13, 29, Cnf.HiLite
  
  REDIM DCTransRec(1) AS DCTransRecType
  DCTransRecLen = LEN(DCTransRec(1))
  DCTransFile = FREEFILE
  OPEN "DCTrans.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCTransFile LEN = DCTransRecLen
  NumOfTRecs& = LOF(DCTransFile) \ DCTransRecLen
  LOCK #DCTransFile
  
  
  FOR TCnt& = NumOfTRecs& TO 1 STEP -1
    GET #DCTransFile, TCnt&, DCTransRec(1)
    
    IF (LEN(QPTrim$(DCTransRec(1).GLInterfaced)) = 0 OR QPTrim$(DCTransRec(1).GLInterfaced) = "N") AND DCTransRec(1).TransType <> 1 THEN
      'Store trans rec numbers and dates in array
      FoundCnt = FoundCnt + 1
      REDIM PRESERVE TranInfo(FoundCnt) AS TranRecInfoType
      TranInfo(FoundCnt).TranDate = DCTransRec(1).TransDate
      TranInfo(FoundCnt).TranRecNo = TCnt&
    ELSE
      NGCnt = NGCnt + 1
    END IF
    RSET P2S$ = QPStrI$(FoundCnt)
    QPrintRC P2S$, 13, 47, Cnf.HiLite
    'SmallPause
    'Allow 1500 Bad Entries Before Exiting
    IF NGCnt >= 1500 THEN EXIT FOR
  NEXT
  
  IF FoundCnt = 0 THEN
    CLOSE
    ClearBox
    PRINT CHR$(7);
    QPrintRC "No Transactions Found To InterFace", 12, 24, 126
    SLEEP 4
    GOTO DCSendExit
  END IF
  
  'Array (1), NumElem, Dir, StructSize, MemOff, MemSize
  SortT TranInfo(1), FoundCnt, 0, 6, 0, -1      'sort'em by date. oldest first
  
  
  
  
  'Open GL InterFace File
  GJFile = FREEFILE
  OPEN "GLTRXED.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS GJFile LEN = GJRecLen
  NumEdTrans = LOF(GJFile) \ GJRecLen
  LOCK #GJFile
  
  'OPEN Decal Catagory Here
  
  REDIM DCCatCodeRec(1) AS DCCatCodeRecType
  DCCatCodeRecLen = LEN(DCCatCodeRec(1))
  DCCatFile = FREEFILE
  OPEN "DCCODE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCCatFile LEN = DCCatCodeRecLen
  NumOFDCCatRecs = LOF(DCCatFile) \ DCCatCodeRecLen
  
  
  FirstTran = 1
  ThisDate = TranInfo(1).TranDate
  WorkDate = ThisDate
  
  FOR Cnt = 1 TO FoundCnt
    IF ThisDate <> TranInfo(Cnt).TranDate THEN
      ThisDate = TranInfo(Cnt).TranDate
      GOSUB DCProcessThisBunch
      DayCount = 0
      WorkDate = ThisDate
    END IF
    
    GET #DCTransFile, TranInfo(Cnt).TranRecNo, DCTransRec(1)
    
    'Find Catagory Record Number So We Can Pull G/L Revenue Account
    CATCODE$ = QPTrim$(DCTransRec(1).DecalCat)
    CatCodeRecord = 0
    
    FOR CCnt! = 1 TO NumOFDCCatRecs
      GET DCCatFile, CCnt!, DCCatCodeRec(1)
      IF QPTrim$(DCCatCodeRec(1).CATCODE) = CATCODE$ THEN
        CatCodeRecord = CCnt!
        EXIT FOR
      END IF
    NEXT CCnt!
    
    
    IF CatCodeRecord > 0 THEN
      
      IF DayCount = 0 THEN
        RevAmt# = DCTransRec(1).TransAmount
        RevAmt# = Round#(RevAmt#)
        IF RevAmt# <> 0 THEN
          'If There Is an Amount get catagory code record
          IF CatCodeRecord >= 1 THEN
            DayCount = DayCount + 1
            Rec#(DayCount) = CatCodeRecord
            RevAmt#(DayCount) = RevAmt#
          END IF
        END IF
        
        
      ELSE
        RevAmt# = DCTransRec(1).TransAmount
        RevAmt# = Round#(RevAmt#)
        DO WHILE RevAmt# <> 0
          FOR FindCount = 1 TO DayCount
            IF Rec#(FindCount) = CatCodeRecord THEN
              RevAmt#(FindCount) = RevAmt#(FindCount) + RevAmt#
              RevAmt# = 0
              EXIT DO
            END IF
          NEXT FindCount
          DayCount = DayCount + 1
          Rec#(DayCount) = CatCodeRecord
          RevAmt#(DayCount) = RevAmt#
          RevAmt# = 0
        LOOP
      END IF
    END IF
    
  NEXT Cnt
  
  GOSUB DCProcessThisBunch
  
  'Mark Transactions as interfaced
  FOR Cnt = 1 TO FoundCnt
    GET #DCTransFile, TranInfo(Cnt).TranRecNo, DCTransRec(1)
    DCTransRec(1).GLInterfaced = "Y"
    PUT #DCTransFile, TranInfo(Cnt).TranRecNo, DCTransRec(1)
  NEXT
  CLOSE
  
DCSendExit:
  EXIT SUB
  
  
DCProcessThisBunch:
  ' Must Combine By Date and Then Do Cash Debit Entry For Total by Fund
  
  IF DayCount <= 0 THEN RETURN
  
  FundCnt = 0   ' Set Funds Used to Zero
  
  FOR Process = 1 TO DayCount
    GET #DCCatFile, Rec#(Process), DCCatCodeRec(1)
    
    acct$ = DCCatCodeRec(1).REVGLNUM
    acct$ = StripAcct$(acct$)
    acct$ = FmtAcct$(acct$, FundLen%, AcctLen%, DetLen%)
    
    GOSUB DCGetAcctName
    
    GJRec(1).AcctRec = 0
    GJRec(1).AcctNum = QPTrim$(acct$)
    GJRec(1).AcctName = AcctName$
    GJRec(1).TrDate = WorkDate
    GJRec(1).Ref = Ref$
    GJRec(1).CrAmt = RevAmt#(Process)
    GJRec(1).DrAmt = 0
    GJRec(1).EType = "C"
    GJRec(1).Desc = "FROM DECAL I/F"
    GJRec(1).Src = "CR"
    PUT #GJFile, , GJRec(1)
    
    'Add Up Fund Total Here for Cash Credit Entry
    IF FundCnt = 0 THEN
      FundCnt = 1
      Fund$(FundCnt) = LEFT$(acct$, FundLen%)
      FundAmt#(FundCnt) = RevAmt#(Process)
    ELSE
      FoundFund = 0
      FOR FCnt = 1 TO FundCnt
        IF Fund$(FCnt) = LEFT$(acct$, FundLen%) THEN
          FoundFund = 1
          FundAmt#(FCnt) = FundAmt#(FCnt) + RevAmt#(Process)
        END IF
      NEXT FCnt
      IF FoundFund = 0 THEN
        FundCnt = FundCnt + 1
        Fund$(FundCnt) = LEFT$(acct$, FundLen%)
        FundAmt#(FundCnt) = RevAmt#(Process)
      END IF
    END IF
    
  NEXT Process
  
  'Now Make Matching Debit Entries to Cash Account
  
  'Open GLSETUP File Here
  SetupFile = FREEFILE
  SetupRecLen = LEN(GLSetup(1))
  OPEN "GLSETUP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS SetupFile LEN = SetupRecLen
  GET SetupFile, 1, GLSetup(1)
  CASHACCT$ = GLSetup(1).CRCashAcct
  CLOSE SetupFile
  
  
  FOR Cash = 1 TO FundCnt
    acct$ = Fund$(Cash) + CASHACCT$
    acct$ = RTRIM$(acct$)
    acct$ = StripAcct$(acct$)
    acct$ = FmtAcct$(acct$, FundLen%, AcctLen%, DetLen%)
    
    GOSUB DCGetAcctName
    
    GJRec(1).AcctRec = 0
    GJRec(1).AcctNum = acct$
    GJRec(1).AcctName = AcctName$
    GJRec(1).TrDate = WorkDate
    GJRec(1).Ref = Ref$
    GJRec(1).DrAmt = FundAmt#(Cash)
    GJRec(1).CrAmt = 0
    GJRec(1).EType = "D"
    GJRec(1).Desc = "FROM DECAL I/F"
    GJRec(1).Src = "CR"
    PUT #GJFile, , GJRec(1)
  NEXT Cash
DCBunchReturn:
  RETURN
  
DCGetAcctName:
  OPEN "glacct.idx" FOR RANDOM ACCESS READ WRITE SHARED AS #13 LEN = 16
  FIELD 13, 14 AS AcctNumber$, 2 AS RecordNumber$
  AcctName$ = "UnDefined"
  FOR Cnt! = 1 TO LOF(13) / 16
    GET 13, Cnt!
    IF RTRIM$(AcctNumber$) = RTRIM$(acct$) THEN
      Record = CVI(RecordNumber$)
      ng = 1
      OPEN "glacct.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #12 LEN = 178
      FIELD 12, 16 AS nu$, 30 AS AName$
      GET 12, Record
      AcctName$ = AName$
      CLOSE 12
      CLOSE 13
      RETURN
    END IF
  NEXT Cnt!
  ng = 0
  CLOSE 12
  CLOSE 13
  RETURN
  
  
  
END SUB

SUB OpenIFEditFile (IFEditFileNum, NumIFTrans)
  
  IFEdLen = LEN(IFEdit)
  IFEditFileNum = FREEFILE
  OPEN "GLTRXED.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS IFEditFileNum LEN = IFEdLen
  NumIFTrans = LOF(IFEditFileNum) \ IFEdLen
END SUB

SUB PostTrans
  
  '--Paint a background
  FOR Cnt = 2 TO 24
    QPrintRC STRING$(80, CHR$(177)), Cnt, 1, 3
  NEXT
  
  '--verify that there are transactions and they are in balance.
  OpenIFEditFile IFEditFileNum, NumIFTrans
  
  '--summarize the file totals
  FOR Cnt = 1 TO NumIFTrans
    GET IFEditFileNum, Cnt, IFEdit
    IF NOT GJEdit.Deleted THEN
      Active = Active + 1
      TotDr# = TotDr# + Round#(IFEdit.DrAmt)
      TotCr# = TotCr# + Round#(IFEdit.CrAmt)
    END IF
  NEXT
  CLOSE
  
  '--if no active transactions tell user and get out
  IF Active = 0 THEN
    Ok = MsgBox%("GL.QSL", "NOTRANS")
    EXIT SUB
  END IF
  
  Ok = MsgBox%("GL.QSL", "POSTGJ")              'Ask user if sure ready to post
  
  IF Ok = 1 THEN EXIT SUB       '1=No
  
  TotDr# = 0    'init totals to zero
  TotCr# = 0
  Active = 0    'Counter for Active Transactions
  
  IF TotDr# <> TotCr# THEN      'Transactions out of balance
    Ok = MsgBox%("GL.QSL", "OUTOFBAL")          'ask user if ok to post
    IF Ok = 1 THEN EXIT SUB     'No = button 1
  END IF
  
  
  Active = 0    'Reset Active counter for posting
  
  OpenIFEditFile IFEditFileNum, NumIFTrans
  
  DIM Tr2Post AS GLTransRecType
  OPEN "GJ2POST.DAT" FOR RANDOM AS #2 LEN = LEN(Tr2Post)
  
  FOR Cnt = 1 TO NumIFTrans     'Assign edit file to trans format
    GET IFEditFileNum, Cnt, GJEdit
    IF NOT GJEdit.Deleted THEN
      Active = Active + 1
      Tr2Post.AcctRec = GJEdit.AcctRec
      Tr2Post.AcctNum = GJEdit.AcctNum
      Tr2Post.TrDate = GJEdit.TrDate
      Tr2Post.Desc = GJEdit.Desc
      Tr2Post.Ref = GJEdit.Ref
      Tr2Post.DrAmt = GJEdit.DrAmt
      Tr2Post.CrAmt = GJEdit.CrAmt
      Tr2Post.Src = "IF" + ConvDateStr$(DATE$)
      PUT #2, Active, Tr2Post
    END IF
  NEXT
  
  CLOSE
  
  Post2GL "GJ2POST.DAT", BadTrans%              'common post & link sub in GLAUX.BAS
  
  KILL "GLTRXED.DAT"            'kill the temp files
  KILL "GJ2POST.DAT"
  
  IF BadTrans <> 0 THEN         'posting problem
    CLS
    PRINT "Error: One or more transactions were not posted."
    PRINT "Press any key to view posting log"
    k$ = INPUT$(1)
    PrintRptFile "Posting Log", "GLUTIL.LOG", LPTNo, RetCode%, 2
  END IF
  
END SUB

SUB PrnEditList
  
  SHARED User$
  SHARED Choice$()

  REDIM Form$(0, 0)             'Holds data from form
  REDIM Fld(0) AS FieldInfo     'Field editing array
  REDIM Frm(1) AS FormInfo      'Dim the form info array
  
  REDIM Choice$(0 TO 3, 1)      'Choices for entry type field
  Choice$(0, 0) = "2"           'Form Field Number for picklist
  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn #1"
  Choice$(3, 0) = "Prn #2"
  
  LibLoadDisplayForm "GL.QSL", "PRNGJ", Form$(), Fld()
  PrintHelp "Print Distribution Register"
  'Set Bad Acct Flag to No Here
  BadAcctFlag = 0
  
  Action = 1
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key '--SaveButton
      '--Fld 2 = OutputDevice
      IF LEN(LTRIM$(RTRIM$(Form$(2, 0)))) = 0 THEN
        Dev$ = "S"
        LPTNo = 1
      ELSE
        Dev$ = LEFT$(Form$(2, 0), 1)
        LPTNo = VAL(RIGHT$(RTRIM$(Form$(2, 0)), 1))
      END IF
      
    CASE ESC
      EXIT SUB
      
    END SELECT
    
  LOOP UNTIL Frm(1).KeyCode = F10Key
  
  LOCATE 1, 1, 0
  
  '--Get a list of active funds
  GJRecLen = LEN(GJRec(1))
  
  REDIM FundList$(1)
  GetFundList FundList$(), NumFunds
  REDIM FundDr#(1 TO NumFunds)
  REDIM FundCr#(1 TO NumFunds)
  
  OpenIFEditFile IFEditFileNum, NumIFTrans
  
  PrnFileNum = FREEFILE
  PrnFileName$ = "GLTRXIF.PRN"
  OPEN PrnFileName$ FOR OUTPUT AS #PrnFileNum
  
  '--Report Variables
  MaxLines = 55
  PAGE = 0
  User$ = QPTrim$(User$)
  RptTitle$ = "General Ledger Interface Report"
  Header$ = "General Ledger Interface Report"
  CRLF$ = CHR$(13) + CHR$(10)
  FF$ = CHR$(12)
  
  GOSUB PrintGJHeader
  CommaFmt$ = "#########,.##"
  
  '--Start of printing loop
  FOR I = 1 TO NumIFTrans
    GET IFEditFileNum, I, GJEdit
    IF NOT GJEdit.Deleted THEN
      IF UCASE$(LEFT$(GJEdit.AcctName, 9)) = "UNDEFINED" THEN BadAcctFlag = 1
      '--First Line
      ToPrint$ = SPACE$(80)
      LSET ToPrint$ = Num2Date(GJEdit.TrDate)
      MID$(ToPrint$, 13) = GJEdit.Desc
      MID$(ToPrint$, 35) = GJEdit.Ref
      PRINT #PrnFileNum, ToPrint$
      LineCnt = LineCnt + 1
      
      '--2nd Line
      
      ToPrint$ = SPACE$(80)
      IF UCASE$(LEFT$(GJEdit.AcctName, 9)) = "UNDEFINED" THEN
        MID$(ToPrint$, 4) = "???"
      END IF
      MID$(ToPrint$, 13) = GJEdit.AcctNum
      MID$(ToPrint$, 27) = GJEdit.AcctName
      MID$(ToPrint$, 53) = FUsing$(STR$(GJEdit.DrAmt), CommaFmt$)
      MID$(ToPrint$, 68) = FUsing$(STR$(GJEdit.CrAmt), CommaFmt$)
      PRINT #PrnFileNum, ToPrint$
      LineCnt = LineCnt + 1
      
      '--3rd line (Blank)
      PRINT #PrnFileNum,
      LineCnt = LineCnt + 1
      IF LineCnt > MaxLines THEN
        PRINT #PrnFileNum, FF$
        GOSUB PrintGJHeader
      END IF
      
      '--Sum total debits and credits
      TotDr# = TotDr# + GJEdit.DrAmt
      TotCr# = TotCr# + GJEdit.CrAmt
      
      '--Sum into proper fund
      Found = False
      FOR Fund = 1 TO NumFunds
        FundNum$ = LEFT$(GJEdit.AcctNum, FundLen)
        IF FundNum$ = FundList$(Fund) THEN
          Found = True
          FundDr#(Fund) = FundDr#(Fund) + Round#(GJEdit.DrAmt)
          FundCr#(Fund) = FundCr#(Fund) + Round#(GJEdit.CrAmt)
          EXIT FOR
        END IF
      NEXT
    END IF
  NEXT
  
  PRINT #PrnFileNum, STRING$(80, "-")
  LineCnt = LineCnt + 1
  IF LineCnt > MaxLines THEN
    PRINT #PrnFileNum, FF$
    GOSUB PrintGJHeader
  END IF
  
  ToPrint$ = SPACE$(80)
  LSET ToPrint$ = "File Totals"
  MID$(ToPrint$, 53) = FUsing$(STR$(TotDr#), CommaFmt$)
  MID$(ToPrint$, 68) = FUsing$(STR$(TotCr#), CommaFmt$)
  PRINT #PrnFileNum, ToPrint$
  LineCnt = LineCnt + 1
  IF LineCnt > MaxLines THEN
    PRINT #PrnFileNum, FF$
    GOSUB PrintGJHeader
  END IF
  
  '--Print Summary by fund
  TranCashTot# = 0
  FundOutofBal = False
  FOR Fund = 1 TO NumFunds
    IF FundDr#(Fund) <> 0 OR FundCr#(Fund) <> 0 THEN
      IF Round(FundDr#(Fund)) <> Round(FundCr#(Fund)) THEN FundOutofBal = True
      ToPrint$ = SPACE$(80)
      MID$(ToPrint$, 4) = "Fund# " + FundList$(Fund)
      MID$(ToPrint$, 53) = FUsing$(STR$(FundDr#(Fund)), CommaFmt$)
      MID$(ToPrint$, 68) = FUsing$(STR$(FundCr#(Fund)), CommaFmt$)
      PRINT #PrnFileNum, ToPrint$
      LineCnt = LineCnt + 1
      IF LineCnt > MaxLines THEN
        PRINT #PrnFileNum, FF$
        GOSUB PrintGJHeader
      END IF
    END IF
    
  NEXT
  
  IF FundOutofBal THEN
    '--skip a line
    ToPrint$ = SPACE$(80)
    PRINT #PrnFileNum, ToPrint$
    LineCnt = LineCnt + 1
    '--Tell user they're screwing up
    PRINT #PrnFileNum, "WARNING:"
    PRINT #PrnFileNum, "Entries are not in balance!"
    PRINT #PrnFileNum, "File WILL NOT POST Due to Bad Account Number"
    PRINT #PrnFileNum, "Please Transfer to the General Journal and Correct"
    LineCnt = LineCnt + 4
  END IF
  IF BadAcctFlag = 1 THEN
    PRINT #PrnFileNum, ""
    PRINT #PrnFileNum, "WARNING:"
    PRINT #PrnFileNum, "File WILL NOT POST Due to Bad Account Number('s)"
    PRINT #PrnFileNum, "Please Transfer to the General Journal and Correct"
    PRINT #PrnFileNum, "Look for ???"
    LineCnt = LineCnt + 4
  END IF
  
  PRINT #PrnFileNum, FF$
  
  CLOSE
  
  SELECT CASE Dev$
  CASE "S"
    EntryPoint = 2
  CASE "P"
    EntryPoint = 5
  END SELECT
  
  PrintRptFile RptTitle$, PrnFileName$, LPTNo, RetCode%, EntryPoint
  KILL PrnFileName$
  EXIT SUB
  
  
PrintGJHeader:
  PAGE = PAGE + 1
  PRINT #PrnFileNum, TAB(40 - (INT(LEN(User$) / 2))); User$
  PRINT #PrnFileNum, TAB(40 - (INT(LEN(Header$) / 2))); Header$
  PRINT #PrnFileNum,
  PRINT #PrnFileNum, "Report Date: "; DATE$; TAB(67); "Page #"; PAGE
  PRINT #PrnFileNum, "Date        Description           Reference"
  PRINT #PrnFileNum, "            G/L Account                                     Debit         Credit"
  PRINT #PrnFileNum, "--------------------------------------------------------------------------------"
  LineCnt = 5
  RETURN
  
  
END SUB

SUB TRGetAcctStruct (GLFundLen%, GLAcctLen%, GLDetLen%)
  
  REDIM GLSetUpRec(1) AS GLSetupRecType
  SetupRecLen = LEN(GLSetUpRec(1))
  
  SetupFile = FREEFILE
  OPEN "GLSETUP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS SetupFile LEN = SetupRecLen
  GET SetupFile, 1, GLSetUpRec(1)
  
  GLFundLen = GLSetUpRec(1).FundLen
  GLAcctLen = GLSetUpRec(1).AcctLen
  GLDetLen = GLSetUpRec(1).DetLen
  
  CLOSE SetupFile
  ERASE GLSetUpRec
END SUB

SUB Trxfer2GJ
  
  GJRecLen = LEN(GJRec(1))
  GJFile = FREEFILE
  OPEN "GLGJED.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS GJFile LEN = GJRecLen
  NumEdTrans = LOF(GJFile) \ GJRecLen
  CLOSE
  
  IF NumEdTrans > 0 THEN
    Ok = MsgBox%("GLUTIL", "GJEDZERO")
    GOTO GJSendExit
  END IF
  
  GJRecLen = LEN(GJRec(1))
  GJFile = FREEFILE
  OPEN "GLTRXED.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS GJFile LEN = GJRecLen
  NumEdTrans = LOF(GJFile) \ GJRecLen
  CLOSE
  
  IF NumEdTrans = 0 THEN
    Ok = MsgBox%("GLUTIL", "GLNOTRAN")
    GOTO GJSendExit
  END IF

  'CLEAR

  'OK to Rename
  DF$ = "Copy GLTRXED.DAT GLGJED.DAT"
  SHELL DF$
  DF$ = "DEL GLTRXED.DAT"
  SHELL DF$
  Ok = MsgBox%("GLUTIL", "TRXOK")
  
GJSendExit:
END SUB

