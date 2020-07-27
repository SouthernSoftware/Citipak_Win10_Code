DEFINT A-Z
DECLARE SUB LoadUBSetUpFile (UBSetUpFileNum%, UBSetUpRecLen%)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Exist (FileName$)
DECLARE FUNCTION FLof& (FileHandle%)
DECLARE SUB FClose (Handle%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FOpenAll (FileName$, AccessMode%, SharedMode%, Handle%)
DECLARE SUB FGetRTA (Handle, UBSetUpRec(), A&, UBSetUpLen%)
DECLARE SUB EdIFEntry ()
DECLARE SUB ExtractCM (ThruDate%)
DECLARE SUB ExtractUB (ThruDate%)
DECLARE SUB ExtractTX ()
DECLARE SUB ExtractBL (ThruDate%)
DECLARE SUB ExtractDC (ThruDate%)
DECLARE SUB Trxfer2GJ ()
DECLARE SUB IFInit ()
DECLARE SUB OpenIFEditFile (IFEditFileNum%, NumIFTrans%)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB BLPrint (LPTPORT%, Text$, ErrCode%)
DECLARE SUB FileView (FileName$, Ky, Action, FVI AS ANY, SEG Array)
DECLARE SUB HideCursor ()
DECLARE SUB MScrnSave (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MScrnRest (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MPaintBox (ULRow%, ULCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB QPrint (X$, Colr%, Page%)
DECLARE SUB QPrintRC (t$, R%, c%, Clr%)
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortStr2 (BYVAL Address%, Size%, Direction%)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE SUB ShowCursor ()
DECLARE SUB MnuTitle (Title$, Row, Col, Wdth, TxtColor)
DECLARE SUB ClrBG ()
DECLARE SUB Post2GL (FileName$, BadTrans%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE FUNCTION QPStrL$ (number&)
DECLARE FUNCTION QPStrI$ (number%)
DECLARE SUB TRGetAcctStruct (GLFundLen%, GLAcctLen%, GLDetLen%)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE FUNCTION FUsing$ (number$, Mask$)
DECLARE FUNCTION MGetKey% (Row%, Col%)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION Peek1% (Segment%, Address%)
DECLARE FUNCTION QPTrim$ (X$)
DECLARE FUNCTION WaitKey% ()
DECLARE FUNCTION Value# (E$, ErrCode%)
DECLARE FUNCTION FldNum% (FldName$, Fld() AS ANY)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPORT%, RetCode%, EntryPoint%)

'$INCLUDE: 'DefCnf.bi'                    'monitor and color info
'$INCLUDE: 'FieldInf.bi'                  'field information
'$INCLUDE: 'FormEdit.bi'                  'form editing information
'$INCLUDE: 'QSCR.BI'                      'QuickScreen Routines
'$INCLUDE: 'GLAUX.BI'                     'Common GL Routines
'$INCLUDE: 'GL.BI'                        'GL File Types
'$INCLUDE: 'cmfiles.bi'
'$INCLUDE: 'ubsetup.bi'
'$INCLUDE: 'ubtrans.bi'
'$INCLUDE: 'AR.bi'                        'A/R FILE LAYOUTS
'$INCLUDE: 'DC.bi'                        'Decal File Layouts
'$INCLUDE: 'SCRAUX.BI'

CONST False = 0, True = NOT False

DECLARE SUB PrnEditList ()
DECLARE SUB DisplayTotals ()
DECLARE SUB PostTrans ()
DECLARE SUB GetFundList (FundList$(), NumFunds%)
DECLARE SUB VertMenuT (Items() AS ANY, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB GetPostDates (LPDate, HPEndDate)
DECLARE FUNCTION EDCheck4Acct% (Form$(), AcctFld%)
DECLARE FUNCTION GetNextRec% (FileNum%, NumRecs%, CurrRec%, NextRec%)


TYPE FLen
  v AS STRING * 42
END TYPE

'REDIM SHARED Choice$(0, 0)                'Used to hold multiple choices

  TYPE TranRecInfoType
    TranDate  AS INTEGER
    TranRecNo AS LONG
  END TYPE




DIM SHARED MCodeRec(1) AS MiscCodeRecType
DIM SHARED GlSetUp(1) AS GLSetupRecType
DIM SHARED GJRec(1) AS TrEditRecType
DIM SHARED GJEdit(1) AS TrEditRecType
DIM SHARED GJEdit AS TrEditRecType
DIM SHARED IFEdit AS TrEditRecType

'$INCLUDE: 'SetCnf.bi'        'Set monitor and color information
'$INCLUDE: 'init.bi'

'changedate
'setfieldvalue

  LOCATE 1, 1, 0
  TextCursor -2, -2             'Define the Mouse cursor
  ShowCursor    'Show it

'***Master Menu

   DIM mChoice$(1 TO 6)

   mChoice$(1) = "Grab Transactions"
   mChoice$(2) = "Print Journal Register"
   mChoice$(3) = "Post Journal Entries"
   mChoice$(4) = "Transfer to General Journal"
   mChoice$(5) = "Initialize Transactions for I/F"
   mChoice$(6) = "Quit"

   MaxLen = 0                           'Set menu width to zero
   BoxBot = 20                            'limit the box to line 20
   Action = 0                             '0 means stay in the menu
   Choice = 1                             'Pre-load choice to highlight

   FOR Cnt = 1 TO UBOUND(mChoice$)        'Find max menu width
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

FUNCTION EDCheck4Acct% (Form$(), AcctFld%)

  List$ = SPACE$(65)

  Lookup$ = QPTrim$(UCASE$(Form$(AcctFld, 0)))
  Lookuplen = LEN(Lookup$)

  REDIM ScrBuf(2000)
  MScrnSave 1, 1, 25, 80, ScrBuf(0)             'Save the existing screen

  OpenAcctIdx AcctIdxFileNum, NumIdxRecs
  CLOSE AcctIdxFileNum
  REDIM TempAcctIdx(1 TO NumIdxRecs) AS GLAcctIndexType

  ElSize = LEN(TempAcctIdx(1))
  FGetAH "GLACCT.IDX", TempAcctIdx(1), ElSize, NumIdxRecs

  OpenAcctFile AcctFileNum, NumAccts

  REDIM AcctList(1 TO NumIdxRecs) AS FLen       'Dim array to hold trans list
  FOR Cnt = 1 TO NumIdxRecs     'Read the file
    IF LEFT$(TempAcctIdx(Cnt).AcctNum, Lookuplen) = Lookup$ THEN
      GET AcctFileNum, TempAcctIdx(Cnt).RecNum, Acct
      MatchAccts = MatchAccts + 1
      LSET AcctList(MatchAccts).v = List$
      AcctList(MatchAccts).v = TempAcctIdx(Cnt).AcctNum
      'MID$(AcctList(MatchAccts).v, TotLen + 4) = Acct.Title
      MID$(AcctList(MatchAccts).v, 41) = MKI$(TempAcctIdx(Cnt).RecNum)
    END IF
  NEXT

  CLOSE AcctFileNum

  SELECT CASE MatchAccts
  CASE 0
    PLAY "L32O3EC"              'Acct NOT found
    Ok = MsgBox("GL.QSL", "NOACCT")
    EDCheck4Acct = False
  CASE 1
    AcctRecNum = CVI(RIGHT$(AcctList(1).v, 2))
    GOSUB GetAcctInfo
    EDCheck4Acct = True
  CASE ELSE
    Row = 5
    Col = 15
    LOCATE Row, Col
    Action = 0
    MaxLen = 39
    BoxBot = 18
    ListDescTxt$ = "Account No.   Description"
    ListDesc$ = SPACE$(MaxLen + 4)
    MID$(ListDesc$, 3) = ListDescTxt$
    QPrintRC ListDesc$, Row - 1, Col, 112       'Cnf.ActivCh
    DO          'call vertmenu to display matching recs to choose from
      VertMenuT AcctList(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        EDCheck4Acct = False
        EXIT DO
      END IF
      IF Choice > 0 THEN
        AcctRecNum = CVI(RIGHT$(AcctList(Choice).v, 2))
        GOSUB GetAcctInfo
        EDCheck4Acct = True
        EXIT DO
      END IF
    LOOP
  END SELECT

  MScrnRest 1, 1, 25, 80, ScrBuf(0)             'Restore Screen

  ERASE AcctList, TempAcctIdx, ScrBuf


  EXIT FUNCTION

GetAcctInfo:
  OpenAcctFile AcctFileNum, NumAccts
  GET AcctFileNum, AcctRecNum, Acct
  CLOSE AcctFileNum
  'LSET Form$(AcctFld, 0) = Acct.Num
  'LSET Form$(AcctFld + 1, 0) = Acct.Title
'  LSET Form$(AcctFld + 2, 0) = MKI$(AcctRecNum)
RETURN



END FUNCTION

SUB EdIFEntry

   GetPostDates LPDate, HPDate

   SHARED Choice$()                          'Multiple choice array
   REDIM Form$(0, 0)                         'Holds data from form
   REDIM Fld(1) AS FieldInfo                 'Field editing array

   REDIM Choice$(0 TO 2, 1)                  'Choices for entry type field
   Choice$(0, 0) = "5"                       'Field Number
   Choice$(1, 0) = "Debit "
   Choice$(2, 0) = "Credit"


   OpenIFEditFile IFEditFileNum, NumIFTrans  'Opens the GJ Edit File


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

   DIM Frm AS FormInfo                       'Dim the form info array

   CLOSE                                     'Close file while doing input

   DO                                        'Edit the form

      EditForm Form$(), Fld(), Frm, Cnf, Action

      SELECT CASE Frm.KeyCode                'Respond to the key presses
         CASE F10Key  'Save

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
                 IF Form$(1, 0) = "TX" THEN ExtractTX: EXIT SUB
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
  DIM Rec#(500), RevAmt#(500), TType%(500), Fund$(100), FundAmt#(100)

  Dash80$ = STRING$(80, "-")
  P2S$ = SPACE$(4)


  TRGetAcctStruct FundLen%, AcctLen%, DetLen%
  GJRecLen = LEN(GJRec(1))


  RPTFile = FREEFILE
  OPEN "GLCMTRX.RPT" FOR OUTPUT AS RPTFile


  GOSUB BLClearBox
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

    IF LEN(QPTrim$(ARTransRec(1).Posted2GL)) = 0 OR QPTrim$(ARTransRec(1).Posted2GL) = "N" THEN
     'Store trans rec numbers and dates in array
      FoundCnt = FoundCnt + 1
      REDIM PRESERVE TranInfo(FoundCnt) AS TranRecInfoType
      TranInfo(FoundCnt).TranDate = ARTransRec(1).TransDate
      TranInfo(FoundCnt).TranRecNo = TCnt&
    ELSE
      NGCnt = NGCnt + 1
    END IF
      RSET P2S$ = QPStrI$(FoundCnt)
      QPrintRC P2S$, 13, 47, Cnf.HiLite
     'SmallPause
     'Allow 500 Bad Entries Before Exiting
    IF NGCnt >= 500 THEN EXIT FOR
  NEXT

  IF FoundCnt = 0 THEN
    CLOSE
    GOSUB BLClearBox
    PRINT CHR$(7);
    QPrintRC "No Transactions Found To InterFace", 12, 24, 126
    SLEEP 4
    GOTO BLSendExit
  END IF

 'Array (1), NumElem, Dir, StructSize, MemOff, MemSize
  SortT TranInfo(1), FoundCnt, 0, 6, 0, -1  'sort'em by date. oldest first



 'Open GL InterFace File
  GJFile = FREEFILE
  OPEN "GLTRXED.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS GJFile LEN = GJRecLen
  NumEdTrans = LOF(GJFile) \ GJRecLen
  LOCK #GJFile

 'OPEN Business License CODES Here

  REDIM ARCatCodeRec(1) AS ARCatCodeRecType
  ARCatCodeRecLen = LEN(ARCatCodeRec(1))
  ARCatFile = FREEFILE
  OPEN "ARCODE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARCatFile LEN = ARCatCodeRecLen
  NumOFARCatRecs = LOF(ARCatFile) \ ARCatCodeRecLen


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
                TType%(DayCount) = ARTransRec(1).TransType
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
            TType%(DayCount) = ARTransRec(1).TransType
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

FundCnt = 0                     ' Set Funds Used to Zero


'Process Payments First Type=2

FOR Process = 1 TO DayCount
 IF TType%(Process) = 2 THEN
  GET #ARCatFile, Rec#(Process), ARCatCodeRec(1)

      IF LEN(QPTrim$(ARCatCodeRec(1).ARGLAcct)) > 0 THEN
       Acct$ = ARCatCodeRec(1).ARGLAcct
       Acct$ = StripAcct$(Acct$)
       Acct$ = FmtAcct$(Acct$, FundLen%, AcctLen%, DetLen%)
        ELSE
       Acct$ = ARCatCodeRec(1).RevGLNum
       Acct$ = StripAcct$(Acct$)
       Acct$ = FmtAcct$(Acct$, FundLen%, AcctLen%, DetLen%)
      END IF

      GOSUB BLGetAcctName

      GJRec(1).AcctRec = 0
      GJRec(1).AcctNum = QPTrim$(Acct$)
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
   SetUpRecLen = LEN(GlSetUp(1))
   OPEN "GLSETUP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS SetupFile LEN = SetUpRecLen
   GET SetupFile, 1, GlSetUp(1)
   CashAcct$ = GlSetUp(1).CRCashAcct
   CLOSE SetupFile

   
   
       Acct$ = LEFT$(Acct$, FundLen%) + CashAcct$
       Acct$ = QPTrim$(Acct$)
       Acct$ = StripAcct$(Acct$)
       Acct$ = FmtAcct$(Acct$, FundLen%, AcctLen%, DetLen%)

       GOSUB BLGetAcctName

       GJRec(1).AcctRec = 0
       GJRec(1).AcctNum = Acct$
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
 IF TType%(Process) = 1 THEN
  GET #ARCatFile, Rec#(Process), ARCatCodeRec(1)
    IF LEN(QPTrim$(ARCatCodeRec(1).ARGLAcct)) > 0 THEN
      Acct$ = ARCatCodeRec(1).RevGLNum
      Acct$ = StripAcct$(Acct$)
      Acct$ = FmtAcct$(Acct$, FundLen%, AcctLen%, DetLen%)

      GOSUB BLGetAcctName

      GJRec(1).AcctRec = 0
      GJRec(1).AcctNum = QPTrim$(Acct$)
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


       Acct$ = ARCatCodeRec(1).ARGLAcct
       Acct$ = StripAcct$(Acct$)
       Acct$ = FmtAcct$(Acct$, FundLen%, AcctLen%, DetLen%)

       GOSUB BLGetAcctName

       GJRec(1).AcctRec = 0
       GJRec(1).AcctNum = Acct$
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
       IF RTRIM$(AcctNumber$) = RTRIM$(Acct$) THEN
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

BLClearBox:
  'Clear Box
  FOR SLine = 11 TO 14: LOCATE SLine, 20: PRINT STRING$(40, 32): NEXT SLine
  FOR SLine = 16 TO 16: LOCATE SLine, 20: PRINT STRING$(41, 32): NEXT SLine
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
  
  GOSUB ClearBox
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
    GOSUB ClearBox
    PRINT CHR$(7);
    QPrintRC "No Transactions Found To InterFace", 12, 24, 126
    SLEEP 4
    GOTO SendExit
  END IF

  SortT TranInfo(1), FoundCnt, 0, 6, 0, -1  'sort'em by date. oldest first
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

FundCnt = 0                     ' Set Funds Used to Zero

FOR Process = 1 TO DayCount
 IF MiscRec#(Process) >= 1 THEN
  GET #MCFile, MiscRec#(Process), MCodeRec(1)
      Acct$ = MCodeRec(1).GlAcctNumb
      Acct$ = StripAcct$(Acct$)
      Acct$ = FmtAcct$(Acct$, FundLen%, AcctLen%, DetLen%)

      GOSUB GetAcctName
      GJRec(1).AcctRec = 0
      GJRec(1).AcctNum = QPTrim$(Acct$)
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
        Fund$(FundCnt) = LEFT$(Acct$, FundLen%)
        FundAmt#(FundCnt) = MiscAmt#(Process)
     ELSE
        FoundFund = 0
          FOR FCnt = 1 TO FundCnt
            IF Fund$(FCnt) = LEFT$(Acct$, FundLen%) THEN
              FoundFund = 1
              FundAmt#(FCnt) = FundAmt#(FCnt) + MiscAmt#(Process)
            END IF
          NEXT FCnt
        IF FoundFund = 0 THEN
          FundCnt = FundCnt + 1
          Fund$(FundCnt) = LEFT$(Acct$, FundLen%)
          FundAmt#(FundCnt) = MiscAmt#(Process)
        END IF
      END IF
  END IF
  NEXT Process

   'Now Make Matching Debit Entries to Cash Account
 'Open GLSETUP File Here
  SetupFile = FREEFILE
  SetUpRecLen = LEN(GlSetUp(1))
  OPEN "GLSETUP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS SetupFile LEN = SetUpRecLen
  GET SetupFile, 1, GlSetUp(1)
  CashAcct$ = GlSetUp(1).CRCashAcct
  CLOSE SetupFile


   FOR Cash = 1 TO FundCnt
       Acct$ = Fund$(Cash) + CashAcct$
       Acct$ = QPTrim$(Acct$)
       Acct$ = StripAcct$(Acct$)
       Acct$ = FmtAcct$(Acct$, FundLen%, AcctLen%, DetLen%)
       GOSUB GetAcctName

       GJRec(1).AcctRec = 0
       GJRec(1).AcctNum = Acct$
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
     FOR Cnt! = 1 TO LOF(13) / 16: GET 13, Cnt!
       IF RTRIM$(AcctNumber$) = RTRIM$(Acct$) THEN
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

ClearBox:
  'Clear Box
  FOR SLine = 11 TO 14: LOCATE SLine, 20: PRINT STRING$(40, 32): NEXT SLine
  FOR SLine = 16 TO 16: LOCATE SLine, 20: PRINT STRING$(41, 32): NEXT SLine
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

  
  GOSUB DCClearBox
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
    GOSUB DCClearBox
    PRINT CHR$(7);
    QPrintRC "No Transactions Found To InterFace", 12, 24, 126
    SLEEP 4
    GOTO DCSendExit
  END IF

 'Array (1), NumElem, Dir, StructSize, MemOff, MemSize
  SortT TranInfo(1), FoundCnt, 0, 6, 0, -1  'sort'em by date. oldest first




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
   CatCode$ = QPTrim$(DCTransRec(1).DecalCat)
   CatCodeRecord = 0

   FOR CCnt! = 1 TO NumOFDCCatRecs
   GET DCCatFile, CCnt!, DCCatCodeRec(1)
    IF QPTrim$(DCCatCodeRec(1).CatCode) = CatCode$ THEN CatCodeRecord = CCnt!: EXIT FOR
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

FundCnt = 0                     ' Set Funds Used to Zero

FOR Process = 1 TO DayCount
  GET #DCCatFile, Rec#(Process), DCCatCodeRec(1)

      Acct$ = DCCatCodeRec(1).RevGLNum
      Acct$ = StripAcct$(Acct$)
      Acct$ = FmtAcct$(Acct$, FundLen%, AcctLen%, DetLen%)

      GOSUB DCGetAcctName

      GJRec(1).AcctRec = 0
      GJRec(1).AcctNum = QPTrim$(Acct$)
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
        Fund$(FundCnt) = LEFT$(Acct$, FundLen%)
        FundAmt#(FundCnt) = RevAmt#(Process)
     ELSE
        FoundFund = 0
          FOR FCnt = 1 TO FundCnt
            IF Fund$(FCnt) = LEFT$(Acct$, FundLen%) THEN
              FoundFund = 1
              FundAmt#(FCnt) = FundAmt#(FCnt) + RevAmt#(Process)
            END IF
          NEXT FCnt
        IF FoundFund = 0 THEN
          FundCnt = FundCnt + 1
          Fund$(FundCnt) = LEFT$(Acct$, FundLen%)
          FundAmt#(FundCnt) = RevAmt#(Process)
        END IF
      END IF

  NEXT Process

   'Now Make Matching Debit Entries to Cash Account

 'Open GLSETUP File Here
  SetupFile = FREEFILE
  SetUpRecLen = LEN(GlSetUp(1))
  OPEN "GLSETUP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS SetupFile LEN = SetUpRecLen
  GET SetupFile, 1, GlSetUp(1)
  CashAcct$ = GlSetUp(1).CRCashAcct
  CLOSE SetupFile


   FOR Cash = 1 TO FundCnt
       Acct$ = Fund$(Cash) + CashAcct$
       Acct$ = RTRIM$(Acct$)
       Acct$ = StripAcct$(Acct$)
       Acct$ = FmtAcct$(Acct$, FundLen%, AcctLen%, DetLen%)
       
       GOSUB DCGetAcctName

       GJRec(1).AcctRec = 0
       GJRec(1).AcctNum = Acct$
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
     FOR Cnt! = 1 TO LOF(13) / 16: GET 13, Cnt!
       IF RTRIM$(AcctNumber$) = RTRIM$(Acct$) THEN
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

DCClearBox:
  'Clear Box
  FOR SLine = 11 TO 14: LOCATE SLine, 20: PRINT STRING$(40, 32): NEXT SLine
  FOR SLine = 16 TO 16: LOCATE SLine, 20: PRINT STRING$(41, 32): NEXT SLine
  RETURN
  

END SUB

SUB ExtractTX
END SUB

SUB ExtractUB (ThruDate%)

  Today$ = DATE$
  Ref$ = "UB" + LEFT$(Today$, 2) + MID$(Today$, 4, 2) + RIGHT$(Today$, 2)

  REDIM TranInfo(1) AS TranRecInfoType

  Dash80$ = STRING$(80, "-")
  P2S$ = SPACE$(4)


  REDIM GJRec1(1 TO 2) AS TrEditRecType
  GJRecLen = LEN(GJRec1(1))
  GJFile = FREEFILE
  OPEN "GLTRXED.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS GJFile LEN = GJRecLen
  NumEdTrans = LOF(GJFile) \ GJRecLen
  CLOSE

  IF NumEdTrans > 0 THEN
    Ok = MsgBox%("UBSENDGL", "JEDNZERO")
    GOTO SendExitUB
  END IF

  Ok = MsgBox%("UBSENDGL", "MUSTEXT2")
  SELECT CASE Ok
  CASE 2
    GOTO SendExitUB
  END SELECT

  REDIM UBSetUpRec(1) AS UBSetUpRecType
  LoadUBSetUpFile UBSetUpFileNum, UBSetUpLen
  GET UBSetUpFileNum, 1, UBSetUpRec(1)
  AcctMeth$ = QPTrim$(UBSetUpRec(1).MethAcct)
  IF (LEN(AcctMeth$) = 0) THEN
    Ok = MsgBox%("UBSENDGL", "NOMETHOD")
    GOTO SendExitUB
  END IF

  SELECT CASE AcctMeth$
  CASE "C"
    InterfaceMethod = 1
  CASE "A"
    InterfaceMethod = 2
  CASE ELSE
    GOTO SendExitUB
  END SELECT

  RPTFile = FREEFILE
  OPEN "UBNOTFND.RPT" FOR OUTPUT AS RPTFile
  GOSUB NotFoundHeader

  'ShowProcessingScrn "Verifying GL Transfer Accounts"

  FOR RevCnt = 1 TO MaxRevsCnt
    TempRev$ = QPTrim$(UBSetUpRec(1).Revenues(RevCnt).REVNAME)
    IF LEN(TempRev$) = 0 THEN
      NumOfRevs = RevCnt - 1
      EXIT FOR
    ELSE
      REDIM PRESERVE GJInfo(1 TO RevCnt) AS GJXferRecType
      GJInfo(RevCnt).RevText = TempRev$
      GJInfo(RevCnt).BAcctInfo.DAcctNo = UBSetUpRec(1).BillAcct(RevCnt).DebitAcct
      GJInfo(RevCnt).BAcctInfo.CAcctNo = UBSetUpRec(1).BillAcct(RevCnt).CreditAcct
      GJInfo(RevCnt).pAcctInfo.DAcctNo = UBSetUpRec(1).PayAcct(RevCnt).DebitAcct
      GJInfo(RevCnt).pAcctInfo.CAcctNo = UBSetUpRec(1).PayAcct(RevCnt).CreditAcct
      IF UBSetUpRec(1).Revenues(RevCnt).UseDep = "Y" THEN
        GJInfo(RevCnt).dacctInfo.DAcctNo = UBSetUpRec(1).DepAcct(RevCnt).DebitAcct
        GJInfo(RevCnt).dacctInfo.CAcctNo = UBSetUpRec(1).DepAcct(RevCnt).CreditAcct
      END IF
    END IF
  NEXT

  'check to see if they are valid GL accounts
  GOSUB ValidateGLAccounts

  IF BadAcct THEN
    GOTO SendExitUB
  END IF

  GOSUB ClearBoxUB
  QPrintRC "Searching Cash Transactions.", 11, 26, 126
  QPrintRC "New Transactions:", 13, 29, Cnf.HiLite

  REDIM UBTransRec(1)  AS UBTransRecType
  UBTransRecLen = LEN(UBTransRec(1))

  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTransRecLen
  NumOfTRecs& = LOF(UBTran) \ UBTransRecLen

  FOR TCnt& = NumOfTRecs& TO 1 STEP -1
    GET #UBTran, TCnt&, UBTransRec(1)
    IF LEN(QPTrim$(UBTransRec(1).Posted2GL)) = 0 OR QPTrim$(UBTransRec(1).Posted2GL) = "N" THEN
      'Store trans rec numbers and dates in array
      IF UBTransRec(1).TransDate <= ThruDate% THEN
      FoundCnt = FoundCnt + 1
      REDIM PRESERVE TranInfo(FoundCnt) AS TranRecInfoType
      TranInfo(FoundCnt).TranDate = UBTransRec(1).TransDate
      TranInfo(FoundCnt).TranRecNo = TCnt&
      END IF
    ELSE
      NGCnt = NGCnt + 1
    END IF
    RSET P2S$ = QPStrI$(FoundCnt)
    QPrintRC P2S$, 13, 47, Cnf.HiLite
    'SmallPause
    IF NGCnt >= 2500 THEN EXIT FOR
  NEXT

  IF FoundCnt = 0 THEN
    CLOSE
    Ok = MsgBox%("UBSENDGL", "NONEFOND")
    GOTO SendExitUB
  END IF

  SortT TranInfo(1), FoundCnt, 0, 6, 0, -1  'sort'em by date. oldest first
  'Array(1), NumElem, Dir, StructSize, MemOff, MemSize


  GJFile = FREEFILE
  OPEN "GLTRXED.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS GJFile LEN = GJRecLen


  FirstTran = 1
  ThisDate = TranInfo(1).TranDate
  FOR Cnt = 1 TO FoundCnt
    IF ThisDate <> TranInfo(Cnt).TranDate THEN
      ThisDate = TranInfo(Cnt).TranDate
      LastTran = Cnt - 1
      GOSUB ProcessThisBunchUB
      FirstTran = Cnt
    END IF
  NEXT
  LastTran = FoundCnt
  GOSUB ProcessThisBunchUB

  'transactions as interfaced
  
  FOR Cnt = 1 TO FoundCnt
    GET #UBTran, TranInfo(Cnt).TranRecNo, UBTransRec(1)
    UBTransRec(1).Posted2GL = "Y"
    PUT #UBTran, TranInfo(Cnt).TranRecNo, UBTransRec(1)
    
    
  NEXT
  CLOSE
  
  
  SLEEP 2
SendExitUB:

EXIT SUB

NotFoundHeader:
  PageNo = PageNo + 1
  PRINT #RPTFile, "Utility Billing GL Transfer Invalid Account Listing."; TAB(70); "Page:"; PageNo
  PRINT #RPTFile, TownName$
  PRINT #RPTFile, "Report Date: "; DATE$
  PRINT #RPTFile, "Revenue           Acct. Type              Debit Acct.          Credit Acct."
  PRINT #RPTFile, Dash80$
  NumPrinted = 0
RETURN

PrintBadAcct:
    IF LEN(QPTrim$(BadCAcct$)) = 0 THEN
      BadCAcct$ = "Undefined"
    END IF

    PRINT #RPTFile, GJInfo(RevCnt).RevText;

    SELECT CASE ActT
    CASE 1
      ActPg$ = "Billing"
    CASE 2
      ActPg$ = "Payment"
    CASE 3
      ActPg$ = "Deposit"
    END SELECT
    PRINT #RPTFile, TAB(22); ActPg$;
    PRINT #RPTFile, TAB(43); BadDAcct$; TAB(64); BadCAcct$
RETURN

ProcessThisBunchUB:
  FOR RevCnt = 1 TO NumOfRevs
    GJInfo(RevCnt).BAcctInfo.CreditAmt = 0
    GJInfo(RevCnt).BAcctInfo.DebitAmt = 0
    GJInfo(RevCnt).pAcctInfo.CreditAmt = 0
    GJInfo(RevCnt).pAcctInfo.DebitAmt = 0
    GJInfo(RevCnt).dacctInfo.CreditAmt = 0
    GJInfo(RevCnt).dacctInfo.DebitAmt = 0
  NEXT

  FOR PCnt = FirstTran TO LastTran
    IF PCnt = FirstTran THEN
      WorkDate = TranInfo(PCnt).TranDate
    END IF
    GET #UBTran, TranInfo(PCnt).TranRecNo, UBTransRec(1)

    SELECT CASE InterfaceMethod
    CASE 1  'Cash Central
      SELECT CASE UBTransRec(1).TransType
        CASE TranUtilityBill          ' 1=Utility bill
          'no action
        CASE TranLateCharge           ' 2=late charge
          'no action
        CASE TranReconnectFee         ' 3=reconnect fee
          'no action
        CASE TranBillPayment          ' 4=Bill Payment
          FOR RevCnt = 1 TO NumOfRevs
            GJInfo(RevCnt).pAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).pAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
            GJInfo(RevCnt).pAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).pAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
          NEXT
        CASE TranAppliedDeposit       ' 5=Applied Deposit
          'no action
        CASE TranPenaltyCharge        ' 6=Penalty Charge
          'no action
        CASE TranDepositPayment       ' 7=Deposit Payment
          FOR RevCnt = 1 TO NumOfRevs
            GJInfo(RevCnt).dacctInfo.CreditAmt = Round#(GJInfo(RevCnt).dacctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
            GJInfo(RevCnt).dacctInfo.DebitAmt = Round#(GJInfo(RevCnt).dacctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
          NEXT
        CASE TranDraftPayment         ' 8=Draft Payment
          FOR RevCnt = 1 TO NumOfRevs
            GJInfo(RevCnt).pAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).pAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
            GJInfo(RevCnt).pAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).dacctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
          NEXT
        CASE TranRefundDeposit        ' 9=Refund Deposit
          'no action
        CASE TranBeginBalance         '10=Beginning Balance
          'no action
        CASE TranUpwardAdjustment     '11=Upward Adjustments
          'no action
        CASE TranDownwardAdjustment   '12=Downward Adjustments
          'no action
      END SELECT

    CASE 2      'Accural
      SELECT CASE UBTransRec(1).TransType
        CASE TranUtilityBill          ' 1=Utility bill
          FOR RevCnt = 1 TO NumOfRevs
            GJInfo(RevCnt).BAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).BAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
            GJInfo(RevCnt).BAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).BAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
          NEXT
        CASE TranLateCharge           ' 2=late charge
          FOR RevCnt = 1 TO NumOfRevs
            GJInfo(RevCnt).BAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).BAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
            GJInfo(RevCnt).BAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).BAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
          NEXT
        CASE TranReconnectFee         ' 3=reconnect fee
          FOR RevCnt = 1 TO NumOfRevs
            GJInfo(RevCnt).BAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).BAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
            GJInfo(RevCnt).BAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).BAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
          NEXT
        CASE TranBillPayment          ' 4=Bill Payment
          FOR RevCnt = 1 TO NumOfRevs
            GJInfo(RevCnt).pAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).pAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
            GJInfo(RevCnt).pAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).pAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
          NEXT
        CASE TranAppliedDeposit       ' 5=Applied Deposit
          'no action
          'FOR RevCnt = 1 TO NumOfRevs
          '  GJInfo(RevCnt).dacctInfo.CreditAmt = Round#(GJInfo(RevCnt).dacctInfo.CreditAmt - UBTransRec(1).RevAmt(RevCnt))
          '  GJInfo(RevCnt).pAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).pAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
          'NEXT

        CASE TranPenaltyCharge        ' 6=Penalty Charge
          FOR RevCnt = 1 TO NumOfRevs
            GJInfo(RevCnt).BAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).BAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
            GJInfo(RevCnt).BAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).BAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
          NEXT
        CASE TranDepositPayment       ' 7=Deposit Payment
          FOR RevCnt = 1 TO NumOfRevs
            GJInfo(RevCnt).dacctInfo.CreditAmt = Round#(GJInfo(RevCnt).dacctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
            GJInfo(RevCnt).dacctInfo.DebitAmt = Round#(GJInfo(RevCnt).dacctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
          NEXT
        CASE TranDraftPayment         ' 8=Draft Payment
          FOR RevCnt = 1 TO NumOfRevs
            GJInfo(RevCnt).pAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).pAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
            GJInfo(RevCnt).pAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).pAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
          NEXT
        CASE TranRefundDeposit        ' 9=Refund Deposit
          'no action
        '  FOR RevCnt = 1 TO NumOfRevs
        '    GJInfo(RevCnt).dacctInfo.CreditAmt = Round#(GJInfo(RevCnt).dacctInfo.CreditAmt - UBTransRec(1).RevAmt(RevCnt))
        '    GJInfo(RevCnt).dacctInfo.DebitAmt = Round#(GJInfo(RevCnt).dacctInfo.DebitAmt - UBTransRec(1).RevAmt(RevCnt))
        '  NEXT
        CASE TranBeginBalance         '10=Beginning Balance
          'no action
        CASE TranUpwardAdjustment     '11=Upward Adjustments
          FOR RevCnt = 1 TO NumOfRevs
            GJInfo(RevCnt).BAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).BAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
            GJInfo(RevCnt).BAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).BAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
          NEXT
        CASE TranDownwardAdjustment   '12=Downward Adjustments
          FOR RevCnt = 1 TO NumOfRevs
            GJInfo(RevCnt).BAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).BAcctInfo.CreditAmt - UBTransRec(1).RevAmt(RevCnt))
            GJInfo(RevCnt).BAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).BAcctInfo.DebitAmt - UBTransRec(1).RevAmt(RevCnt))
          NEXT
       END SELECT
    END SELECT
    'SmallPause
    
  NEXT


'NOTE: Journal Rec 1 is the credit, Rec 2 is the debit
  FOR RevCnt = 1 TO NumOfRevs
    REDIM GJRec1(1 TO 2) AS TrEditRecType
    IF GJInfo(RevCnt).BAcctInfo.CreditAmt <> 0 THEN
      GJRec1(1).AcctRec = GJInfo(RevCnt).BAcctInfo.CRecNo
      GJRec1(1).AcctNum = GJInfo(RevCnt).BAcctInfo.CAcctNo
      GJRec1(1).AcctName = GJInfo(RevCnt).BAcctInfo.CTitle
      GJRec1(1).TrDate = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).CrAmt = GJInfo(RevCnt).BAcctInfo.CreditAmt
      GJRec1(1).EType = "C"
      GJRec1(1).Desc = "FROM UTILITIES"
      GJRec1(1).Src = "IF"
      PUT #GJFile, , GJRec1(1)
    END IF
    IF GJInfo(RevCnt).BAcctInfo.DebitAmt <> 0 THEN
      GJRec1(2).AcctRec = GJInfo(RevCnt).BAcctInfo.DRecNo
      GJRec1(2).AcctNum = GJInfo(RevCnt).BAcctInfo.DAcctNo
      GJRec1(2).AcctName = GJInfo(RevCnt).BAcctInfo.DTitle
      GJRec1(2).TrDate = WorkDate
      GJRec1(2).Ref = Ref$
      GJRec1(2).DrAmt = GJInfo(RevCnt).BAcctInfo.DebitAmt
      GJRec1(2).EType = "D"
      GJRec1(2).Desc = "FROM UTILITIES"
      GJRec1(2).Src = "IF"
      PUT #GJFile, , GJRec1(2)
    END IF
  NEXT

  FOR RevCnt = 1 TO NumOfRevs
    REDIM GJRec1(1 TO 2) AS TrEditRecType
    IF GJInfo(RevCnt).pAcctInfo.CreditAmt <> 0 THEN
      GJRec1(1).AcctRec = GJInfo(RevCnt).pAcctInfo.CRecNo
      GJRec1(1).AcctNum = GJInfo(RevCnt).pAcctInfo.CAcctNo
      GJRec1(1).AcctName = GJInfo(RevCnt).pAcctInfo.CTitle
      GJRec1(1).TrDate = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).CrAmt = GJInfo(RevCnt).pAcctInfo.CreditAmt
      GJRec1(1).EType = "C"
      GJRec1(1).Desc = "FROM UTILITIES"
      GJRec1(1).Src = "IF"
      PUT #GJFile, , GJRec1(1)
    END IF
    IF GJInfo(RevCnt).pAcctInfo.DebitAmt <> 0 THEN
      GJRec1(2).AcctRec = GJInfo(RevCnt).pAcctInfo.DRecNo
      GJRec1(2).AcctNum = GJInfo(RevCnt).pAcctInfo.DAcctNo
      GJRec1(2).AcctName = GJInfo(RevCnt).pAcctInfo.DTitle
      GJRec1(2).TrDate = WorkDate
      GJRec1(2).Ref = Ref$
      GJRec1(2).DrAmt = GJInfo(RevCnt).pAcctInfo.DebitAmt
      GJRec1(2).EType = "D"
      GJRec1(2).Desc = "FROM UTILITIES"
      GJRec1(2).Src = "IF"
      PUT #GJFile, , GJRec1(2)
    END IF
  NEXT

  FOR RevCnt = 1 TO NumOfRevs
    REDIM GJRec1(1 TO 2) AS TrEditRecType
    IF GJInfo(RevCnt).dacctInfo.CreditAmt <> 0 THEN
      GJRec1(1).AcctRec = GJInfo(RevCnt).dacctInfo.CRecNo
      GJRec1(1).AcctNum = GJInfo(RevCnt).dacctInfo.CAcctNo
      GJRec1(1).AcctName = GJInfo(RevCnt).dacctInfo.CTitle
      GJRec1(1).TrDate = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).CrAmt = GJInfo(RevCnt).dacctInfo.CreditAmt
      GJRec1(1).EType = "C"
      GJRec1(1).Desc = "FROM UTILITIES"
      GJRec1(1).Src = "IF"
      PUT #GJFile, , GJRec1(1)
    END IF
    IF GJInfo(RevCnt).dacctInfo.DebitAmt <> 0 THEN
      GJRec1(2).AcctRec = GJInfo(RevCnt).dacctInfo.DRecNo
      GJRec1(2).AcctNum = GJInfo(RevCnt).dacctInfo.DAcctNo
      GJRec1(2).AcctName = GJInfo(RevCnt).dacctInfo.DTitle
      GJRec1(2).TrDate = WorkDate
      GJRec1(2).Ref = Ref$
      GJRec1(2).DrAmt = GJInfo(RevCnt).dacctInfo.DebitAmt
      GJRec1(2).EType = "D"
      GJRec1(2).Desc = "FROM UTILITIES"
      GJRec1(2).Src = "IF"
      PUT #GJFile, , GJRec1(2)
    END IF
  NEXT

UBBunchReturn:
RETURN

ValidateGLAccounts:
  BadAcct = False
  FOR RevCnt = 1 TO NumOfRevs
    'Billing Accounts
    IF InterfaceMethod = 2 THEN
'NOTE: We Only check billing accounts if Accural method
      ActT = 1
      ThisAcct = FindAcct(GJInfo(RevCnt).BAcctInfo.DAcctNo)
      IF ThisAcct <= 0 THEN
        BadDAcct$ = GJInfo(RevCnt).BAcctInfo.DAcctNo
        BadAcct = True
      ELSE
        GJInfo(RevCnt).BAcctInfo.DRecNo = ThisAcct
        GJInfo(RevCnt).BAcctInfo.DTitle = GetAcctTitle$(ThisAcct)
        BadDAcct$ = "     OK"
      END IF
      ThisAcct = FindAcct(GJInfo(RevCnt).BAcctInfo.CAcctNo)
      IF ThisAcct <= 0 THEN
        BadCAcct$ = GJInfo(RevCnt).BAcctInfo.CAcctNo
        BadAcct = True
      ELSE
        GJInfo(RevCnt).BAcctInfo.CRecNo = ThisAcct
        GJInfo(RevCnt).BAcctInfo.CTitle = GetAcctTitle$(ThisAcct)
        BadCAcct$ = "     OK"
      END IF
      GOSUB PrintBadAcct
    END IF

    'Payment Accounts
    ActT = 2
    ThisAcct = FindAcct(GJInfo(RevCnt).pAcctInfo.DAcctNo)
    IF ThisAcct <= 0 THEN
      BadDAcct$ = GJInfo(RevCnt).pAcctInfo.DAcctNo
      BadAcct = True
    ELSE
      GJInfo(RevCnt).pAcctInfo.DRecNo = ThisAcct
      GJInfo(RevCnt).pAcctInfo.DTitle = GetAcctTitle$(ThisAcct)
      BadDAcct$ = "     OK"
    END IF
    ThisAcct = FindAcct(GJInfo(RevCnt).pAcctInfo.CAcctNo)
    IF ThisAcct <= 0 THEN
      BadCAcct$ = GJInfo(RevCnt).pAcctInfo.CAcctNo
      BadAcct = True
    ELSE
      GJInfo(RevCnt).pAcctInfo.CRecNo = ThisAcct
      GJInfo(RevCnt).pAcctInfo.CTitle = GetAcctTitle$(ThisAcct)
      BadCAcct$ = "     OK"
    END IF
    GOSUB PrintBadAcct

    'Deposit Accounts
    ActT = 3
    IF UBSetUpRec(1).Revenues(RevCnt).UseDep = "Y" THEN
      ThisAcct = FindAcct(GJInfo(RevCnt).dacctInfo.DAcctNo)
      IF ThisAcct <= 0 THEN
        BadDAcct$ = GJInfo(RevCnt).dacctInfo.DAcctNo
        BadAcct = True
      ELSE
        GJInfo(RevCnt).dacctInfo.DRecNo = ThisAcct
        GJInfo(RevCnt).dacctInfo.DTitle = GetAcctTitle$(ThisAcct)
        BadDAcct$ = "     OK"
      END IF
    ELSE
      BadDAcct$ = "    N/A"
    END IF
    IF UBSetUpRec(1).Revenues(RevCnt).UseDep = "Y" THEN
      ThisAcct = FindAcct(GJInfo(RevCnt).dacctInfo.CAcctNo)
      IF ThisAcct <= 0 THEN
        BadCAcct$ = GJInfo(RevCnt).dacctInfo.CAcctNo
        BadAcct = True
      ELSE
        GJInfo(RevCnt).dacctInfo.CRecNo = ThisAcct
        GJInfo(RevCnt).dacctInfo.CTitle = GetAcctTitle$(ThisAcct)
        BadCAcct$ = "     OK"
      END IF
    ELSE
      BadCAcct$ = "    N/A"
    END IF
    GOSUB PrintBadAcct
    NEXT
  CLOSE RPTFile

  IF BadAcct THEN
    Ok = MsgBox%("UBSENDGL", "UBADACTS")
    EntryPoint = 0
    PrintRptFile "GL Transfer Invalid Account List.", "UBNOTFND.RPT", 1, RetCode, EntryPoint
  END IF
  KILL "UBNOTFND.RPT"
RETURN
ClearBoxUB:
  'Clear Box
  FOR SLine = 11 TO 14: LOCATE SLine, 20: PRINT STRING$(40, 32): NEXT SLine
  FOR SLine = 16 TO 16: LOCATE SLine, 20: PRINT STRING$(41, 32): NEXT SLine
  RETURN

END SUB

'
'***************************************************************************
'Finds the next undeleted record.
'Call with NextRec value of -1 for previous record, +1 for the next record.
'If a record is not found, the function returns the value of CurrRec.
'***************************************************************************
'
FUNCTION GetNextRec (FileNum, NumRecs, CurrRec, NextRec)

   Found = 0
   Rec = CurrRec

   DO

      Rec = Rec + NextRec                'Set file pointer to next record

      IF Rec > NumRecs OR Rec <= 0 THEN  'test for beg or end of file
         Found = 0                       'if no more records then get out
         EXIT DO
      END IF

      GET FileNum, Rec, GJEdit           'Get the record

      IF GJEdit.Deleted = 0 THEN         'Ok if not deleted
         Found = 1
         EXIT DO                         'Get out of loop when we find one
      END IF

   LOOP

   IF Found = 0 THEN
      GetNextRec = CurrRec
   ELSE
      GetNextRec = Rec
   END IF

END FUNCTION

SUB IFInit

   REDIM Form$(0, 0)                      'Form array holds data while editing
   REDIM Fld(0) AS FieldInfo              'Field info Array
   DIM Frm AS FormInfo                    'Form editing info

   LibLoadDisplayForm "GL.QSL", "SETUPIF", Form$(), Fld()
   Form$(1, 0) = "N"
   Form$(2, 0) = "N"
   Form$(3, 0) = "N"
   Form$(4, 0) = "N"
   Form$(5, 0) = "N"
   Action = 1

   DO

      EditForm Form$(), Fld(), Frm, Cnf, Action

         SELECT CASE Frm.KeyCode

         CASE F10Key  '--SaveButton
            '--Fld 2 = OutputDevice
        F1$ = Form$(1, 0)
        F2$ = Form$(2, 0)
        F3$ = Form$(3, 0)
        F4$ = Form$(4, 0)
        F5$ = Form$(5, 0)
        GOSUB InitIF
        DONE = True

         CASE ESC
         EXIT SUB

         END SELECT

   LOOP UNTIL DONE
         EXIT SUB

InitIF:
 LibLoadDisplayForm "GL.QSL", "PROCESIF", Form$(), Fld()
 IF F1$ = "Y" THEN
  QPrintRC "INITIALIZING CASH TRANSACTIONS.        ", 12, 23, 15

  REDIM CMTransRec(1)  AS CMTransRecType
  CMTransRecLen = LEN(CMTransRec(1))
  CMTran = FREEFILE
  OPEN "CMTRANS.DAT" FOR RANDOM SHARED AS CMTran LEN = CMTransRecLen
  NumOfTRecs& = LOF(CMTran) \ CMTransRecLen
  FOR TCnt& = NumOfTRecs& TO 1 STEP -1
    GET #CMTran, TCnt&, CMTransRec(1)
    CMTransRec(1).Trans2GL = "Y"
    PUT #CMTran, TCnt&, CMTransRec(1)
  NEXT TCnt&
  CLOSE CMTran
END IF

 IF F2$ = "Y" THEN
  QPrintRC "INITIALIZING UTILITY TRANSACTIONS.     ", 12, 23, 15

  REDIM UBTransRec(1)  AS UBTransRecType
  UBTransRecLen = LEN(UBTransRec(1))
  
  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTransRecLen
  NumOfTRecs& = LOF(UBTran) \ UBTransRecLen

  FOR TCnt& = NumOfTRecs& TO 1 STEP -1
    GET #UBTran, TCnt&, UBTransRec(1)
    UBTransRec(1).Posted2GL = "Y"
    PUT #UBTran, TCnt&, UBTransRec(1)
  NEXT TCnt&
  CLOSE UBTran

 END IF
 IF F3$ = "Y" THEN
  QPrintRC "INITIALIZING TAX TRANSACTIONS.         ", 12, 23, 15
  SLEEP 2
 END IF
 IF F4$ = "Y" THEN
  QPrintRC "INITIALIZING BUS. LICENSE TRANSACTIONS.", 12, 23, 15
  REDIM ARTransRec(1) AS ARTransRecType
  ARTransRecLen = LEN(ARTransRec(1))
  ARTransFile = FREEFILE
  OPEN "ARTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARTransFile LEN = ARTransRecLen
  NumOfTransRecs& = LOF(ARTransFile) \ ARTransRecLen
  
  FOR TCnt& = 1 TO NumOfTransRecs&
  GET ARTransFile, TCnt&, ARTransRec(1)
   ARTransRec(1).Posted2GL = "Y"
  PUT ARTransFile, TCnt&, ARTransRec(1)
  NEXT TCnt&
  CLOSE ARTransFile
 END IF
 IF F5$ = "Y" THEN
  QPrintRC "INITIALIZING DECAL TRANSACTIONS.       ", 12, 23, 15
  SLEEP 2
 END IF
  PRINT CHR$(7);
  QPrintRC "DONE!!!                                ", 12, 23, 15
  SLEEP 3


 RETURN

END SUB

SUB LoadUBSetUpFile (UBSetUpFileNum, UBSetUpLen)

   REDIM UBSetUpRec(1) AS UBSetUpRecType
   UBSetUpLen = LEN(UBSetUpRec(1))
   UBSetUpFileNum = FREEFILE
   OPEN "UBSETUP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBSetUpFileNum LEN = UBSetUpLen
   
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

   Ok = MsgBox%("GL.QSL", "POSTGJ")       'Ask user if sure ready to post

   IF Ok = 1 THEN EXIT SUB   '1=No
   
   TotDr# = 0                             'init totals to zero
   TotCr# = 0
   Active = 0                             'Counter for Active Transactions

   IF TotDr# <> TotCr# THEN               'Transactions out of balance
      Ok = MsgBox%("GL.QSL", "OUTOFBAL")  'ask user if ok to post
      IF Ok = 1 THEN EXIT SUB             'No = button 1
   END IF

   
   Active = 0                             'Reset Active counter for posting

   OpenIFEditFile IFEditFileNum, NumIFTrans

   DIM Tr2Post AS GLTransRecType
   OPEN "GJ2POST.DAT" FOR RANDOM AS #2 LEN = LEN(Tr2Post)

   FOR Cnt = 1 TO NumIFTrans              'Assign edit file to trans format
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
   
   Post2GL "GJ2POST.DAT", BadTrans%       'common post & link sub in GLAUX.BAS
   
   KILL "GLTRXED.DAT"                     'kill the temp files
   KILL "GJ2POST.DAT"

   IF BadTrans <> 0 THEN                  'posting problem
   CLS
      PRINT "Error: One or more transactions were not posted."
      PRINT "Press any key to view posting log"
      k$ = INPUT$(1)
      PrintRptFile "Posting Log", "GLUTIL.LOG", LPTNo, RetCode%, 2
   END IF

END SUB

SUB PrnEditList

   SHARED User$
   SHARED Choice$()                          'Multiple choice array
   REDIM Form$(0, 0)                           'Holds data from form
   REDIM Fld(0) AS FieldInfo                 'Field editing array
   DIM Frm AS FormInfo                       'Dim the form info array

   REDIM Choice$(0 TO 3, 1)                  'Choices for entry type field
   Choice$(0, 0) = "2"                       'Form Field Number for picklist
   Choice$(1, 0) = "Screen"
   Choice$(2, 0) = "Prn #1"
   Choice$(3, 0) = "Prn #2"
   
   LibLoadDisplayForm "GL.QSL", "PRNGJ", Form$(), Fld()
   PrintHelp "Print Distribution Register"
   'Set Bad Acct Flag to No Here
   BadAcctFlag = 0

   Action = 1

   DO

      EditForm Form$(), Fld(), Frm, Cnf, Action

         SELECT CASE Frm.KeyCode

         CASE F10Key  '--SaveButton
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

   LOOP UNTIL Frm.KeyCode = F10Key

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
   Page = 0
   User$ = QPTrim$(User$)
   RptTitle$ = "General Ledger Interface Report"
   Header$ = "General Ledger Interface Report"
   CrLf$ = CHR$(13) + CHR$(10)
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
Page = Page + 1
  PRINT #PrnFileNum, TAB(40 - (INT(LEN(User$) / 2))); User$
  PRINT #PrnFileNum, TAB(40 - (INT(LEN(Header$) / 2))); Header$
  PRINT #PrnFileNum,
  PRINT #PrnFileNum, "Report Date: "; DATE$; TAB(67); "Page #"; Page
  PRINT #PrnFileNum, "Date        Description           Reference"
  PRINT #PrnFileNum, "            G/L Account                                     Debit         Credit"
  PRINT #PrnFileNum, "--------------------------------------------------------------------------------"
  LineCnt = 5
RETURN


END SUB

SUB TRGetAcctStruct (GLFundLen%, GLAcctLen%, GLDetLen%)

  REDIM GLSetUpRec(1) AS GLSetupRecType
  SetUpRecLen = LEN(GLSetUpRec(1))

  SetupFile = FREEFILE
  OPEN "GLSETUP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS SetupFile LEN = SetUpRecLen
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

 'OK to Rename
  DF$ = "Copy GLTRXED.DAT GLGJED.DAT"
  SHELL DF$
  DF$ = "DEL GLTRXED.DAT"
  SHELL DF$
  Ok = MsgBox%("GLUTIL", "TRXOK")

GJSendExit:
END SUB

