DEFINT A-Z
DECLARE FUNCTION FindFund% (FundNum$)
DECLARE FUNCTION FindAcct% (AcctNum$)
DECLARE FUNCTION FindDept% (DeptNum$)
DECLARE FUNCTION FmtAcct$ (AcctNum$, FundLen%, AcctLen%, DetLen%)
DECLARE FUNCTION StripAcct$ (AcctNum$)
DECLARE FUNCTION GetFundTitle$ (FundRecNum%)
DECLARE FUNCTION GetAcctTitle$ (AcctRecNum)
DECLARE FUNCTION GetAcctType$ (AcctRecNum)
DECLARE FUNCTION GetDeptTitle$ (DeptRecNum%)
DECLARE FUNCTION FmtDate$ (d$)
DECLARE FUNCTION ConvDateStr$ (d$)
DECLARE FUNCTION ConvDateStr1$ (d$)
DECLARE FUNCTION GetBegDate% ()
DECLARE FUNCTION GetEndDate% ()
DECLARE FUNCTION Round# (N#)
DECLARE FUNCTION GetBankGLAcct$ (BankNum%)
DECLARE FUNCTION EDCheck4Acct2% (Form$(), AcctFld%)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION FUsing$ (Number$, Mask$)
DECLARE FUNCTION MGetKey% (Row%, Col%)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION Peek1% (Segment%, Address%)
DECLARE FUNCTION QPTrim$ (X$)
DECLARE FUNCTION WaitKey% ()

DECLARE SUB OpenBankFile (BankRecLen%, BankFile%, NumBanks%)
DECLARE SUB OpenDeptFile (DeptFile%, NumDepts%)
DECLARE SUB OpenPOTransFile (TransFileNum%, NumTrans&)
DECLARE SUB StuffBuf (Ky$)
DECLARE SUB ButtonPress (ButNo%, Down%, Presses%, X%, Y%)
DECLARE SUB GetCursor (X, Y, Button)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB BLPrint (LPTPort%, Text$, ErrCode%)
DECLARE SUB FileView (FileName$, Ky, Action, FVI AS ANY, SEG Array)
DECLARE SUB HideCursor ()
DECLARE SUB MScrnSave (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MScrnRest (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MPaintBox (ULRow%, ULCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB QPrint (X$, Colr%, page%)
DECLARE SUB QPrintRC (t$, r%, c%, Clr%)
DECLARE SUB ShowCursor ()
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE SUB OpenFundIdx (FundIdxFileNum%, NumFIdxRecs%)
DECLARE SUB OpenFundFile (FundFileNum%, NumFunds%)
DECLARE SUB OpenAcctIdx (AcctIdxFileNum%, NumIdxRecs%)
DECLARE SUB OpenAcctFile (AcctFileNum%, NumGLAccts%)
DECLARE SUB OpenDeptIdx (DeptIdxFile%, NumActiveDepts%)
DECLARE SUB OpenTransFile (TransFileNum%, NumTrans&)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PrintHelp (Help$)
DECLARE SUB GetFundCodes (FirstFund$, LastFund$)
DECLARE SUB Post2GL (FileName$, BadTrans%)
DECLARE SUB GetFYDates (FY1BegDate, FY1EndDate, FY2BegDate, FY2EndDate)
DECLARE SUB GetPostDates (LPDate, HPDate)

''$INCLUDE: 'GL.BI'
''$INCLUDE: 'Fileview.bi'
''$INCLUDE: 'DefCnf.bi'        'Defines a TYPE for monitor and color info..
''$INCLUDE: 'FieldInf.bi'      'field information
''$INCLUDE: 'FormEdit.bi'      'form editing information
''$INCLUDE: 'QSCR.BI'          'QuickScreen Declarations
''$INCLUDE: 'SCRAUX.BI'

'$INCLUDE: 'fileview.bi'
'$INCLUDE: 'GL.BI'
'$INCLUDE: 'AP.BI'

DECLARE SUB LibLoadDisplayForm (LibName$, FormName$, Form$(), Fld() AS ANY)
DECLARE SUB EditForm (Form$(), Fld() AS ANY, Frm AS ANY, Cnf AS ANY, Action%)
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE FUNCTION InputBox$ (LibName$, FormName$, Cancel%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION GetPassword$ (Row%, Col%, Max%, Abort%)
DECLARE SUB PrintHelp (H$)
DECLARE SUB DispLoad ()

DIM SHARED FundIdx AS GLFundIndexType
DIM SHARED Fund AS GLFundRecType
DIM SHARED AcctIdx AS GLAcctIndexType
DIM SHARED Acct AS GLAcctRecType
DIM SHARED Trans AS GLTransRecType
DIM SHARED DeptIdx AS GLDeptIndexType
DIM SHARED Dept AS GLDeptRecType

'****************************************************************************
'  D$ is a date string in the format 03-15-1995.
'  Returns: 6 character string in the format 031595
'
'****************************************************************************
FUNCTION ConvDateStr$ (d$)
   ConvDateStr$ = LEFT$(d$, 2) + MID$(d$, 4, 2) + RIGHT$(d$, 2)
END FUNCTION

'****************************************************************************
'  D$ is a date string in the format 03-15-1995.
'  Returns: 6 character date string in the format 03-15-95
'****************************************************************************
FUNCTION ConvDateStr1$ (d$)
   ConvDateStr1$ = LEFT$(d$, 2) + "-" + MID$(d$, 4, 2) + "-" + RIGHT$(d$, 2)
END FUNCTION

'*****************************************************************************
'Searches the acct index for a matching account number and returns the record
'number of the account
'
'    Input: AcctNum$ as a formatted G/L account number string
'  Returns: Record number of the account
'*****************************************************************************
'
FUNCTION FindAcct (AcctNum$)

   OpenAcctIdx AcctIdxFileNum, NumIdxRecs

   IF NumIdxRecs = 0 THEN
      RecordNum = 0
      CLOSE AcctIdxFileNum
      EXIT FUNCTION
   END IF

   Match = FALSE
   FirstRec = 1
   LastRec = NumIdxRecs
   Lookfor$ = LTRIM$(RTRIM$(AcctNum$))

   DO UNTIL LastRec < FirstRec

      MiddleRec = (LastRec + FirstRec) \ 2

      GET AcctIdxFileNum, MiddleRec, AcctIdx

      IF Lookfor$ = RTRIM$(AcctIdx.AcctNum) THEN
         Match = -1
         EXIT DO
      ELSEIF Lookfor$ < AcctIdx.AcctNum THEN
         LastRec = MiddleRec - 1
      ELSE
         FirstRec = MiddleRec + 1
      END IF

   LOOP

   IF Match THEN
      RecNum = AcctIdx.RecNum
   ELSE
      RecNum = 0
   END IF

   FindAcct = RecNum

   CLOSE AcctIdxFileNum

END FUNCTION

FUNCTION FindDept% (DeptNum$)
'*****************************************************************************
'Searches the Dept index for a matching account number and returns the record
'number of the account
'
'    Input: DeptNum$ as a formatted G/L account number string
'  Returns: Record number oaf the account
'*****************************************************************************
'
'FUNCTION FindDept (DeptNum$)

   OpenDeptIdx DeptIdxFileNum, NumIdxRecs

   IF NumIdxRecs = 0 THEN
      RecordNum = 0
      CLOSE DeptIdxFileNum
      EXIT FUNCTION
   END IF

   Match = FALSE
   FirstRec = 1
   LastRec = NumIdxRecs
   Lookfor$ = LTRIM$(RTRIM$(DeptNum$))

   DO UNTIL LastRec < FirstRec

      MiddleRec = (LastRec + FirstRec) \ 2

      GET DeptIdxFileNum, MiddleRec, DeptIdx

      IF Lookfor$ = RTRIM$(DeptIdx.DeptNum) THEN
         Match = -1
         EXIT DO
      ELSEIF Lookfor$ < DeptIdx.DeptNum THEN
         LastRec = MiddleRec - 1
      ELSE
         FirstRec = MiddleRec + 1
      END IF

   LOOP

   IF Match THEN
      RecNum = DeptIdx.RecNum
   ELSE
      RecNum = 0
   END IF

   FindDept = RecNum

   CLOSE DeptIdxFileNum


END FUNCTION

'****************************************************************************
'formats an account number string with dashes.
'****************************************************************************
'
FUNCTION FmtAcct$ (AN$, FundLen%, AcctLen%, DetLen%) STATIC

  FmtTotAcctLen = FundLen + AcctLen + DetLen

  AN$ = QPTrim$(AN$)
  ANLen = LEN(AN$)

  IF ANLen > FmtTotAcctLen THEN
    AN$ = LEFT$(AN$, FmtTotAcctLen)
    ANLen = FmtTotAcctLen
  END IF

  SELECT CASE ANLen
    CASE IS < FundLen
      FmtAcct$ = AN$
    CASE FundLen
      FmtAcct$ = AN$ + "-"
    CASE (FundLen + 1) TO (AcctLen + FundLen) - 1
      FmtAcct$ = LEFT$(AN$, FundLen) + "-" + MID$(AN$, FundLen + 1)
    CASE (AcctLen + FundLen)
      FmtAcct$ = LEFT$(AN$, FundLen) + "-" + MID$(AN$, FundLen + 1, AcctLen) + "-"
    CASE (AcctLen + FundLen + 1) TO (AcctLen + FundLen + DetLen) - 1
      FmtAcct$ = LEFT$(AN$, FundLen) + "-" + MID$(AN$, FundLen + 1, AcctLen) + "-" + MID$(AN$, FundLen + AcctLen + 1)
    CASE (AcctLen + FundLen + DetLen)
      FmtAcct$ = LEFT$(AN$, FundLen) + "-" + MID$(AN$, FundLen + 1, AcctLen) + "-" + MID$(AN$, FundLen + AcctLen + 1, DetLen) 'RIGHT$(AN$, DetLen)
  END SELECT

END FUNCTION

SUB GetAcctStruct (GLFundLen%, GLAcctLen%, GLDetLen%)

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

'****************************************************************************
'Retrieves the GL account title from the account data file.
'****************************************************************************
'
FUNCTION GetAcctTitle$ (RecordNumber)

   OpenAcctFile AcctFileNum, NumAccts
   GET AcctFileNum, RecordNumber, Acct
   GetAcctTitle$ = Acct.Title
   CLOSE AcctFileNum
   
END FUNCTION

FUNCTION GetBankGLAcct$ (BankNum)

  REDIM Bank(1) AS BankRecType
  BankRecLen = LEN(Bank(1))
  OpenBankFile BankRecLen, BankFile, NumBanks
  GET BankFile, BankNum, Bank(1)
  CLOSE BankFile
  GetBankGLAcct$ = Bank(1).GLAcct

END FUNCTION

SUB GetBankList (BankList$(), NumBanks)

  REDIM Bank(1) AS BankRecType
  BankRecLen = LEN(Bank(1))
  OpenBankFile BankRecLen, BankFile, NumBanks
  
  IF NumBanks = 0 THEN
    CLOSE BankFile
    EXIT SUB
  END IF
  
  REDIM BankList$(1 TO NumBanks)
  BankNum$ = SPACE$(2)

  FOR Cnt = 1 TO NumBanks
    GET BankFile, Cnt, Bank(1)
    IF NOT Bank(1).Deleted THEN
      BCnt = BCnt + 1
      RSET BankNum$ = QPTrim$(STR$(Bank(1).BankNum))
      BankList$(BCnt) = BankNum$ + " " + QPTrim$(Bank(1).BankName)
      LSET BankNum$ = ""
    ELSE

    END IF
  NEXT

  CLOSE BankFile
  
END SUB

FUNCTION GetBegDate

   DIM GLSetup AS GLSetupRecType
   SetUpRecLen = LEN(GLSetup)

   SetupFile = FREEFILE
   OPEN "GLSETUP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS SetupFile LEN = SetUpRecLen

   GET SetupFile, 1, GLSetup

   GetBegDate = GLSetup.FYBeg

   CLOSE SetupFile


END FUNCTION

FUNCTION GetDeptTitle$ (DeptRecNum)

  DIM DeptRec AS GLDeptRecType
  OpenDeptFile DeptFileNum, NumDepts
  GET DeptFileNum, DeptRecNum, DeptRec
  GetDeptTitle$ = DeptRec.Title
  CLOSE DeptFileNum

END FUNCTION

FUNCTION GetEndDate

   DIM GLSetup AS GLSetupRecType
   SetUpRecLen = LEN(GLSetup)

   SetupFile = FREEFILE
   OPEN "GLSETUP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS SetupFile LEN = SetUpRecLen

   GET SetupFile, 1, GLSetup

   GetEndDate = GLSetup.FYEnd

   CLOSE SetupFile

END FUNCTION

SUB GetFundList (FundList$(), NumFunds)

  DIM FundIndex AS GLFundIndexType
  OpenFundIdx FundIdxFile, NumFunds

  IF NumFunds = 0 THEN
    Ok = MsgBox("GL", "NOFUNDS")
    CLOSE
    EXIT SUB
  END IF

  REDIM FundList$(1 TO NumFunds)

  FOR Cnt = 1 TO NumFunds
    GET FundIdxFile, Cnt, FundIndex
    FundList$(Cnt) = QPTrim$(FundIndex.FundNum)
  NEXT

  CLOSE FundIdxFileNum

END SUB

SUB GetFYDates (FY1BegDate, FY1EndDate, FY2BegDate, FY2EndDate)

   DIM GLSetup AS GLSetupRecType
   SetUpRecLen = LEN(GLSetup)

   SetupFile = FREEFILE
   OPEN "GLSETUP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS SetupFile LEN = SetUpRecLen

   GET SetupFile, 1, GLSetup

   FY1BegDate = GLSetup.FYBeg
   FY1EndDate = GLSetup.FYEnd
   FY2BegDate = GLSetup.NYBeg
   FY2EndDate = GLSetup.NYEnd

   CLOSE SetupFile



END SUB

SUB GetPostDates (LPDate, HPDate)

   DIM GLSetup AS GLSetupRecType
   SetUpRecLen = LEN(GLSetup)
   SetupFile = FREEFILE
   OPEN "GLSETUP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS SetupFile LEN = SetUpRecLen

   GET SetupFile, 1, GLSetup

   LPDate = GLSetup.LPDate
   HPDate = GLSetup.HPDate

   CLOSE SetupFile

END SUB

SUB ooPressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
  
  Cols = ButtonRCol - ButtonLCol + 1
  REDIM Btn(Cols)
  MScrnSave ButtonRow, ButtonLCol, ButtonRow, ButtonRCol, Btn(1)

  MPaintBox ButtonRow, ButtonLCol, ButtonRow, ButtonRCol, 15
  
'  Pause 2
  DO
    GetCursor 0, 0, i
  LOOP WHILE i
  ButtonPress 0, 0, 0, 0, 0

  MScrnRest ButtonRow, ButtonLCol, ButtonRow, ButtonRCol, Btn(1)
  
  IF KeyCode < 0 THEN
    StuffBuf CHR$(0) + CHR$(ABS(KeyCode))
  ELSEIF KeyCode > 0 THEN
    StuffBuf CHR$(KeyCode)
  END IF

  ERASE Btn

END SUB

SUB OpenAcctFile (AcctFileNum, NumAccts)

   AcctRecLen = LEN(Acct)
   AcctFileNum = FREEFILE
   OPEN "GLACCT.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS AcctFileNum LEN = AcctRecLen
   NumAccts = LOF(AcctFileNum) \ AcctRecLen

END SUB

SUB OpenAcctIdx (AcctIdxFileNum, NumIdxRecs)

   AcctIdxLen = LEN(AcctIdx)
   AcctIdxFileNum = FREEFILE
   OPEN "GLACCT.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS AcctIdxFileNum LEN = AcctIdxLen
   NumIdxRecs = LOF(AcctIdxFileNum) \ AcctIdxLen

END SUB

SUB OpenBankFile (BankRecLen, BankFile, NumBanks)

   BankFile = FREEFILE
   OPEN "GLBANK.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS BankFile LEN = BankRecLen
   NumBanks = LOF(BankFile) \ BankRecLen

END SUB

SUB OpenDeptFile (DeptFile%, NumDepts%)

   DeptRecLen = LEN(Dept)
   DeptFile = FREEFILE
   OPEN "GLDEPT.DAT " FOR RANDOM ACCESS READ WRITE SHARED AS DeptFile LEN = DeptRecLen
   NumDepts = LOF(DeptFile) \ DeptRecLen

END SUB

SUB OpenDeptIdx (DeptIdxFile, NumActiveDepts)

   DeptIdxLen = LEN(DeptIdx)
   DeptIdxFile = FREEFILE
   OPEN "GLDEPT.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS DeptIdxFile LEN = DeptIdxLen
   NumActiveDepts = LOF(DeptIdxFile) \ DeptIdxLen

END SUB

SUB OpenFundFile (FundFileNum%, NumFunds%)

   FundRecLen = LEN(Fund)
   FundFileNum = FREEFILE
   OPEN "GLFUND.DAT " FOR RANDOM ACCESS READ WRITE SHARED AS FundFileNum LEN = FundRecLen
   NumFunds = LOF(FundFileNum) \ FundRecLen
   
END SUB

SUB OpenFundIdx (FundIdxFileNum, NumFIdxRecs)

   FundIdxLen = LEN(FundIdx)
   FundIdxFileNum = FREEFILE
   OPEN "GLFUND.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS FundIdxFileNum LEN = FundIdxLen
   NumFIdxRecs = LOF(FundIdxFileNum) \ FundIdxLen

END SUB

SUB OpenPOTransFile (TransFileNum, NumTrans&)

   REDIM TempPOTrans(1) AS GLTransRecType
   TransRecLen = LEN(TempPOTrans(1))
   TransFileNum = FREEFILE
   OPEN "POTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS TransFileNum LEN = TransRecLen
   NumTrans& = LOF(TransFileNum) \ TransRecLen
   ERASE TempPOTrans

END SUB

SUB OpenTransFile (TransFileNum, NumTrans&)

   TransRecLen = LEN(Trans)
   TransFileNum = FREEFILE
   OPEN "GLTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS TransFileNum LEN = TransRecLen
   NumTrans& = LOF(TransFileNum) \ TransRecLen

END SUB

'****************************************************************************
' Input: FileName$ is the edit file to be posted, which is in the same type
'        as the transaction history (GLTRANS.DAT) file
' BadTrans returns the record number of a transaction which was not posted
'****************************************************************************
SUB Post2GL (FileName$, BadTrans%)

   SHARED Acct AS GLAcctRecType, Trans AS GLTransRecType
   
   DIM Tran2Post AS GLTransRecType        'Dim a buffer for the edit file
   TrRecLen = LEN(Tran2Post)              'Determine the rec length
   File2Post = FREEFILE                   'Get a handle

   OPEN FileName$ FOR RANDOM AS File2Post LEN = TrRecLen

   Num2Post = LOF(File2Post) \ TrRecLen   'Find the num of transactions
   
   OpenAcctFile AcctFileNum, NumAccts     'Open & lock GL files
   'LOCK AcctFileNum

   OpenTransFile TransFileNum, NumTrans&
   'LOCK TransFileNum

   '--update the posting log file
   GLLogFileName$ = "GLUTIL.LOG"
   GLLogFile = FREEFILE
   OPEN GLLogFileName$ FOR APPEND AS GLLogFile
   PRINT #GLLogFile, "Post to General Ledger initiated on " + DATE$ + " @ " + TIME$
   Log$ = SPACE$(132)
   
   FOR Cnt = 1 TO Num2Post                'Start processing transactions
      GET File2Post, Cnt, Tran2Post

      IF Tran2Post.Marked = FALSE THEN

        RecNum = FindAcct(Tran2Post.AcctNum)   'Verify account is in G/L
        IF RecNum > 0 THEN                  'if valid acct then proceed

           '--tell user what's going on
           QPrintRC " Posting Account Number: ", 25, 1, 112
           QPrintRC Tran2Post.AcctNum, 25, 26, 112

           GET AcctFileNum, RecNum, Acct    'Get the account

           '--depending on account type, update running balance
         
           SELECT CASE Acct.Typ
              CASE "A", "E"                 'asset, exp accts
                 Acct.Bal = Round#(Acct.Bal) + Round#(Tran2Post.DrAmt) - Round#(Tran2Post.CrAmt)
                 PUT AcctFileNum, RecNum, Acct
           
              CASE "L", "R"                 'liab, rev accts
                 Acct.Bal = Round#(Acct.Bal) + Round#(Tran2Post.CrAmt) - Round#(Tran2Post.DrAmt)
                 PUT AcctFileNum, RecNum, Acct

           END SELECT

           DrPosted# = Round#(DrPosted#) + Round#(Tran2Post.DrAmt)
           CrPosted# = Round#(CrPosted#) + Round#(Tran2Post.CrAmt)
           NumTrans& = NumTrans& + 1          'increment record pointer

           GET TransFileNum, NumTrans&, Trans
           Trans.AcctNum = Tran2Post.AcctNum 'Assign editfile to trans history
           Trans.TRDate = Tran2Post.TRDate
           Trans.Desc = Tran2Post.Desc
           Trans.CrAmt = Tran2Post.CrAmt
           Trans.DrAmt = Tran2Post.DrAmt
           Trans.Ref = Tran2Post.Ref
           Trans.Src = Tran2Post.Src
           Trans.NextTran = 0
           PUT TransFileNum, NumTrans&, Trans
           Posted = Posted + 1

           Tran2Post.Marked = True
           PUT File2Post, Cnt, Tran2Post

           '---------------------------------Start linking here
           '--if first trans for this acct,
           IF Acct.FrstTran = 0 THEN
              Acct.FrstTran = NumTrans&      'assign first & last pointers to
              Acct.LastTran = NumTrans&      'this transaction
              PUT AcctFileNum, RecNum, Acct

           '--Prior Transactions have been posted to this acct
           ELSE
                                                'in the account file..
              Prev& = Acct.LastTran             'remember the prev trans pointer,
              Acct.LastTran = NumTrans&         'reset last trans to this trans
              PUT AcctFileNum, RecNum, Acct

                                                'In the trans file...
              GET TransFileNum, Prev&, Trans    'Get the last transaction
              Trans.NextTran = NumTrans&        'reset pointer to this trans
              PUT TransFileNum, Prev&, Trans

           END IF

           TransPosted = TransPosted + 1

        ELSE                                'Account NOT found!
           BadTrans = BadTrans + 1          'Pass info back to caller
           GOSUB LogGLPostErr

        END IF
      END IF  '--marked test
   NEXT

   IF BadTrans = 0 THEN
     PRINT #GLLogFile, USING "No Posting Errors. Posted Transaction Count:####"; TransPosted
   END IF

   PRINT #GLLogFile, USING "Debits Posted  : ########,.##"; DrPosted#
   PRINT #GLLogFile, USING "Credits Posted : ########,.##"; CrPosted#
   PRINT #GLLogFile, STRING$(78, "-")
   
  CLOSE

EXIT SUB

'was printing register and deleteing edit file here.
'Now do this in module that called this sub
GotErr:
   SELECT CASE ERR
      CASE 70
         CLS
         CLOSE
         LOCATE 10, 1: COLOR 11
         PRINT "Another user has the file locked, Please try again later."
         PRINT "Press any key to exit."
         Bye$ = INPUT$(1)
         EXIT SUB
      CASE ELSE
   END SELECT
RETURN

LogGLPostErr:
   PRINT #GLLogFile, "Unposted Transaction"
   PRINT #GLLogFile, "Record Number  :"; STR$(Cnt)
   PRINT #GLLogFile, "Account Number :"; Tran2Post.AcctNum
   PRINT #GLLogFile, "Date           :"; Num2Date(Tran2Post.TRDate)
   PRINT #GLLogFile, "Description    :"; Tran2Post.Desc
   PRINT #GLLogFile, "Debit          :"; STR$(Tran2Post.CrAmt)
   PRINT #GLLogFile, "Credit         :"; STR$(Tran2Post.DrAmt)
   PRINT #GLLogFile,
RETURN

END SUB

SUB Post2PO (FileName$, BadTrans%)

   SHARED Acct AS GLAcctRecType
   DIM POTrans AS GLTransRecType          '--PO Transactions
   DIM Tran2Post AS GLTransRecType        '--Dim a buffer for the edit file

   TrRecLen = LEN(Tran2Post)              'Determine the rec length
   File2Post = FREEFILE                   'Get a handle on the Interface file
   OPEN FileName$ FOR RANDOM AS File2Post LEN = TrRecLen
   Num2Post = LOF(File2Post) \ TrRecLen   'Find the num of transactions

   OpenAcctFile AcctFileNum, NumAccts     'Open & lock GL files
   'LOCK AcctFileNum

   OpenPOTransFile TransFileNum, NumTrans&
   'LOCK TransFileNum

   '--update the posting log file
   POLogFileName$ = "GLUTIL.LOG"
   POLogFile = FREEFILE
   OPEN POLogFileName$ FOR APPEND AS POLogFile
   PRINT #POLogFile, "Purchase Order initiated on " + DATE$ + " @ " + TIME$
   Log$ = SPACE$(132)

   FOR Cnt = 1 TO Num2Post                'Start processing transactions

      GET File2Post, Cnt, Tran2Post

      RecNum = FindAcct(Tran2Post.AcctNum)   'Verify account is in G/L

      IF RecNum > 0 THEN                     'if valid acct then proceed

         'tell user what's going on
         QPrintRC " Posting Account Number: ", 25, 1, 112
         QPrintRC Tran2Post.AcctNum, 25, 26, 112

         GET AcctFileNum, RecNum, Acct    'Get the account

         '--Update encumbrace field
         SELECT CASE Acct.Typ
            CASE "A", "E"                 'asset, exp accts
               Acct.Encumb = Round#(Acct.Encumb) + Round#(Tran2Post.DrAmt) - Round#(Tran2Post.CrAmt)
               PUT AcctFileNum, RecNum, Acct

            CASE "L", "R"                 'liab, rev accts
               Acct.Encumb = Round#(Acct.Encumb) + Round#(Tran2Post.CrAmt) - Round#(Tran2Post.DrAmt)
               PUT AcctFileNum, RecNum, Acct

         END SELECT

         NumTrans& = NumTrans& + 1          'increment record pointer

         GET TransFileNum, NumTrans&, POTrans

         POTrans.AcctNum = Tran2Post.AcctNum 'Assign editfile to trans history
         POTrans.TRDate = Tran2Post.TRDate
         POTrans.Desc = Tran2Post.Desc
         POTrans.CrAmt = Tran2Post.CrAmt
         POTrans.DrAmt = Tran2Post.DrAmt
         POTrans.Ref = Tran2Post.Ref
         POTrans.Src = Tran2Post.Src
         POTrans.NextTran = 0

         PUT TransFileNum, NumTrans&, POTrans

         Posted = Posted + 1

         '---------------------------------Start linking here
         IF Acct.FrstPTran = 0 THEN        'if first trans for this acct,
            Acct.FrstPTran = NumTrans&      'assign first & last pointers to
            Acct.LastPTran = NumTrans&      'this transaction
            PUT AcctFileNum, RecNum, Acct

         ELSE                                 'otherwise
                                              'in the account file..
            Prev& = Acct.LastPTran            'remember the prev trans pointer,
            Acct.LastPTran = NumTrans&        'reset last trans to this trans
            PUT AcctFileNum, RecNum, Acct

                                               'In the POTrans file...
            GET TransFileNum, Prev&, POTrans   'Get the last transaction
            POTrans.NextTran = NumTrans&       'reset pointer to this trans
            PUT TransFileNum, Prev&, POTrans

         END IF

         TransPosted = TransPosted + 1

      ELSE                                'Account NOT found!
         BadTrans = BadTrans + 1          'Pass info back to caller
         '--how about an error log here.
         GOSUB LogPOPostErr

      END IF

   NEXT

   'UNLOCK AcctFileNum
   'UNLOCK TransFileNum
    IF BadTrans = 0 THEN
      PRINT #POLogFile, USING "No Posting Errors. Posted Transaction Count:####"; TransPosted
      PRINT #POLogFile, STRING$(78, "-")
    END IF

  CLOSE AcctFileNum
  CLOSE TransFileNum
  CLOSE File2Post
  CLOSE POLogFile
'Clean up editfile in calling program in case not posted
EXIT SUB



POGotErr:
   SELECT CASE ERR
      CASE 70
         CLS
         CLOSE
         LOCATE 10, 1: COLOR 11
         PRINT "Another user has the file locked, Please try again later."
         PRINT "Press any key to exit."
         Bye$ = INPUT$(1)
         EXIT SUB
      CASE ELSE
   END SELECT
RETURN

LogPOPostErr:
   PRINT #POLogFile, "Error: Unposted Transaction "
   PRINT #POLogFile, "Record Number  :"; STR$(Cnt)
   PRINT #POLogFile, "Account Number :"; Tran2Post.AcctNum
   PRINT #POLogFile, "Date           :"; Num2Date(Tran2Post.TRDate)
   PRINT #POLogFile, "Description    :"; Tran2Post.Desc
   PRINT #POLogFile, "Debit          :"; STR$(Tran2Post.CrAmt)
   PRINT #POLogFile, "Credit         :"; STR$(Tran2Post.DrAmt)
   PRINT #POLogFile,
RETURN

END SUB

'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
FUNCTION Round# (N#)
   Round# = INT(N# * 100 + .5) / 100
END FUNCTION

FUNCTION StripAcct$ (AcctNum$)

   X$ = QPTrim$(AcctNum$)  '(Form$(AcctNum, 0))
   DO
      DashPos = INSTR(X$, "-")
      IF DashPos > 0 THEN
         X$ = LEFT$(X$, DashPos - 1) + MID$(X$, DashPos + 1)
      END IF
    LOOP WHILE DashPos

    StripAcct$ = X$

END FUNCTION

