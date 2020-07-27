DEFINT A-Z

'$INCLUDE: 'PRO7.BI'
'$INCLUDE: 'GL.BI'
'$INCLUDE: 'Fileview.bi'
'$INCLUDE: 'DefCnf.bi'        'Defines a TYPE for monitor and color info..
'$INCLUDE: 'FieldInf.bi'      'field information
'$INCLUDE: 'FormEdit.bi'      'form editing information
'$INCLUDE: 'QSCR.BI'          'QuickScreen Declarations
'$INCLUDE: 'SCRAUX.BI'

'Subs in this module
DECLARE SUB OpenFundIdx (FundIdxFileNum%, NumFIdxRecs%)
DECLARE SUB OpenFundFile (FundFileNum%, NumFunds%)
DECLARE SUB OpenAcctIdx (AcctIdxFileNum%, NumIdxRecs%)
DECLARE SUB OpenAcctFile (AcctFileNum%, NumGLAccts%)
DECLARE SUB OpenDeptIdx (DeptIdxFile%, NumActiveDepts%)
DECLARE SUB OpenDeptFile (DeptFile%, NumDepts%)
DECLARE SUB OpenTransFile (TransFileNum%, NumTrans&)
DECLARE SUB SortFundIndex ()
DECLARE SUB SortAcctIndex ()
DECLARE SUB Post2GL (FileName$, BadTrans%)
DECLARE SUB CalcAcctBal (StartFund$, EndFund$, FundLen, EndDate)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PrintHelp (Help$)
DECLARE SUB GetFundCodes (FirstFund$, LastFund$)

DECLARE FUNCTION Round# (N#)
DECLARE FUNCTION FindFund% (FundNum$)
DECLARE FUNCTION FindAcct% (AcctNum$)
DECLARE FUNCTION FindDept% (DeptNum$)
DECLARE FUNCTION GetFundTitle$ (FundRecNum%)
DECLARE FUNCTION GetAcctTitle$ (AcctRecNum)
DECLARE FUNCTION GetAcctType$ (AcctRecNum)
DECLARE FUNCTION GetDeptTitle$ (DeptRecNum%)
DECLARE FUNCTION GetDeptOffsets$ ()
DECLARE FUNCTION FmtAcct$ (AcctNum$, FundLen%, AcctLen%, DetLen%)
DECLARE FUNCTION StripAcct$ (AcctNum$)
DECLARE FUNCTION FmtDate$ (D$)
DECLARE FUNCTION ConvDateStr$ (D$)
DECLARE FUNCTION ConvDateStr1$ (D$)
DECLARE FUNCTION GetBegDate% ()
DECLARE FUNCTION GetEndDate% ()
DECLARE SUB GetFYDates (FY1BegDate, FY1EndDate, FY2BegDate, FY2EndDate)
DECLARE SUB GetPostDates (LPDate, HPDate)

DECLARE FUNCTION GetPct$ (N1#, N2#)

DIM SHARED FundIdx AS GLFundIndexType
DIM SHARED Fund AS GLFundRecType
DIM SHARED AcctIdx AS GLAcctIndexType
DIM SHARED Acct AS GLAcctRecType
DIM SHARED Trans AS GLTransRecType
DIM SHARED DeptIdx AS GLDeptIndexType
DIM SHARED Dept AS GLDeptRecType

''$INCLUDE: 'SetCnf.BI'

SUB CalcAcctBal (StartFund$, EndFund$, FundLen, EndDate)

  SHARED AcctIdx AS GLAcctIndexType
  SHARED Acct AS GLAcctRecType
  SHARED Trans AS GLTransRecType

  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum, NumGLAcctRecs
  OpenTransFile TransFileNum, NumTrans&
  
  FOR Cnt = 1 TO NumGLAccts
    GET AcctIdxFileNum, Cnt, AcctIdx
    GET AcctFileNum, AcctIdx.RecNum, Acct

    FundCode$ = LEFT$(Acct.Num, FundLen)
    IF FundCode$ >= StartFund$ AND FundNumber$ <= EndFund$ THEN

      CalcBal# = Round#(Acct.BegBal)            'get the beginning balance
      NextTr& = Acct.FrstTran   'get the first trans for this acct

      QPrintRC Acct.Num, 25, 14, -1

      DO UNTIL NextTr& = 0      'keep going 'til we run out of trans
        GET TransFileNum, NextTr&, Trans

        IF Trans.TRDate <= EndDate THEN

          SELECT CASE Acct.Typ
          CASE "A", "E"
            CalcBal# = CalcBal# + Round#(Trans.DrAmt) - Round#(Trans.CrAmt)
          CASE "L", "R"
            CalcBal# = CalcBal# + Round#(Trans.CrAmt) - Round#(Trans.DrAmt)
          END SELECT

        END IF

        NextTr& = Trans.NextTran                'Get the next transaction

      LOOP

      Acct.Bal = CalcBal#
      PUT AcctFileNum, AcctIdx.RecNum, Acct

    END IF   'test for account in fund range
  NEXT       'next account

  CLOSE

END SUB

'****************************************************************************
'  D$ is a date string in the format 03-15-1995.
'  Returns: 6 character string in the format 031595
'
'****************************************************************************
FUNCTION ConvDateStr$ (D$)
   ConvDateStr$ = LEFT$(D$, 2) + MID$(D$, 4, 2) + RIGHT$(D$, 2)
END FUNCTION

'
'****************************************************************************
'  D$ is a date string in the format 03-15-1995.
'  Returns: 6 character date string in the format 03-15-95
'****************************************************************************
'
FUNCTION ConvDateStr1$ (D$)
   ConvDateStr1$ = LEFT$(D$, 2) + "-" + MID$(D$, 4, 2) + "-" + RIGHT$(D$, 2)
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

FUNCTION FindFund (FundNum$)
'*****************************************************************************
'Searches the acct index for a matching fund number and returns the record
'number of the fund
'
'    Input: AcctNum$ as a formatted G/L account number string
'  Returns: Record number of the account
'*****************************************************************************
'
   OpenFundIdx FundIdxFileNum, NumFIdxRecs

   IF NumFIdxRecs = 0 THEN
      RecordNum = 0
      CLOSE FundIdxFileNum
      EXIT FUNCTION
   END IF

   Match = FALSE
   FirstRec = 1
   LastRec = NumFIdxRecs
   Lookfor$ = FundNum$ 'LTRIM$(RTRIM$(FundNum$))

   DO UNTIL LastRec < FirstRec

      MiddleRec = (LastRec + FirstRec) \ 2

      GET FundIdxFileNum, MiddleRec, FundIdx

      IF Lookfor$ = RTRIM$(FundIdx.FundNum) THEN
         Match = -1
         EXIT DO
      ELSEIF Lookfor$ < FundIdx.FundNum THEN
         LastRec = MiddleRec - 1
      ELSE
         FirstRec = MiddleRec + 1
      END IF

   LOOP

   IF Match THEN
      RecNum = FundIdx.RecNum
   ELSE
      RecNum = 0
   END IF

   FindFund = RecNum

   CLOSE FundIdxFileNum


END FUNCTION

'****************************************************************************
'formats an account number string with dashes.
'****************************************************************************
'
FUNCTION FmtAcct$ (AN$, FundLen%, AcctLen%, DetLen%)

'   AN$ = QPTrim$(AN$)
'   FmtAcct$ = LEFT$(AN$, FundLen) + "-" + MID$(AN$, FundLen + 1, AcctLen) + "-" + RIGHT$(AN$, DetLen) 'MID$(AN$, FundLen + AcctLen + 2, DetLen) 'RIGHT$(AN$, DetLen)

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

'****************************************************************************
'Converts a date string in the format "######" to ##-##-##
'****************************************************************************
FUNCTION FmtDate$ (D$)
   D$ = LTRIM$(RTRIM$(D$))
   FmtDate = LEFT$(D$, 2) + "-" + MID$(D$, 3, 2) + "-" + RIGHT$(D$, 2)
END FUNCTION

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

'****************************************************************************
'Retrieves the GL account type from the account data file.
'****************************************************************************
'
FUNCTION GetAcctType$ (AcctRecNum)

   OpenAcctFile AcctFileNum, NumAccts
   GET AcctFileNum, AcctRecNum, Acct
   GetAcctType$ = Acct.Typ
   CLOSE AcctFileNum

END FUNCTION

FUNCTION GetBegDate

   DIM GLSetup AS GLSetupRecType
   SetUpRecLen = LEN(GLSetup)

   SetupFile = FREEFILE
   OPEN "GLSETUP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS SetupFile LEN = SetUpRecLen

   GET SetupFile, 1, GLSetup

   GetBegDate = GLSetup.FYBeg

   CLOSE SetupFile


END FUNCTION

FUNCTION GetDeptOffsets$
'---------------------------------------------------
'Gets the Dept Code from the setup file.  The
'code determines where to look for the dept code in
'the acct code
'---------------------------------------------------
   DIM GLSetup AS GLSetupRecType
   SetUpRecLen = LEN(GLSetup)

   SetupFile = FREEFILE
   OPEN "GLSETUP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS SetupFile LEN = SetUpRecLen

   GET SetupFile, 1, GLSetup

   GetDeptOffsets$ = GLSetup.DeptCode

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

SUB GetFundCodes (FirstFund$, LastFund$)
   
   OpenFundIdx FundIdxFileNum, NumFIdxRecs

   IF NumFIdxRecs = 0 THEN
      FirstFund = 0
      LastFund = 0
      EXIT SUB
   END IF
   
   GET FundIdxFileNum, 1, FundIdx
   FirstFund$ = QPTrim$(FundIdx.FundNum)

   GET FundIdxFileNum, NumFIdxRecs, FundIdx
   LastFund$ = QPTrim$(FundIdx.FundNum)
   
   CLOSE FundIdxFileNum
   
END SUB

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

'****************************************************************************
'Retrieves the fund title from the fund data file.
'****************************************************************************
'
FUNCTION GetFundTitle$ (FundRecNum)

   DIM FundRec AS GLFundRecType
   OpenFundFile FundFileNum, NumFunds
   GET FundFileNum, FundRecNum, FundRec
   GetFundTitle$ = FundRec.Title
   CLOSE FundFileNum

END FUNCTION

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

FUNCTION GetPct$ (N1#, N2#)
  
  Pct$ = SPACE$(5)
  IF N1# > 0 AND N2# > 0 THEN
    P# = Round#((N1# / N2#) * 100)
    P$ = STR$(INT(P#)) + "%"
    RSET Pct$ = P$
  END IF
  GetPct$ = Pct$
  
END FUNCTION

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

SUB OpenAPCtrlFile (CtrlFileNum)

   CtrlFileNum = FREEFILE
   CtrlRecLen = LEN(CtrlFile)
   OPEN "APCONTRL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS CtrlFileNum LEN = CtrlRecLen

END SUB

SUB OpenDeptFile (DeptFile%, NumDepts%)

   DeptRecLen = LEN(Dept)
   DeptFile = FREEFILE
   OPEN "GLDEPT.DAT " FOR RANDOM ACCESS READ WRITE SHARED AS DeptFile LEN = DeptRecLen
   NumDepts = LOF(DeptFile) \ DeptRecLen


'SUB OpenFundFile (FundFileNum%, NumFunds%)

'   FundRecLen = LEN(Fund)
'   FundFileNum = FREEFILE
'   OPEN "GLFUND.DAT " FOR RANDOM ACCESS READ WRITE SHARED AS FundFileNum LEN = FundRecLen
'   NumFunds = LOF(FundFileNum) \ FundRecLen

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

SUB OpenTransFile (TransFileNum, NumTrans&)

   TransRecLen = LEN(Trans)
   TransFileNum = FREEFILE
   OPEN "GLTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS TransFileNum LEN = TransRecLen
   NumTrans& = LOF(TransFileNum) \ TransRecLen

END SUB

'****************************************************************************
' Input: FileName$ is the edit file to be posted, which is in the same type
'        as the transaction history (BATRANS.DAT) file
' BadTrans returns the record number of a transaction which was not posted
'****************************************************************************
'
SUB Post2GL (FileName$, BadTrans%)

   SHARED Acct AS GLAcctRecType, Trans AS GLTransRecType

   DIM Tran2Post AS GLTransRecType        'Dim a buffer for the edit file
   TrRecLen = LEN(Tran2Post)              'Determine the rec length
   File2Post = FREEFILE                   'Get a handle

   OPEN FileName$ FOR RANDOM AS File2Post LEN = TrRecLen

   Num2Post& = LOF(File2Post) \ TrRecLen   'Find the num of transactions

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

   FOR Cnt& = 1 TO Num2Post&                'Start processing transactions
      GET File2Post, Cnt&, Tran2Post

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
           PUT File2Post, Cnt&, Tran2Post

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
              Acct.LastTran = NumTrans&        'reset last trans to this trans
              PUT AcctFileNum, RecNum, Acct

                                            'In the trans file...
              GET TransFileNum, Prev&, Trans    'Get the last transaction
              Trans.NextTran = NumTrans&       'reset pointer to this trans
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
     PRINT #GLLogFile, USING "Debits Posted  : ########,.##"; DrPosted#
     PRINT #GLLogFile, USING "Credits Posted : ########,.##"; CrPosted#
     PRINT #GLLogFile, STRING$(78, "-")
   END IF

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
   PRINT #GLLogFile, "Record Number  :"; STR$(Cnt&)
   PRINT #GLLogFile, "Account Number :"; Tran2Post.AcctNum
   PRINT #GLLogFile, "Date           :"; Num2Date(Tran2Post.TRDate)
   PRINT #GLLogFile, "Description    :"; Tran2Post.Desc
   PRINT #GLLogFile, "Debit          :"; STR$(Tran2Post.CrAmt)
   PRINT #GLLogFile, "Credit         :"; STR$(Tran2Post.DrAmt)
   PRINT #GLLogFile,
RETURN



END SUB

'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
FUNCTION Round# (N#)
   Round# = INT(N# * 100 + .5) / 100
END FUNCTION

SUB SortAcctIndex

   IF Exist("GLACCT.IDX") THEN                'Delete index file if it exits
      KILL "GLACCT.IDX"
   END IF

   OpenAcctIdx AcctIdxFileNum, NumIdxRecs     'Open the Index file
 
   OpenAcctFile AcctFileNum, NumAccts         'Open the Acct file

   IF LOF(AcctFileNum) = 0 THEN               'get out if nothing to do
      CLOSE
      EXIT SUB
   END IF
   
   REDIM IdxBuff(1 TO NumAccts) AS GLAcctIndexType  'dim the buffer array

   FOR Cnt = 1 TO NumAccts                    'Load the buffer with accts
      GET AcctFileNum, Cnt, Acct
      IF Acct.Deleted = 0 THEN                'Get only active accounts
         GoodAccts = GoodAccts + 1
         IdxBuff(GoodAccts).AcctNum = Acct.Num
         IdxBuff(GoodAccts).RecNum = Cnt
      END IF
   NEXT

   CLOSE AcctFileNum                          'Close the acct file

   REDIM PRESERVE IdxBuff(1 TO GoodAccts) AS GLAcctIndexType  'redim with
                                                              'just good accts

   SortT IdxBuff(1), GoodAccts, 0, 16, 0, 14 'sort the $#&+ out of it!

   FOR Cnt = 1 TO GoodAccts                   'write out to the index file
      GET AcctIdxFileNum, Cnt, AcctIdx
      AcctIdx.AcctNum = IdxBuff(Cnt).AcctNum
      AcctIdx.RecNum = IdxBuff(Cnt).RecNum
      PUT AcctIdxFileNum, Cnt, AcctIdx
   NEXT

   CLOSE AcctIdxFileNum                       'close index
 

END SUB

SUB SortFundIndex

   IF Exist("GLFUND.IDX") THEN                'Delete index file if it exits
      KILL "GLFUND.IDX"
   END IF

   OpenFundIdx FundIdxFileNum, NumFFIdxRecs   'Open the Index file
 
   OpenFundFile FundFileNum, NumFunds         'Open the Fund file

   IF LOF(FundFileNum) = 0 THEN               'get out if nothing to do
      CLOSE
      EXIT SUB
   END IF
   
   REDIM IdxBuff(1 TO NumFunds) AS GLFundIndexType  'dim the buffer array

   FOR Cnt = 1 TO NumFunds                    'Load the buffer with Funds
      GET FundFileNum, Cnt, Fund
      IF Fund.Deleted = 0 THEN                'Get only active accounts
         GoodFunds = GoodFunds + 1
         IdxBuff(GoodFunds).FundNum = Fund.FundNum
         IdxBuff(GoodFunds).RecNum = Cnt
      END IF
   NEXT

   CLOSE FundFileNum                          'Close the Fund file

   IF GoodFunds = 0 THEN                      'Exit if all are deleted
      CLOSE FundIdxFileNum
      EXIT SUB
   END IF

   REDIM PRESERVE IdxBuff(1 TO GoodFunds) AS GLFundIndexType  'redim with
   SortT IdxBuff(1), GoodFunds, 0, 6, 0, 4 'sort the $#&+ out of it!

   FOR Cnt = 1 TO GoodFunds                   'write out to the index file
      GET FundIdxFileNum, Cnt, FundIdx
      FundIdx.FundNum = IdxBuff(Cnt).FundNum
      FundIdx.RecNum = IdxBuff(Cnt).RecNum
      PUT FundIdxFileNum, Cnt, FundIdx
   NEXT

   CLOSE FundIdxFileNum                       'close index
 

END SUB

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

