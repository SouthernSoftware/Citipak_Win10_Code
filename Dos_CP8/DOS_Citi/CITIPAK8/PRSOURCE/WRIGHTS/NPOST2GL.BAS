DEFINT A-Z

DECLARE SUB StuffBuf (ky$)
DECLARE FUNCTION Round# (N#)
DECLARE SUB ButtonPress (ButNo%, Down%, Presses%, X%, Y%)
DECLARE SUB GetCursor (X, Y, Button)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB BLPrint (LPTPort%, Text$, ErrCode%)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE SUB FileView (FileName$, ky, Action, FVI AS ANY, SEG Array)
DECLARE FUNCTION FUsing$ (Number$, Mask$)
DECLARE SUB HideCursor ()
DECLARE FUNCTION MGetKey% (Row%, Col%)
DECLARE FUNCTION Monitor% ()
DECLARE SUB MScrnSave (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MScrnRest (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MPaintBox (ULRow%, ULCol%, LRRow%, LRCol%, Colr%)
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION Peek1% (Segment%, Address%)
DECLARE SUB QPrint (X$, Colr%, page%)
DECLARE SUB QPrintRC (t$, r%, c%, Clr%)
DECLARE FUNCTION QPTrim$ (X$)
DECLARE SUB ShowCursor ()
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE FUNCTION WaitKey% ()
'Subs in this module
DECLARE SUB OpenFundIdx (FundIdxFileNum%, NumFIdxRecs%)
DECLARE SUB OpenFundFile (FundFileNum%, NumFunds%)
DECLARE SUB OpenAcctIdx (AcctIndexName$, AcctIdxFileNum%, NumIdxRecs%)
DECLARE SUB OpenAcctFile (AcctFileName$, AcctFileNum%, NumGLAccts%)
DECLARE SUB OpenDeptIdx (DeptIdxFile%, NumActiveDepts%)
DECLARE FUNCTION FindDept% (DeptNum$)
DECLARE SUB OpenTransFile (TransFileName$, TransFileNum%, NumTrans&)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PrintHelp (Help$)
DECLARE FUNCTION FindFund% (FundNum$)
DECLARE FUNCTION FindAcct% (AcctIndexName$, AcctNum$)
DECLARE SUB GetFundCodes (FirstFund$, LastFund$)
DECLARE FUNCTION GetFundTitle$ (FundRecNum%)
DECLARE FUNCTION GetAcctTitle$ (AcctRecNum)
DECLARE FUNCTION GetAcctType$ (AcctRecNum)
DECLARE FUNCTION FmtAcct$ (AcctNum$, FundLen%, AcctLen%, DetLen%)
DECLARE FUNCTION FmtDate$ (D$)
DECLARE FUNCTION ConvDateStr$ (D$)
DECLARE FUNCTION ConvDateStr1$ (D$)
DECLARE FUNCTION GetBegDate% ()
DECLARE FUNCTION GetEndDate% ()
DECLARE FUNCTION Round# (N#)
DECLARE SUB Post2GL (FileName$, PSysRec() AS ANY, BadTrans%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)

''$INCLUDE: 'GL.BI'
''$INCLUDE: 'Fileview.bi'
''$INCLUDE: 'DefCnf.bi'        'Defines a TYPE for monitor and color info..
''$INCLUDE: 'FieldInf.bi'      'field information
''$INCLUDE: 'FormEdit.bi'      'form editing information
''$INCLUDE: 'QSCR.BI'          'QuickScreen Declarations
''$INCLUDE: 'SCRAUX.BI'

'$INCLUDE: 'fileview.bi'
'$INCLUDE: 'GL.BI'
''$INCLUDE: 'SCRAUX.BI'
'$INCLUDE: 'PRSYSCTR.Bi'
'$INCLUDE: 'PRFILES.Bi'

DECLARE SUB LibLoadDisplayForm (LibName$, FormName$, Form$(), Fld() AS ANY)
DECLARE SUB EditForm (Form$(), Fld() AS ANY, Frm AS ANY, Cnf AS ANY, Action%)
'DECLARE SUB LibGetFldDef (LibName$, FrmName$, StartEl%, Fld() AS ANY, Form$(), ErrCode%)
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE FUNCTION InputBox$ (LibName$, FormName$, Cancel%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION GetPassword$ (Row%, Col%, Max%, Abort%)
DECLARE SUB PrintHelp (H$)
DECLARE SUB DispLoad ()


DIM SHARED AcctIdx AS GLAcctIndexType
DIM SHARED Acct AS GLAcctRecType
DIM SHARED Trans AS GLTransRecType

'*****************************************************************************
'Searches the acct index for a matching account number and returns the record
'number of the account
'
'    Input: AcctNum$ as a formatted G/L account number string
'  Returns: Record number of the account
'*****************************************************************************
'
FUNCTION FindAcct (AcctIndexName$, AcctNum$)

   OpenAcctIdx AcctIndexName$, AcctIdxFileNum, NumIdxRecs

   IF NumIdxRecs = 0 THEN
      RecordNum = 0
      CLOSE AcctIdxFileNum
      EXIT FUNCTION
   END IF

   Match = FALSE
   FirstRec = 1
   LastRec = NumIdxRecs

   Lookfor$ = QPTrim$(AcctNum$)

   DO UNTIL LastRec < FirstRec

      MiddleRec = (LastRec + FirstRec) \ 2

      GET AcctIdxFileNum, MiddleRec, AcctIdx

      TLoAcct$ = QPTrim$(AcctIdx.AcctNum)

      IF TLoAcct$ = Lookfor$ THEN
        Match = -1
        EXIT DO
      ELSEIF Lookfor$ < TLoAcct$ THEN
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

'****************************************************************************
'formats an account number string with dashes.
'****************************************************************************
FUNCTION FmtAcct$ (AN$, FundLen%, AcctLen%, DetLen%)

  AN$ = QPTrim$(AN$)

  FmtAcct$ = LEFT$(AN$, FundLen) + "-" + MID$(AN$, FundLen + 1, AcctLen) + "-" + MID$(AN$, FundLen + AcctLen + 1, DetLen)

  'RIGHT$(AN$, DetLen) 'MID$(AN$, FundLen + AcctLen + 2, DetLen) 'RIGHT$(AN$, DetLen)

END FUNCTION

SUB OpenAcctFile (AcctFileName$, AcctFileNum, NumAccts)

   AcctRecLen = LEN(Acct)
   AcctFileNum = FREEFILE

   OPEN AcctFileName$ FOR RANDOM ACCESS READ WRITE SHARED AS AcctFileNum LEN = AcctRecLen
   NumAccts = LOF(AcctFileNum) \ AcctRecLen

END SUB

SUB OpenAcctIdx (AcctIndexName$, AcctIdxFileNum, NumIdxRecs)

   AcctIdxLen = LEN(AcctIdx)
   AcctIdxFileNum = FREEFILE
   OPEN AcctIndexName$ FOR RANDOM ACCESS READ WRITE SHARED AS AcctIdxFileNum LEN = AcctIdxLen
   NumIdxRecs = LOF(AcctIdxFileNum) \ AcctIdxLen

END SUB

SUB OpenTransFile (TransFileName$, TransFileNum, NumTrans&)

   TransRecLen = LEN(Trans)
   TransFileNum = FREEFILE
   OPEN TransFileName$ FOR RANDOM ACCESS READ WRITE SHARED AS TransFileNum LEN = TransRecLen
   NumTrans& = LOF(TransFileNum) \ TransRecLen

END SUB

'****************************************************************************
' Input: FileName$ is the edit file to be posted, which is in the same type
'        as the transaction history (BATRANS.DAT) file
' BadTrans returns the record number of a transaction which was not posted
'****************************************************************************
'
SUB Post2GL (FileName$, PSysRec() AS RegDSysFileRecType, BadTrans%)

   SHARED Acct AS GLAcctRecType, Trans AS GLTransRecType
   
   SysDir$ = QPTrim$(PSysRec(1).CITIDIR)

   IF RIGHT$(SysDir$, 1) <> "\" THEN
     SysDir$ = SysDir$ + "\"
   END IF
'STOP
   AcctFileName$ = SysDir$ + "GLACCT.DAT"
   TransFileName$ = SysDir$ + "GLTRANS.DAT"
   AcctIndexName$ = SysDir$ + "GLACCT.IDX"

   DIM Tran2Post AS GLTransRecType        'Dim a buffer for the edit file
   TrRecLen = LEN(Tran2Post)              'Determine the rec length

   File2Post = FREEFILE                   'Get a handle

   OPEN FileName$ FOR RANDOM AS File2Post LEN = TrRecLen

   Num2Post = LOF(File2Post) \ TrRecLen   'Find the num of transactions
   
   OpenAcctFile AcctFileName$, AcctFileNum, NumAccts    'Open & lock GL files
   'LOCK AcctFileNum

   OpenTransFile TransFileName$, TransFileNum, NumTrans&
   'LOCK TransFileNum

   FOR Cnt = 1 TO Num2Post                'Start processing transactions

      GET File2Post, Cnt, Tran2Post
      
      RecNum = FindAcct(AcctIndexName$, Tran2Post.AcctNum)  'Verify account is in G/L

      IF RecNum > 0 THEN                  'if valid acct then proceed

         'tell user what's going on
         QPrintRC " Posting Account Number: ", 25, 1, 112
         QPrintRC Tran2Post.AcctNum, 25, 26, 112

         GET AcctFileNum, RecNum, Acct    'Get the account

         'depending on account type, update running balance
         'Nick was updating MTD & YTD fields here also.

         SELECT CASE Acct.Typ
            CASE "A", "E"                 'asset, exp accts
               Acct.Bal = Round#(Acct.Bal) + Round#(Tran2Post.DrAmt) - Round#(Tran2Post.CrAmt)
               PUT AcctFileNum, RecNum, Acct
           
            CASE "L", "R"                 'liab, rev accts
               Acct.Bal = Round#(Acct.Bal) + Round#(Tran2Post.CrAmt) - Round#(Tran2Post.DrAmt)
               PUT AcctFileNum, RecNum, Acct

         END SELECT

         NumTrans& = NumTrans& + 1          'increment record pointer
         GET TransFileNum, NumTrans&, Trans
         
         Trans.AcctNum = Tran2Post.AcctNum 'Assign editfile to trans history
         Trans.TrDate = Tran2Post.TrDate
         Trans.Desc = Tran2Post.Desc
         Trans.CrAmt = Tran2Post.CrAmt
         Trans.DrAmt = Tran2Post.DrAmt
         Trans.Ref = "" 'Tran2Post.Ref
         Trans.Src = Tran2Post.Src
         Trans.NextTran = 0

         PUT TransFileNum, NumTrans&, Trans

         Posted = Posted + 1

         '---------------------------------Start linking here
         IF Acct.FrstTran = 0 THEN        'if first trans for this acct,
            Acct.FrstTran = NumTrans&      'assign first & last pointers to
            Acct.LastTran = NumTrans&      'this transaction
            PUT AcctFileNum, RecNum, Acct
         
         ELSE                             'otherwise
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
                                          'how about an error log here.
      END IF

   NEXT

   'UNLOCK AcctFileNum
   'UNLOCK TransFileNum

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

END SUB

