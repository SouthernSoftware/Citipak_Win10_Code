DEFINT A-Z
DECLARE FUNCTION FindAcct% (FileName$, AcctNumber$)
DECLARE SUB OpenAcctFile (FileName$, BAAcctFile%, NumAccts%)
DECLARE SUB OpenAcctIdx (FileName$, BAAcctIdx%, ActiveAccts%)
DECLARE SUB OpenTransFile (FileName$, BATransFile%, NumTrans&)
DECLARE SUB Post2GL (FileName$, BadTrans%)
DECLARE FUNCTION Round# (N#)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE SUB WaitForAction ()
DECLARE SUB DisplayMiscScrn (ScrnName$)
DECLARE SUB BlockClear ()

'$INCLUDE: 'barecs.bi'
'$INCLUDE: 'prfiles.bi'


DIM SHARED AcctIdx AS BAAcctNdxType
DIM SHARED Acct AS BAAcctRecType
DIM SHARED Trans AS BATransRecType




 'CLS

 'Post2BA "c:\citipak\prif.dat", BadTrans%
 
 'PRINT "Number of bad trans: "; BadTrans%

FUNCTION FindAcct (FileName$, AcctNumber$)

   SHARED AcctIdx AS BAAcctNdxType

   AcctNum# = VAL(LEFT$(AcctNumber$, 12))

   IF AcctNum# < 999999 THEN AcctNum# = AcctNum# * 100

   OpenAcctIdx FileName$, BAAcctIdxFileNum, NumIdxRecs

   IF NumIdxRecs = 0 THEN
      RecordNum = 0
      CLOSE BAAcctIdxFileNum
      EXIT FUNCTION
   END IF

   Match = FALSE
   FirstRec = 1
   LastRec = NumIdxRecs
   Lookfor# = AcctNum#

   DO UNTIL LastRec < FirstRec

      MiddleRec = (LastRec + FirstRec) \ 2

      GET BAAcctIdxFileNum, MiddleRec, AcctIdx

      IF Lookfor# = AcctIdx.AcctNum THEN
         Match = -1
         EXIT DO
      ELSEIF Lookfor# < AcctIdx.AcctNum THEN
         LastRec = MiddleRec - 1
      ELSE
         FirstRec = MiddleRec + 1
      END IF

   LOOP

   IF Match THEN
      FindAcct = AcctIdx.RecNum
   ELSE
      FindAcct = 0
   END IF

   CLOSE BAAcctIdxFileNum

END FUNCTION

SUB OpenAcctFile (FileName$, BAAcctFile, NumAccts)

   TempFileName$ = QPTrim$(FileName$) + "\baacct.dat"

   SHARED Acct AS BAAcctRecType

   AcctLen = LEN(Acct)
   BAAcctFile = FREEFILE
   'OPEN "baacct.dat" FOR RANDOM ACCESS READ WRITE SHARED AS BAAcctFile LEN = AcctLen
   OPEN TempFileName$ FOR RANDOM ACCESS READ WRITE SHARED AS BAAcctFile LEN = AcctLen
   'OPEN TempFileName$ FOR RANDOM LOCK WRITE AS BAAcctFile LEN = AcctLen
   'LOCK AcctFileNum

   NumAccts = LOF(BAAcctFile) \ AcctLen
   
END SUB

SUB OpenAcctIdx (FileName$, BAAcctIdx, ActiveAccts)

   TempFileName$ = QPTrim$(FileName$) + "\baacctdx.dat"

   SHARED Ndx AS BAAcctNdxType

   IndexLen = LEN(Ndx)
   BAAcctIdx = FREEFILE
   OPEN TempFileName$ FOR RANDOM ACCESS READ WRITE SHARED AS BAAcctIdx LEN = IndexLen

   ActiveAccts = LOF(BAAcctIdx) \ IndexLen

END SUB

SUB OpenTransFile (FileName$, BATransFile, NumTrans&)
   TempFileName$ = QPTrim$(FileName$) + "\batrans.dat"

   SHARED Trans AS BATransRecType

   TransLen = LEN(Trans)
   BATransFile = FREEFILE

   OPEN TempFileName$ FOR RANDOM ACCESS READ WRITE SHARED AS BATransFile LEN = TransLen'85
   
   NumTrans& = LOF(BATransFile) \ TransLen
   
END SUB

'****************************************************************************
' Input: FileName$ is the edit file to be posted, which is in the same type
'        as the transaction history (BATRANS.DAT) file
' BadTrans returns the record number of a transaction which was not posted
'****************************************************************************
'
SUB Post2BA (FileName$, BadTrans%)

'   ON LOCAL ERROR GOTO GotErr

   TempFileName$ = QPTrim$(FileName$)

   'PRIFName$ = TempFileName$ + "\TEMPIF.DAT"

   PRIFName$ = TempFileName$ + "\PRIF.DAT"

   SHARED Acct AS BAAcctRecType, Trans AS BATransRecType
   
   DIM Tran2Post AS BATransRecType     'Dim a buffer for the edit file

   TrRecLen = LEN(Tran2Post)              'Determine the rec length
   File2Post = FREEFILE                   'Get a handle

   'oPEN PRIFName$ FOR RANDOM AS File2Post LEN = TrRecLen

   OPEN "TempIf.dat" FOR RANDOM AS File2Post LEN = TrRecLen
'Changed this to fix the problem with the payroll posting

   Num2Post = LOF(File2Post) \ TrRecLen   'Find the num of transactions
   
   OpenAcctFile FileName$, AcctFileNum, NumAccts     'Open & lock GL files

   AtFlag = 1
   'LOCK AcctFileNum

   OpenTransFile FileName$, TransFileNum, NumTrans&

   AtFlag = 2
   'LOCK TransFileNum

   FOR Cnt = 1 TO Num2Post                'Start processing transactions

      GET File2Post, Cnt, Tran2Post
      
      RecNum = FindAcct(FileName$, Tran2Post.AcctNum)  'Verify account is in G/L

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
         Trans.Ref = Tran2Post.Ref
         Trans.Source = Tran2Post.Source
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

ExitPost:

EXIT SUB


'was printing register and deleteing edit file here.
'Now do this in module that called this sub

GotErr:
  SELECT CASE AtFlag
    CASE 2
      'UNLOCK AcctFileNum
    CASE ELSE
  END SELECT
  CLOSE
  BlockClear
  DisplayMiscScrn LockError
  WaitForAction
  BadTrans% = -1
  RESUME ExitPost

END SUB

