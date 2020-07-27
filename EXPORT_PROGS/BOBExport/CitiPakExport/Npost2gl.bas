Attribute VB_Name = "Module1"

Dim AcctIdx As GLAcctIndexType
Dim Acct As GLAcctRecType
Dim Trans As GLTransRecType

'*****************************************************************************
'Searches the acct index for a matching account number and returns the record
'number of the account
'
'    Input: AcctNum$ as a formatted G/L account number string
'  Returns: Record number of the account
'*****************************************************************************
'
Function FindAcct(AcctIndexName$, AcctNum$)
   Dim NumIdxRecs As Long
   Dim Match As Boolean
   Dim FirstRec As Integer
   Dim LastRec As Long, Lookfor$
   Dim MiddleRec As Long, TLoAcct$
   Dim RecdNum As Long
   Dim AcctIdxFileNum As Integer
   
   OpenGLAcctIdx AcctIdxFileNum
   NumIdxRecs = LOF(AcctIdxFileNum) / Len(AcctIdx)
   If NumIdxRecs = 0 Then
      recordNum = 0
      Close AcctIdxFileNum
      Exit Function
   End If

   Match = False
   FirstRec = 1
   LastRec = NumIdxRecs

   Lookfor$ = QPTrim$(AcctNum$)

   Do Until LastRec < FirstRec

      MiddleRec = (LastRec + FirstRec) \ 2

      Get AcctIdxFileNum, MiddleRec, AcctIdx

      TLoAcct$ = QPTrim$(AcctIdx.AcctNum)

      If TLoAcct$ = Lookfor$ Then
        Match = -1
        Exit Do
      ElseIf Lookfor$ < TLoAcct$ Then
        LastRec = MiddleRec - 1
      Else
         FirstRec = MiddleRec + 1
      End If

   Loop

   If Match Then
      RecdNum = AcctIdx.RecNum
   Else
      RecdNum = 0
   End If

   FindAcct = RecdNum

   Close AcctIdxFileNum

End Function

'****************************************************************************
'formats an account number string with dashes.
'****************************************************************************
Function FmtAcct$(AN$, FundLen%, AcctLen%, DetLen%)

  AN$ = QPTrim$(AN$)

  FmtAcct$ = Left$(AN$, FundLen) + "-" + Mid$(AN$, FundLen + 1, AcctLen) + "-" + Mid$(AN$, FundLen + AcctLen + 1, DetLen)

  'RIGHT$(AN$, DetLen) 'MID$(AN$, FundLen + AcctLen + 2, DetLen) 'RIGHT$(AN$, DetLen)

End Function

'****************************************************************************
' Input: FileName$ is the edit file to be posted, which is in the same type
'        as the transaction history (BATRANS.DAT) file
' BadTrans returns the record number of a transaction which was not posted
'****************************************************************************
'
Sub Post2GL(FileName$, PSysRec() As RegDSysFileRecType, BadTrans%)

   Dim SysDir$, AcctFileName$, TransFileName$
   Dim AcctIndexName$
   Dim Tran2Post As GLTransRecType        'Dim a buffer for the edit file
   Dim TrRecLen As Long
   Dim File2Post As Integer
   Dim Num2POst As Long
   Dim TransFileNum As Integer
   Dim NumAccts As Long
   Dim AcctFileNum As Integer
   Dim Acct As GLAcctRecType
   Dim AcctRecLen As Long
   Dim cnt As Long, Prev&
   Dim Posted As Long, NumTrans&
   Dim TransPosted As Long
   
'   SysDir$ = QPTrim$(PSysRec(1).CITIDIR)
   SysDir$ = CurrCitiPath

   If Right$(SysDir$, 1) <> "\" Then
     SysDir$ = SysDir$ + "\"
   End If

   TrRecLen = Len(Tran2Post)              'Determine the rec length

   File2Post = FreeFile                   'Get a handle

   Open FileName$ For Random As File2Post Len = TrRecLen

   Num2POst = LOF(File2Post) \ TrRecLen   'Find the num of transactions
   
   'LOCK AcctFileNum
   OpenGLAcctFile AcctFileNum 'GLACCT.DAT
   NumAccts = LOF(AcctFileNum) \ Len(Acct)

   'LOCK TransFileNum
   OpenGLTransFile TransFileNum 'GLTRANS.DAT
   NumTrans& = LOF(TransFileNum) / Len(Tran2Post)
   For cnt = 1 To Num2POst                'Start processing transactions
     Get File2Post, cnt, Tran2Post
     RecdNum = FindAcct(AcctIndexName$, Tran2Post.AcctNum)  'Verify account is in G/L
     If RecdNum > 0 Then                  'if valid acct then proceed
       Get AcctFileNum, RecdNum, Acct    'Get the account
       'depending on account type, update running balance
       'Nick was updating MTD & YTD fields here also.
  
       Select Case Acct.Typ
         Case "A", "E"                 'asset, exp accts
           Acct.Bal = OldRound#(Acct.Bal) + OldRound#(Tran2Post.DrAmt) - OldRound#(Tran2Post.CrAmt)
           Put AcctFileNum, RecdNum, Acct
         
         Case "L", "R"                 'liab, rev accts
           Acct.Bal = OldRound#(Acct.Bal) + OldRound#(Tran2Post.CrAmt) - OldRound#(Tran2Post.DrAmt)
           Put AcctFileNum, RecdNum, Acct
  
       End Select
       NumTrans& = NumTrans& + 1          'increment record pointer
       Get TransFileNum, NumTrans&, Trans
       Trans.AcctNum = Tran2Post.AcctNum 'Assign editfile to trans history
       Trans.TrDate = Tran2Post.TrDate
       Trans.Desc = Tran2Post.Desc
       Trans.CrAmt = Tran2Post.CrAmt
       Trans.DrAmt = Tran2Post.DrAmt
       Trans.Ref = "" 'Tran2Post.Ref
       Trans.Src = Tran2Post.Src
       Trans.NextTran = 0
  
       Put TransFileNum, NumTrans&, Trans
  
       Posted = Posted + 1
  
       '---------------------------------Start linking here
       If Acct.FrstTran = 0 Then        'if first trans for this acct,
         Acct.FrstTran = NumTrans&      'assign first & last pointers to
         Acct.LastTran = NumTrans&      'this transaction
         Put AcctFileNum, RecdNum, Acct
       Else                             'otherwise
         Prev& = Acct.LastTran             'remember the prev trans pointer,
         Acct.LastTran = NumTrans&        'reset last trans to this trans
         Put AcctFileNum, RecdNum, Acct
                                        'In the trans file...
         Get TransFileNum, Prev&, Trans    'Get the last transaction
         Trans.NextTran = NumTrans&       'reset pointer to this trans
         Put TransFileNum, Prev&, Trans
       End If
       TransPosted = TransPosted + 1
     Else                                'Account NOT found!
       BadTrans = BadTrans + 1          'Pass info back to caller
                                        'how about an error log here.
     End If

   Next

Close

Exit Sub

'was printing register and deleteing edit file here.
'Now do this in module that called this sub

End Sub

