VERSION 5.00
Begin VB.Form ChkCMTrans 
   Caption         =   "ChkCMTrans"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   2070
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "CM Transactions Checker"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   765
      TabIndex        =   5
      Top             =   360
      Width           =   3150
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   765
      TabIndex        =   4
      Top             =   705
      Width           =   3150
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Click 'OK' to Continue"
      Height          =   240
      Left            =   765
      TabIndex        =   3
      Top             =   1020
      Width           =   3150
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   765
      TabIndex        =   2
      Top             =   1305
      Width           =   3150
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1485
      TabIndex        =   1
      Top             =   1710
      Width           =   1725
   End
End
Attribute VB_Name = "ChkCMTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Dim TxTRFile As String
Dim CmTRFile As String
Dim TxCuFile As String

Dim blnDoneFix As Boolean
Dim blnInProcs As Boolean
Dim tCnt As Long
Dim DoRelink As Boolean
Dim BadCnt As Integer

Private Sub Form_Load()
  Call InitStuff
End Sub

Private Sub InitStuff()
  TxTRFile = "TAXTRANS.DAT"
  TxCuFile = "TaxCust.Dat"
  CmTRFile = "CMTrans.Dat"
  
  blnDoneFix = False
  blnInProcs = False
End Sub

Private Sub Command1_Click()
  If blnDoneFix Then
    End
  Else
    Command1.Enabled = False
    Call CheckCMTransactions
  End If
End Sub

Private Sub CheckCMTransactions()
  'Label2="Processing: "+
  blnInProcs = True
  DoRelink = False
  Label3.Caption = "Processing. . . "
  DoEvents
  
  Dim tstDate As Integer
  
  Dim TxTR As TaxTransactionType
  Dim CmTR As CMTransRecType
  Dim TxCust As TaxCustType

  Dim TxTRLen As Integer
  Dim CmTRLen As Integer
  Dim TxCuLen As Integer

  Dim TxTRCnt As Long
  Dim CmTRCnt As Long
  Dim TxCuCnt As Long
  
  TxTRLen = Len(TxTR)
  CmTRLen = Len(CmTR)
  TxCuLen = Len(TxCust)
  
  tstDate = Date2Num("01/01/2000")
  
  Open TxTRFile For Random As #1 Len = TxTRLen
  TxTRCnt = LOF(1) / TxTRLen
  Open TxCuFile For Random As #2 Len = TxCuLen
  TxCuCnt = LOF(2) / TxCuLen
  
  Open CmTRFile For Random As #3 Len = CmTRLen
  CmTRCnt = LOF(3) / CmTRLen
  
  Open "NotTheSameNames.txt" For Output As #4
  
  For tCnt = 1 To CmTRCnt
    Get #3, tCnt, CmTR
    Select Case CmTR.TransSource
    Case 30 To 39, 131, 231, 161, 261, 171, 271
      If CmTR.TransDate >= tstDate Then
        Get #2, CmTR.TransAcctNum, TxCust
        If InStr(TxCust.CustName, QPTrim(CmTR.TransName)) <= 0 Then
          Print #4, CmTR.TransAcctNum, TxCust.CustName; "<-->   "; CmTR.TransName; Num2Date(CmTR.TransDate)
          'BadCnt = BadCnt + 1
        End If
      End If
    
    End Select
  
  Next
  Print BadCnt
'ByeByeNow:
'
Close
Shell "notepad.exe NotTheSameNames.txt", vbNormalFocus
'
GoTo AllDoneExit:

DoneAll:
'  Call RelinkTransactions
  Label2 = ""
  Label4 = "" ' CStr(F2Cnt)
  Label5 = ""
  Label3 = "Correction Complete."

AllDoneExit:
  If DoRelink Then
    Call RelinkTransactions
  End If
  
  Label2 = ""
  Label4 = "" '"Removed: " + CStr(trTypeCnt)
  Label5 = ""
  Label3 = "Processing Complete."
  
  Command1.Caption = "EXIT"
  blnInProcs = False
  blnDoneFix = True
  Command1.Enabled = True
  
End Sub

'  Put #1, 200776, TaxTR
'  For TRCnt = 1 To TRRecCnt 'To 1 Step -1
'    Get #1, TRCnt, TaxTR
'    If TaxTR.TransDate = TRDATE And TaxTR.TranType = 1 Then
'      If TaxTR.DiscAmt > 0 Then
'        F1Cnt = F1Cnt + 1
'        'Print TaxTR.CustomerRec
'      Else
'        F2Cnt = F2Cnt + 1
'      End If
'    End If
'  Next
'  Print "f1: " & F1Cnt & "   f2: " & F2Cnt



'  Open TaxPersFile For Random As #1 Len = PPLen
'  TPCnt = LOF(1) / PPLen
'  For TRCnt = 1 To TPCnt
'    Get #1, TRCnt, PersPropRec
'    If PersPropRec.LastYrPrinted = 2015 Then
'      PersPropRec.LastYrPrinted = 2014
'      Put #1, TRCnt, PersPropRec
'    End If
'  Next
'  Close #1
  
'  Open TaxPropFile For Random As #1 Len = RPLen
'  TPCnt = LOF(1) / RPLen
'  For TRCnt = 1 To TPCnt
'    Get #1, TRCnt, RealPropRec
'    If RealPropRec.LastYrPrinted = 2015 Then
'      RealPropRec.LastYrPrinted = 2014
'      Put #1, TRCnt, RealPropRec
'    End If
'  Next
'  Close #1
  
'  Name "TAXTrans.dat" As "TRans.BAK"
'  Open "TRans.BAK" For Random As #1 Len = TRLen
  

'  TRRecCnt = LOF(1) / TRLen
'  For TRCnt = TRRecCnt To 1 Step -1
'    Get #1, TRCnt, IntTR
'    If IntTR.TranType = 4 Then
'      Get #1, IntTR.BelongTo, TaxTR
'      TaxTR.Revenue.Interest = wRound(TaxTR.Revenue.Interest - IntTR.Amount)
'      Put #1, IntTR.BelongTo, TaxTR
'      IntTR.Amount = 0
'      IntTR.Revenue.Interest = 0
'      IntTR.CustomerRec = -Abs(IntTR.CustomerRec)
'      Put #1, TRCnt, IntTR
'      trTypeCnt = trTypeCnt + 1
'    Else
'      Exit For
'    End If
'  Next
  
'  Get #1, 78045, TaxTR
'  TaxTR.Amount = 5.76
'  TaxTR.Revenue.Interest = 5.76
'  Put #1, 78045, TaxTR
'
'  Get #1, 74894, TaxTR
'  'TaxTR.Amount = 5.76
'  TaxTR.Revenue.Interest = 17.8
'  TaxTR.Revenue.InterestPd = 17.8
'  TaxTR.Revenue.Collection = 0
'  TaxTR.Revenue.CollectionPd = 0
'  Put #1, 74894, TaxTR
'
'  'TaxTR.Revenue.Collection = 0
'  Get #1, 63533, TaxTR
'  'TaxTR.Revenue.Collection = 0
'  TaxTR.Revenue.CollectionPd = 0
'  TaxTR.Amount = 122.19
'  'TaxTR.CustomerRec = Abs(TaxTR.CustomerRec)
'  Put #1, 63533, TaxTR
'
'  Get #1, 65493, TaxTR
'  TaxTR.Revenue.Principle1Pd = 173.82
'  Put #1, 65493, TaxTR
'
'  Get #1, 57149, TaxTR
'  TaxTR.Revenue.Collection = 0
'  TaxTR.Revenue.CollectionPd = 0
'  'TaxTR.CustomerRec = Abs(TaxTR.CustomerRec)
'  Put #1, 57149, TaxTR
'
'
'  Get #1, 184688, TaxTR
'  TaxTR.Amount = 31.58
'  TaxTR.Revenue.Principle1Pd = 31.58
'  Put #1, 184688, TaxTR
'
  'TRRecCnt = LOF(1) / TRLen
'  For TRCnt = 1 To TRRecCnt
'    Get #1, TRCnt, TaxTR

'    If TaxTR.TransDate = TRDATE Then
'      F2Cnt = F2Cnt + 1
'    Else
'      Put #2, , TaxTR
'    End If
'  Next
  
  'Print F2Cnt

'      If TaxTR.TranType = 1 Then
'        If TaxTR.TaxYear <= 2004 Then
'          Print #100, TRCnt, TaxTR.Revenue.Principle1, TaxTR.Revenue.Principle1Pd
'          Print #100, , TaxTR.Revenue.Interest, TaxTR.Revenue.InterestPd
'          Print #100, , TaxTR.Revenue.Penalty, TaxTR.Revenue.PenaltyPd
'
'          'Stop
'        End If
'      End If
'    End If
'
'
''      If TaxTR.Revenue.Principle1Pd > TaxTR.Revenue.Principle1 Then
''        TaxTR.Revenue.Principle1Pd = TaxTR.Revenue.Principle1
''        Put #1, TRCnt, TaxTR
''        F2Cnt = F2Cnt + 1
''      End If

Private Sub KillTransBAK()
  On Error GoTo WeDontCare
  Kill "TRANS.BAK"
WeDontCare:
End Sub

Private Sub Form_Unload(Cancel As Integer)
  'On Error Resume Next
  If blnInProcs Then
    Cancel = 1
  End If
End Sub

Private Sub RelinkTransactions()
  Dim TaxCust As TaxCustType
  Dim NumOfTCRecs As Long
  Dim TaxTran As TaxTransactionType
  Dim TCHandle As Integer, TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim Cnt As Long
  Dim TaxCustLen As Integer
  Dim TaxTransLen As Integer
  Dim TaxTransRate As TaxTransactionType

  TaxCustLen = Len(TaxCust)
  TaxTransLen = Len(TaxTransRate)
  
  TCHandle = FreeFile
  Open TaxCustFile For Random Shared As TCHandle Len = TaxCustLen
  NumOfTCRecs& = LOF(TCHandle) / Len(TaxCust)

  TTHandle = FreeFile
  Open TaxTransFile For Random Shared As TTHandle Len = TaxTransLen
  NumOfTTRecs& = LOF(TTHandle) / Len(TaxTransRate)
  
'Clear lasttrans pointer in customer records.
  For Cnt& = 1 To NumOfTCRecs&
    Get TCHandle, Cnt&, TaxCust
    TaxCust.LastTrans = 0
    Put TCHandle, Cnt&, TaxCust
  Next

  For Cnt& = 1 To NumOfTTRecs&
    Get TTHandle, Cnt&, TaxTran
    If TaxTran.CustomerRec > 0 And TaxTran.CustomerRec <= NumOfTCRecs& Then
      Get TCHandle, TaxTran.CustomerRec, TaxCust
      TaxTran.LastTrans = TaxCust.LastTrans
      TaxCust.LastTrans = Cnt&
      Put TCHandle, TaxTran.CustomerRec, TaxCust
      Put TTHandle, Cnt&, TaxTran
    End If
    If Cnt& Mod 100 = 0 Then 'update display every 100 recs
      Label4.Caption = MakePctComp(Cnt&, NumOfTTRecs&) + " Complete."
      DoEvents
    End If
  Next Cnt
  Close
End Sub

Private Sub keepme()
  'TaxTR.Revenue.InterestPd = 0
  'Put #1, 81583, TaxTR
  'Get #1, 104165, TaxTR
  'TaxTR.Revenue.Principle1Pd = 0
  'TaxTR.Revenue.InterestPd = 0
  'Put #1, 104165, TaxTR
'  For TRCnt = 1 To TRRecCnt
'    Get #1, TRCnt, TaxTR
'    If TaxTR.CustomerRec = 1552 Then
'      If TaxTR.TranType = 1 Then
'        If TaxTR.TaxYear <= 2004 Then
'          Print #100, TRCnt, TaxTR.Revenue.Principle1, TaxTR.Revenue.Principle1Pd
'          Print #100, , TaxTR.Revenue.Interest, TaxTR.Revenue.InterestPd
'          Print #100, , TaxTR.Revenue.Penalty, TaxTR.Revenue.PenaltyPd
'
'          'Stop
'        End If
'      End If
'    End If
'
'
''      If TaxTR.Revenue.Principle1Pd > TaxTR.Revenue.Principle1 Then
''        TaxTR.Revenue.Principle1Pd = TaxTR.Revenue.Principle1
''        Put #1, TRCnt, TaxTR
''        F2Cnt = F2Cnt + 1
''      End If
'
'  Next
'
'  Get #1, 306682, TaxTR
'  TaxTR.Revenue.InterestPd = TaxTR.Revenue.InterestPd - 0.38
'  TaxTR.Revenue.Principle1Pd = TaxTR.Revenue.Principle1Pd + 0.38
'  Put #1, 306682, TaxTR
'    If TRCnt = 276925 Then Stop
'    If TaxTR.TranType = 1 And TaxTR.CustomerRec > 0 And TaxTR.TransDate > tstDate Then
'      If TaxTR.Revenue.InterestPd > TaxTR.Revenue.Interest Then
'        Dif2Add = wRound(TaxTR.Revenue.InterestPd - TaxTR.Revenue.Interest)
'        TaxTR.Revenue.Principle1Pd = wRound(TaxTR.Revenue.Principle1Pd + Dif2Add)
'        TaxTR.Revenue.InterestPd = TaxTR.Revenue.Interest
'        Put #1, TRCnt, TaxTR
'        Corrected = Corrected + 1
'      End If
'    End If
'    If TRCnt Mod 1000 = 0 Then
'      Label2.Caption = "Corrected: " + CStr(Corrected)
'      Label4.Caption = MakePctComp(TRCnt, TRRecCnt) + " Complete."
'      DoEvents
'    End If
'
'
''
'GoTo AllDone
'
'  For TRCnt = 1 To TRRecCnt
'    Get #1, TRCnt, TaxTR
'    If TRCnt Mod 1000 = 0 Then
'      Label4.Caption = MakePctComp(TRCnt, TRRecCnt) + " Complete."
'      DoEvents
'    End If
'    If TaxTR.TranType = 1 And TaxTR.CustomerRec > 0 And TaxTR.TransDate > tstDate Then
'      'If TaxTR.Revenue.Interest <> TaxTR.Revenue.InterestPd Then
'        TotInt = 0
'        BillTR = TRCnt
'        For TR2Cnt = BillTR + 1 To TRRecCnt
'          Get #1, TR2Cnt, IntTR
'          If IntTR.BelongTo = BillTR Then
'            If IntTR.TranType = 4 Then
'              TotInt = wRound(TotInt + IntTR.Revenue.Interest)
'            End If
'            If IntTR.TranType = 3 Then
'              If IntTR.Revenue.Interest <> 0 Then
'                TotInt = wRound(TotInt - IntTR.Revenue.Interest)
'              End If
'            End If
'            If IntTR.TranType = 13 Then
'              If IntTR.Revenue.Interest <> 0 Then
'                TotInt = wRound(TotInt - IntTR.Revenue.Interest)
'              End If
'            End If
'            If IntTR.TranType = 14 Then
'              If IntTR.Revenue.Interest <> 0 Then
'                TotInt = wRound(TotInt + IntTR.Revenue.Interest)
'              End If
'            End If
'            'End Select
'          End If
'          If TR2Cnt Mod 10000 = 0 Then
'            Label5.Caption = MakePctComp(TR2Cnt, TRRecCnt) + " Complete."
'            DoEvents
'          End If
'        Next
'
'        If TotInt <> TaxTR.Revenue.Interest Then
'          TaxTR.Revenue.Interest = TotInt
'          Corrected = Corrected + 1
'          Label2.Caption = "Corrected: " + CStr(Corrected)
'          Put #1, BillTR, TaxTR
'        End If
'    End If
'  Next
'  Close
'
'
'ResumeHere:
'
''Open "translist.txt" For Output As #100
''
''For xCnt = 1 To trTypeCnt
''Print #100, trType(xCnt)
''
''Next
''Close
'DownHereMan:
'  Label3.Caption = "Relinking Transactions."
'  DoEvents
'
'  Call RelinkTransactions
'  'Label4.Caption = "Removed:" + CStr(F2Cnt)
'  DoEvents
'
'AllDone:
'  Label3.Caption = "Relinking Transactions."
'  DoEvents
'
'  'Label4.Caption = "Removed:" + CStr(F2Cnt)
'  DoEvents
'
'  Close #1
'  Label2 = ""
'  Label4 = ""
'  Label5 = ""
'  Label3 = "Correction Complete."
End Sub


