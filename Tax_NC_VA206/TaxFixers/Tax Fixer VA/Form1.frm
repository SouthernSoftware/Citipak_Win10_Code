VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VA Taxes "
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1673
      TabIndex        =   1
      Top             =   2040
      Width           =   1335
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
      Height          =   255
      Left            =   1020
      TabIndex        =   4
      Top             =   1590
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Click 'OK' to Continue"
      Height          =   255
      Left            =   1020
      TabIndex        =   3
      Top             =   1275
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   1020
      TabIndex        =   2
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "VA Taxes"
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
      Left            =   1020
      TabIndex        =   0
      Top             =   600
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefInt A-Z

Dim TRFile As String
Dim PPFile As String
Dim REFile As String
Dim TaxTR As TaxTransactionType
Dim TaxTR2 As TaxTransactionType
Dim PersRec As PersonalRecType
Dim TXCustRec As TaxCustType
'dim RealRec as Real
Dim TRLen As Integer
Dim RecCnt As Long
Dim CustLen As Integer
Dim blnDoneFix As Boolean
Dim blnInProcs As Boolean
  

Private Sub Form_Load()
  Call InitStuff
End Sub

Private Sub InitStuff()
  TRFile = "TAXTRANS.DAT"
  PPFile = "TAXPERS.DAT"
  REFile = "TAXPROP.DAT"
  blnDoneFix = False
  blnInProcs = False
End Sub

Private Sub Command1_Click()
  If blnDoneFix Then
    End
  Else
    Command1.Enabled = False
    Call FixTransactions
  End If
End Sub

Private Sub FixTransactions()

  Dim RealRec As PropertyRecType
  Dim PersRec As PersonalRecType
  Dim atDate As Integer
  Dim CngCnt As Integer
  Dim RecLen As Integer
  Dim LopCnt As Long
  Dim TRRecCnt As Long
  Dim RealLen As Integer
  Dim RateRec As OptRevRateTablesType
  Dim RateLen As Integer
 ' Dim RealRec As PropertyRecType
  Dim RateCnt As Integer
  Dim DoRelink As Boolean
  RealLen = Len(RealRec)
  CustLen = Len(TXCustRec)
  RecLen = Len(TaxTR)
  RateLen = Len(RateRec)
  Dim tgDate  As Integer
  Dim tDate As Integer
  Dim eDate As Integer
  'sDate = Date2Num("01/30/2016")
  eDate = Date2Num("04/24/2019")
  
  Label3.Caption = ""
 
  Dim FoundIt As Boolean
  Dim LBRec As Integer
  Dim PCnt As Long
  Dim tInStr As String
  Dim OutStr As String * 80
  Dim ChkStr As String
  Dim AcctPins(1 To 8000) As PinsRecType
  Dim AcctStr As String * 19
  
  Open "TaxTrans.Dat" For Random As #1 Len = RecLen
  TRRecCnt = LOF(1) / RecLen
  For LopCnt = 1 To TRRecCnt
    Get #1, LopCnt, TaxTR
    Select Case LopCnt
    Case 23065, 23070, 23075, 23101
      TaxTR.Revenue.RevOpt3Pd = 0
      TaxTR.Revenue.Principle1Pd = 0
      TaxTR.Amount = 0
      TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
      Put #1, LopCnt, TaxTR
    Case 22374
      TaxTR.Revenue.RevOpt3Pd = 40
      TaxTR.Revenue.Principle1Pd = 43.53
      Put #1, LopCnt, TaxTR
    End Select
    
'    Select Case TaxTR.TaxYear
'    Case "2015", "2016"
'      If TaxTR.TransDate = eDate Then
'        If TaxTR.TranType = 5 Then
'          Get #1, TaxTR.BelongTo, TaxTR2
'          TaxTR2.Revenue.Penalty = (TaxTR2.Revenue.Penalty - TaxTR.Revenue.Penalty)
'          Put #1, TaxTR.BelongTo, TaxTR2
'          TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
'          Put #1, LopCnt, TaxTR
'        End If
'      End If
'    End Select
  Next
  Close
  DoRelink = True
'  Open "AcctRealPINS.txt" For Input As #1
'  Do
'    Line Input #1, tInStr
'    If Len(tInStr) > 5 Then
'        PCnt = PCnt + 1
'        AcctPins(PCnt).AcctID = LTrim(RTrim(Left$(tInStr, 38)))
'        AcctPins(PCnt).RealPin = LTrim(RTrim(Mid$(tInStr, 38, 40)))
'    End If
'  Loop Until EOF(1)
'  Close
'
'  Open "LockBox.Txt" For Input As #1
'  Open "FixedLockBox.Txt" For Output As #2
'  'Open "BadRecs.txt" For Output As #10
'  Do
'    Line Input #1, tInStr
'    LBRec = LBRec + 1
'    FoundIt = False
'    ChkStr = LTrim(RTrim(Left$(tInStr, 19)))
'    If Len(ChkStr) > 1 Then
'      For LopCnt = 2 To PCnt
'        If InStr(AcctPins(LopCnt).RealPin, ChkStr) > 0 Then
'          FoundIt = True
'          Exit For
'        End If
'      Next
'      LSet AcctStr = LTrim(RTrim(AcctPins(LopCnt).AcctID))
'      Mid$(tInStr, 1) = AcctStr
'      Mid$(tInStr, 41) = "11/15/2018"
'      Print #2, tInStr
'
''      If Not FoundIt Then
''        Print #10, LBRec, ChkStr
''      End If
'    End If
'
'  Loop Until EOF(1)
'
'  If Len(tInStr) > 20 Then
'    Print #2, tInStr
'  End If
'
'  Close
    'AcctID As String * 40
    'RealPin As String * 40
  
  
  'Open PPFile For Random As #1 Len = RecLen
  'Name TRFile As "TRANS.BAK"
  'On Error Resume Next
  'Kill "TAXRLOP3.DAT"
  'On Error GoTo 0
  
  'End
  
  
  'Call RelinkTransactions
'  Open REFile For Random As #1 Len = TRLen
'  'Open "TaxTrans.NEW" For Random As #2 Len = TRLen
'  TRRecCnt = LOF(1) / TRLen
'
'  For LopCnt = 1 To TRRecCnt
'    Get #1, LopCnt, RealRec
'    RealRec.PropAddr = ""
'    Put #1, LopCnt, RealRec
'
'
'NextLoop:
'  Next

'  Close
  
'  TRLen = Len(PersRec)
'  Open "TaxPers.dat" For Random As #1 Len = TRLen
'  TRRecCnt = LOF(1) / TRLen
'
'  For LopCnt = 1 To TRRecCnt
'    Get #1, LopCnt, PersRec
'    PersRec.DMVSubmitted = "N"
'    Put #1, LopCnt, PersRec
'    CngCnt = CngCnt + 1
'  Next

DownHere:
 Close
 Label3.Caption = "Relinking Transactions."
 If DoRelink Then
   Call RelinkTransactions
 End If
OverHere:
'  Print RealLen, RateLen
  
  Label3 = "Processing Complete."
  Command1.Caption = "EXIT"
  'Label3.Caption = "Detached: " + CStr(CngCnt)
  blnInProcs = False
  blnDoneFix = True
  Command1.Enabled = True
  
End Sub

Private Sub RelinkTransactions()
  Dim TaxCust As TaxCustType
  Dim TaxTran As TaxTransactionType
  Dim TCHandle As Integer
  Dim TTHandle As Integer
  
  Dim NumOfTCRecs As Long
  Dim NumOfTTRecs As Long

  Dim Cnt As Long
  Dim TaxCustLen As Integer
  Dim TaxTransLen As Integer
  
  TaxCustLen = Len(TaxCust)
  TCHandle = FreeFile
  Open TaxCustFile For Random Shared As TCHandle Len = TaxCustLen
  NumOfTCRecs = LOF(TCHandle) / TaxCustLen
  
  TaxTransLen = Len(TaxTran)
  TTHandle = FreeFile
  Open TaxTransFile For Random Shared As TTHandle Len = TaxTransLen
  NumOfTTRecs = LOF(TTHandle) / TaxTransLen
  
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
    If Cnt& Mod 10 = 0 Then
      Label4.Caption = MakePctComp(Cnt&, NumOfTTRecs&) + " Complete."
      DoEvents
    End If
  Next Cnt
  Close
  Label4.Caption = "100% Complete."
  DoEvents

End Sub

Public Function Date2Num(TheDate$) As Integer
 'useful function throughout program...
 'takes a string date and converts into a number based on 12/31/1979
  Date2Num% = DateDiff("d", "12/31/1979", (TheDate$))
End Function

Public Function MakeRegDate(ByVal DateNumb) As String
  Dim Month As Integer, ThisDate As String
  'function does the opposite of Date2Num
  If DateNumb = -32767 Then
    MakeRegDate = "%%%%%%%%%% "
  Else
    MakeRegDate = Format(DateAdd("d", (DateNumb), "12-31-1979"), "mm/dd/yyyy")
  End If
End Function


Private Sub Form_Unload(Cancel As Integer)
  'On Error Resume Next
  If blnInProcs Then
    Cancel = 1
'  Else
'    If blnDoneFix Then
'      Name "VAFixTrans.EXE" As "VAFixTra.EXE"
'      Kill "VAFixTra.EXE"
'    End If
  End If
End Sub

Public Function MakePctComp(ByVal Cnt As Long, ByVal TotalCnt As Long) As String
  Dim PctComp As Long
  Dim RetStr As String
  PctComp = Int((Cnt / TotalCnt) * 100)
  RetStr = CStr(PctComp) + "%"
  MakePctComp = RetStr
End Function

Private Sub remmedstuff()
'    Get #1, 34744, TaxTR
'    TaxTR.TranType = 2
'    TaxTR.BelongTo = 35319
'    Put #1, 34744, TaxTR
'
'    Get #1, 34732, TaxTR
'    TaxTR.TranType = 2
'    TaxTR.BelongTo = 35195
'    Put #1, 34732, TaxTR
'
'    Get #1, 34733, TaxTR
'    TaxTR.TranType = 2
'    TaxTR.BelongTo = 35298
'    Put #1, 34733, TaxTR
'  Print CngCnt
'  Get #1, TRRecCnt, TaxTR
'  Print TaxTR.TranType
'  Get #1, TRRecCnt - 300, TaxTR
'  Print TaxTR.TranType

  '1=Bill 2=Payment 3=Release 4=Interest
  '5=Penalty 6=Collection/Ad Cost Billing
  '7=AdjustmentDwnBill 8=MiscCost 9=AdjUpBill
  '10=DwnAdjPay 11=UpAdjPay
  '22=PrePayment 23=Refund Prepayment added 3-25-03
  'Open "TXRTTBLS.DAT" For Random Shared As #1 Len = RateLen
'  RateCnt = LOF(1) / RateLen
'  For LopCnt = 1 To RateCnt
'    Get #1, LopCnt, RateRec
'    If InStr(RateRec.Desc, "15.00") > 0 Then
'      RateRec.Desc = "VLF 10.00"
'      Put #1, LopCnt, RateRec
'      Exit For
'    End If
'  Next
'  Close
'
'  GoTo OverHere
'  Open TRFile For Random As #1 Len = RecLen
'  TRRecCnt = LOF(1) / RecLen
'
'  Get #1, 61715, TaxTR
'  TaxTR.BelongTo = 0
'  TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
'  Put #1, 61715, TaxTR
'
'  Get #1, 80691, TaxTR
'  TaxTR.BelongTo = 0
'  TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
'  Put #1, 80691, TaxTR
'
'  Get #1, 119796, TaxTR
'  TaxTR.BelongTo = 0
'  TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
'  Put #1, 119796, TaxTR
'
'  Get #1, 120476, TaxTR
'  TaxTR.BelongTo = 0
'  TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
'  Put #1, 120476, TaxTR
'
'  Get #1, 120875, TaxTR
'  TaxTR.BelongTo = 0
'  TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
'  Put #1, 120875, TaxTR
'
'  Get #1, 126136, TaxTR
'  TaxTR.BelongTo = 0
'  TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
'  Put #1, 126136, TaxTR
'
'  Get #1, 128614, TaxTR
'  TaxTR.BelongTo = 0
'  TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
'  Put #1, 128614, TaxTR
'
'  Get #1, 124932, TaxTR
'  TaxTR.BelongTo = 0
'  TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
'  Put #1, 124932, TaxTR
'
'
'  'Open "taxcust.dat" For Random As #1 Len = CustLen
'  'TRRecCnt = LOF(1) / CustLen
'
'  'For LopCnt = 1 To TRRecCnt
'    'Get #1, LopCnt, TXCustRec
'    ''? TXCustRec.Op
''    If TaxTR.CustomerRec = 1822 Then
''      If TaxTR.TransDate <= atDate Then
''        TaxTR.BelongTo = 0
''        TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
''        Put #1, LopCnt, TaxTR
''      End If
''    End If
'  'Next
'
''  Get #1, 7156, TaxTR
''  TaxTR.BelongTo = 0
''  TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
''  Put #1, 7156, TaxTR
''
''  Get #1, 9173, TaxTR
''  TaxTR.BelongTo = 0
''  TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
''  Put #1, 9173, TaxTR
''
''  Get #1, 9874, TaxTR
''  TaxTR.BelongTo = 0
''  TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
''  Put #1, 9874, TaxTR
''
''  Get #1, 11421, TaxTR
''  TaxTR.BelongTo = 0
''  TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
''  Put #1, 11421, TaxTR
''
''  Get #1, 14559, TaxTR
''  TaxTR.BelongTo = 0
''  TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
''  Put #1, 14559, TaxTR
''
''  For LopCnt = 1 To TRRecCnt
''    Get #1, LopCnt, TaxTR
''    If TaxTR.CustomerRec = 1822 Then
''      If TaxTR.TransDate = atDate Then
''        'Print TaxTR.Description
''        Print TaxTR.BillType
''      End If
'''      If InStr(TaxTR.Description, "-911") > 0 Then
''''        Print TaxTR.Description
'''      End If
''    End If
''  Next
''  Get #1, 37002, TaxTR
''  TaxTR.BelongTo = 0
''  TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
''  Put #1, 37002, TaxTR
''
''  Get #1, 36967, TaxTR
''  TaxTR.BelongTo = 0
''  TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
''  Put #1, 36967, TaxTR
''
''  Get #1, 35319, TaxTR
''  TaxTR.Revenue.Principle1Pd = 55.33
''  TaxTR.Revenue.RevOpt1Pd = 50
''  Put #1, 35319, TaxTR
''
''  Get #1, 37004, TaxTR
''  TaxTR.BelongTo = 0
''  TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
''  Put #1, 37004, TaxTR
''
''  Get #1, 36944, TaxTR
''  TaxTR.BelongTo = 0
''  TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
''  Put #1, 36944, TaxTR
''
''  Get #1, 35195, TaxTR
''  TaxTR.Revenue.Principle1Pd = 8.82
''  TaxTR.Revenue.RevOpt1Pd = 25
''  Put #1, 35195, TaxTR
''
''  Get #1, 37003, TaxTR
''  TaxTR.BelongTo = 0
''  TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
''  Put #1, 37003, TaxTR
''
''  Get #1, 36964, TaxTR
''  TaxTR.BelongTo = 0
''  TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
''  Put #1, 36964, TaxTR
''
''  Get #1, 35298, TaxTR
''  TaxTR.Revenue.Principle1Pd = 22.53
''  TaxTR.Revenue.RevOpt1Pd = 50
''  Put #1, 35298, TaxTR
''
''Stop
''  TaxTR.Revenue.RevOpt2Pd = 0
''  Put #1, 41103, TaxTR
'
''  Get #1, 41103, TaxTR
''  TaxTR.Revenue.Principle1Pd = 0
''  TaxTR.Revenue.RevOpt1Pd = 0
''  Put #1, 41103, TaxTR
''
''  Get #1, 40619, TaxTR
''  TaxTR.Revenue.Principle1Pd = 0
''  Put #1, 40619, TaxTR
'
''    'Put #2, RecCnt, TaxTR
'  'Open "TaxCust.Dat" For Random As #2 Len = CustLen
'  'Get #2, 2494, TXCustRec
'  'TXCustRec.LastTrans = 0
'  'TXCustRec.FirstPersRec = 0
'  'Put #2, 2494, TXCustRec
'  'Label3.Caption = "Scanning Personl Property"
'
''  For RecCnt = TRRecCnt - 4 To TRRecCnt
''    Get #1, RecCnt, TaxTR
''    'Put #2, RecCnt, TaxTR
''      'If TaxTR.TranType <> 1 Then
'''         Debug.Print RecCnt, TaxTR.TranType, TaxTR.CustomerRec
''      'End If
'''      TaxTR.CustomerRec = -TaxTR.CustomerRec
'''      CngCnt = CngCnt + 1
'''      Put #1, RecCnt, TaxTR
'''    End If
''
''  Next
''    Get #1, 41583, TaxTR
''    TaxTR.CustomerRec = -TaxTR.CustomerRec
'    ''''TaxTR.Revenue.Principle1Pd = 0
'    'Print TaxTR.Revenue.Principle1
'    'TaxTR.Revenue.Principle1Pd = TaxTR.Revenue.Principle1
'    'TaxTR.Revenue.Principle1 = 0
''    Put #1, 41583, TaxTR
'
''     TaxTR.Amount = 0
'
'    'Print TaxTR.Revenue.PrePaidAmt '= 0
'    'Print TaxTR.Revenue.PrePaidUsed '= 0
''     TaxTR.Revenue.PrePaidBal = 0
'
'    'Print TaxTR.FromPrePay
'    ''''Put #1, 24420, TaxTR
'
''     Get #1, 24203, TaxTR
''     TaxTR.Amount = 0
''     TaxTR.Revenue.PrePaidAmt = 0
''     TaxTR.Revenue.PrePaidUsed = 0
''     TaxTR.Revenue.PrePaidBal = 0
''     Put #1, 24203, TaxTR
'
''    Get #1, 18111, TaxTR
''    TaxTR.Revenue.Principle1 = 0
''    TaxTR.Amount = 0
''    Put #1, 18111, TaxTR
''
''    Select Case TaxTR.CustomerRec
''    Case 14, 440, 441, 596
''      If TaxTR.TransDate = atDate Then
''        F1Cnt = F1Cnt + 1
''        TaxTR.CustomerRec = -TaxTR.CustomerRec
''        Put #1, TRCnt, TaxTR
''      End If
''    Case Else
''    End Select
''SkipToHere:
''  Next
''
'ThisHere:
'  '2848
'  Get #1, 75112, TaxTR
'  'Print TaxTR.Amount
'  TaxTR.Revenue.Principle1Pd = TaxTR.Revenue.Principle1
'  Put #1, 75112, TaxTR
'
'  Get #1, 77035, TaxTR
'  TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
'  Put #1, 77035, TaxTR
'
'  Get #1, 86354, TaxTR
'  TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
'  Put #1, 86354, TaxTR
'
'  Get #1, 87418, TaxTR
'  TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
'  Put #1, 87418, TaxTR
'
'  Get #1, 86353, TaxTR
'  TaxTR.Revenue.Principle1Pd = TaxTR.Amount
'  TaxTR.Revenue.Principle1 = TaxTR.Amount
'  Put #1, 86353, TaxTR
'
'  '2849
'  Get #1, 77036, TaxTR
'  TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
'  Put #1, 77036, TaxTR
'
'  Get #1, 86352, TaxTR
'  TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
'  Put #1, 86352, TaxTR
'
'  Get #1, 87417, TaxTR
'  TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
'  Put #1, 87417, TaxTR
'
'  Get #1, 86351, TaxTR
'  TaxTR.Revenue.Principle1Pd = TaxTR.Amount
'  TaxTR.Revenue.Principle1 = TaxTR.Amount
'  Put #1, 86351, TaxTR

'GoTo DownHere
  '2942
'  Get #1, 77029, TaxTR
'  TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
'  Put #1, 77029, TaxTR
'
'  Get #1, 85946, TaxTR
'  TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
'  Put #1, 85946, TaxTR
'
'  Get #1, 74723, TaxTR
'  TaxTR.Revenue.Principle1Pd = TaxTR.Amount
'  TaxTR.Revenue.Principle1 = TaxTR.Amount
'  Put #1, 74723, TaxTR

  '3170
'  Get #1, 77031, TaxTR
'  TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
'  Put #1, 77031, TaxTR
'
'  Get #1, 74740, TaxTR
'  TaxTR.Revenue.Principle1Pd = TaxTR.Amount
'  TaxTR.Revenue.Principle1 = TaxTR.Amount
'  Put #1, 74740, TaxTR
'
'  '3242
'  Get #1, 30248, TaxTR
'  TaxTR.Revenue.Principle1Pd = TaxTR.Revenue.Principle1
'  TaxTR.Revenue.InterestPd = TaxTR.Revenue.Interest
'  Put #1, 30248, TaxTR
'
'  Get #1, 77027, TaxTR
'  TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
'  Put #1, 77027, TaxTR
'
'  '3376
'  Get #1, 77033, TaxTR
'  TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
'  Put #1, 77033, TaxTR
'
'  Get #1, 75103, TaxTR
'  TaxTR.Revenue.Principle1Pd = TaxTR.Amount
'  TaxTR.Revenue.Principle1 = TaxTR.Amount
'  Put #1, 75103, TaxTR
'
'  Get #1, 88019, TaxTR
'  TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
'  Put #1, 88019, TaxTR
'
'  Get #1, 86344, TaxTR
'  TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
'  Put #1, 86344, TaxTR
'
'  Get #1, 86343, TaxTR
'  TaxTR.Revenue.Principle1Pd = TaxTR.Amount
'  TaxTR.Revenue.Principle1 = TaxTR.Amount
'  Put #1, 86343, TaxTR

End Sub
