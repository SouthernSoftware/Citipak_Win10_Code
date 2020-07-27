VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NC Taxes"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
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
      Left            =   1478
      TabIndex        =   5
      Top             =   1680
      Width           =   1725
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
      TabIndex        =   4
      Top             =   1275
      Width           =   3150
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Click 'OK' to Continue"
      Height          =   240
      Left            =   765
      TabIndex        =   3
      Top             =   990
      Width           =   3150
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   765
      TabIndex        =   2
      Top             =   675
      Width           =   3150
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "NC Taxes"
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
      TabIndex        =   0
      Top             =   330
      Width           =   3150
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Dim TRFile As String
Dim TRBack As String
Dim TaxTR As TaxTransactionType
Dim TaxBL As TaxTransactionType
Dim IntTR As TaxTransactionType

Dim TRLen As Integer
Dim TRCnt As Long
Dim blnDoneFix As Boolean
Dim blnInProcs As Boolean
Dim TCnt As Long

Private Sub Form_Load()
  Call InitStuff
End Sub

Private Sub InitStuff()
  TRFile = "TAXTRANS.DAT"
  TRBack = "TAXTRANS.BAK"
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
  'Label2="Processing: "+
  blnInProcs = True
  Label3.Caption = "Processing. . . "
  DoEvents
  
  Dim tstDate As Integer
  Dim CMTrans As CMTransRecType
  Dim TRRecCnt As Long
  Dim atDate As Integer
  Dim F1Cnt As Integer
  Dim F2Cnt As Integer
  Dim discDate As Integer
  Dim BillDate As Integer
  Dim TRDATE As Integer
  Dim SwpCnt As Integer
  Dim LastTR As Long
  Dim TR2Cnt As Long
  Dim TRCount As Long
  Dim Corrected As Integer
  Dim BillTR As Long
  Dim TotInt As Double
  ReDim trType(1) As Integer
  Dim xCnt As Integer
  Dim trTypeCnt As Integer
  Dim foundIT As Boolean
  Dim DidFirstOne As Boolean
  Dim Dif2Add As Double
  Dim xRCnt As Long
  Dim PersPropRec As PersonalRecType
  Dim PPLen As Integer
  Dim RPLen As Integer
  Dim RealPropRec As PropertyRecType
  Dim TPCnt As Integer
  Dim thisBill As Long
      
  Dim doRelink As Boolean
  
  TRDATE = Date2Num("12/01/2017")

  Call KillTransBAK
  
  DidFirstOne = False
  TRLen = Len(TaxTR)
  PPLen = Len(PersPropRec)
  RPLen = Len(RealPropRec)
  
  'frmTaxMasterBalList.Show
  
  'Call LoadMe
  
  Open TRFile For Random As #1 Len = TRLen
'
'  TRRecCnt = LOF(1) / TRLen
'
   Get #1, 308047, TaxTR
   TaxTR.RealPin = "-1-1-1"
   Put #1, 308047, TaxTR
   Close
'  TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
'  Put #1, 377361, TaxTR
'  Get #1, 376053, TaxTR
'  TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
'  Put #1, 376053, TaxTR
'
'  Get #1, 361490, TaxTR
'  TaxTR.Revenue.Interest = 8.97
'  TaxTR.Revenue.InterestPd = 8.97
'  Put #1, 361490, TaxTR
'
'  Get #1, 377230, TaxTR
'  TaxTR.Revenue.InterestPd = 8.97
'  TaxTR.Revenue.PrePaidAmt = 13.5
'  TaxTR.Revenue.PrePaidBal = 13.5
'  Put #1, 377230, TaxTR
'
''  BillTR = TaxTR.BelongTo
''  Get #1, BillTR, TaxBL
''  'Print TaxBL.Description
''
''  TaxBL.Revenue.Principle1Pd = TaxBL.Revenue.Principle1Pd - TaxTR.Revenue.Principle1Pd
''  TaxBL.Revenue.RevOpt3Pd = TaxBL.Revenue.RevOpt3Pd - TaxTR.Revenue.RevOpt3Pd
''  Put #1, BillTR, TaxBL
''  TaxTR.CustomerRec = -TaxTR.CustomerRec
''  Put #1, 421723, TaxTR
''
''  Get #1, 421722, TaxTR
''  BillTR = TaxTR.BelongTo
''  Get #1, BillTR, TaxBL
''  TaxBL.Revenue.Principle1Pd = TaxBL.Revenue.Principle1Pd - TaxTR.Revenue.Principle1Pd
''  TaxBL.Revenue.RevOpt3Pd = TaxBL.Revenue.RevOpt3Pd - TaxTR.Revenue.RevOpt3Pd
''  Put #1, BillTR, TaxBL
''  TaxTR.CustomerRec = -TaxTR.CustomerRec
''  Put #1, 421722, TaxTR
''
''  Get #1, 421721, TaxTR
''  BillTR = TaxTR.BelongTo
''  Get #1, BillTR, TaxBL
''  TaxBL.Revenue.RevOpt3Pd = TaxBL.Revenue.RevOpt3Pd - TaxTR.Revenue.RevOpt3Pd
''  Put #1, BillTR, TaxBL
''  TaxTR.CustomerRec = -TaxTR.CustomerRec
''  Put #1, 421721, TaxTR
''
''  'Print TaxBL.Revenue.Principle2Pd, TaxTR.Revenue.Principle2Pd
''  'Print TaxBL.Revenue.Principle3Pd, TaxTR.Revenue.Principle3Pd
''  'Print TaxBL.Revenue.Principle4Pd, TaxTR.Revenue.Principle4Pd
''  'Print TaxBL.Revenue.Principle5Pd, TaxTR.Revenue.Principle5Pd
''
''  'TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
'''  Put #1, 337393, TaxTR
''
''
'  Close
  doRelink = True
'  doRelink = False
  
'  For BillTR = 1 To TRRecCnt
'    Get #1, BillTR, TaxTR
'    If TaxTR.CustomerRec <= 0 Then
'      If Len(QPTrim(TaxTR.RealPin)) > 0 Then
'        TaxTR.RealPin = "z0z0z0z0z0z0z0'"
'        Put #1, BillTR, TaxTR
'        xRCnt = xRCnt + 1
'      End If
'    End If
'  Next
'  Close

'  For xRCnt = 1 To TRRecCnt

'GoTo Good2Here
'    Get #1, 76361, TaxTR
'    TaxTR.Revenue.Principle1Pd = 73.3
'    Put #1, 76361, TaxTR
'
'    Get #1, 79314, TaxTR
'    TaxTR.Amount = 73.3
'    TaxTR.Revenue.Principle1Pd = 73.3
'    Put #1, 79314, TaxTR
'
'    Get #1, 94307, TaxTR
'    TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
'    Put #1, 94307, TaxTR
'
'    Get #1, 96184, TaxTR
'    TaxTR.Amount = 53.9
'    TaxTR.Revenue.Principle1Pd = 53.9
'    Put #1, 96184, TaxTR
'
'    Get #1, 95045, TaxTR
'    TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
'    Put #1, 95045, TaxTR
'
'    Get #1, 130309, TaxTR
'    TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
'    Put #1, 130309, TaxTR
'
'    Get #1, 79439, TaxTR
'    TaxTR.Amount = 482.25
'    TaxTR.Revenue.Principle1Pd = 482.25
'    Put #1, 79439, TaxTR
'
'    Get #1, 97034, TaxTR
'    TaxTR.Amount = 726.92
'    TaxTR.Revenue.Principle1Pd = 726.92
'    Put #1, 97034, TaxTR
'
'    Get #1, 115765, TaxTR
'    TaxTR.Amount = 734.16
'    'Print TaxTR.Revenue.Principle1Pd '= 726.92
'    Put #1, 115765, TaxTR
'
'    Get #1, 139051, TaxTR
'    TaxTR.Amount = 734.16
'    TaxTR.Revenue.Principle1Pd = 726.92
'    Put #1, 139051, TaxTR
'
'    Get #1, 95105, TaxTR
'    TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
'    Put #1, 95105, TaxTR
'
'    Get #1, 79293, TaxTR
'    TaxTR.Amount = 915.7
'    TaxTR.Revenue.Principle1Pd = 915.7
'    Put #1, 79293, TaxTR
'
'    Get #1, 97125, TaxTR
'    TaxTR.Amount = 1738.74
'    TaxTR.Revenue.Principle1Pd = 1738.74
'    Put #1, 97125, TaxTR
'
'    Get #1, 95100, TaxTR
'    TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
'    Put #1, 95100, TaxTR
'
'    Get #1, 97124, TaxTR
'    TaxTR.Amount = 13.08
'    TaxTR.Revenue.Principle1Pd = 13.08
'    Put #1, 97124, TaxTR
'
'    Get #1, 79292, TaxTR
'    TaxTR.Amount = 23.03
'    TaxTR.Revenue.Principle1Pd = 23.03
'    Put #1, 79292, TaxTR
'
'    Get #1, 79420, TaxTR
'    TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
'    Put #1, 79420, TaxTR
'

'Good2Here:
'    Get #1, 87717, TaxTR
'    TaxTR.Amount = 13.52
'    TaxTR.Revenue.Principle1Pd = 13.16
'    Put #1, 87717, TaxTR
'
'    Get #1, 188800, TaxTR
'    TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
'    Put #1, 188800, TaxTR


'    Exit For
'      Put #1, xRCnt, TaxTR
'    End If
'  Next
  
'  Get #1, 104519, TaxTR
  'Print
  'TaxTR.DiscAmt = 0
'  Put #1, 104519, TaxTR
  
'  TaxTR.Revenue.Interest = 5.76
'  Put #1, 302017, TaxTR

   
'  For xRCnt = 1 To TRRecCnt
'    Get #1, xRCnt, TaxTR
'    If TaxTR.TransDate = TRDATE Then
'      If TaxTR.CustomerRec > 0 Then
'        Select Case TaxTR.TranType
'        Case 2   'payments  'Print TaxTR.Revenue.Principle1Pd
'          Get #1, TaxTR.BelongTo, IntTR 'get the bill
'          IntTR.Revenue.Principle1Pd = IntTR.Revenue.Principle1Pd - TaxTR.Revenue.Principle1Pd
'          IntTR.Revenue.InterestPd = IntTR.Revenue.InterestPd - TaxTR.Revenue.InterestPd
'          IntTR.Revenue.CollectionPd = IntTR.Revenue.CollectionPd - TaxTR.Revenue.CollectionPd
'          Put #1, TaxTR.BelongTo, IntTR 'put bill back
'          TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
'          Put #1, xRCnt, TaxTR
'        Case 6
'          Get #1, TaxTR.BelongTo, IntTR 'get the bill
'          IntTR.Revenue.Collection = IntTR.Revenue.Collection - TaxTR.Revenue.Collection
'          Put #1, TaxTR.BelongTo, IntTR 'put bill back
'          TaxTR.CustomerRec = -Abs(TaxTR.CustomerRec)
'          Put #1, xRCnt, TaxTR
'          F1Cnt = F1Cnt + 1
'        Case Else
'          Stop
'        End Select
'      End If
'    End If
'  Next
'
  
ByeByeNow:

Close

GoTo AllDoneExit:

DoneAll:
'  Call RelinkTransactions
  Label2 = ""
  Label4 = "" ' CStr(F2Cnt)
  Label5 = ""
  Label3 = "Correction Complete."

AllDoneExit:
  If doRelink Then
    Call RelinkTransactions
  End If
  
  Label2 = "" 'Str$(xRCnt)
  'Label4 = " Removed: " + CStr(F1Cnt)
  Label3 = "Correction Complete."
  Label4 = "" ' CStr(F2Cnt)
  Label5 = ""

  
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

Public Function wRound(n As Double) As Double
  wRound = Int(n * 100 + 0.500000001) / 100
End Function

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

Public Function MakePctComp(ByVal Cnt As Long, ByVal TotalCnt As Long) As String
  Dim PctComp As Long
  Dim RetStr As String
  PctComp = Int((Cnt / TotalCnt) * 100)
  RetStr = CStr(PctComp) + "%"
  MakePctComp = RetStr
End Function

'Public Function ParseBillNum$(Text$)
'  Dim BillNum$
'  Dim BNumLen As Integer
'  Dim thischar$
'  Dim GoodPos As Integer
'  Dim Cnt As Integer
'
'  BillNum$ = QPTrim$(Text$)
'  BNumLen = Len(BillNum$)
'  If BNumLen > 0 Then
'    For Cnt = BNumLen To 1 Step -1
'      thischar$ = Mid$(BillNum$, Cnt, 1)
'      If InStr("0123456789", thischar$) <= 0 Then
'        Exit For
'      End If
'    Next
'    GoodPos = Cnt + 1
'    BillNum$ = Mid$(BillNum$, GoodPos)
'  End If
'  If Not IsNumeric(BillNum$) Then
'    BillNum = "-911"
'  End If
'  ParseBillNum$ = BillNum$
'End Function

'Public Function QPTrim$(Text As String)
'  Dim StrLen As Long
'  Dim Cnt As Long
'  Dim thischar As Integer
'  StrLen = Len(Text)
'  For Cnt = 1 To StrLen
'    thischar = Asc(Mid$(Text, Cnt, 1))
'    If thischar = 0 Then
'      Mid$(Text$, Cnt, 1) = " "
'    End If
'  Next
'  QPTrim$ = Trim$(Text)
'End Function

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

