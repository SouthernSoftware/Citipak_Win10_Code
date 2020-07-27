VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Preconversion Checker"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4530
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
      Left            =   1598
      TabIndex        =   0
      Top             =   2070
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "VA Taxes Transaction Checker"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   608
      TabIndex        =   4
      Top             =   405
      Width           =   3315
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   938
      TabIndex        =   3
      Top             =   810
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Click 'OK' to Continue"
      Height          =   255
      Left            =   938
      TabIndex        =   2
      Top             =   1170
      Width           =   2655
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
      Height          =   285
      Left            =   938
      TabIndex        =   1
      Top             =   1485
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
'Dim TRFile As String
Dim blnDoneFix As Boolean
Dim blnInProcs As Boolean

Private Sub Command1_Click()
  If blnDoneFix Then
    'Shell "notepad TXReport.txt", vbNormalFocus
    End
  Else
    Command1.Enabled = False
    Call RelinkTransactions
    Call CheckTransactions
  End If
End Sub

Private Sub CheckTransactions()

  Dim TaxCust As TaxCustType
  Dim TaxTR As TaxTransactionType
  Dim TRLen As Integer
  Dim TCLen As Integer
  Dim TRCnt As Long
  Dim TCCnt As Long
  Dim CustLastTR As Long
  Dim Padd As String * 15
  Dim dPad As String * 10
  Dim tCustName As String * 30
  Dim BillBal As Double
  blnDoneFix = False
  blnInProcs = True
  Label3.Caption = "Zero Amount Transactions."
  DoEvents
  
  ReDim TRInfo(0 To 1, 0 To 1) As Long
  
  Dim CustCnt As Integer
  Dim TRBCnt As Integer
  Dim gotCust As Boolean
  Dim TRRecCnt As Long
  Dim TCustCnt As Long

  Dim F1Cnt As Integer
  Dim F2Cnt As Integer
  Dim zCnt As Integer
  Dim BigCnt As Long
  Dim WhatTR As Long
  Dim TRAmt$
  Dim tPrinciplePad As Double
  Dim tPrincipleRev As Double
  
  'Dim tBalDiff As Double
  Dim Balance As Double
'  Dim pdPri As Double: Dim rvPri As Double
'  Dim pdInt As Double: Dim rvInt As Double
'  Dim pdLat As Double: Dim rvLat As Double
  
  CustCnt = 0
  gotCust = False

  TCLen = Len(TaxCust)
  TRLen = Len(TaxTR)
  
  Open TaxTransFile For Random As #1 Len = TRLen
  TRRecCnt = LOF(1) / TRLen
  Open TaxCustFile For Random As #2 Len = TCLen
  TCustCnt = LOF(2) / TCLen

  Open "TXReport.txt" For Output As #3
  For F2Cnt = 1 To TCustCnt
    Get #2, F2Cnt, TaxCust
    WhatTR = TaxCust.LastTrans
    Do While WhatTR > 0
      Get #1, WhatTR, TaxTR
      If WhatTR Mod 10 = 0 Then
        Label4.Caption = MakePctComp(F2Cnt, TCustCnt) + " Complete."
        DoEvents
      End If
      TRAmt$ = CStr(TaxTR.Amount)
      If InStr(TRAmt$, "N") <= 0 Then
        If TaxTR.Amount <= 0 Then
          If CustCnt = 0 Then
            CustCnt = 1
            TRInfo(0, 1) = TaxTR.CustomerRec
            TRInfo(1, 1) = 1
          Else
            For zCnt = 1 To CustCnt
              If TRInfo(0, zCnt) = TaxTR.CustomerRec Then
                TRInfo(1, zCnt) = TRInfo(1, zCnt) + 1
                gotCust = True
                Exit For
              End If
            Next
            If gotCust Then '
              gotCust = False
              GoTo DoneThisTR
            Else
              CustCnt = CustCnt + 1
              ReDim Preserve TRInfo(0 To 1, 0 To CustCnt) As Long
              TRInfo(0, CustCnt) = TaxTR.CustomerRec
              TRInfo(1, CustCnt) = 1
            End If
          End If
          BigCnt = BigCnt + 1
        End If
      End If
DoneThisTR:
      WhatTR = TaxTR.LastTrans
    Loop
  Next
  
  Print #3, "VA Tax Preconversion Transaction Checker."
  Print #3, ""

  Print #3, "  Zero transactions. "
  Print #3, "        Name                    Account Number      TR Count"
  Print #3, "------------------------------------------------------------"
  For zCnt = 1 To CustCnt
    Get #2, TRInfo(0, zCnt), TaxCust
    LSet tCustName = TaxCust.CustName
    Print #3, tCustName;
    RSet Padd = CStr(TRInfo(0, zCnt))
    Print #3, Padd;
    RSet Padd = CStr(TRInfo(1, zCnt))
    Print #3, Padd
  Next
  Print #3, ""
  Print #3, " Possible balance suspect transactions."
  Print #3, "        Name                    Account Number      Bill Bal   TRN Bal"
  Print #3, "-----------------------------------------------------------------------"

'**********************************************
'GoTo AllDoneNow:

DownHere:
  BigCnt = 0
  Label3.Caption = "Suspect Balances."
  For F2Cnt = 1 To TCustCnt
    Get #2, F2Cnt, TaxCust
    CustLastTR = TaxCust.LastTrans
    WhatTR = TaxCust.LastTrans
    Balance# = 0
    BillBal = 0
    If F2Cnt Mod 10 = 0 Then
      Label4.Caption = MakePctComp(F2Cnt, TCustCnt) + " Complete."
      DoEvents
    End If
    
    Balance# = GetCustBalance(F2Cnt, 0)
    BillBal = GetThisBillBal(CustLastTR, WhatTR, CInt(TaxTR.TaxYear))
    
    If BillBal <> Balance# Then
      LSet tCustName = TaxCust.CustName
      Print #3, tCustName;
      RSet Padd = CStr(F2Cnt)
      Print #3, Padd;
      RSet Padd = Using$("#####.00", Balance#)
      Print #3, Padd;
      RSet dPad = Using$("#####.00", BillBal)
      Print #3, dPad
      BigCnt = BigCnt + 1
    End If
  Next
  
  Label3.Caption = "Orphan Transactions."
  BigCnt = 0
  For WhatTR = 1 To TRRecCnt
    Get #1, WhatTR, TaxTR
    Label4.Caption = MakePctComp(WhatTR, TRRecCnt) + " Complete."
    DoEvents
    If TaxTR.CustomerRec <= 0 Then
      BigCnt = BigCnt + 1
    End If
  Next
  
  Print #3, ""
  Print #3, ""
  Print #3, "-----------------------------------------------------------------------"
  Print #3, "  Orp TR Count: "; BigCnt
  
AllDoneNow:
  Close
  
OverThis:

AllDone:
  
  Label3 = "Analysis Complete."
  Label4 = ""
  DoEvents
  Command1.Caption = "EXIT"
  Shell "notepad TXReport.txt", vbNormalFocus
  blnInProcs = False
  blnDoneFix = True
  Command1.Enabled = True
  

End Sub

Private Function GetThisBillBal(CustLastTR&, ByVal ThisTR&, WhatYear%) As Double
  Dim TaxTrans As TaxTransactionType
  Dim TestBal As Double
  Dim ThisRec As Long
  Dim TRecLen As Integer
  Dim PPTRADisc As Double
  'ThisRec = ThisTR ' TaxCust.LastTrans
  PPTRADisc = 0
  TestBal = 0
  CreditUsed = 0
  OverPaid = 0
  
  TRecLen = Len(TaxTrans)
  
  Open TaxTransFile For Random As #10 Len = TRecLen
  ThisRec = CustLastTR
  Do While ThisRec > 0
    Get #10, ThisRec, TaxTrans
'    If TaxTrans.CustomerRec = 1538 Then Stop
'    If WhatYear <> TaxTrans.TaxYear Then
'      GoTo SkipItYear
'    End If
    
    If TaxTrans.TranType = 10 Then 'added 10/3/06 adjust bill down affecting credit
      TestBal = TestBal + TaxTrans.Revenue.PrePaidUsed
      GoTo SkipItYear:
    End If
    If TaxTrans.TranType = 11 Then 'added 10/3/06 prepay adjust down
      TestBal# = wRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo SkipItYear:
    End If
    If TaxTrans.TranType = 24 Then 'added 10/3/06 adjust bill up affecting credit
      TestBal# = wRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo SkipItYear:
    End If
    If TaxTrans.TranType = 21 Or TaxTrans.TranType = 22 Then
'      TestBal# = wRound#(TestBal# - TaxTrans.Revenue.PrePaidAmt)'remmed on 2/2/07
      OverPaid = wRound(OverPaid + TaxTrans.Revenue.PrePaidAmt) 'added 2/2/07
      GoTo SkipItYear
    End If
    If TaxTrans.TranType = 12 Then 'refund on prepay 7/17/06
      TestBal# = wRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo SkipItYear
    End If
    If TaxTrans.TranType = 9 Then 'credit applied at billing  'added 2/2/07
      CreditUsed = wRound(CreditUsed + TaxTrans.Revenue.PrePaidUsed)
      'Jumps everywhare
      GoTo SkipItYear
    End If
    If TaxTrans.TranType = 1 Then
      TestBal# = wRound#(TestBal# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      TestBal# = wRound#(TestBal# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
      TestBal# = wRound#(TestBal# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
      TestBal# = wRound#(TestBal# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
      TestBal# = wRound#(TestBal# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt))
      TestBal# = wRound#(TestBal# - TaxTrans.PPTRADisc)
    End If
SkipItYear:
    ThisRec = TaxTrans.LastTrans
  Loop
  If OverPaid = 0 Then CreditUsed = 0 'added 2/20/07
  
  TestBal = wRound(TestBal - (OverPaid - CreditUsed)) 'added 2/2/07
  
  Close #10
  
  GetThisBillBal = TestBal

End Function

Public Function wRound(n As Double) As Double
  wRound = Int(n * 100 + 0.500000001) / 100
End Function

Public Function Date2Num(TheDate$) As Integer
 'useful function throughout program...
 'takes a string date and converts into an integer number based on 12/31/1979
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
  On Error GoTo WhatEver
  Kill "TRANS.BAK"
WhatEver:
  On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
  'On Error Resume Next
  If blnInProcs Then
    Cancel = 1
  End If
End Sub

Public Function MakePctComp(ByVal Cnt As Long, ByVal TotalCnt As Long) As String
  Dim PctComp As Long
  Dim RetStr As String
  PctComp = Int((Cnt / TotalCnt) * 100)
  RetStr = CStr(PctComp) + "%"
  MakePctComp = RetStr
End Function

Private Sub RelinkTransactions()
  Dim TaxCust As TaxCustType
  Dim NumOfTCRecs As Long
  Dim TaxTran As TaxTransactionType
  Dim TCHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TTHandle As Integer
  Dim Cnt As Long
  
  Dim TaxCustLen As Integer
  Dim TaxCustRec As TaxCustType
  Dim TaxTransLen As Integer
  Dim TaxTransRate As TaxTransactionType

  TaxCustLen = Len(TaxCustRec)
  TaxTransLen = Len(TaxTransRate)
  Label3.Caption = "Relinking Transactions."
  TCHandle = FreeFile
  Open TaxCustFile For Random Shared As TCHandle Len = TaxCustLen
  NumOfTCRecs& = LOF(TCHandle) / Len(TaxCustRec)

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
    If Cnt& Mod 10 = 0 Then
      Label4.Caption = MakePctComp(Cnt&, NumOfTTRecs&) + " Complete."
      DoEvents
    End If
  Next Cnt
  
  Close
 ' Exit Sub
End Sub

Private Function GetCustBalance(ByVal RecNo&, TaxYear As Integer) As Double
  Dim TaxTran As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim TaxCustRec As TaxCustType
  Dim TaxCustLen As Integer
  Dim TaxTransLen As Integer
  Dim TaxTransRate As TaxTransactionType

  Dim PrevTranRec&
  Dim GTOwed#
  Dim TPaid#
  Dim GTPaid#

  TaxCustLen = Len(TaxCustRec)
  TaxCustHandle = FreeFile
  Open TaxCustFile For Random Shared As TaxCustHandle Len = TaxCustLen
  'NumOfTaxCustRec = LOF(TaxCustHandle) / Len(TaxCustRec)
  Get TaxCustHandle, RecNo&, TaxCustRec
  Close TaxCustHandle

  TaxTransLen = Len(TaxTransRate)
  THandle = FreeFile
  Open TaxTransFile For Random Shared As THandle Len = TaxTransLen
'  NumOfTaxTransRecs = LOF(TaxTransHandle) / Len(TaxTransRate)

  PrevTranRec& = TaxCustRec.LastTrans
  GTOwed = 0
  TPaid = 0
  GTPaid = 0

'  TaxYear = 2005
  If PrevTranRec& > 0 Then
    Do While PrevTranRec& > 0
      Get THandle, PrevTranRec&, TaxTran
      TaxTran.OperNum = TaxTran.OperNum
      Select Case TaxTran.TranType
      Case 1    'bill
        GTOwed# = wRound#(GTOwed# + TaxTran.Amount)
      Case 2    'payment
'        TPaid# = wRound#(TPaid# + TaxTran.Amount)
'        GTPaid# = wRound#(GTPaid# + TaxTran.Amount)
        TPaid# = wRound#(TPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        GTPaid# = wRound#(GTPaid# + TaxTran.Amount + TaxTran.DiscAmt)
      Case 3    'release
        GTOwed# = wRound#(GTOwed# - TaxTran.Amount)
      Case 4    'interest
        GTOwed# = wRound#(GTOwed# + TaxTran.Amount)
      Case 5
'        Stop
        GTOwed# = wRound#(GTOwed# + TaxTran.Amount)
      Case 6    'collect/add cost
        GTOwed# = wRound#(GTOwed# + TaxTran.Amount)
      Case 7    'adjust paid down
        If TaxTran.CustPin = 0 Then
          TPaid# = wRound#(TPaid# + TaxTran.Amount)
          GTPaid# = wRound#(GTPaid# + TaxTran.Amount)
        Else
          TPaid# = wRound#(TPaid# - TaxTran.Amount)
          GTPaid# = wRound#(GTPaid# - TaxTran.Amount)
        End If
      Case 8    'misc cost
        GTOwed# = wRound#(GTOwed# + TaxTran.Amount)
      Case 9    'credit applied at billing
        TPaid# = wRound#(TPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        GTPaid# = wRound#(GTPaid# + TaxTran.Amount + TaxTran.DiscAmt)
      Case 13 'adjust bill down
        GTOwed# = wRound#(GTOwed# - TaxTran.Amount)
      Case 14 'adjust bill up
        GTOwed# = wRound#(GTOwed# + TaxTran.Amount)
      Case 21    'payment plus overpayment
        TPaid# = wRound#(TPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        GTPaid# = wRound#(GTPaid# + TaxTran.Amount + TaxTran.DiscAmt)
      Case 22    'overpayment only
        TPaid# = wRound#(TPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        GTPaid# = wRound#(GTPaid# + TaxTran.Amount + TaxTran.DiscAmt)
      Case 10    'adjust pay down affecting credit balance
        TPaid# = wRound#(TPaid - TaxTran.Amount)
        GTPaid# = wRound#(GTPaid - TaxTran.Amount)
      Case 11    'adjust prepay down
        TPaid# = wRound#(TPaid - TaxTran.Amount)
        GTPaid# = wRound#(GTPaid - TaxTran.Amount)
      Case 12    'refund total prepay balance
        TPaid# = wRound#(TPaid - TaxTran.Amount)
        GTPaid# = wRound#(GTPaid - TaxTran.Amount)
      Case 24    'adjust bill up affecting credit balance
        GTOwed# = wRound#(GTOwed# + TaxTran.Amount)
      Case Else
        Print "TRType: "; TaxTran.TranType
        Stop
'        BillType$ = "?????"
      End Select
MoveAlong:
      PrevTranRec& = TaxTran.LastTrans
    Loop

    GetCustBalance# = wRound#(GTOwed# - GTPaid#)
  Else
    GetCustBalance# = 0
  End If

  Close THandle

End Function

Public Static Function Using$(ByVal fmt As String, ByVal Number As Double)
  Dim TempNumber As String
  Dim FmtNumber As String
  Dim TempLen As Integer
  Dim BuckPos As Integer, FmtLen As Integer
  FmtLen = Len(fmt)
  BuckPos = InStr(fmt, "$")
  If BuckPos = 1 Then
    fmt = Right$(fmt, FmtLen - 1)
  ElseIf BuckPos > 1 Then
    fmt = Left$(fmt, BuckPos - 1) + Mid$(fmt, BuckPos + 1)
  End If
  FmtNumber = Space$(FmtLen)
  TempNumber = Format(Number, fmt)
  TempLen = Len(TempNumber)
  If TempLen >= 2 Then
    If Mid$(TempNumber, (TempLen - 1), 1) = "." Then
      TempNumber = TempNumber + "0"
    End If
  End If
  If Right$(TempNumber, 1) = "." Then
    TempNumber = TempNumber + "00"
  End If
  If BuckPos > 0 Then
    TempNumber = "$" + TempNumber
  End If
  RSet FmtNumber = TempNumber
  Using = FmtNumber
  
End Function

