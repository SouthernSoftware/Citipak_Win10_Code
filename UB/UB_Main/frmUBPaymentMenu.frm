VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUBPaymentMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payments, Deposits Menu"
   ClientHeight    =   8865
   ClientLeft      =   3930
   ClientTop       =   1890
   ClientWidth     =   12210
   Icon            =   "frmUBPaymentMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPrintJournalN 
      Caption         =   "Print Transaction Journal &Name Order"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3864
      TabIndex        =   5
      Top             =   5313
      Width           =   4524
   End
   Begin VB.CommandButton cmdGetLockbox 
      Caption         =   "&Read LockBox Payment File"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3870
      TabIndex        =   6
      Top             =   5940
      Width           =   4524
   End
   Begin VB.CommandButton cmdDelDep 
      Caption         =   "De&lete Deposit Transaction"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3864
      TabIndex        =   3
      Top             =   4071
      Width           =   4524
   End
   Begin VB.CommandButton cmdExitMenu 
      Caption         =   "E&xit to Previous Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3864
      TabIndex        =   8
      Top             =   7176
      Width           =   4524
   End
   Begin VB.CommandButton cmdDelPay 
      Caption         =   "D&elete Payment Transaction"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3864
      TabIndex        =   2
      Top             =   3450
      Width           =   4524
   End
   Begin VB.CommandButton cmdPostPayments 
      Caption         =   "P&ost Transaction Entries"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3864
      TabIndex        =   7
      Top             =   6555
      Width           =   4524
   End
   Begin VB.CommandButton cmdPrintJournalE 
      Caption         =   "P&rint Transaction Journal Entry Order"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3864
      TabIndex        =   4
      Top             =   4692
      Width           =   4524
   End
   Begin VB.CommandButton cmdDeposits 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Deposit Transaction Entry/Edit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3870
      TabIndex        =   1
      Top             =   2829
      Width           =   4524
   End
   Begin VB.CommandButton cmdPayment 
      BackColor       =   &H008F8265&
      Caption         =   "&Payment Transaction Entry/Edit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3864
      MaskColor       =   &H8000000F&
      TabIndex        =   0
      Top             =   2208
      Width           =   4524
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   9
      Top             =   8508
      Width           =   12216
      _ExtentX        =   21537
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7144
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "11:26 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "3/14/2019"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   1092
      Left            =   1788
      Top             =   744
      Width           =   8652
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2508
      X2              =   2508
      Y1              =   2064
      Y2              =   7944
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2508
      X2              =   3228
      Y1              =   7944
      Y2              =   7944
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2388
      X2              =   3348
      Y1              =   2064
      Y2              =   2064
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   2400
      X2              =   3360
      Y1              =   1944
      Y2              =   1944
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2388
      X2              =   2388
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   3348
      X2              =   3348
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   1
      X1              =   8988
      X2              =   8988
      Y1              =   2064
      Y2              =   7944
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   1
      X1              =   8988
      X2              =   9708
      Y1              =   7944
      Y2              =   7944
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   2
      X1              =   8868
      X2              =   9828
      Y1              =   2064
      Y2              =   2064
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   3
      X1              =   8868
      X2              =   9828
      Y1              =   1944
      Y2              =   1944
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   8868
      X2              =   8868
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   9828
      X2              =   9828
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Payments, Deposits Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3540
      TabIndex        =   10
      Top             =   1104
      Width           =   5148
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2508
      Top             =   2064
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2388
      Top             =   1824
      Width           =   972
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   8988
      Top             =   2064
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   1
      Left            =   8868
      Top             =   1824
      Width           =   972
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1788
      Top             =   624
      Width           =   8652
   End
End
Attribute VB_Name = "frmUBPaymentMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim DefPayDate As String
Public BoxFileName As String, InvalidDate As Boolean
Dim DistArray() As DistArrayType
Dim NumofRevs As Integer
Public Sub setstuff(dt As String)
DefPayDate = dt
End Sub

Private Sub cmdDelDep_Click()
'  If Not OPERNUM = 98 Then
    frmPaymentDelete.Wherefrom 2
    'Load frmPaymentDelete
    DoEvents
    frmPaymentDelete.Show
 ' End If
End Sub

Private Sub cmdDelPay_Click()
'  If Not OPERNUM = 98 Then
    frmPaymentDelete.Wherefrom 1
    'Load frmPaymentDelete
    DoEvents
    frmPaymentDelete.Show
'  End If
End Sub

Private Sub cmdDeposits_Click()
  Dim FntSize As Integer, RecpPort As String
  Dim RP As Integer, lenRP As Integer
  Dim RcptPrnFile As ReceiptPRNType
'  If Not OPERNUM = 98 Then
  frmInfo.Label1 = "Verifying Receipt Printer..."
  frmInfo.Show
  DoEvents
  If Not Exist(RcptFileName$) Then
    Unload frmInfo
    ReDim MsgText(0 To 5) As String
    FntSize = frmMsgDialog.Label(1).FontSize
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    MsgText(0) = "WARNING:"
    MsgText(1) = ""
    MsgText(2) = "RECEIPT SETUP FILE NOT FOUND!"
    MsgText(3) = "If you continue receipt printing"
    MsgText(4) = "will be disabled."
    MsgText(5) = "Receipt setup option is on CitiPak Main Menu."
    If GetOKorNot(MsgText()) Then
     UBLog "USER WANTS TO CONTINUE!"
    Else
     UBLog "USER ABORTED."
     Exit Sub
    End If
  Else
    RP = FreeFile
    lenRP = Len(RcptPrnFile)
    Open RcptFileName$ For Random Shared As RP Len = lenRP
    Get RP, 1, RcptPrnFile
    RecpPort = QPTrim(RcptPrnFile.RcpPort)
    Close
    If RcptPrnFile.PrnDefYN = 1 Then
      On Local Error GoTo noprnfound
      Open RecpPort For Output As RP
      Close RP
     End If
  End If
  frmDepositPayment.Wheretogo frmUBPaymentMenu, frmUBPaymentMenu, , DefPayDate
  DoEvents
  frmDepositPayment.Show
  Unload frmInfo '  End If
  'Unload frmUBPaymentMenu
Exit Sub
noprnfound:
        Unload frmInfo
        ReDim MsgText(0 To 5) As String
        FntSize = frmMsgDialog.Label(1).FontSize
        frmMsgDialog.Label(1).FontSize = (FntSize + 2)
        MsgText(0) = "WARNING:"
        MsgText(1) = ""
        MsgText(2) = "RECEIPT PRINTER NOT FOUND!"
        MsgText(3) = "If you continue receipt printing"
        MsgText(4) = "will be disabled."
        MsgText(5) = "Receipt setup option is on CitiPak Main Menu."
        If GetOKorNot(MsgText()) Then
          UBLog "USER WANTS TO CONTINUE!"
          frmDepositPayment.Wheretogo frmUBPaymentMenu, frmUBPaymentMenu, , DefPayDate
          DoEvents
          frmDepositPayment.Show
        Else
          UBLog "USER ABORTED."
          Exit Sub
        End If
End Sub

'Private Sub cmdCustConsumpRpt_Click()
'  frmCustEditLookUP.Caption = "Customer Consumption History"
'  frmCustEditLookUP.Label1.Caption = "Customer Consumption History"
'  frmCustEditLookUP.Wheretogo frmUBCustMenu, frmRptCustConsHist
'  DoEvents
'  frmCustEditLookUP.Show
'  Unload frmUBCustMenu
''  Load frmRptCustHistory
''  frmRptCustHistory.RptType = True
''  DoEvents
''  frmRptCustHistory.Caption = "Customer Consumption History"
''  frmRptCustHistory.Label1 = frmRptCustHistory.Caption
''  frmRptCustHistory.fpDetailFlag.Visible = False
''  frmRptCustHistory.DetailLabel.Visible = False
''  frmRptCustHistory.Wheretogo frmUBCustMenu, frmRptCustHistory
''  frmRptCustHistory.Show
''  Unload frmUBCustMenu
'End Sub
'
'Private Sub cmdCustTransRpt_Click()
'  frmCustEditLookUP.Caption = "Customer Transaction History"
'  frmCustEditLookUP.Label1.Caption = "Customer Transaction History"
'  frmCustEditLookUP.Wheretogo frmUBCustMenu, frmRptCustTranHist
'  DoEvents
'  frmCustEditLookUP.Show
'  Unload frmUBCustMenu
'
''  Load frmRptCustHistory
''  frmRptCustHistory.RptType = False
''  DoEvents
''  frmRptCustHistory.Caption = "Customer Transaction History"
''  frmRptCustHistory.Label1 = frmRptCustHistory.Caption
''  frmRptCustHistory.Wheretogo frmUBCustMenu, frmRptCustHistory
''  frmRptCustHistory.Show
''  Unload frmUBCustMenu
'End Sub
'
'Private Sub cmdDeleteCustomer_Click()
'  frmCustEditLookUP.Caption = "Customer Delete Search"
'  frmCustEditLookUP.Label1.Caption = "Customer Delete Search"
'  frmCustEditLookUP.Wheretogo frmUBCustMenu, frmCustDelete, , 1
'  'Load frmCustEditLookUP
'  DoEvents
'  frmCustEditLookUP.Show
'  Unload frmUBCustMenu
'
'End Sub
'
'Private Sub cmdEditCustomer_Click()
'  frmCustEditLookUP.Wheretogo frmUBCustMenu, frmCustAddEdit
''  Load frmCustEditLookUP
'  DoEvents
'  frmCustEditLookUP.Show
'  Unload frmUBCustMenu
'End Sub

Private Sub cmdExitMenu_Click()
  Load frmUBMainMenu
  DoEvents
  frmUBMainMenu.Show
  Unload Me
End Sub


Private Sub cmdGetLockbox_Click()
  Dim FntSize As Integer, dotheLB As Boolean, lbx As Integer
  dotheLB = False
  If OPERNUM = 98 Then
    dotheLB = True
'  ElseIf PWcnt = 0 And PWUser = "Sosoft Support" Then
'    dotheLB = True
  End If
  lbx = GetDefaultLockbox%
  If lbx <> 6 Then
    If lbx <> 8 Then
      MsgBox "Error with Lock Box Payment File Default.", vbOKOnly, "Procedure Canceled"
      Exit Sub
    End If
  End If

  If dotheLB = True Then
    BoxFileName$ = ""
    frmLockBoxPay.Label1 = "Enter the Name of the LockBox File."
    frmLockBoxPay.Show 1, Me
    If frmLockBoxPay.Exout <> 1 Then
      DeActivateControls Me
      If lbx = 6 Then
        GetLockBoxPayments6
      ElseIf lbx = 8 Then
        GetLockBoxPayments8
      End If
      ActivateControls Me
    Else
      Exit Sub
    End If
  Else
    frmLBWarning.Label1.Caption = "Access Denied"
    frmLBWarning.Label3.Caption = "Operator 98 Access Only"
    frmLBWarning.Label4.Caption = "Sign In with Operator 98 Password."
    frmLBWarning.Show 1
  End If
End Sub

Private Sub cmdPayment_Click()
  Dim FntSize As Integer, RecpPort As String
  Dim RP As Integer, lenRP As Integer
  Dim RcptPrnFile As ReceiptPRNType
  'If Not OPERNUM = 98 Then
    frmInfo.Label1 = "Verifying Receipt Printer..."
    frmInfo.Show
    DoEvents
    If Not Exist(RcptFileName$) Then
      Unload frmInfo
      ReDim MsgText(0 To 5) As String
      FntSize = frmMsgDialog.Label(1).FontSize
      frmMsgDialog.Label(1).FontSize = (FntSize + 2)
      MsgText(0) = "WARNING:"
      MsgText(1) = ""
      MsgText(2) = "RECEIPT SETUP FILE NOT FOUND!"
      MsgText(3) = "If you continue receipt printing"
      MsgText(4) = "will be disabled."
      MsgText(5) = "Receipt setup option is on CitiPak Main Menu."
      If GetOKorNot(MsgText()) Then
        UBLog "USER WANTS TO CONTINUE!"
      Else
        UBLog "USER ABORTED."
        Exit Sub
      End If
    Else
      RP = FreeFile
      lenRP = Len(RcptPrnFile)
      Open RcptFileName$ For Random Shared As RP Len = lenRP
      Get RP, 1, RcptPrnFile
      RecpPort = QPTrim(RcptPrnFile.RcpPort)
      Close
      If RcptPrnFile.PrnDefYN = 1 Then
        On Local Error GoTo noprnfound
        Open RecpPort For Output As RP
        Close RP
       End If
    End If
    frmPaymentEntry.Wheretogo frmUBPaymentMenu, frmUBPaymentMenu, , DefPayDate
    DoEvents
    frmPaymentEntry.Show
    Unload frmInfo
    'Unload frmUBPaymentMenu
 'End If
Exit Sub
noprnfound:
        Unload frmInfo
        ReDim MsgText(0 To 5) As String
        FntSize = frmMsgDialog.Label(1).FontSize
        frmMsgDialog.Label(1).FontSize = (FntSize + 2)
        MsgText(0) = "WARNING:"
        MsgText(1) = ""
        MsgText(2) = "RECEIPT PRINTER NOT FOUND!"
        MsgText(3) = "If you continue receipt printing"
        MsgText(4) = "will be disabled."
        MsgText(5) = "Receipt setup option is on CitiPak Main Menu."
        If GetOKorNot(MsgText()) Then
          UBLog "USER WANTS TO CONTINUE!"
          frmPaymentEntry.Wheretogo frmUBPaymentMenu, frmUBPaymentMenu, , DefPayDate
          DoEvents
          frmPaymentEntry.Show
        Else
          UBLog "USER ABORTED."
          Exit Sub
        End If
End Sub

Private Sub cmdPostPayments_Click()
  Dim FntSize As Integer, PayBillName As String, PayDepoName As String
  'Dim OPERNUM As Integer
  
  PayBillName$ = UBPath$ + "UBPAY" + QPTrim$(Str$(OPERNUM)) + ".DAT"
  PayDepoName$ = UBPath$ + "UBDEP" + QPTrim$(Str$(OPERNUM)) + ".DAT"

  If FileSize&(PayBillName$) <= 0 And FileSize&(PayDepoName$) <= 0 Then
  ReDim MsgText(0 To 5) As String
   
     frmMsgDialog.RetLabel = "-2"
     FntSize = frmMsgDialog.Label(2).FontSize
     
     frmMsgDialog.Label(2).FontSize = (FntSize + 2)
     MsgText(0) = "ERROR:"
     MsgText(1) = ""
     MsgText(2) = ""
     MsgText(3) = "NO UNPOSTED PAYMENT TRANSACTIONS!"
     MsgText(4) = ""
     MsgText(5) = ""
     GetOKorNot MsgText(), True
    GoTo Exitthis
  End If
  UBLog " IN: UB POST PAYMENTS,  OPER:" + Str$(OPERNUM)
  ChkTranDate
  If InvalidDate = True Then
    UBLog "Invalid Date found ubpaypost, give opt to cancel- OPER:" + Str$(OPERNUM)
    ReDim MsgText(0 To 5) As String
    FntSize = frmMsgDialog.Label(1).FontSize
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
    frmMsgDialog.Label(4).FontSize = (FntSize + 2)
    frmMsgDialog.Label(2).FontSize = (FntSize + 2)
    MsgText(0) = "WARNING:"
    MsgText(1) = ""
    MsgText(2) = "Date of one or more payments"
    MsgText(3) = "is NOT within monthly date range."
    MsgText(4) = ""
    MsgText(5) = "OK to continue, or Cancel."
    If GetOKorNot(MsgText()) Then
      UBLog "Continue pay post with out of range dates."
    Else
      UBLog "Cancel pay post so can check dates."
      GoTo Exitthis
    End If
  End If

  DoItFlag = False
    frmNoOperatorsWarning.Label(5).Caption = "Post Payment Transactions"
    Load frmNoOperatorsWarning
    frmNoOperatorsWarning.Show vbModal
    If Not DoItFlag Then
      GoTo Exitthis
    End If
  DeActivateControls Me
  PostPayments
  ActivateControls Me
  MsgBox "Posting Payments Completed.", vbOKOnly, "Procedure Complete"
Exitthis:

End Sub

Private Sub cmdPrintJournalE_Click()
Dim FntSize As Integer, PayFileName As String, DepFileName As String
  Dim Tot As Integer
  Tot = 0

  PayFileName$ = UBPath$ + "UBPAY" + QPTrim$(Str$(OPERNUM)) + ".DAT"
  DepFileName$ = UBPath$ + "UBDEP" + QPTrim$(Str$(OPERNUM)) + ".DAT"
  If Exist(PayFileName$) Then Tot = Tot + 1
  If Exist(DepFileName$) Then Tot = Tot + 1
  ReDim MsgText(0 To 5) As String
   If Tot < 1 Then
     frmMsgDialog.RetLabel = "-2"
     FntSize = frmMsgDialog.Label(2).FontSize
     
     frmMsgDialog.Label(2).FontSize = (FntSize + 2)
     MsgText(0) = "ERROR:"
     MsgText(1) = ""
     MsgText(2) = ""
     MsgText(3) = "NO UNPOSTED PAYMENT TRANSACTIONS!"
     MsgText(4) = ""
     MsgText(5) = ""
     GetOKorNot MsgText(), True
  Else
    frmReportOpt.Show 1
    DeActivateControls Me
    If rptopt > 0 Then
     PrintTransJournal rptopt, 0
    Else
      ActivateControls Me
    End If
    
  End If
End Sub

Private Sub cmdPrintJournalN_Click()
Dim FntSize As Integer, PayFileName As String, DepFileName As String
  Dim Tot As Integer
  Tot = 0

  PayFileName$ = UBPath$ + "UBPAY" + QPTrim$(Str$(OPERNUM)) + ".DAT"
  DepFileName$ = UBPath$ + "UBDEP" + QPTrim$(Str$(OPERNUM)) + ".DAT"
  If Exist(PayFileName$) Then Tot = Tot + 1
  If Exist(DepFileName$) Then Tot = Tot + 1
  ReDim MsgText(0 To 5) As String
   If Tot < 1 Then
     frmMsgDialog.RetLabel = "-2"
     FntSize = frmMsgDialog.Label(2).FontSize
     
     frmMsgDialog.Label(2).FontSize = (FntSize + 2)
     MsgText(0) = "ERROR:"
     MsgText(1) = ""
     MsgText(2) = ""
     MsgText(3) = "NO UNPOSTED PAYMENT TRANSACTIONS!"
     MsgText(4) = ""
     MsgText(5) = ""
     GetOKorNot MsgText(), True
  Else
    frmReportOpt.Show 1
    DeActivateControls Me
    If rptopt > 0 Then
     PrintTransJournal rptopt, 1
    Else
      ActivateControls Me
    End If
    
  End If
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  'screenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpPaymentsDeposits
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExitMenu.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via PaymentMenu by " + PWUser$
        CitiTerminate
      End If
    End If
  End If
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
   ' Me.Visible = False
    Temp_Class.ResizeControls Me
   ' Me.Visible = True
   ' Me.SetFocus
  End If
  DoEvents
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyHome
      cmdPayment.SetFocus
    Case vbKeyEnd
      cmdExitMenu.SetFocus
    Case Else:
  End Select
End Sub

Private Sub PrintTransJournal(rptopt As Integer, Order As Integer) '(OPERNUM, PostDate$)
  Dim NumofRevs As Integer, UBSetupLen As Integer, RevCnt As Integer
  Dim InvRev As Integer, cnt As Long, x As Integer, Dash1 As String
  Dim UBCustRecLen As Integer, LastRev As Integer, IndexRecLen As Integer
  Dim TempRev As String, Operator As String, Page As Integer
  Dim PayFileName As String, DepFileName As String, PayJourName As String
  Dim Header As String, CMOperRecLen As Integer, UBPayRecLen As Integer
  Dim CMFile As Integer, NumRecs As Long, PayOKFlag As Boolean
  Dim DepOKFlag As Boolean, CustFile As Integer, TotalRecs As Long
  Dim RptHandle As Integer, PHandle As Integer, NumOfRecs As Long
  Dim DoneCnt As Integer, TaxExempt As Boolean, Pmnt As String
  Dim CustBook As Integer, TotalCash As Double, TotalCheck As Double
  Dim TotalAmount As Double, TotalChange As Double, TotalReceipts As Integer
  Dim RCnt As Integer, Diff As Double, Tax As Double, DepositTot As Double
  Dim GTotal As Double, TTax As Double, PostDate As String, ToPrint As String
  Dim Graph As Boolean, ReportSum1 As String, ReportSum2 As String
  Dim SumRpt1 As Integer, SumRpt2 As Integer, TotalChrge As Double
  Dim tmp As DistArrayType, SumPrnt As String, TotalChks As Integer
  Dim lngCurLow As Long, lngCurHigh As Long, AcctNo As Long, CHandle As Integer
  Dim IndexName As String, IHandle As Integer, dcnt As Long, CRec As Long
  Dim DHandle As Integer
  NumofRevs = MaxRevsCnt
  ReDim UBSetUpRec(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  
  ToPrint$ = ""
  If rptopt = 1 Then
    Graph = True
  Else
    Graph = False
  End If
  FrmShowPctComp.Label1 = "Creating Payment Journal Report"
  FrmShowPctComp.Show

  PostDate$ = Format(Now, "mm/dd/yyyy")
'  If InStr(UBSetUpRec(1).UTILNAME, "HARRISBURG") > 0 Then
'    BlockClear
'    OK = MsgBox("UBSETUP", "FINONLY")
'    Select Case OK
'    Case 1
'      StatusFlag = False
'    Case Else
'      StatusFlag = True
'    End Select
'  Else
'    StatusFlag = False
'  End If
  ReDim RevText$(1 To MaxRevsCnt)

  ReDim TaxRates(1 To 15) As Single
  ReDim TaxAmt(1 To 15) As Double

  ReDim BookTotals(0 To 99) As BookTotalType

  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  'IF StatusFlag THEN


  'IF INSTR(UBSetUp(1).UTILNAME, "AUTRY") > 0 THEN
  '  LptPort = 2
  'ELSE
  '  LPTPORT = 1
  'END IF

  LastRev = 15
   For cnt = 1 To MaxRevsCnt
    TempRev$ = QPTrim$(UBSetUpRec(1).Revenues(cnt).RevName)
    If Len(TempRev$) = 0 Then
      LastRev = cnt - 1
      Exit For
    Else
      RevText$(cnt) = TempRev$
      TaxRates(cnt) = UBSetUpRec(1).Revenues(cnt).TAXRATE
    End If
  Next

  ReDim RevAmts(1 To LastRev) As Double
  ReDim DepRevs(1 To LastRev) As Double

  'CursorOff
  Operator$ = QPTrim$(Str$(OPERNUM))
  FF$ = Chr$(12)
  Page = 0
  LineCnt = 0
  MaxLines = 55
  Dash1$ = String$(120, "-")
  PayFileName$ = UBPath$ + "UBPAY" + QPTrim$(Str$(OPERNUM)) + ".DAT"
  DepFileName$ = UBPath$ + "UBDEP" + QPTrim$(Str$(OPERNUM)) + ".DAT"
  PayJourName$ = UBPath$ + "UBPAY" + QPTrim$(Str$(OPERNUM)) + ".RPT"
  If Order = 1 Then
    Header$ = "Utility Payment/Deposit Journal(Name Order)"
  Else
    Header$ = "Utility Payment/Deposit Journal(Entry Order)"
  End If
  ReDim CMOperRec(1) As CMOperRecType
  CMOperRecLen = Len(CMOperRec(1))

  ReDim UBPaymentRec(1) As UBPaymentRecType
  UBPayRecLen = Len(UBPaymentRec(1))

'  CMFile = FreeFile
'  Open "CMOPER.DAT" For Random Shared As CMFile Len = CMOperRecLen
'  NumRecs = LOF(CMFile) \ CMOperRecLen
'
'  For cnt = 1 To NumRecs
'    Get CMFile, cnt, CMOperRec(1)
'    If CMOperRec(1).OperatorNumber = OPERNUM Then
'      Operator$ = QPTrim$(CMOperRec(1).OperatorName)
'      Exit For
'    End If
'  Next
'  Close CMFile

  If Exist(PayFileName$) And FileSize(PayFileName$) > 0 Then
    PayOKFlag = True
  End If
  If Exist(DepFileName$) And FileSize(DepFileName$) > 0 Then
    DepOKFlag = True
  End If

  If Not DepOKFlag And Not PayOKFlag Then
'    BlockClear
'    DisplayUBScrn "NOPAYJUR"
'    QPrintRC Str$(OPERNUM), 12, 34, 79
'    WaitForAction
    GoTo ExitJournal
  End If
  If Order = 1 Then
  If PayOKFlag Then
    CHandle = FreeFile
    Open PayFileName$ For Random Shared As CHandle Len = UBPayRecLen
    NumOfRecs& = LOF(CHandle) \ UBPayRecLen

    ReDim ServIndex(1 To NumOfRecs) As UBServiceAddressIndexType
    IndexRecLen = Len(ServIndex(1))
    For cnt& = 1 To NumOfRecs&
      Get CHandle, cnt&, UBPaymentRec(1)
      ServIndex(cnt).ServiceAddress = QPStripLast$(UBPaymentRec(1).CustName)
      ServIndex(cnt).RecNum = cnt
    'ShowPctComp cnt, NumCustRecs                'show user percentage complete
    Next
    Close CHandle

    lngCurLow = LBound(ServIndex)
    lngCurHigh = UBound(ServIndex)
    AddrQSort ServIndex(), lngCurLow, lngCurHigh
    IndexName$ = "ubPTemp.IDX"
    KillFile IndexName$
    IHandle = FreeFile
    Open IndexName$ For Random Shared As IHandle Len = 4
    For cnt = 1 To lngCurHigh
      CRec& = ServIndex(cnt).RecNum
      Put IHandle, cnt, CRec&
    Next
    Close IHandle
    End If
  End If
  CustFile = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As CustFile Len = UBCustRecLen
  TotalRecs& = (FileSize(DepFileName$) + FileSize(PayFileName$)) \ UBPayRecLen
  RptHandle = FreeFile
  Open PayJourName$ For Output As RptHandle
  If Graph Then
    ReportSum1$ = UBPath$ + "UBPJSUM1.RPT"
    SumRpt1 = FreeFile
    Open ReportSum1$ For Output As SumRpt1
    ReportSum2$ = UBPath$ + "UBPJSUM2.RPT"
    SumRpt2 = FreeFile
    Open ReportSum2$ For Output As SumRpt2
  End If
  If Not Graph Then
    GoSub PrintRptHeader
  End If
  If PayOKFlag Then
    'FOpenS PayFileName$, PHandle
    PHandle = FreeFile
    Open PayFileName$ For Random Shared As PHandle Len = UBPayRecLen
    NumOfRecs& = LOF(PHandle) \ UBPayRecLen
    For cnt& = 1 To NumOfRecs&
      FrmShowPctComp.ShowPctComp cnt&, NumOfRecs&
      If FrmShowPctComp.Out = True Then
        Close
        FrmShowPctComp.Out = False
        Close
        Exit Sub
      End If
      If Order = 1 Then
        AcctNo& = ServIndex(cnt).RecNum
      Else
        AcctNo& = cnt&
      End If
      Get PHandle, AcctNo&, UBPaymentRec(1)
 '*&*&*&*&*&*&**&*&
      If UBPaymentRec(1).CustAcct <= 0 Then GoTo OnlyFinalSkip
'*&%$#$%%$#$%$%$%#%#%#%
      Get CustFile, UBPaymentRec(1).CustAcct, UBCustRec(1)
'      If StatusFlag Then
'        If UBPaymentRec(1).Status <> "F" Then
'          GoTo OnlyFinalSkip
'        End If
'      End If

      GoSub GetCustBook
      DoneCnt = DoneCnt + 1
      If UBPaymentRec(1).TaxExempt = "Y" Then
        TaxExempt = True
      Else
        TaxExempt = False
      End If
      If Not Graph Then
        If LineCnt >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintRptHeader
        End If
      End If
      If UBPaymentRec(1).CASHAMT < 0 Then UBPaymentRec(1).CASHAMT = 0
      If UBPaymentRec(1).CHKAMT < 0 Then UBPaymentRec(1).CHKAMT = 0
      If Not Graph Then
        Print #RptHandle, Num2Date(UBPaymentRec(1).PAYDATE);
        Print #RptHandle, Tab(13); Using("#####", UBPaymentRec(1).CustAcct);
      Else
        ToPrint$ = Num2Date(UBPaymentRec(1).PAYDATE) + "~"
        ToPrint$ = ToPrint$ + Using("#####", UBPaymentRec(1).CustAcct) + "~"
      End If
      If OPERNUM = 99 Then
        Pmnt$ = " DFT"
      Else
        Pmnt$ = " PMT"
      End If

      'IF UBPaymentRec(1).Status = "F" THEN
      '  Pmnt$ = Pmnt$ + "*F"
      'END IF
      If Not Graph Then
        Print #RptHandle, Tab(20); UBPaymentRec(1).CustName; Pmnt$;
        Print #RptHandle, Tab(55); Using("######.##", UBPaymentRec(1).CASHAMT);
        If QPTrim(UBPaymentRec(1).TENDERTY) = "Charge" Then
          Print #RptHandle, Tab(70); Using("######.##", 0);
          Print #RptHandle, Tab(84); Using("######.##", UBPaymentRec(1).CHKAMT);
        Else
          Print #RptHandle, Tab(70); Using("######.##", UBPaymentRec(1).CHKAMT);
          Print #RptHandle, Tab(84); Using("######.##", 0);
        End If
        Print #RptHandle, Tab(99); Using("######.##", Round#(Round#(UBPaymentRec(1).CHKAMT + UBPaymentRec(1).CASHAMT) - UBPaymentRec(1).Change));
        Print #RptHandle, Tab(112); Using("######.##", UBPaymentRec(1).Change)
      Else
        ToPrint$ = ToPrint$ + QPTrim(UBPaymentRec(1).CustName) + "~" + Pmnt$ + "~"
        ToPrint$ = ToPrint$ + Using("#####,#.##", UBPaymentRec(1).CASHAMT) + "~"
        If QPTrim(UBPaymentRec(1).TENDERTY) = "Charge" Then
          ToPrint$ = ToPrint$ + "0~"
        End If
        ToPrint$ = ToPrint$ + Using("#####,#.##", UBPaymentRec(1).CHKAMT) + "~"
        If QPTrim(UBPaymentRec(1).TENDERTY) <> "Charge" Then
          ToPrint$ = ToPrint$ + "0~"
        End If
        ToPrint$ = ToPrint$ + Using("#####,#.##", Round#(Round#(UBPaymentRec(1).CHKAMT + UBPaymentRec(1).CASHAMT) - UBPaymentRec(1).Change)) + "~"
        ToPrint$ = ToPrint$ + Using("#####,#.##", UBPaymentRec(1).Change)
        Print #RptHandle, ToPrint$
        ToPrint$ = ""
      End If
      If QPTrim(UBPaymentRec(1).TENDERTY) = "Charge" Then
        BookTotals(CustBook).Charge = Round#(BookTotals(CustBook).Charge + UBPaymentRec(1).CHKAMT)
        TotalChrge# = Round#(TotalChrge# + UBPaymentRec(1).CHKAMT)
      Else
        BookTotals(CustBook).Check = Round#(BookTotals(CustBook).Check + UBPaymentRec(1).CHKAMT)
        TotalCheck# = Round#(TotalCheck# + UBPaymentRec(1).CHKAMT)
        If UBPaymentRec(1).CHKAMT > 0 Then
          TotalChks = TotalChks + 1
        End If
      End If
      BookTotals(CustBook).Cash = Round#(BookTotals(CustBook).Cash + UBPaymentRec(1).CASHAMT)
      BookTotals(CustBook).Change = Round#(BookTotals(CustBook).Change + UBPaymentRec(1).Change)
      
      TotalCash# = Round#(TotalCash# + UBPaymentRec(1).CASHAMT)
      'TotalCheck# = Round#(TotalCheck# + UBPaymentRec(1).CHKAMT)
      TotalAmount# = Round#(TotalAmount# + UBPaymentRec(1).AMTPAID)
      TotalChange# = Round#(TotalChange# + UBPaymentRec(1).Change)
      TotalReceipts = TotalReceipts + 1

      LineCnt = LineCnt + 1

      For RCnt = 1 To LastRev
        If Not TaxExempt Then
          If TaxRates(RCnt) > 0 Then
            Diff# = Round#(UBPaymentRec(1).PaidOwed(RCnt).AMTPD1 / (1 + TaxRates(RCnt)))
            Tax# = Round#(UBPaymentRec(1).PaidOwed(RCnt).AMTPD1 - Diff#)
            TaxAmt(RCnt) = Round#(TaxAmt(RCnt) + Tax#)
            RevAmts(RCnt) = Round#(RevAmts(RCnt) + (UBPaymentRec(1).PaidOwed(RCnt).AMTPD1 - Tax#))
          Else
            RevAmts(RCnt) = Round#(RevAmts(RCnt) + UBPaymentRec(1).PaidOwed(RCnt).AMTPD1)
            'IF UBPaymentRec(1).PaidOwed(RCnt).AMTPD1 < -10000 OR UBPaymentRec(1).PaidOwed(RCnt).AMTPD1 > 10000 THEN
            '  STOP
            'END IF
          End If
        Else
          RevAmts(RCnt) = Round#(RevAmts(RCnt) + UBPaymentRec(1).PaidOwed(RCnt).AMTPD1)
        End If
DebugSkip2Here:
      Next
      '*********************
      'ShowPctComp DoneCnt, TotalRecs&
OnlyFinalSkip:
    Next
    Close PHandle
  Else
    FrmShowPctComp.ShowPctComp 100, 100
  End If
    If DepOKFlag Then
      If Order = 1 Then
    CHandle = FreeFile
    Open DepFileName$ For Random Shared As CHandle Len = UBPayRecLen
    NumOfRecs& = LOF(CHandle) \ UBPayRecLen

    ReDim ServIndex(1 To NumOfRecs) As UBServiceAddressIndexType
    IndexRecLen = Len(ServIndex(1))
    For cnt& = 1 To NumOfRecs&
      Get CHandle, cnt&, UBPaymentRec(1)
      ServIndex(cnt).ServiceAddress = QPStripLast$(UBPaymentRec(1).CustName)
      ServIndex(cnt).RecNum = cnt
    'ShowPctComp cnt, NumCustRecs                'show user percentage complete
    Next
    Close CHandle

    lngCurLow = LBound(ServIndex)
    lngCurHigh = UBound(ServIndex)
    AddrQSort ServIndex(), lngCurLow, lngCurHigh
    IndexName$ = "ubDTemp.IDX"
    KillFile IndexName$
    IHandle = FreeFile
    Open IndexName$ For Random Shared As IHandle Len = 4
    For cnt = 1 To lngCurHigh
      CRec& = ServIndex(cnt).RecNum
      Put IHandle, cnt, CRec&
    Next
    Close IHandle
  End If

    PHandle = FreeFile
    Open DepFileName$ For Random Shared As PHandle Len = UBPayRecLen
    NumOfRecs& = LOF(PHandle) \ UBPayRecLen
    For cnt& = 1 To NumOfRecs&
      If Order = 1 Then
        AcctNo& = ServIndex(cnt).RecNum
      Else
        AcctNo& = cnt&
      End If

      Get PHandle, AcctNo&, UBPaymentRec(1)
      Get CustFile, UBPaymentRec(1).CustAcct, UBCustRec(1)
      GoSub GetCustBook

      DoneCnt = DoneCnt + 1
      If Not Graph Then
        If LineCnt >= MaxLines Then
          Print #RptHandle, FF$
          GoSub PrintRptHeader
        End If
      End If
      If UBPaymentRec(1).CASHAMT < 0 Then UBPaymentRec(1).CASHAMT = 0
      If UBPaymentRec(1).CHKAMT < 0 Then UBPaymentRec(1).CHKAMT = 0
      If Not Graph Then
        Print #RptHandle, Num2Date(UBPaymentRec(1).PAYDATE);
        Print #RptHandle, Tab(13); Using("#####", UBPaymentRec(1).CustAcct);
        Print #RptHandle, Tab(20); UBPaymentRec(1).CustName; ; " DEP";
        Print #RptHandle, Tab(55); Using("######.##", UBPaymentRec(1).CASHAMT);
        If QPTrim(UBPaymentRec(1).TENDERTY) = "Charge" Then
          Print #RptHandle, Tab(70); Using("######.##", 0);
          Print #RptHandle, Tab(84); Using("######.##", UBPaymentRec(1).CHKAMT);
        Else
          Print #RptHandle, Tab(70); Using("######.##", UBPaymentRec(1).CHKAMT);
          Print #RptHandle, Tab(84); Using("######.##", 0);
        End If
        Print #RptHandle, Tab(112); Using("######.##", UBPaymentRec(1).Change)
      Else
        ToPrint$ = Num2Date(UBPaymentRec(1).PAYDATE) + "~"
        ToPrint$ = ToPrint$ + Using("#####", UBPaymentRec(1).CustAcct) + "~"
        ToPrint$ = ToPrint$ + QPTrim(UBPaymentRec(1).CustName) + "~ DEP~"
        ToPrint$ = ToPrint$ + Using("#####,#.##", UBPaymentRec(1).CASHAMT) + "~"
        If QPTrim(UBPaymentRec(1).TENDERTY) = "Charge" Then
          ToPrint$ = ToPrint$ + "0~"
        End If
        ToPrint$ = ToPrint$ + Using("#####,#.##", UBPaymentRec(1).CHKAMT) + "~"
        If QPTrim(UBPaymentRec(1).TENDERTY) <> "Charge" Then
          ToPrint$ = ToPrint$ + "0~"
        End If
        ToPrint$ = ToPrint$ + " ~"   'Using("#####,#.##", Round#(Round#(UBPaymentRec(1).CHKAMT + UBPaymentRec(1).CASHAMT) - UBPaymentRec(1).CHANGE)) + "~"
        ToPrint$ = ToPrint$ + Using("#####,#.##", UBPaymentRec(1).Change)
        Print #RptHandle, ToPrint$
        ToPrint$ = ""
      End If
      If QPTrim(UBPaymentRec(1).TENDERTY) = "Charge" Then
        BookTotals(CustBook).Charge = Round#(BookTotals(CustBook).Charge + UBPaymentRec(1).CHKAMT)
      Else
        BookTotals(CustBook).Check = Round#(BookTotals(CustBook).Check + UBPaymentRec(1).CHKAMT)
      End If
      BookTotals(CustBook).Cash = Round#(BookTotals(CustBook).Cash + UBPaymentRec(1).CASHAMT)
      BookTotals(CustBook).Change = Round#(BookTotals(CustBook).Change + UBPaymentRec(1).Change)
      TotalChange# = Round#(TotalChange# + UBPaymentRec(1).Change)
      TotalCash# = Round#(TotalCash# + UBPaymentRec(1).CASHAMT)
      If QPTrim(UBPaymentRec(1).TENDERTY) = "Charge" Then
        TotalChrge# = Round#(TotalChrge# + UBPaymentRec(1).CHKAMT)
      Else
        TotalCheck# = Round#(TotalCheck# + UBPaymentRec(1).CHKAMT)
      End If
      TotalReceipts = TotalReceipts + 1
      LineCnt = LineCnt + 1
      For RCnt = 1 To LastRev
        DepositTot# = Round#(DepositTot# + UBPaymentRec(1).PaidOwed(RCnt).AMTPD1)
        DepRevs(RCnt) = Round#(DepRevs(RCnt) + UBPaymentRec(1).PaidOwed(RCnt).AMTPD1)
      Next
      'ShowPctComp DoneCnt, TotalRecs&
    Next
    Close PHandle
  Else
    FrmShowPctComp.ShowPctComp 100, 100
  End If

  GoSub PrintRptEnding

  Close
  'PrintRptFile Header$, PayJourName$, LPTPORT, -1, EntryPoint
  If Not Graph Then
    ViewPrint PayJourName$, Header$, True
    ActivateControls Me
  Else
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmUBPaymentMenu
    ARptPaymentJournal.Title = Header$
    ARptPaymentJournal.txtDate = Now
    ARptPaymentJournal.txtTown = TOWNNAME$
    ARptPaymentJournal.lblOperator = PWUser
    ARptPaymentJournal.GetName PayJourName$, ReportSum1$, ReportSum2$
    ARptPaymentJournal.startrpt
  End If
  'KillFile PayJourName$

ExitJournal:
  Exit Sub

GetCustBook:
  CustBook = Val(UBCustRec(1).Book)
  BookTotals(CustBook).Count = BookTotals(CustBook).Count + 1
Return

PrintRptHeader:
  If Not Graph Then
  Page = Page + 1
'  If StatusFlag Then
'    Print #RptHandle, "Utility Payment Receipts Journal (Final Only)"
'  Else

    Print #RptHandle, "Utility Payment/Deposit Receipts Journal"
'  End If
  Print #RptHandle, "Posting Date: "; PostDate$
  Print #RptHandle, "    Operator: "; PWUser; Tab(89); "Page #"; Page
  Print #RptHandle, ""
  Print #RptHandle, "       "; Tab(12); "        "; Tab(44); "           "; Tab(61); "            "; Tab(98); "Amount Paid"
  Print #RptHandle, " Date"; Tab(11); "Acct No      Customer"; Tab(60); "Cash"; Tab(74); "Check"; Tab(87); "Charge"; Tab(98); " on Account"; Tab(115); "Change"
  Print #RptHandle, Dash1$
  LineCnt = 6
 End If
 Return

PrintRptEnding:
  If Not Graph Then
    Print #RptHandle, Dash1$
    Print #RptHandle, "                  Totals: ";
    Print #RptHandle, Tab(54); Using("###,###.##", TotalCash#);
    Print #RptHandle, Tab(69); Using("###,###.##", TotalCheck#);
    Print #RptHandle, Tab(83); Using("###,###.##", TotalChrge#);
    Print #RptHandle, Tab(98); Using("###,###.##", TotalAmount#);
    Print #RptHandle, Tab(111); Using("###,###.##", TotalChange#)
    Print #RptHandle, "Total Number of Receipts: "; Using("##,###", TotalReceipts)
    Print #RptHandle, "  Total Number of Checks: "; Using("##,###", TotalChks)
    Print #RptHandle, FF$
  '  If StatusFlag Then
  '    Print #RptHandle, "Utility Payment Receipts Journal (Final Only)"
  '  Else
       Print #RptHandle, "Utility Payment/Deposit Receipts Journal"
  '  End If
    Print #RptHandle, "Revenue Summary"
    Print #RptHandle, ""
    Print #RptHandle, "    Revenue"; Tab(33); "Payments         Deposits               Tax"
    Print #RptHandle, Dash1$
    GTotal# = 0
    For RCnt = 1 To LastRev
      Print #RptHandle, Tab(5); RevText$(RCnt); Tab(31); Using("$###,###.##", RevAmts(RCnt)); Tab(48); Using("$###,###.##", DepRevs(RCnt)); Tab(66); Using("$###,###.##", TaxAmt(RCnt))
      GTotal# = Round#(GTotal# + RevAmts(RCnt))
      TTax# = Round#(TTax# + TaxAmt(RCnt))
    Next
    Print #RptHandle, Dash1$
    Print #RptHandle, "Revenue Totals:"; Tab(29); Using("$#,###,###.##", GTotal#); Tab(46); Using("$#,###,###.##", DepositTot#); Tab(64); Using("$#,###,###.##", TTax#)
    Print #RptHandle,             '"Deposit Total:"; TAB(35); USING "$$#####,#.##"; DepositTot#
    Print #RptHandle, "   Grand Total:"; Tab(29); Using("$#,###,###.##", Round#(GTotal# + DepositTot# + TTax#))
    Print #RptHandle, FF$
  '  If StatusFlag Then
  '    Print #RptHandle, "Utility Payment Receipts Journal (Final Only)"
  '  Else
      Print #RptHandle, "Utility Payment/Deposit Receipts Journal"
  '  End If
    Print #RptHandle, "Posting Date: "; PostDate$; "    Operator: "; Operator$
    Print #RptHandle, "Books Summary"
    Print #RptHandle, "Book#     Count            Cash          Check            Charge              Total            Change"
    Print #RptHandle, Dash1$
    For cnt = 0 To 99
      If BookTotals(cnt).Check <> 0 Or BookTotals(cnt).Cash <> 0 Or BookTotals(cnt).Charge <> 0 Then
        If cnt < 10 Then
          Print #RptHandle, Tab(2); Using("0#", cnt);
        Else
          Print #RptHandle, Tab(2); Using("##", cnt);
        End If
        Print #RptHandle, Tab(12); Using("####", BookTotals(cnt).Count);
        Print #RptHandle, Tab(22); Using("#######.##", BookTotals(cnt).Cash); Tab(37); Using("#######.##", BookTotals(cnt).Check); Tab(55); Using("#######.##", BookTotals(cnt).Charge); Tab(74); Using("#######.##", Round#(BookTotals(cnt).Cash + BookTotals(cnt).Check + BookTotals(cnt).Charge));
        Print #RptHandle, Tab(92); Using("#######.##", BookTotals(cnt).Change)
      End If
    Next
    Print #RptHandle, Dash1$
    Print #RptHandle, "Totals:"; Tab(12); Using("####", TotalReceipts);
    Print #RptHandle, Tab(22); Using("#######.##", TotalCash#);
    Print #RptHandle, Tab(37); Using("#######.##", TotalCheck#);
    Print #RptHandle, Tab(55); Using("#######.##", TotalChrge#);
    Print #RptHandle, Tab(74); Using("#######.##", Round#(TotalCash# + TotalCheck# + TotalChrge#));
    Print #RptHandle, Tab(92); Using("#######.##", TotalChange#)
    Print #RptHandle, FF$
  Else
    GTotal# = 0
    SumPrnt$ = ""
    For RCnt = 1 To LastRev
      SumPrnt$ = SumPrnt$ + RevText$(RCnt) + "~" + Using("$###,###.##", RevAmts(RCnt)) + "~" + Using("$###,###.##", DepRevs(RCnt)) + "~" + Using("$###,###.##", TaxAmt(RCnt)) + "~"
      GTotal# = Round#(GTotal# + RevAmts(RCnt))
      TTax# = Round#(TTax# + TaxAmt(RCnt))
      Print #SumRpt1, SumPrnt$
      SumPrnt$ = ""
    Next
    For cnt = 0 To 99
      If BookTotals(cnt).Check <> 0 Or BookTotals(cnt).Cash <> 0 Or BookTotals(cnt).Charge <> 0 Then
        If cnt < 10 Then
          SumPrnt$ = SumPrnt$ + Using("0#", cnt) + "~"
        Else
          SumPrnt$ = SumPrnt$ + Using("##", cnt) + "~"
        End If
        SumPrnt$ = SumPrnt$ + Using("####", BookTotals(cnt).Count) + "~"
        SumPrnt$ = SumPrnt$ + Using("#######.##", BookTotals(cnt).Cash) + "~"
        SumPrnt$ = SumPrnt$ + Using("#######.##", BookTotals(cnt).Check) + "~"
        SumPrnt$ = SumPrnt$ + Using("#######.##", BookTotals(cnt).Charge) + "~"
        SumPrnt$ = SumPrnt$ + Using("#######.##", Round#(BookTotals(cnt).Cash + BookTotals(cnt).Check + BookTotals(cnt).Charge)) + "~"
        SumPrnt$ = SumPrnt$ + Using("#######.##", BookTotals(cnt).Change)
        Print #SumRpt2, SumPrnt$
        SumPrnt$ = ""
      End If
    Next
  End If
  Return
End Sub

Private Sub PostPayments()
' OPERNUM , PostDate$
  Dim PayBillName As String, PayDepoName As String
  Dim UBCustRecLen As Integer, UBPayRecLen As Integer, UBTransRecLen As Integer
  Dim TranFile As Integer, CHandle As Integer, thandle As Integer
  Dim PHandle As Integer, NumPayRecs As Long, cnt As Long
  Dim RevAmts As Integer, NextTransRec As Long, TotalCustBalance As Double
  Dim CustChCnt As Integer
  Dim GTot As Double
  
  PayBillName$ = UBPath$ + "UBPAY" + QPTrim$(Str$(OPERNUM)) + ".DAT"
  PayDepoName$ = UBPath$ + "UBDEP" + QPTrim$(Str$(OPERNUM)) + ".DAT"
  FrmShowPctComp.Label1 = "Posting Deposit Transactions"
  FrmShowPctComp.Show

  UBLog "POSTING TRANSACTIONS START"

  ReDim UBTransRec(1) As UBTransRecType
  ReDim TUBTransRec(1) As UBTransRecType
  ReDim UBCustRec(1) As NewUBCustRecType
  ReDim UBPaymentRec(1) As UBPaymentRecType

  UBCustRecLen = Len(UBCustRec(1))
  UBPayRecLen = Len(UBPaymentRec(1))
  UBTransRecLen = Len(UBTransRec(1))

  TranFile = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As TranFile Len = UBTransRecLen
  Close TranFile


  UBLog "POSTING: DEPOSITS"
  If FileSize&(PayDepoName$) > 0 Then
    CHandle = FreeFile
    Open UBPath$ + "UBCUST.DAT" For Random Shared As CHandle Len = UBCustRecLen
    thandle = FreeFile
    Open UBPath$ + "UBTRANS.DAT" For Random Shared As thandle Len = UBTransRecLen
    PHandle = FreeFile
    Open PayDepoName$ For Random Shared As PHandle Len = UBPayRecLen

    NumPayRecs& = LOF(PHandle) \ UBPayRecLen

    'ShowProcessingScrn "Posting Deposit Transactions"
    For cnt& = 1 To NumPayRecs&
      FrmShowPctComp.ShowPctComp cnt&, NumPayRecs&
      If FrmShowPctComp.Out = True Then
        Close
        FrmShowPctComp.Out = False
        Close
        Exit Sub
      End If
      LSet UBTransRec(1) = TUBTransRec(1)
      Get PHandle, cnt&, UBPaymentRec(1) ',  UBPayRecLen
      Get CHandle, UBPaymentRec(1).CustAcct, UBCustRec(1) ',  UBCustRecLen
      UBTransRec(1).TransDate = UBPaymentRec(1).PAYDATE
      UBTransRec(1).TransType = TranDepositPayment

      '022098 Added
      If Len(QPTrim$(UBPaymentRec(1).Desc)) = 0 Then
        UBTransRec(1).TransDesc = "DEPOSIT PAYMENT"
      Else
        UBTransRec(1).TransDesc = UBPaymentRec(1).Desc
        UBTransRec(1).BillMsg = "DEPOSIT PAYMENT"
      End If
      '^^This holds the Payment Description

      'UBTransRec(1)CustLocation = UBPaymentRec(1).CUSTACCT
      UBTransRec(1).OperatorNumber = OPERNUM
      UBTransRec(1).CustAcctNo = UBPaymentRec(1).CustAcct
      UBTransRec(1).CustStatus = UBCustRec(1).Status
      UBTransRec(1).Transamt = UBPaymentRec(1).AMTPAID
      UBTransRec(1).CheckAmount = UBPaymentRec(1).CHKAMT
      UBTransRec(1).CashAmount = UBPaymentRec(1).CASHAMT

      If UBTransRec(1).CheckAmount > 0 And UBTransRec(1).CashAmount > 0 Then
        UBTransRec(1).PayTypeCode = 3
      ElseIf UBTransRec(1).CashAmount > 0 Then
        UBTransRec(1).PayTypeCode = 1
      ElseIf UBTransRec(1).CheckAmount > 0 Then
        If QPTrim(UBPaymentRec(1).TENDERTY) = "Charge" Then
          UBTransRec(1).PayTypeCode = 4
        Else
          UBTransRec(1).PayTypeCode = 2
        End If
      End If

      For RevAmts = 1 To MaxRevsCnt
        UBTransRec(1).RevAmt(RevAmts) = UBPaymentRec(1).PaidOwed(RevAmts).AMTPD1
      Next
   '05-05-97 FIX added run balance to deposit trans
   UBTransRec(1).RunBalance = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)

   UBTransRec(1).PrevTrans = UBCustRec(1).LastTrans
   NextTransRec& = (LOF(thandle) \ UBTransRecLen) + 1

   If NextTransRec& <= 0 Then
     NextTransRec& = 1
   End If
   Put thandle, NextTransRec&, UBTransRec(1) ',  UBTransRecLen


   '04-14-98 Testing
   UBCustRec(1).DepositAmt = Round#(UBCustRec(1).DepositAmt + UBTransRec(1).Transamt)

   UBCustRec(1).LastTrans = NextTransRec&
   Put CHandle, UBPaymentRec(1).CustAcct, UBCustRec(1) ',  UBCustRecLen
   'ShowPctComp cnt&, NumPayRecs&
 Next
    Close CHandle
    Close thandle
    Close PHandle

    KillFile PayDepoName$
  End If
  UBLog "POSTED:" + Str$(NumPayRecs&)
  '**********
  UBLog "POSTING: PAYMENTS"
  If FileSize&(PayBillName$) > 0 Then
    CHandle = FreeFile
    Open UBPath$ + "UBCUST.DAT" For Random Shared As CHandle Len = UBCustRecLen
    thandle = FreeFile
    Open UBPath$ + "UBTRANS.DAT" For Random Shared As thandle Len = UBTransRecLen
    PHandle = FreeFile
    Open PayBillName$ For Random Shared As PHandle Len = UBPayRecLen

    NumPayRecs& = LOF(PHandle) \ UBPayRecLen
    FrmShowPctComp.Label1 = "Posting Payment Transactions"
    FrmShowPctComp.Show

    'ShowProcessingScrn "Posting Payment Transactions"
    For cnt& = 1 To NumPayRecs&
      FrmShowPctComp.ShowPctComp cnt&, NumPayRecs&
      If FrmShowPctComp.Out = True Then
        Close
        FrmShowPctComp.Out = False
        Close
        Exit Sub
      End If

      LSet UBTransRec(1) = TUBTransRec(1)
      Get PHandle, cnt&, UBPaymentRec(1) ',  UBPayRecLen
      Get CHandle, UBPaymentRec(1).CustAcct, UBCustRec(1) ',  UBCustRecLen
      UBTransRec(1).TransDate = UBPaymentRec(1).PAYDATE
      UBTransRec(1).TransType = TranBillPayment

      '052698 Added tax exempt flag to trans rec. For payment summary report
      UBTransRec(1).TaxExempt = UBPaymentRec(1).TaxExempt

      '022098 Added
      If Len(QPTrim$(UBPaymentRec(1).Desc)) = 0 Then
        UBTransRec(1).TransDesc = "BILLING-PAYMENT"
            Else
        UBTransRec(1).TransDesc = UBPaymentRec(1).Desc
        UBTransRec(1).BillMsg = "BILLING-PAYMENT"
      End If
      '^^This holds the Payment Description
      'UBTransRec(1)CustLocation = UBPaymentRec(1).CUSTACCT
      UBTransRec(1).OperatorNumber = OPERNUM
      UBTransRec(1).CustAcctNo = UBPaymentRec(1).CustAcct
      UBTransRec(1).CustStatus = UBCustRec(1).Status
      UBTransRec(1).Transamt = UBPaymentRec(1).AMTPAID
      UBTransRec(1).CheckAmount = UBPaymentRec(1).CHKAMT
      UBTransRec(1).CashAmount = UBPaymentRec(1).CASHAMT

      If UBTransRec(1).CheckAmount > 0 And UBTransRec(1).CashAmount > 0 Then
        UBTransRec(1).PayTypeCode = 3
      ElseIf UBTransRec(1).CashAmount > 0 Then
        UBTransRec(1).PayTypeCode = 1
      ElseIf UBTransRec(1).CheckAmount > 0 Then
        If QPTrim(UBPaymentRec(1).TENDERTY) = "Charge" Then
          UBTransRec(1).PayTypeCode = 4
        Else
          UBTransRec(1).PayTypeCode = 2
        End If
      End If

      'IF UBCustRec(1).PrevBalance > 0 THEN
      '050597 changed to zero if <> zero
      If UBCustRec(1).PrevBalance <> 0 Then
        If UBTransRec(1).Transamt >= UBCustRec(1).PrevBalance Then
          UBCustRec(1).PrevBalance = 0
        ElseIf UBTransRec(1).Transamt < UBCustRec(1).PrevBalance Then
          UBCustRec(1).PrevBalance = Round#(UBCustRec(1).PrevBalance - UBTransRec(1).Transamt)
        End If
      End If
      GTot = 0
      For RevAmts = 1 To MaxRevsCnt
        UBTransRec(1).RevAmt(RevAmts) = UBPaymentRec(1).PaidOwed(RevAmts).AMTPD1
        UBCustRec(1).CurrRevAmts(RevAmts) = Round#(UBCustRec(1).CurrRevAmts(RevAmts) - UBTransRec(1).RevAmt(RevAmts))
        GTot = GTot + Round#(UBPaymentRec(1).PaidOwed(RevAmts).AMTPD1)
      Next
      'If UBPaymentRec(1).AMTRECD <> GTot Then Stop
   '  If GTot <> Round#(UBPaymentRec(1).AMTPAID) Then Stop
    
      TotalCustBalance# = 0
      For RevAmts = 1 To MaxRevsCnt
        TotalCustBalance# = Round#(TotalCustBalance# + UBCustRec(1).CurrRevAmts(RevAmts))
      Next
      UBCustRec(1).CurrBalance = Round#(TotalCustBalance# - UBCustRec(1).PrevBalance)
      '02-26-97 Was not adding prev bal
      UBTransRec(1).RunBalance = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
      UBTransRec(1).PrevTrans = UBCustRec(1).LastTrans
      UBTransRec(1).VoidFlag = False
      UBTransRec(1).FromCMFlag = False
      NextTransRec& = (LOF(thandle) \ UBTransRecLen) + 1
      If NextTransRec& <= 0 Then
        NextTransRec& = 1
      End If
      Put thandle, NextTransRec&, UBTransRec(1) ',  UBTransRecLen
      UBCustRec(1).LastTrans = NextTransRec&
      If Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance) = 0 Then
        If UBCustRec(1).Status = "B" Then
          UBCustRec(1).Status = "I"
          CustChCnt = CustChCnt + 1
          UBLog "POSTING: SET CUST STATUS to I. Acct:" + Str$(UBPaymentRec(1).CustAcct)
        End If
      End If
      Put CHandle, UBPaymentRec(1).CustAcct, UBCustRec(1) ',  UBCustRecLen
      'ShowPctComp cnt&, NumPayRecs&
    Next
    Close CHandle
    Close thandle
    Close PHandle
    KillFile PayBillName$
  End If
  UBLog "POSTED:" + Str$(NumPayRecs&)
  If CustChCnt > 0 Then
    UBLog "POSTING: CUST STATUS CHANGED:" + Str$(CustChCnt)
  End If
  UBLog "POSTING TRANSACTIONS FINISH"

'  BlockClear
'  DisplayUBScrn "UPDATEOK"
'  WaitForAction

  Erase UBTransRec, TUBTransRec, UBCustRec, UBPaymentRec

ExitPayPost:
  UBLog "OUT: UB POST PAYMENTS,  OPER:" + Str$(OPERNUM)
End Sub
'Public Function CheckPayDate(ValCheck As String)
'  Dim Month As Integer, Day As Integer, Year As Integer
'  Month = Val(Mid(ValCheck, 1, 2))
'  Day = Val(Mid(ValCheck, 4, 2))
'  Year = Val(Mid(ValCheck, 7, 4))
'  'Checks date if Blank then won't check for valid date
'  'and then checks each section, month, day and year
'  'if any section wrong then returns false value
'  If InStr(ValCheck, "_") <= 0 Then
'    If ((Month > 0) And (Month < 13)) Then
'      If Day > 0 And Day < 32 Then
'        If Year > 1979 And Year < 2099 Then
'          CheckValDate = True
'        End If
'      End If
'    End If
'  End If
'End Function
Private Sub GetLockBoxPayments6()
  Dim BoxRecLen As Integer, NumOfAcct As Long, AmtLen As Integer
  Dim UBCustRecLen As Integer, BoxFile As Integer, TTLen As Integer
  Dim NumBoxRecs As Long, PayFileName As String, CalcTotal As Double
  Dim LockTotal As Double, cnt As Double, AcctLen As Integer
  Dim CustRec As Long, AcctError As Boolean, UBCust As Integer
  Dim Msg1 As String, Msg2 As String, Msg3 As String, TAmtPaid As Double
  Dim BackName As String, CustBal As Double, AMTPAID As Double
  Dim ZZCnt As Integer, Pay98File As Integer, PayRecLen As Integer
  ReDim LockBoxRec(1) As LockBoxRecType
  BoxRecLen = Len(LockBoxRec(1))
  ReDim UBCustRec(1) As NewUBCustRecType
  Dim ThisAmt As Double, UBTransRecLen As Integer
  Dim WhatRev As Integer, UBTran As Integer
  Dim CustFile As Integer, ThisTran As Long
  Dim DZCnt As Integer
 
  
  UBCustRecLen = Len(UBCustRec(1))
  NumBoxRecs = FileSize(BoxFileName$) / BoxRecLen

  PayFileName$ = "UBPAY98.DAT"

  If NumBoxRecs < 1 Then
    'BlockClear
    'Ok = MsgBox%("UBSETUP", "NOLOCKPY")
    Msg1$ = "No LockBox Payments found."
    Msg2$ = "Process Terminated"
    GoTo ExitBoxPaymentsM
  End If

  If Not Check98File Then
    GoTo ExitBoxPayments
  End If
  LoadRevs
  NumOfAcct& = FileSize("UBCUST.DAT") / UBCustRecLen

  'ShowProcessingScrn "Checking LockBox File Totals."

'check calculated total with last record total
  BoxFile = FreeFile
  Open BoxFileName$ For Random Shared As BoxFile Len = BoxRecLen
'get last rec
  Get BoxFile, NumBoxRecs, LockBoxRec(1)
  LockTotal# = GetBoxAmount#(LockBoxRec(1).Amount)
'sum 1 to lastrec-1 total and compare
  For cnt = 1 To NumBoxRecs - 1
    Get BoxFile, cnt, LockBoxRec(1)

    AcctLen = Len(QPTrim$(LockBoxRec(1).AcctNum))
    AmtLen = Len(QPTrim$(LockBoxRec(1).Amount))
    TTLen = Len(QPTrim$(LockBoxRec(1).TenderType))

    If AcctLen = 0 And AmtLen = 0 And TTLen = 0 Then
      GoTo BlankRecSkip
    End If

    CalcTotal# = Round#(CalcTotal# + GetBoxAmount#(LockBoxRec(1).Amount))

    If CalcTotal# > LockTotal# Then
      'Exit For 'Stop
    End If
    CustRec& = GetBoxCustRec&(LockBoxRec(1).AcctNum)

'Make sure the customer record pointer is in range
    If CustRec& < 1 Or CustRec& > NumOfAcct& Then
      AcctError = True
      Exit For
    End If


BlankRecSkip:
 '   ShowPctComp cnt, NumBoxRecs
  Next
  Close BoxFile

  If AcctError Then
   ' BlockClear
   ' Ok = MsgBox%("UBSETUP", "BDLOCKCU")
    Msg1$ = "An invalid account record has been detected in this transfer file."
    Msg2$ = "Processing Terminated."
    Msg3$ = "Please contact the software support staff for instructions."
    GoTo ExitBoxPaymentsM
  End If

  If CalcTotal# <> LockTotal# Then
   ' BlockClear
   ' Ok = MsgBox%("UBSETUP", "BDLOCKPY")
    Msg1$ = "Totals are NOT in balance."
    Msg2$ = "Processing Terminated."
    Msg3$ = "Please contact the software support staff for instructions."
    GoTo ExitBoxPaymentsM
  End If

 ' QPrintRC "Creating LockBox Payment File.", 9, 26, 126

  GoSub Setup98File

  BoxFile = FreeFile
  Open BoxFileName$ For Random Shared As BoxFile Len = BoxRecLen
  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  For cnt = 1 To NumBoxRecs - 1
    Get BoxFile, cnt, LockBoxRec(1)
    AcctLen = Len(QPTrim$(LockBoxRec(1).AcctNum))
    AmtLen = Len(QPTrim$(LockBoxRec(1).Amount))
    TTLen = Len(QPTrim$(LockBoxRec(1).TenderType))

    If AcctLen = 0 And AmtLen = 0 And TTLen = 0 Then
      GoTo BlankRecSkip2
    End If

    GoSub Make98Rec
    'SmallPause
BlankRecSkip2:
  '  ShowPctComp cnt, NumBoxRecs - 1
  Next
  Close

  If Not Exist("LBBACKUP\" + BoxFileName$) Then
    Kill "LBBACKUP\" + BoxFileName$
    Name BoxFileName$ As "LBBACKUP\" + BoxFileName$
  Else
    For cnt = 1 To 26
      BackName$ = BoxFileName$
      Mid$(BackName$, Len(BackName$), 1) = Chr$(cnt + 64)
      If Not Exist("LBBACKUP\" + BackName$) Then
        Kill "LBBACKUP\" + BackName$
        Name BoxFileName$ As "LBBACKUP\" + BackName$
        Exit For
      End If
    Next
  End If
  'BlockClear
 ' DisplayUBScrn "UPDATEOK"
 ' WaitForAction
  MsgBox "UBPay98.Dat file has been created.", vbOKOnly, "Procedure Complete"
Exit Sub

Make98Rec:
'MsgBox "Got to opened 98file"
''  UBPaymentRec(1).CUSTCMNT = QPTrim(Label4.Caption)
  ReDim PayRec(1) As UBPaymentRecType
  CustRec& = GetBoxCustRec&(LockBoxRec(1).AcctNum)
  Get UBCust, CustRec&, UBCustRec(1)
  CustBal# = UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance
  PayRec(1).OPERNUM = 98
  PayRec(1).PAYDATE = Date2Num(LockBoxRec(1).PaymentDate)
  PayRec(1).CustAcct = CustRec&
  PayRec(1).CustName = UBCustRec(1).CustName
  PayRec(1).CUSTADDR = UBCustRec(1).ADDR1
  PayRec(1).AMTOWED = CustBal#
  PayRec(1).TENDERTY = "LOCK BOX"
  PayRec(1).Status = UBCustRec(1).Status
  PayRec(1).TaxExempt = UBCustRec(1).TAXEXPT
  AMTPAID# = GetBoxAmount#(LockBoxRec(1).Amount)
  TAmtPaid# = AMTPAID#
  If LockBoxRec(1).TenderType = "CH" Then
    PayRec(1).CHKAMT = AMTPAID#
    PayRec(1).CASHAMT = 0
  ElseIf LockBoxRec(1).TenderType = "CC" Then
    PayRec(1).CHKAMT = AMTPAID#
    PayRec(1).CASHAMT = 0
    PayRec(1).TENDERTY = "Charge"
  Else
    PayRec(1).CHKAMT = 0
    PayRec(1).CASHAMT = AMTPAID#
    PayRec(1).TENDERTY = "Cash"
  End If
  PayRec(1).AMTRECD = AMTPAID#
  PayRec(1).Change = 0
  PayRec(1).Desc = "LOCKBOX PAYMENT TRANS"
 'NumofRevs = MaxRevsCnt
'  For ZZCnt = 1 To NumofRevs
'    WhatRev = DistArray(ZZCnt).DistCnt
'    If WhatRev >= 0 Then
'    ThisAmt# = UBCustRec(1).CurrRevAmts(WhatRev)
'    If ThisAmt# < 0 Then
'      AMTPAID# = Round#(AMTPAID# - ThisAmt#)
'    End If
'    End If
'  Next
  
  For ZZCnt = 1 To NumofRevs
    WhatRev = DistArray(ZZCnt).DistCnt
    If WhatRev >= 0 Then
      ThisAmt# = UBCustRec(1).CurrRevAmts(WhatRev)
      If ThisAmt# <> 0 Then
        If AMTPAID# >= ThisAmt# Then
          PayRec(1).PaidOwed(WhatRev).AMTOWE1 = UBCustRec(1).CurrRevAmts(WhatRev)
          PayRec(1).PaidOwed(WhatRev).AMTPD1 = UBCustRec(1).CurrRevAmts(WhatRev)
          AMTPAID# = Round#(AMTPAID# - ThisAmt#)
        Else
          ThisAmt# = AMTPAID#
          PayRec(1).PaidOwed(WhatRev).AMTPD1 = ThisAmt#
          AMTPAID# = 0
        End If
      ElseIf AMTPAID# = 0 Then
        PayRec(1).PaidOwed(WhatRev).AMTPD1 = 0
      ElseIf ThisAmt# = 0 Then
        PayRec(1).PaidOwed(WhatRev).AMTPD1 = 0
      End If
    End If
  Next
  If AMTPAID# <> 0 Then  'if they over payed, dump it in the first rev
    PayRec(1).PaidOwed(1).AMTPD1 = (PayRec(1).PaidOwed(1).AMTPD1 + AMTPAID#)
  End If
'  For ZZCnt = 1 To 15
'    If AMTPAID# >= UBCustRec(1).CurrRevAmts(ZZCnt) Then
'      PayRec(1).PaidOwed(ZZCnt).AMTOWE1 = UBCustRec(1).CurrRevAmts(ZZCnt)
'      AMTPAID# = Round#(AMTPAID# - UBCustRec(1).CurrRevAmts(ZZCnt))
'      PayRec(1).PaidOwed(ZZCnt).AMTPD1 = PayRec(1).PaidOwed(ZZCnt).AMTOWE1
'      If AMTPAID# <= 0 Then Exit For
'    Else
'      PayRec(1).PaidOwed(ZZCnt).AMTOWE1 = AMTPAID#
'      PayRec(1).PaidOwed(ZZCnt).AMTPD1 = PayRec(1).PaidOwed(ZZCnt).AMTOWE1
'      AMTPAID# = 0
'      Exit For
'    End If
'  Next
'
'  If AMTPAID# <> 0 Then  'if they over payed, dump it in the first rev
'    PayRec(1).PaidOwed(1).AMTPD1 = (PayRec(1).PaidOwed(1).AMTPD1 + AMTPAID#)
'  End If

  PayRec(1).TOTOWED = PayRec(1).AMTOWED
  PayRec(1).AMTPAID = TAmtPaid#

  Put #Pay98File, , PayRec(1)

Return

Setup98File:
  ReDim PayRec(1) As UBPaymentRecType
  PayRecLen = Len(PayRec(1))

  Pay98File = FreeFile
  Open PayFileName$ For Output As #Pay98File
  Close Pay98File

  Pay98File = FreeFile
  Open PayFileName$ For Random Shared As Pay98File Len = PayRecLen

Return
ExitBoxPaymentsM:
  frmLBWarning.Label1.Caption = Msg1$
  frmLBWarning.Label3.Caption = Msg2$
  frmLBWarning.Label4.Caption = Msg3$
  frmLBWarning.Show 1

ExitBoxPayments:
  Close
End Sub

Private Sub GetLockBoxPayments8()
  Dim BoxRecLen As Integer, NumOfAcct As Long, AmtLen As Integer
  Dim UBCustRecLen As Integer, BoxFile As Integer, TTLen As Integer
  Dim NumBoxRecs As Long, PayFileName As String, CalcTotal As Double
  Dim LockTotal As Double, cnt As Double, AcctLen As Integer
  Dim CustRec As Long, AcctError As Boolean, UBCust As Integer
  Dim Msg1 As String, Msg2 As String, Msg3 As String, TAmtPaid As Double
  Dim BackName As String, CustBal As Double, AMTPAID As Double
  Dim ZZCnt As Integer, Pay98File As Integer, PayRecLen As Integer
  ReDim LockBoxRec(1) As LockBoxRecTypeFC
  BoxRecLen = Len(LockBoxRec(1))
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  Dim ThisAmt As Double, UBTransRecLen As Integer
  Dim WhatRev As Integer, UBTran As Integer
  Dim CustFile As Integer, ThisTran As Long
  Dim DZCnt As Integer


      'LSet Form$(1, 0) = "30501dep"
'
''*******************************************
'  If AbortFlag Then
'    GoTo ExitBoxPayments
'  End If

  NumBoxRecs = FileSize(BoxFileName$) / BoxRecLen

  PayFileName$ = "UBPAY98.DAT"

  If NumBoxRecs < 1 Then
    'BlockClear
    'Ok = MsgBox%("UBSETUP", "NOLOCKPY")
    Msg1$ = "No LockBox Payments found."
    Msg2$ = "Process Terminated"
    GoTo ExitBoxPaymentsM
  End If

  If Not Check98File Then
    GoTo ExitBoxPayments
  End If
  LoadRevs
  NumOfAcct& = FileSize("UBCUST.DAT") / UBCustRecLen

  'ShowProcessingScrn "Checking LockBox File Totals."

'check calculated total with last record total
  BoxFile = FreeFile
  Open BoxFileName$ For Random Shared As BoxFile Len = BoxRecLen
'get last rec
  Get BoxFile, NumBoxRecs, LockBoxRec(1)
  LockTotal# = GetBoxAmount#(LockBoxRec(1).Amount)
'sum 1 to lastrec-1 total and compare
  For cnt = 1 To NumBoxRecs - 1
    Get BoxFile, cnt, LockBoxRec(1)

    AcctLen = Len(QPTrim$(LockBoxRec(1).AcctNum))
    AmtLen = Len(QPTrim$(LockBoxRec(1).Amount))
    TTLen = Len(QPTrim$(LockBoxRec(1).TenderType))

    If AcctLen = 0 And AmtLen = 0 And TTLen = 0 Then
      GoTo BlankRecSkip
    End If

    CalcTotal# = Round#(CalcTotal# + GetBoxAmount#(LockBoxRec(1).Amount))

    If CalcTotal# > LockTotal# Then
      'Exit For 'Stop
    End If
    CustRec& = GetBoxCustRec&(LockBoxRec(1).AcctNum)

'Make sure the customer record pointer is in range
    If CustRec& < 1 Or CustRec& > NumOfAcct& Then
      AcctError = True
      Exit For
    End If


BlankRecSkip:
 '   ShowPctComp cnt, NumBoxRecs
  Next
  Close BoxFile

  If AcctError Then
   ' BlockClear
   ' Ok = MsgBox%("UBSETUP", "BDLOCKCU")
    Msg1$ = "An invalid account record has been detected in this transfer file."
    Msg2$ = "Processing Terminated."
    Msg3$ = "Please contact the software support staff for instructions."
    GoTo ExitBoxPaymentsM
  End If

  If CalcTotal# <> LockTotal# Then
   ' BlockClear
   ' Ok = MsgBox%("UBSETUP", "BDLOCKPY")
    Msg1$ = "Totals are NOT in balance."
    Msg2$ = "Processing Terminated."
    Msg3$ = "Please contact the software support staff for instructions."
    GoTo ExitBoxPaymentsM
  End If

 ' QPrintRC "Creating LockBox Payment File.", 9, 26, 126

  GoSub Setup98File

  BoxFile = FreeFile
  Open BoxFileName$ For Random Shared As BoxFile Len = BoxRecLen

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen

  For cnt = 1 To NumBoxRecs - 1
    Get BoxFile, cnt, LockBoxRec(1)
    AcctLen = Len(QPTrim$(LockBoxRec(1).AcctNum))
    AmtLen = Len(QPTrim$(LockBoxRec(1).Amount))
    TTLen = Len(QPTrim$(LockBoxRec(1).TenderType))

    If AcctLen = 0 And AmtLen = 0 And TTLen = 0 Then
      GoTo BlankRecSkip2
    End If

    GoSub Make98Rec
    'SmallPause
BlankRecSkip2:
  '  ShowPctComp cnt, NumBoxRecs - 1
  Next
  Close

  If Not Exist("LBBACKUP\" + BoxFileName$) Then
    Kill "LBBACKUP\" + BoxFileName$
    Name BoxFileName$ As "LBBACKUP\" + BoxFileName$
  Else
    For cnt = 1 To 26
      BackName$ = BoxFileName$
      Mid$(BackName$, Len(BackName$), 1) = Chr$(cnt + 64)
      If Not Exist("LBBACKUP\" + BackName$) Then
        Kill "LBBACKUP\" + BackName$
        Name BoxFileName$ As "LBBACKUP\" + BackName$
        Exit For
      End If
    Next
  End If
  'BlockClear
 ' DisplayUBScrn "UPDATEOK"
 ' WaitForAction
  MsgBox "UBPay98.Dat file has been created.", vbOKOnly, "Procedure Complete"
Exit Sub

Make98Rec:
''  UBPaymentRec(1).CUSTCMNT = QPTrim(Label4.Caption)
  ReDim PayRec(1) As UBPaymentRecType
  CustRec& = GetBoxCustRec&(LockBoxRec(1).AcctNum)
  Get UBCust, CustRec&, UBCustRec(1)
  CustBal# = UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance
  PayRec(1).OPERNUM = 98
  PayRec(1).PAYDATE = Date2Num(LockBoxRec(1).PaymentDate)
  PayRec(1).CustAcct = CustRec&
  PayRec(1).CustName = UBCustRec(1).CustName
  PayRec(1).CUSTADDR = UBCustRec(1).ADDR1
  PayRec(1).AMTOWED = CustBal#
  'PayRec(1).TENDERTY = "LOCK BOX"
  PayRec(1).Status = UBCustRec(1).Status
  PayRec(1).TaxExempt = UBCustRec(1).TAXEXPT
  AMTPAID# = GetBoxAmount#(LockBoxRec(1).Amount)
  TAmtPaid# = AMTPAID#
  If LockBoxRec(1).TenderType = "CH" Then
    PayRec(1).CHKAMT = AMTPAID#
    PayRec(1).CASHAMT = 0
    PayRec(1).TENDERTY = "Check"
  ElseIf LockBoxRec(1).TenderType = "CC" Then
    PayRec(1).CHKAMT = AMTPAID#
    PayRec(1).CASHAMT = 0
    PayRec(1).TENDERTY = "Charge"
  Else
    PayRec(1).CHKAMT = 0
    PayRec(1).CASHAMT = AMTPAID#
    PayRec(1).TENDERTY = "Cash"
  End If
  PayRec(1).AMTRECD = AMTPAID#
  PayRec(1).Change = 0
  PayRec(1).Desc = "LOCKBOX PAYMENT TRANS"

  
'  For ZZCnt = 1 To NumofRevs
'    WhatRev = DistArray(ZZCnt).DistCnt
'    If WhatRev >= 0 Then
'    ThisAmt# = UBCustRec(1).CurrRevAmts(WhatRev)
'    If ThisAmt# < 0 Then
'      AMTPAID# = Round#(AMTPAID# - ThisAmt#)
'    End If
'    End If
'  Next
  
  For ZZCnt = 1 To NumofRevs
    WhatRev = DistArray(ZZCnt).DistCnt
    If WhatRev >= 0 Then
      ThisAmt# = UBCustRec(1).CurrRevAmts(WhatRev)
      If ThisAmt# <> 0 Then
        If AMTPAID# >= ThisAmt# Then
          PayRec(1).PaidOwed(WhatRev).AMTOWE1 = UBCustRec(1).CurrRevAmts(WhatRev)
          PayRec(1).PaidOwed(WhatRev).AMTPD1 = UBCustRec(1).CurrRevAmts(WhatRev)
          AMTPAID# = Round#(AMTPAID# - ThisAmt#)
        Else
          ThisAmt# = AMTPAID#
          PayRec(1).PaidOwed(WhatRev).AMTPD1 = ThisAmt#
          AMTPAID# = 0
        End If
      ElseIf AMTPAID# = 0 Then
        PayRec(1).PaidOwed(WhatRev).AMTPD1 = 0
      ElseIf ThisAmt# = 0 Then
        PayRec(1).PaidOwed(WhatRev).AMTPD1 = 0
      End If
    End If
  Next
  If AMTPAID# <> 0 Then  'if they over payed, dump it in the first rev
    PayRec(1).PaidOwed(1).AMTPD1 = (PayRec(1).PaidOwed(1).AMTPD1 + AMTPAID#)
  End If


'  For ZZCnt = 1 To 15
'    If AMTPAID# >= UBCustRec(1).CurrRevAmts(ZZCnt) Then
'      PayRec(1).PaidOwed(ZZCnt).AMTOWE1 = UBCustRec(1).CurrRevAmts(ZZCnt)
'      AMTPAID# = Round#(AMTPAID# - UBCustRec(1).CurrRevAmts(ZZCnt))
'      PayRec(1).PaidOwed(ZZCnt).AMTPD1 = PayRec(1).PaidOwed(ZZCnt).AMTOWE1
'      If AMTPAID# <= 0 Then Exit For
'    Else
'      PayRec(1).PaidOwed(ZZCnt).AMTOWE1 = AMTPAID#
'      PayRec(1).PaidOwed(ZZCnt).AMTPD1 = PayRec(1).PaidOwed(ZZCnt).AMTOWE1
'      AMTPAID# = 0
'      Exit For
'    End If
'  Next
'
'  If AMTPAID# <> 0 Then  'if they over payed, dump it in the first rev
'    PayRec(1).PaidOwed(1).AMTPD1 = (PayRec(1).PaidOwed(1).AMTPD1 + AMTPAID#)
'  End If

  PayRec(1).TOTOWED = PayRec(1).AMTOWED
  PayRec(1).AMTPAID = TAmtPaid#

  Put #Pay98File, , PayRec(1)

Return

Setup98File:
  ReDim PayRec(1) As UBPaymentRecType
  PayRecLen = Len(PayRec(1))

  Pay98File = FreeFile
  Open PayFileName$ For Output As #Pay98File
  Close Pay98File

  Pay98File = FreeFile
  Open PayFileName$ For Random Shared As Pay98File Len = PayRecLen

Return
ExitBoxPaymentsM:
  frmLBWarning.Label1.Caption = Msg1$
  frmLBWarning.Label3.Caption = Msg2$
  frmLBWarning.Label4.Caption = Msg3$
  frmLBWarning.Show 1

ExitBoxPayments:
  Close
End Sub
Private Sub LoadRevs()
  Dim UBSetupLen As Integer, RevCnt As Integer
  Dim InvRev As Integer, OutOfOrder As Boolean, x As Integer
  Dim tmp As DistArrayType
  NumofRevs = MaxRevsCnt
  ReDim UBSetUpRec(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  ReDim Preserve DistArray(1 To NumofRevs) As DistArrayType

  For RevCnt = 1 To MaxRevsCnt
    DistArray(RevCnt).DistOrder = UBSetUpRec(1).Revenues(RevCnt).DistOr
    DistArray(RevCnt).DistCnt = RevCnt
    If Len(QPTrim$(UBSetUpRec(1).Revenues(RevCnt).RevName)) = 0 Then
      NumofRevs = RevCnt - 1
      Exit For
    End If
  Next

  Do
    OutOfOrder = False          'assume it's sorted
    For x = 1 To NumofRevs - 1
      If DistArray(x).DistOrder > DistArray(x + 1).DistOrder Then
        'SWAP DistArray(x), DistArray(x + 1)     'if we had to swap
        tmp = DistArray(x)
        DistArray(x) = DistArray(x + 1)
        DistArray(x + 1) = tmp
        OutOfOrder = True       'we're not done yet
      End If
    Next
  Loop While OutOfOrder


End Sub
'Private Sub Autodist()
'  Dim cnt As Integer, ThisAmt As Double, UBTransRecLen As Integer
'  Dim NumofRevs As Integer, WhatRev As Integer, UBTran As Integer
'  Dim CustFile As Integer, UBCustRecLen As Integer, ThisTran As Long
'  Dim DZCnt As Integer
'  ReDim UBCustRec(1) As NewUBCustRecType
'
'  NumofRevs = MaxRevsCnt
'  For cnt = 1 To NumofRevs
'    WhatRev = DistArray(cnt).DistCnt - 1
'    If WhatRev >= 0 Then
'    ThisAmt# = Val(fpAmtOwed(WhatRev))
'    If ThisAmt# < 0 Then
'      TempAmtRecv# = Round#(TempAmtRecv# - ThisAmt#)
'    End If
'    End If
'  Next
'
'  For cnt = 1 To NumofRevs
'    WhatRev = DistArray(cnt).DistCnt - 1
'    If WhatRev >= 0 Then
'      ThisAmt# = fpAmtOwed(WhatRev)
'      If ThisAmt# <> 0 Then
'        If TempAmtRecv# >= ThisAmt# Then
'          fpAmtPaid(WhatRev) = fpAmtOwed(WhatRev)
'          TempAmtRecv# = Round#(TempAmtRecv# - ThisAmt#)
'        Else
'          ThisAmt# = TempAmtRecv#
'          fpAmtPaid(WhatRev) = ThisAmt#
'          TempAmtRecv# = 0
'        End If
'      ElseIf TempAmtRecv# = 0 Then
'        fpAmtPaid(WhatRev) = 0
'      ElseIf ThisAmt# = 0 Then
'        fpAmtPaid(WhatRev) = 0
'      End If
'    End If
'  Next
' End Sub

Private Function Check98File()
  Dim FntSize As Integer, zz As Integer, Ext As String
  Dim NewName As String
  
  If FileSize("UBPAY98.DAT") > 0 Then
    UBLog "LB: UBPAY98.DAT ALLREADY EXISTS!"
    'BlockClear
    'Ok = MsgBox%("UBSETUP", "KILL98FL")
    ReDim MsgText(0 To 5) As String
    FntSize = frmMsgDialog.Label(1).FontSize
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    MsgText(0) = "WARNING:"
    MsgText(1) = ""
    MsgText(2) = "UBPAY98.DAT ALLREADY EXISTS!"
    MsgText(3) = "If you continue this file"
    MsgText(4) = "will be renamed."
    MsgText(5) = "OK to continue, or Esc to Cancel."
    If GetOKorNot(MsgText()) Then
      UBLog "USER WANTS TO CONTINUE!"
      Check98File = True
      GoSub RenameOld98File
    Else
      Check98File = False
    End If

  Else
    Check98File = True
  End If

Exit Function

RenameOld98File:
  For zz = 1 To 999
    Ext$ = "000" + QPTrim$(Str$(zz))
    Ext$ = Right$(Ext$, 3)
    NewName$ = "UBPAY98." + Ext$
    If Not Exist(NewName$) Then
      Kill NewName$
      UBLog "LB: RENAMED UBPAY98.DAT TO " + "UBPAY98." + Ext$
      Name "UBPAY98.DAT" As NewName$
      UBLog "LB: PAYMENT FILE RENAMED SUCCESSFULLY"
      Exit For
    End If
  Next
Return
End Function
Private Function GetBoxAmount#(Number$)
  Dim BoxAmt As Double
  Number$ = QPTrim$(Number$)
  BoxAmt# = Val(Number$)
  If BoxAmt# > 0 Then
    BoxAmt# = Round#(BoxAmt# / 100)
  End If
  GetBoxAmount# = BoxAmt#

End Function
Private Function GetBoxCustRec&(Number$)
  GetBoxCustRec& = Val(Number$)
End Function
Private Sub ChkTranDate()
  Dim PayBillName As String, PayDepoName As String, Today As String
  Dim UBPayRecLen As Integer, UBTransRecLen As Integer
  Dim CHandle As Integer, thandle As Integer, chkthedate As Integer
  Dim PHandle As Integer, NumPayRecs As Long, cnt As Long
  InvalidDate = False
  PayBillName$ = UBPath$ + "UBPAY" + QPTrim$(Str$(OPERNUM)) + ".DAT"
  PayDepoName$ = UBPath$ + "UBDEP" + QPTrim$(Str$(OPERNUM)) + ".DAT"
  UBLog "Check Payment Date BP"
  FrmShowPctComp.Label1 = "Checking Transaction Dates"
  FrmShowPctComp.Show
  Today = Format(Now, "mm/dd/yyyy")
  Dim UBPaymentRec As UBPaymentRecType
  chkthedate = Date2Num(Today)
  UBPayRecLen = Len(UBPaymentRec)
  If FileSize&(PayDepoName$) > 0 Then
    PHandle = FreeFile
    Open PayDepoName$ For Random Shared As PHandle Len = UBPayRecLen
    NumPayRecs& = LOF(PHandle) \ UBPayRecLen
    For cnt& = 1 To NumPayRecs&
      FrmShowPctComp.ShowPctComp cnt&, NumPayRecs&
      If FrmShowPctComp.Out = True Then
        Close
        FrmShowPctComp.Out = False
        Close
        Exit Sub
      End If
      Get PHandle, cnt&, UBPaymentRec ',  UBPayRecLen
      If UBPaymentRec.PAYDATE > (chkthedate + 30) Or UBPaymentRec.PAYDATE < (chkthedate - 30) Then
        InvalidDate = True
        Unload FrmShowPctComp
        Close
        Exit Sub
      End If
    Next
    Close
  End If
  FrmShowPctComp.Label1 = "Checking Transaction Dates"
  FrmShowPctComp.Show
  If FileSize&(PayBillName$) > 0 Then
    PHandle = FreeFile
    Open PayBillName$ For Random Shared As PHandle Len = UBPayRecLen
    NumPayRecs& = LOF(PHandle) \ UBPayRecLen
    For cnt& = 1 To NumPayRecs&
      FrmShowPctComp.ShowPctComp cnt&, NumPayRecs&
      If FrmShowPctComp.Out = True Then
        Close
        FrmShowPctComp.Out = False
        Close
        Exit Sub
      End If
      Get PHandle, cnt&, UBPaymentRec ',  UBPayRecLen
      If UBPaymentRec.PAYDATE > (chkthedate + 30) Or UBPaymentRec.PAYDATE < (chkthedate - 30) Then
        InvalidDate = True
        Unload FrmShowPctComp
        Close
        Exit Sub
      End If
    Next
    Close
  End If
End Sub

