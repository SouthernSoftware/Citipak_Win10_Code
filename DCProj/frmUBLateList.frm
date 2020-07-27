VERSION 5.00
Begin VB.Form frmUBSta 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Utility Billing Statistical Reports"
   ClientHeight    =   8868
   ClientLeft      =   48
   ClientTop       =   324
   ClientWidth     =   12216
   Icon            =   "frmUBLateList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExportFiles 
      Caption         =   "&Export Files"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4320
      TabIndex        =   9
      Top             =   7080
      Width           =   3612
   End
   Begin VB.CommandButton cmdQueryGLTrans 
      Caption         =   "&Query G/L Transactions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4320
      TabIndex        =   8
      Top             =   6600
      Width           =   3612
   End
   Begin VB.CommandButton cmdExitUBStatReportsMenu 
      Caption         =   "E&xit Statistical Reports Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4320
      TabIndex        =   10
      Top             =   7560
      Width           =   3612
   End
   Begin VB.CommandButton cmdBudvsAct 
      Caption         =   "Bud&get vs Actual"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4320
      TabIndex        =   6
      Top             =   5640
      Width           =   3612
   End
   Begin VB.CommandButton cmdDeptBudvsAct 
      Caption         =   "&Department Budget vs Actual"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4320
      TabIndex        =   7
      Top             =   6120
      Width           =   3612
   End
   Begin VB.CommandButton cmdAcctBalSummary 
      Caption         =   "&Account Balance Summary"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4320
      TabIndex        =   2
      Top             =   3720
      Width           =   3612
   End
   Begin VB.CommandButton cmdBalSheet 
      Caption         =   "Balance &Sheet"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4320
      TabIndex        =   5
      Top             =   5160
      Width           =   3612
   End
   Begin VB.CommandButton cmdBudgHistory 
      Caption         =   "&Budget History"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4320
      TabIndex        =   4
      Top             =   4680
      Width           =   3612
   End
   Begin VB.CommandButton cmdAcctHistory 
      Caption         =   "Account &History"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4320
      TabIndex        =   3
      Top             =   4200
      Width           =   3612
   End
   Begin VB.CommandButton cmdCashBalance 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Cash Balance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4320
      TabIndex        =   1
      Top             =   3240
      Width           =   3612
   End
   Begin VB.CommandButton cmdTrialBalance 
      BackColor       =   &H008F8265&
      Caption         =   "&Trial Balance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4320
      MaskColor       =   &H8000000F&
      TabIndex        =   0
      Top             =   2760
      Width           =   3612
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   1092
      Left            =   1800
      Top             =   1080
      Width           =   8652
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2400
      X2              =   2400
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   3360
      X2              =   3360
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   2400
      X2              =   3360
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2400
      X2              =   3360
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2400
      Top             =   2160
      Width           =   972
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2520
      X2              =   2520
      Y1              =   2400
      Y2              =   8280
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   9840
      X2              =   9840
      Y1              =   2304
      Y2              =   2400
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   3
      X1              =   8880
      X2              =   9840
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   8880
      X2              =   8880
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   2
      X1              =   8880
      X2              =   9840
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   1
      Left            =   8880
      Top             =   2160
      Width           =   972
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2520
      X2              =   3216
      Y1              =   8304
      Y2              =   8304
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   1
      X1              =   9000
      X2              =   9000
      Y1              =   2400
      Y2              =   8280
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   1
      X1              =   9000
      X2              =   9720
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Statistical Reports Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3720
      TabIndex        =   11
      Top             =   1440
      Width           =   4692
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   9000
      Top             =   2400
      Width           =   732
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1800
      Top             =   960
      Width           =   8652
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5916
      Index           =   0
      Left            =   2496
      Top             =   2400
      Width           =   732
   End
End
Attribute VB_Name = "frmUBSta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TownName$
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExitUBRep.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        'ClearInUse PWcnt
      End If
    End If
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      cmdExitUBStatRep_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub
Private Sub cmdExitUBStatRep_Click()
  frmUBMainMenu.Show
  Unload frmUBStatReportsMenu
End Sub
'Private Sub cmdBillPayTax_Click()
'  frmRptBillPayTax.Show
'  Unload frmUBReportsMenu
'End Sub
'
'Private Sub cmdCustFlatRates_Click()
'  frmRptFlatRate.Show
'  Unload frmUBReportsMenu
'End Sub
'
'Private Sub cmdCustInquiry_Click()
'  frmCustEditLookUP.Caption = "Customer Inquiry Search"
'  frmCustEditLookUP.Label1.Caption = "Customer Inquiry Search"
'  frmCustEditLookUP.Wheretogo frmUBReportsMenu, frmRptCustInq
'  'Load frmCustEditLookUP
'  DoEvents
'  frmCustEditLookUP.Show
'  Unload frmUBReportsMenu
'
'End Sub
'
'Private Sub cmdCustStreetList_Click()
'  frmRptStreetList.Show
'  Unload frmUBReportsMenu
'End Sub
'
'Private Sub cmdCutoffList_Click()
'  frmRptCutOff.Show
'  Unload frmUBReportsMenu
'End Sub
'
'Private Sub cmdCycleCntSum_Click()
'  frmRptCycleSum.Show
'  Unload frmUBReportsMenu
'End Sub
'
'Private Sub cmdDetTransJournal_Click()
'  frmRptTransJournal.Show
'  Unload frmUBReportsMenu
'End Sub
'
'
'Private Sub cmdMailingLabels_Click()
'  frmRptMailLabels.Show
'  Unload frmUBReportsMenu
'End Sub
'
'Private Sub cmdMastBalList_Click()
'  frmRptMastBal.Show
'  Unload frmUBReportsMenu
'End Sub
'
'Private Sub cmdMastCustList_Click()
'  frmRptMastCust.Show
'  Unload frmUBReportsMenu
'End Sub
'
'Private Sub cmdMastCustList_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeyLeft Then
'    cmdCustInquiry.SetFocus
'    KeyCode = 0
'  End If
'End Sub
'
'Private Sub cmdMastDepList_Click()
'  frmRptMastDep.Show
'  Unload frmUBReportsMenu
'End Sub
'
'Private Sub cmdMembFeeRpt_Click()
'  frmRptMembFee.Show
'  Unload frmUBReportsMenu
'End Sub
'
'Private Sub cmdMeterInstall_Click()
'  frmRptMtrInstDate.Show
'  Unload frmUBReportsMenu
'End Sub
'
'Private Sub cmdPaySummary_Click()
'  frmRptPaymSum.Show
'  Unload frmUBReportsMenu
'End Sub
'

'
'
