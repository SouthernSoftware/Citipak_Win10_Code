VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUBReportsMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Utility Billing Reports"
   ClientHeight    =   8868
   ClientLeft      =   48
   ClientTop       =   324
   ClientWidth     =   12216
   Icon            =   "frmUBReports.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCustStreetList 
      BackColor       =   &H008F8265&
      Caption         =   "Customer &Street Listing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3414
      MaskColor       =   &H8000000F&
      TabIndex        =   8
      Top             =   4812
      Width           =   2556
   End
   Begin VB.CommandButton cmdMembFeeRpt 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Membership &Fees Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6246
      TabIndex        =   7
      Top             =   4248
      Width           =   2556
   End
   Begin VB.CommandButton cmdCustInquiry 
      Caption         =   "Customer &Inquiry"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3414
      TabIndex        =   0
      Top             =   2568
      Width           =   2556
   End
   Begin VB.CommandButton cmdMastCustList 
      Caption         =   "M&aster Customer Listing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6240
      TabIndex        =   1
      Top             =   2568
      Width           =   2556
   End
   Begin VB.CommandButton cmdMastBalList 
      Caption         =   "Master &Balance Listing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6246
      TabIndex        =   3
      Top             =   3132
      Width           =   2556
   End
   Begin VB.CommandButton cmdCutoffList 
      Caption         =   "Customer Cut-&Off List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3414
      TabIndex        =   6
      Top             =   4248
      Width           =   2556
   End
   Begin VB.CommandButton cmdMastDepList 
      Caption         =   "Master &Deposit Listing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6264
      TabIndex        =   5
      Top             =   3684
      Width           =   2556
   End
   Begin VB.CommandButton cmdCustFlatRates 
      Caption         =   "Flat &Rate Customers"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6246
      TabIndex        =   9
      Top             =   4812
      Width           =   2556
   End
   Begin VB.CommandButton cmdDetTransJournal 
      Caption         =   "Transaction &Journal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3414
      TabIndex        =   2
      Top             =   3132
      Width           =   2556
   End
   Begin VB.CommandButton cmdPaySummary 
      Caption         =   "&Payment Summary"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3414
      TabIndex        =   4
      Top             =   3684
      Width           =   2556
   End
   Begin VB.CommandButton cmdMailingLabels 
      Caption         =   "Mailing &Labels"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6246
      TabIndex        =   11
      Top             =   5364
      Width           =   2556
   End
   Begin VB.CommandButton cmdCycleCntSum 
      Caption         =   "&Cycle Count Summary"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3414
      TabIndex        =   12
      Top             =   5928
      Width           =   2556
   End
   Begin VB.CommandButton cmdBillPayTax 
      Caption         =   "Bill/Payment &Tax Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3432
      TabIndex        =   10
      Top             =   5364
      Width           =   2556
   End
   Begin VB.CommandButton cmdMeterInstall 
      Caption         =   "&Meter Installed Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6246
      TabIndex        =   13
      Top             =   5928
      Width           =   2556
   End
   Begin VB.CommandButton cmdExitUBRep 
      Caption         =   "E&XIT to Previous Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4800
      TabIndex        =   14
      Top             =   6480
      Width           =   2652
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   15
      Top             =   8508
      Width           =   12216
      _ExtentX        =   21548
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7154
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "4:35 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "6/3/2003"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      X1              =   2388
      X2              =   3348
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
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   1092
      Left            =   1788
      Top             =   744
      Width           =   8652
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "UTILITY BILLING REPORTS MENU"
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
      Left            =   3348
      TabIndex        =   16
      Top             =   1176
      Width           =   5292
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
      Index           =   0
      Left            =   2508
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
End
Attribute VB_Name = "frmUBReportsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmdBillPayTax_Click()
  frmRptBillPayTax.Show
  Unload frmUBReportsMenu
End Sub

Private Sub cmdCustFlatRates_Click()
  frmRptFlatRate.Show
  Unload frmUBReportsMenu
End Sub

Private Sub cmdCustInquiry_Click()
  frmCustEditLookUP.Caption = "Customer Inquiry Search"
  frmCustEditLookUP.Label1.Caption = "Customer Inquiry Search"
  frmCustEditLookUP.Wheretogo frmUBReportsMenu, frmRptCustInq
  'Load frmCustEditLookUP
  DoEvents
  frmCustEditLookUP.Show
  Unload frmUBReportsMenu

End Sub

Private Sub cmdCustStreetList_Click()
  frmRptStreetList.Show
  Unload frmUBReportsMenu
End Sub

Private Sub cmdCutoffList_Click()
  frmRptCutOff.Show
  Unload frmUBReportsMenu
End Sub

Private Sub cmdCycleCntSum_Click()
  frmRptCycleSum.Show
  Unload frmUBReportsMenu
End Sub

Private Sub cmdDetTransJournal_Click()
  frmRptTransJournal.Show
  Unload frmUBReportsMenu
End Sub

Private Sub cmdExitUBRep_Click()
  frmUBMainMenu.Show
  Unload frmUBReportsMenu
End Sub

Private Sub cmdMailingLabels_Click()
  frmRptMailLabels.Show
  Unload frmUBReportsMenu
End Sub

Private Sub cmdMastBalList_Click()
  frmRptMastBal.Show
  Unload frmUBReportsMenu
End Sub

Private Sub cmdMastCustList_Click()
  frmRptMastCust.Show
  Unload frmUBReportsMenu
End Sub

Private Sub cmdMastCustList_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyLeft Then
    cmdCustInquiry.SetFocus
    KeyCode = 0
  End If
End Sub

Private Sub cmdMastDepList_Click()
  frmRptMastDep.Show
  Unload frmUBReportsMenu
End Sub

Private Sub cmdMembFeeRpt_Click()
  frmRptMembFee.Show
  Unload frmUBReportsMenu
End Sub

Private Sub cmdMeterInstall_Click()
  frmRptMtrInstDate.Show
  Unload frmUBReportsMenu
End Sub

Private Sub cmdPaySummary_Click()
  frmRptPaymSum.Show
  Unload frmUBReportsMenu
End Sub

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
      cmdExitUBRep_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

