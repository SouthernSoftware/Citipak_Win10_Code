VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUBDepositMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Deposit Menu"
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   1884
   ClientWidth     =   12216
   Icon            =   "frmUBDepositMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdDepVoid 
      Caption         =   "Deposit Payment &Void"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   564
      Left            =   3876
      TabIndex        =   3
      Top             =   5322
      Width           =   4524
   End
   Begin VB.CommandButton cmdDepCreditRem 
      Caption         =   "&Deposit Credit Removal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   564
      Left            =   3876
      TabIndex        =   2
      Top             =   4476
      Width           =   4524
   End
   Begin VB.CommandButton cmdApplyDep 
      BackColor       =   &H008F8265&
      Caption         =   "&Apply Deposit to Balance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   564
      Left            =   3876
      MaskColor       =   &H8000000F&
      TabIndex        =   0
      Top             =   2784
      Width           =   4524
   End
   Begin VB.CommandButton cmdRefundDep 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Refund Customer Deposit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   564
      Left            =   3876
      TabIndex        =   1
      Top             =   3630
      Width           =   4524
   End
   Begin VB.CommandButton cmdExitMenu 
      Caption         =   "E&xit to Previous Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   564
      Left            =   3876
      TabIndex        =   4
      Top             =   6168
      Width           =   4524
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
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
            TextSave        =   "12:30 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "6/23/2005"
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Deposit Menu"
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
      Left            =   3540
      TabIndex        =   6
      Top             =   1104
      Width           =   5148
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   9828
      X2              =   9828
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   8868
      X2              =   8868
      Y1              =   1944
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
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   2
      X1              =   8868
      X2              =   9828
      Y1              =   2064
      Y2              =   2064
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
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   1
      X1              =   8988
      X2              =   8988
      Y1              =   2064
      Y2              =   7944
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   3348
      X2              =   3348
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2388
      X2              =   2388
      Y1              =   1944
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
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2388
      X2              =   3348
      Y1              =   2064
      Y2              =   2064
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
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2508
      X2              =   2508
      Y1              =   2064
      Y2              =   7944
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   1092
      Left            =   1788
      Top             =   744
      Width           =   8652
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
Attribute VB_Name = "frmUBDepositMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
'Dim FormOver As clsFormOverRider
Private Temp_Class As Resize_Class

Private Sub cmdApplyDep_Click()
  frmCustEditLookUP.Caption = "Apply Customer Deposit Search"
  frmCustEditLookUP.Label1.Caption = "Apply Customer Deposit Search"
  frmCustEditLookUP.Wheretogo frmUBDepositMenu, frmDepositApply, , 3
  DoEvents
  frmCustEditLookUP.Show
  Unload Me

End Sub

Private Sub cmdDepCreditRem_Click()
  frmCustEditLookUP.Caption = "Customer Credit Removal Search"
  frmCustEditLookUP.Label1.Caption = "Customer Credit Removal Search"
  frmCustEditLookUP.Wheretogo frmUBDepositMenu, frmDepCreditRemoval, , 4
  DoEvents
  frmCustEditLookUP.Show
  'NO NO Stay on screen
  'Unload Me
End Sub

Private Sub cmdDepVoid_Click()
  frmCustEditLookUP.Caption = "Customer Deposit Reversal Search"
  frmCustEditLookUP.Label1.Caption = "Customer Deposit Reversal Search"
  frmCustEditLookUP.Wheretogo frmUBDepositMenu, frmDepositVoid, , 6
  DoEvents
  frmCustEditLookUP.Show
End Sub

Private Sub cmdExitMenu_Click()
  Load frmUBBillingMenu
  DoEvents
  frmUBBillingMenu.Show
  Unload Me
End Sub

Private Sub cmdRefundDep_Click()
  frmCustEditLookUP.Caption = "Customer Deposit Refund Search"
  frmCustEditLookUP.Label1.Caption = "Customer Deposit Refund Search"
  frmCustEditLookUP.Wheretogo frmUBDepositMenu, frmDepositRefund, , 5
  DoEvents
  frmCustEditLookUP.Show
End Sub

'Private Sub cmdReprint_Click()
'  Dim FntSize As Integer
'  ReDim MsgText(0 To 5) As String
'  If Not Exist("UBFBILLS.PRN") Then
'    frmMsgDialog.RetLabel = "-2"
'    UBLog "ERROR: NO PRN FILE. Reprint Final"
'    FntSize = frmMsgDialog.Label(3).FontSize
'    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
'    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
'    MsgText(0) = "ERROR:"
'    MsgText(1) = ""
'    MsgText(2) = "NO BILL PRINT FILE!"
'    MsgText(3) = ""
'    MsgText(4) = ""
'    MsgText(5) = ""
'    GetOKorNot MsgText(), True
'    Exit Sub
'  End If
'
'  If Not Exist(UBFinBillsFile) Then
'    frmMsgDialog.RetLabel = "-2"
'    UBLog "ERROR: NO BILL FILE! Reprint Final"
'    FntSize = frmMsgDialog.Label(3).FontSize
'    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
'    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
'    MsgText(0) = "ERROR:"
'    MsgText(1) = ""
'    MsgText(2) = "NO BILL FILE!"
'    MsgText(3) = ""
'    MsgText(4) = ""
'    MsgText(5) = ""
'    GetOKorNot MsgText(), True
'    Exit Sub
'  End If
'  frmBillPrinting.REPRN True, True
'  Load frmBillPrinting
'  DoEvents
'  frmBillPrinting.Show
'  Unload frmUBFinalBillPrintMenu
'
'End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  Me.HelpContextID = hlpCustomerDeposit
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExitMenu.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via UBDepositMenu by " + PWUser$
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

'Private Sub cmdPrnAllUBBills_Click()
'  Dim FntSize As Integer
'  ReDim MsgText(0 To 5) As String
'
'  If Not Exist(UBFinBillsFile) Then
'    frmMsgDialog.RetLabel = "-2"
'    UBLog "ERROR: NO BILL FILE! Final"
'    FntSize = frmMsgDialog.Label(3).FontSize
'    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
'    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
'    MsgText(0) = "ERROR:"
'    MsgText(1) = ""
'    MsgText(2) = "NO BILL FILE!"
'    MsgText(3) = ""
'    MsgText(4) = ""
'    MsgText(5) = ""
'    GetOKorNot MsgText(), True
'    Exit Sub
'  End If
'  frmBillPrinting.REPRN False, True
'  Load frmBillPrinting
'  DoEvents
'  frmBillPrinting.Show
'  Unload frmUBFinalBillPrintMenu
'End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyHome
      cmdApplyDep.SetFocus
    Case vbKeyEnd
      cmdExitMenu.SetFocus
    Case Else:
  End Select
End Sub

