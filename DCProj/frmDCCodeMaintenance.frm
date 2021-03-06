VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDCCodeMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Decal Code Maintenance "
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   1884
   ClientWidth     =   12216
   ClipControls    =   0   'False
   Icon            =   "frmDCCodeMaintenance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   888
      Top             =   2784
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   888
      Top             =   3216
   End
   Begin VB.CommandButton cmdDCSetupMenu 
      BackColor       =   &H008F8265&
      Caption         =   "System &Maintenance Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3846
      MaskColor       =   &H8000000F&
      TabIndex        =   4
      Top             =   5376
      Width           =   4524
   End
   Begin VB.CommandButton cmdListDecals 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Print Decal Listing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3846
      TabIndex        =   3
      Top             =   4008
      Width           =   4524
   End
   Begin VB.CommandButton cmdAddCategory 
      Caption         =   "&Add New Decal Catagory"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3846
      TabIndex        =   0
      Top             =   2640
      Width           =   4524
   End
   Begin VB.CommandButton cmdEditCategory 
      Caption         =   "&Edit Existing Decal Category"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3846
      TabIndex        =   1
      Top             =   3324
      Width           =   4524
   End
   Begin VB.CommandButton cmdPurchaseDecals 
      Caption         =   "&Purchase Decals"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3846
      TabIndex        =   2
      Top             =   4692
      Width           =   4524
   End
   Begin VB.CommandButton cmdExitDC 
      Caption         =   "E&XIT Vehicle Decals"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3846
      TabIndex        =   5
      Top             =   6060
      Width           =   4524
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
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
            TextSave        =   "10:37 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "6/6/2005"
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
      Caption         =   "DECAL CODE MAINTENANCE"
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
      TabIndex        =   6
      Top             =   1176
      Width           =   5292
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
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   1
      Left            =   8868
      Top             =   1824
      Width           =   972
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
      X1              =   2388
      X2              =   3348
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
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2388
      Top             =   1824
      Width           =   972
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
End
Attribute VB_Name = "frmDCCodeMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
'Dim FormOver As clsFormOverRider
Private Temp_Class As Resize_Class

'LevelPass 1 is Full Access, 2 is Payments, 3 is Reports Only
Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  If DelayExit = True Then
    DelayExit = False
    Timer2.Enabled = True
  Else
    cmdExitDC.Enabled = True
  End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    If cmdExitDC.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        Call DCLog("Close via Main DC Menu" + PWUser$)
        CitiTerminate
      End If
    End If
  End If
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
    Temp_Class.ResizeControls Me
    'Me.Visible = True
    'Me.SetFocus
    DoEvents
  End If
End Sub

Private Sub cmdPurchaseDecals_Click()
'  Dim FntSize As Integer
'  ReDim MsgText(0 To 5) As String
'  If LevelPass < 3 Then
'    If Not Exist(UBPath$ + "UBCust.dat") Then
'      frmMsgDialog.RetLabel = "-2"
'      UBLog "ERROR: NO Cust Info"
'      FntSize = frmMsgDialog.Label(3).FontSize
'      frmMsgDialog.Label(1).FontSize = (FntSize + 2)
'      frmMsgDialog.Label(3).FontSize = (FntSize + 2)
'      MsgText(0) = "ERROR:"
'      MsgText(1) = ""
'      MsgText(2) = "NO CUSTOMER INFORMATION!"
'      MsgText(3) = ""
'      MsgText(4) = ""
'      MsgText(5) = ""
'      GetOKorNot MsgText(), True
'      Exit Sub
'    End If
'
'    Load frmPaymentDate
'    DoEvents
'    frmPaymentDate.Show
'    Unload frmUBMainMenu
'  Else
'    MsgBox "Your password does not allow access to this option.", vbOKOnly, "Access Denied"
'  End If
End Sub


Private Sub cmdCustomerMenu_Click()
'  If LevelPass = 1 Then
'   Load frmUBCustMenu
'   DoEvents
'   frmUBCustMenu.Show
'   Unload frmUBMainMenu
'  Else
'    MsgBox "Your password does not allow access to this option.", vbOKOnly, "Access Denied"
'  End If
End Sub
Private Sub cmdReportsMenu_Click()
'  Load frmUBReportsMenu
'  DoEvents
'  frmUBReportsMenu.Show
'  Unload frmUBMainMenu
End Sub

Private Sub cmdExitDC_Click()
  DCTerminate
  frmDCMainMenu.Enabled = False
  Timer1.Enabled = True
End Sub

'Private Sub cmdStaticReportsMenu_Click()
''  Load frmUBStatReportsMenu
''  DoEvents
''  frmUBStatReportsMenu.Show
''  Unload frmUBMainMenu
'End Sub

Private Sub cmdDCSetupMenu_Click()
'  If LevelPass = 1 Then
    Load frmDCSetupMenu
    DoEvents
    frmDCSetupMenu.Show
    Unload Me
'  Else
'    MsgBox "Your password does not allow access to this option.", vbOKOnly, "Access Denied"
'  End If
End Sub

'  Dim FntSize As Integer
'  ReDim MsgText(0 To 5) As String
'  If LevelPass = 1 Then
'    If Not Exist(UBPath$ + "UBCust.dat") Then
'      frmMsgDialog.RetLabel = "-2"
'      UBLog "ERROR: NO Cust Info"
'      FntSize = frmMsgDialog.Label(3).FontSize
'      frmMsgDialog.Label(1).FontSize = (FntSize + 2)
'      frmMsgDialog.Label(3).FontSize = (FntSize + 2)
'      MsgText(0) = "ERROR:"
'      MsgText(1) = ""
'      MsgText(2) = "NO CUSTOMER INFORMATION!"
'      MsgText(3) = ""
'      MsgText(4) = ""
'      MsgText(5) = ""
'      GetOKorNot MsgText(), True
'      Exit Sub
'    End If
'    Load frmUBWorkOrderMenu
'    DoEvents
'    frmUBWorkOrderMenu.Show
'    Unload frmUBMainMenu
'  Else
'    MsgBox "Your password does not allow access to this option.", vbOKOnly, "Access Denied"
'  End If

'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'  Select Case KeyCode
'    Case vbKeyEscape:
'      SendKeys "%X"
'      KeyCode = 0
'    Case vbKeyHome
'      cmdCustomerMenu.SetFocus
'    Case vbKeyEnd
'      cmdExitUB.SetFocus
'    Case Else:
'  End Select
'End Sub
Private Sub Timer1_Timer()
  Unload Me
End Sub

Private Sub Timer2_Timer()
  cmdExitDC.Enabled = True
End Sub

