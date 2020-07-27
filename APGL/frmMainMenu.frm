VERSION 5.00
Begin VB.Form frmMainMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "v2.01 CitiPak Main Menu"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   12225
   Icon            =   "frmMainMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPasswords 
      Caption         =   "Password &Maintenance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4308
      TabIndex        =   9
      Top             =   7152
      Width           =   3612
   End
   Begin VB.CommandButton cmdCashManagement 
      Caption         =   "&Cash Management"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4308
      TabIndex        =   6
      Top             =   5556
      Width           =   3612
   End
   Begin VB.CommandButton cmdPropertyTaxes 
      Caption         =   "Property &Taxes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4308
      TabIndex        =   5
      Top             =   5028
      Width           =   3612
   End
   Begin VB.CommandButton cmdFixedAssets 
      Caption         =   "&Fixed Assets"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4308
      TabIndex        =   4
      Top             =   4500
      Width           =   3612
   End
   Begin VB.CommandButton cmdExitCitipak 
      Caption         =   "E&xit Citipak "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4308
      TabIndex        =   10
      Top             =   7680
      Width           =   3612
   End
   Begin VB.CommandButton cmdVehicleDecals 
      Caption         =   "&Vehicle Decals"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4308
      TabIndex        =   8
      Top             =   6612
      Width           =   3612
   End
   Begin VB.CommandButton cmdPayroll 
      Caption         =   "&Payroll"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4308
      TabIndex        =   3
      Top             =   3972
      Width           =   3612
   End
   Begin VB.CommandButton cmdGenLedgMenu 
      Caption         =   "&General Ledger"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4308
      TabIndex        =   2
      Top             =   3432
      Width           =   3612
   End
   Begin VB.CommandButton cmdAcctPayMenu 
      Caption         =   "&Accounts Payable"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4308
      TabIndex        =   1
      Top             =   2904
      UseMaskColor    =   -1  'True
      Width           =   3612
   End
   Begin VB.CommandButton cmdBusinessLicense 
      Caption         =   "&Business License"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4308
      TabIndex        =   0
      Top             =   2376
      Width           =   3612
   End
   Begin VB.CommandButton cmdUtilityBilling 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Utility Billing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4308
      TabIndex        =   7
      Top             =   6084
      Width           =   3612
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   2520
      X2              =   2520
      Y1              =   2424
      Y2              =   8280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   2532
      X2              =   3240
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   2
      X1              =   9000
      X2              =   9000
      Y1              =   2400
      Y2              =   8256
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   4
      X1              =   8880
      X2              =   9840
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   5
      X1              =   8880
      X2              =   9840
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   2
      X1              =   8880
      X2              =   8880
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   2
      X1              =   9840
      X2              =   9840
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   1092
      Left            =   1800
      Top             =   1080
      Width           =   8652
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2400
      X2              =   3360
      Y1              =   2400
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
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   9000
      X2              =   9720
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CITIPAK MAIN MENU"
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
      Left            =   3942
      TabIndex        =   11
      Top             =   1440
      Width           =   4332
   End
   Begin VB.Shape Shape6 
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
      Height          =   5892
      Index           =   2
      Left            =   9000
      Top             =   2400
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   8880
      Top             =   2160
      Width           =   972
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
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2520
      Top             =   2400
      Width           =   732
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmdBusinessLicense_Click()
  If Exist("armenu.EXE") Then
    Shell "armenu.exe", vbMaximizedFocus
    DoEvents
    Unload frmMainMenu
  Else
    frmModuleInfo.Label4 = "Business License"
    frmModuleInfo.Show
  End If
End Sub

Private Sub cmdCashManagement_Click()
  If Exist("cmmenu.EXE") Then
    Shell "cmmenu.exe", vbMaximizedFocus
    DoEvents
    Unload frmMainMenu
  Else
    frmModuleInfo.Label4 = "Cash Management"
    frmModuleInfo.Show
  End If
End Sub

Private Sub cmdFixedAssets_Click()
  If Exist("FixedAssets.EXE") Then
     Shell "FixedAssets.EXE", vbMaximizedFocus
     Unload frmMainMenu
  ElseIf Exist("famenu.EXE") Then
    Shell "famenu.exe", vbMaximizedFocus
    DoEvents
    Unload frmMainMenu
  Else
    frmModuleInfo.Label4 = "Fixed Assets"
    frmModuleInfo.Show
  End If
End Sub

'Private Sub cmdInventoryControl_Click()
'  If Exist("icmenu.EXE") Then
'    Shell "icmenu.exe", vbMaximizedFocus
'    DoEvents
'    Unload frmMainMenu
'  Else
'    frmModuleInfo.Label4 = "Inventory Control"
'    frmModuleInfo.Show
'  End If
'End Sub

Private Sub cmdPasswords_Click()
''Lock the file so nobody can use during maintenance
'  Dim CitiPass As CitiPassType
'  Dim PassRecLen As Integer, NumPassRecs As Integer
'
' ' On Local Error GoTo PassError
'  If Exist("CitiPass.dat") Then
'    PassRecLen = Len(CitiPass)
'    CPAdminhand = FreeFile
'    Open "CitiPass.dat" For Random Lock Read Write As CPAdminhand Len = PassRecLen
'    NumPassRecs = LOF(CPAdminhand) \ PassRecLen
'  End If
'  MainLog "In Password Maintenance"
'  frmMainPassWord.Show 1
'  Exit Sub
'PassError:
'  CPAdminhand = -1
'  MsgBox "Password Maintenance Already Open.", vbOKOnly, "Access Denied"
  MainLog "In Password Maintenance"
  frmMainPassWord.Show 1
End Sub

Private Sub cmdPropertyTaxes_Click()
    If Exist("CitiTaxes.exe") Then
      Shell "CitiTaxes.exe", vbMaximizedFocus
      DoEvents
      Unload frmMainMenu
    ElseIf Exist("VACitiTax.EXE") Then
      Shell "VACitiTax.exe", vbMaximizedFocus
      DoEvents
      Unload frmMainMenu
    Else
      frmModuleInfo.Label4 = "Property Tax Management"
      frmModuleInfo.Show
    End If
  'End If
End Sub

Private Sub cmdUtilityBilling_Click()
  If Exist("Ubcopyrg.exe") Then
    Shell "Ubcopyrg.exe", vbMaximizedFocus
    DoEvents
    Unload frmMainMenu
  Else
    If Exist("Ubmenu.exe") Then
      Shell "Ubmenu.exe", vbMaximizedFocus
      DoEvents
      Unload frmMainMenu
    ElseIf Exist("wbmenu.EXE") Then
      Shell "wbmenu.exe", vbMaximizedFocus
      DoEvents
      Unload frmMainMenu
    Else
    frmModuleInfo.Label4 = "Utility Billing"
    frmModuleInfo.Show
    End If
  End If
  
End Sub

Private Sub cmdVehicleDecals_Click()
  If Exist("dcmenu.EXE") Then
    Shell "dcmenu.exe", vbMaximizedFocus
    DoEvents
    Unload frmMainMenu
  Else
    frmModuleInfo.Label4 = "Vehicle Decals"
    frmModuleInfo.Show
  End If
End Sub


Private Sub cmdAcctPayMenu_Click()
  If Exist("AP.EXE") Then
    If Exist("GLSETUP.DAT") And Exist("GLACCT.DAT") Then
      Shell "AP.exe", vbMaximizedFocus
      Unload frmMainMenu
    Else
      MsgBox "You Must First Enter GL Setup Information And Accounts.", vbOKOnly, "Missing Setup Info."
    End If
  Else
    frmModuleInfo.Label4 = "Accounts Payable"
    frmModuleInfo.Show
  End If
End Sub

Private Sub cmdExitCitipak_Click()
  Unload frmMainMenu
End Sub

Private Sub cmdGenLedgMenu_Click()
  If Exist("GL.EXE") Then
    Shell "GL.exe", vbMaximizedFocus
    Unload frmMainMenu
  Else
    frmModuleInfo.Label4 = "General Ledger"
    frmModuleInfo.Show
  End If
End Sub


Private Sub cmdPayroll_Click()
  If Exist("PAYROLL.EXE") Then
  '  If Not Exist("prrun.opn") Then
      Shell "Payroll.EXE", vbMaximizedFocus
      Unload frmMainMenu
'      Open "prrun.opn" For Output As #250
'      Print #250, ComputerName$
'      Close #250
'      frmMainPassWord.Caption = "Payroll Main Menu"
'      frmMainPassWord.Callingfrm = 2
'      frmMainPassWord.Show 1, Me
'    Else
'      MsgBox "Payroll Already Open - You May Not Continue.", vbOKOnly, "Access Denied"
'    End If
  Else
    frmModuleInfo.Label4 = "Payroll"
    frmModuleInfo.Show
  End If
End Sub

Private Sub Form_Load()
  Dim cnt&, dl&
  If App.PrevInstance Then
    ActivatePrevInstance
  End If
  
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  cnt& = 199
  ComputerName$ = String$(200, 0)
  dl& = GetUserName(ComputerName$, cnt)
  ComputerName$ = QTR$(ComputerName$)

End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      cmdExitCitipak_Click
      KeyCode = 0
      DoEvents
    Case Else:
  End Select
End Sub


