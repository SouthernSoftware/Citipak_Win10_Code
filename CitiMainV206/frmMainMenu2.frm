VERSION 5.00
Begin VB.Form frmMainMenu2 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "v2.04 CitiPak Main Menu"
   ClientHeight    =   8868
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   12216
   Icon            =   "frmMainMenu2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdReceiptPrinter 
      Caption         =   "&Receipt Printer Setup"
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
      Top             =   7236
      Width           =   3612
   End
   Begin VB.CommandButton cmdPasswords 
      Caption         =   "Password &Maintenance"
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
      Top             =   6750
      Width           =   3612
   End
   Begin VB.CommandButton cmdCashManagement 
      Caption         =   "&Cash Management"
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
      Top             =   5292
      Width           =   3612
   End
   Begin VB.CommandButton cmdPropertyTaxes 
      Caption         =   "Property &Taxes"
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
      Top             =   4806
      Width           =   3612
   End
   Begin VB.CommandButton cmdFixedAssets 
      Caption         =   "&Fixed Assets"
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
      Top             =   4320
      Width           =   3612
   End
   Begin VB.CommandButton cmdExitCitipak 
      Caption         =   "E&xit Citipak "
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
      TabIndex        =   11
      Top             =   7728
      Width           =   3612
   End
   Begin VB.CommandButton cmdVehicleDecals 
      Caption         =   "&Vehicle Decals"
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
      Top             =   6264
      Width           =   3612
   End
   Begin VB.CommandButton cmdPayroll 
      Caption         =   "&Payroll"
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
      Top             =   3834
      Width           =   3612
   End
   Begin VB.CommandButton cmdGenLedgMenu 
      Caption         =   "&General Ledger"
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
      Top             =   3348
      Width           =   3612
   End
   Begin VB.CommandButton cmdAcctPayMenu 
      Caption         =   "&Accounts Payable"
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
      Top             =   2862
      UseMaskColor    =   -1  'True
      Width           =   3612
   End
   Begin VB.CommandButton cmdBusinessLicense 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Business License"
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
      TabIndex        =   0
      Top             =   2376
      Width           =   3612
   End
   Begin VB.CommandButton cmdUtilityBilling 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Utility Billing"
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
      Top             =   5778
      Width           =   3612
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      Height          =   156
      Left            =   8880
      Top             =   2256
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      Height          =   156
      Left            =   2400
      Top             =   2256
      Width           =   972
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   2520
      X2              =   2520
      Y1              =   2400
      Y2              =   8280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   2532
      X2              =   3240
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   2
      X1              =   9000
      X2              =   9000
      Y1              =   2400
      Y2              =   8280
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   1092
      Left            =   1800
      Top             =   1080
      Width           =   8652
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
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
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3942
      TabIndex        =   12
      Top             =   1440
      Width           =   4332
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1800
      Top             =   960
      Width           =   8652
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   2
      Left            =   9000
      Top             =   2400
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   8880
      Top             =   2160
      Width           =   972
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2400
      Top             =   2160
      Width           =   972
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2520
      Top             =   2400
      Width           =   732
   End
End
Attribute VB_Name = "frmMainMenu2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmdBusinessLicense_Click()
  If Exist(UBPath$ + "BusinessLicense.EXE") Then
    Shell UBPath$ + "BusinessLicense.EXE", vbMaximizedFocus
    DoTheTime
    DoEvents
    Unload frmMainMenu
  ElseIf Exist(UBPath$ + "armenu.EXE") Then
    Shell UBPath$ + "armenu.exe", vbMaximizedFocus
    DoEvents
    Unload frmMainMenu
  Else
    frmModuleInfo.Label4 = "Business License"
    frmModuleInfo.Show
  End If
End Sub
Private Sub cmdCashManagement_Click()
  If Exist(UBPath$ + "CM.EXE") Then
    Shell UBPath$ + "CM.exe", vbMaximizedFocus
    DoTheTime
    Unload frmMainMenu
  ElseIf Exist(UBPath$ + "cmmenu.EXE") Then
    Shell UBPath$ + "cmmenu.exe", vbMaximizedFocus
    DoEvents
    Unload frmMainMenu
  Else
    frmModuleInfo.Label4 = "Cash Management"
    frmModuleInfo.Show
  End If
End Sub

Private Sub cmdFixedAssets_Click()
  If Exist(UBPath$ + "FixedAssets.EXE") Then
     Shell UBPath$ + "FixedAssets.EXE", vbMaximizedFocus
     DoTheTime
     Unload frmMainMenu
  ElseIf Exist(UBPath$ + "famenu.EXE") Then
    Shell UBPath$ + "famenu.exe", vbMaximizedFocus
    DoEvents
    Unload frmMainMenu
  Else
    frmModuleInfo.Label4 = "Fixed Assets"
    frmModuleInfo.Show
  End If
End Sub
Private Sub cmdPropertyTaxes_Click()
  If Exist(UBPath$ + "TAXCPYRG.EXE") Then
    Shell UBPath$ + "Taxcpryg.exe", vbMaximizedFocus
    DoTheTime
    DoEvents
    Unload frmMainMenu
  Else
    If Exist(UBPath$ + "Taxmenu.exe") Then
      Shell UBPath$ + "Taxmenu.exe", vbMaximizedFocus
      DoEvents
      Unload frmMainMenu
    ElseIf Exist(UBPath$ + "tbmenu.EXE") Then
      Shell UBPath$ + "tbmenu.exe", vbMaximizedFocus
      DoEvents
      Unload frmMainMenu
    Else
      frmModuleInfo.Label4 = "Property Tax Management"
      frmModuleInfo.Show
    End If
  End If
End Sub

Private Sub cmdUtilityBilling_Click()
  If Exist(UBPath$ + "UB.EXE") Then
    Shell UBPath$ + "UB.exe", vbMaximizedFocus
    DoTheTime
    DoEvents
    Unload frmMainMenu
  ElseIf Exist(UBPath$ + "Ubcopyrg.exe") Then
    Shell UBPath$ + "Ubcopyrg.exe", vbMaximizedFocus
    DoEvents
    Unload frmMainMenu
  Else
    If Exist(UBPath$ + "Ubmenu.exe") Then
      Shell UBPath$ + "Ubmenu.exe", vbMaximizedFocus
      DoEvents
      Unload frmMainMenu
    ElseIf Exist(UBPath$ + "wbmenu.EXE") Then
      Shell UBPath$ + "wbmenu.exe", vbMaximizedFocus
      DoEvents
      Unload frmMainMenu
    Else
    frmModuleInfo.Label4 = "Utility Billing"
    frmModuleInfo.Show
    End If
  End If
  
End Sub

Private Sub cmdVehicleDecals_Click()
  If Exist(UBPath$ + "dcmenu.EXE") Then
    Shell UBPath$ + "dcmenu.exe", vbMaximizedFocus
    DoEvents
    Unload frmMainMenu
  Else
    frmModuleInfo.Label4 = "Vehicle Decals"
    frmModuleInfo.Show
  End If
End Sub


Private Sub cmdAcctPayMenu_Click()
  If Exist(UBPath$ + "AP.EXE") Then
    If Exist(UBPath$ + "GLSETUP.DAT") And Exist(UBPath$ + "GLACCT.DAT") Then
      Shell UBPath$ + "AP.exe", vbMaximizedFocus
      DoTheTime
      DoEvents
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
  If Exist(UBPath$ + "GL.EXE") Then
    Shell UBPath$ + "GL.exe", vbMaximizedFocus
    DoTheTime
    DoEvents
    Unload frmMainMenu
  Else
    frmModuleInfo.Label4 = "General Ledger"
    frmModuleInfo.Show
  End If
End Sub


Private Sub cmdPayroll_Click()
  If Exist(UBPath$ + "PAYROLL.EXE") Then
  '  If Not Exist("prrun.opn") Then
      Shell UBPath$ + "Payroll.EXE", vbMaximizedFocus
      DoTheTime
      DoEvents
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
  ElseIf Exist(UBPath$ + "prcopyrg.exe") Then
    Shell UBPath$ + "prcopyrg.exe", vbMaximizedFocus
    DoEvents
    Unload frmMainMenu
  ElseIf Exist(UBPath$ + "pr.exe") Then
    Shell UBPath$ + "pr.exe", vbMaximizedFocus
    DoEvents
    Unload frmMainMenu
  Else
    frmModuleInfo.Label4 = "Payroll"
    frmModuleInfo.Show
  End If
End Sub
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
Private Sub cmdReceiptPrinter_Click()
  Load frmReceiptSetup
  DoEvents
  frmReceiptSetup.Show
  Unload frmMainMenu
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
  UBPath$ = QPTrim$(App.Path)    'start up path
  If Right$(UBPath$, 1) <> "\" Then
    UBPath$ = UBPath$ + "\"
  End If

End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
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

