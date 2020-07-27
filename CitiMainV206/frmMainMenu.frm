VERSION 5.00
Begin VB.Form frmMainMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "v2.05 CitiPak Main Menu"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   12225
   ClipControls    =   0   'False
   Icon            =   "frmMainMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   504
      Top             =   1152
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   480
      Top             =   528
   End
   Begin VB.CommandButton cmdReceiptPrinter 
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Receipt Printer Setup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4320
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7236
      Width           =   3612
   End
   Begin VB.CommandButton cmdPassLogin 
      BackColor       =   &H00D0D0D0&
      Caption         =   "Return to Pass/&Login"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4320
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6750
      Width           =   3612
   End
   Begin VB.CommandButton cmdCashManagement 
      BackColor       =   &H00D0D0D0&
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
      Height          =   372
      Left            =   4320
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5292
      Width           =   3612
   End
   Begin VB.CommandButton cmdPropertyTaxes 
      BackColor       =   &H00D0D0D0&
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
      Height          =   372
      Left            =   4320
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4806
      Width           =   3612
   End
   Begin VB.CommandButton cmdFixedAssets 
      BackColor       =   &H00D0D0D0&
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
      Height          =   372
      Left            =   4320
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
      Width           =   3612
   End
   Begin VB.CommandButton cmdExitCitipak 
      BackColor       =   &H00D0D0D0&
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
      Height          =   372
      Left            =   4320
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7728
      Width           =   3612
   End
   Begin VB.CommandButton cmdVehicleDecals 
      BackColor       =   &H00D0D0D0&
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
      Height          =   372
      Left            =   4320
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6264
      Width           =   3612
   End
   Begin VB.CommandButton cmdPayroll 
      BackColor       =   &H00D0D0D0&
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
      Height          =   372
      Left            =   4320
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3840
      Width           =   3612
   End
   Begin VB.CommandButton cmdGenLedgMenu 
      BackColor       =   &H00D0D0D0&
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
      Height          =   372
      Left            =   4320
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3348
      Width           =   3612
   End
   Begin VB.CommandButton cmdAcctPayMenu 
      BackColor       =   &H00D0D0D0&
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
      Height          =   372
      Left            =   4320
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2862
      UseMaskColor    =   -1  'True
      Width           =   3612
   End
   Begin VB.CommandButton cmdBusinessLicense 
      BackColor       =   &H00D0D0D0&
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
      Height          =   372
      Left            =   4320
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2376
      Width           =   3612
   End
   Begin VB.CommandButton cmdUtilityBilling 
      BackColor       =   &H00D0D0D0&
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
      Height          =   372
      Left            =   4320
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
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
   Begin VB.Shape Shape5 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   1092
      Left            =   1800
      Top             =   1080
      Width           =   8652
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
         Size            =   13.5
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
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1800
      Top             =   960
      Width           =   8652
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   2
      Left            =   9000
      Top             =   2400
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   8880
      Top             =   2160
      Width           =   972
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2400
      Top             =   2160
      Width           =   972
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
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
Dim Mdl As Integer, Oktogo As Boolean

Private Sub cmdBusinessLicense_Click()
    On Error GoTo BLTestExit
    Mdl = BL
    Findsettings
    If Oktogo Then
        DoEvents
        Shell UBPath$ + "BusinessLicense.EXE", vbMaximizedFocus
        frmMainMenu.Enabled = False
        Timer1.Enabled = True
    Else
        MsgBox "Your Password does not allow access to this Module.", vbOKOnly, "Access Denied"
    End If
    GoTo BlNormalExit
BLTestExit:
    If Err = 53 Then
        frmModuleInfo.Label4 = "Business License"
        frmModuleInfo.Show
    Else
        MsgBox "Error Code Was " + Err.Description + Str$(Err)
    End If
BlNormalExit:
    On Error GoTo 0
End Sub

Private Sub cmdCashManagement_Click()
    On Error GoTo CMTestExit
    Mdl = CM
    Findsettings
    If Oktogo Then
        DoEvents
        Shell UBPath$ + "CM.exe", vbMaximizedFocus
        frmMainMenu.Enabled = False
        Timer1.Enabled = True
    Else
        MsgBox "Your Password does not allow access to this Module.", vbOKOnly, "Access Denied"
    End If
    GoTo CMNormalExit
CMTestExit:
    If Err = 53 Then
        frmModuleInfo.Label4 = "Cash Management"
        frmModuleInfo.Show
    Else
        MsgBox "Error Code Was " + Err.Description + Str$(Err)
    End If
CMNormalExit:
    On Error GoTo 0
End Sub

Private Sub cmdFixedAssets_Click()
    On Error GoTo FATestExit
    Mdl = FA
    Findsettings
    If Oktogo Then
        DoEvents
        Shell UBPath$ + "FixedAssets.EXE", vbMaximizedFocus
        frmMainMenu.Enabled = False
        Timer1.Enabled = True
    Else
        MsgBox "Your Password does not allow access to this Module.", vbOKOnly, "Access Denied"
    End If
    GoTo FANormalExit
FATestExit:
    If Err = 53 Then
        frmModuleInfo.Label4 = "Fixed Assets"
        frmModuleInfo.Show
    Else
        MsgBox "Error Code Was " + Err.Description + Str$(Err) + " (CitipakMain CMMenuClick - Line:" & Erl & ")"
    End If
FANormalExit:
    On Error GoTo 0
End Sub

Private Sub cmdPropertyTaxes_Click()
    Dim blnNOTaxes As Boolean
    On Error Resume Next
    blnNOTaxes = False
    Mdl = TX
    Findsettings
    If Oktogo Then
        DoEvents
        Shell UBPath$ + "CitiTaxes.EXE", vbMaximizedFocus
        If Err = 53 Then
          Err.Clear
          GoTo TryVATaxes
        End If
        frmMainMenu.Enabled = False
        Timer1.Enabled = True
    Else
        MsgBox "Your Password does not allow access to this Module.", vbOKOnly, "Access Denied"
    End If
    GoTo TaxNormalExit
TryVATaxes:
    DoEvents
    Shell UBPath$ + "VACitiTax.EXE", vbMaximizedFocus
    If Err = 53 Then
        blnNOTaxes = True
        GoTo TaxErrorExit
    End If
    frmMainMenu.Enabled = False
    Timer1.Enabled = True
    GoTo TaxNormalExit
TaxErrorExit:
    If blnNOTaxes And Err = 53 Then
        frmModuleInfo.Label4 = "Property Tax Management"
        frmModuleInfo.Show
    ElseIf Err = 53 Then
        MsgBox "Error Code Was " + Err.Description + Str$(Err)
    End If
TaxNormalExit:
    On Error GoTo 0
End Sub

Private Sub cmdUtilityBilling_Click()
    On Error GoTo UBTestExit
    Mdl = UB
    Findsettings
    If Oktogo Then
        DoEvents
        Shell UBPath$ + "UB.exe", vbMaximizedFocus
        frmMainMenu.Enabled = False
        Timer1.Enabled = True
        GoTo UBNormalExit
    Else
        MsgBox "Your Password does not allow access to this Module.", vbOKOnly, "Access Denied"
    End If
    GoTo UBNormalExit
UBTestExit:
    If Err = 53 Then
        frmModuleInfo.Label4 = "Utility Billing"
        frmModuleInfo.Show
    Else
        MsgBox "Error Code Was " + Err.Description + Str$(Err)
    End If
UBNormalExit:
    On Error GoTo 0
End Sub

Private Sub cmdVehicleDecals_Click()
    On Error GoTo DCTestExit
    Mdl = DC
    Findsettings
    If Oktogo Then
        DoEvents
        Shell UBPath$ + "DC.exe", vbMaximizedFocus
        frmMainMenu.Enabled = False
        Timer1.Enabled = True
        GoTo DCNormalExit
    Else
        MsgBox "Your Password does not allow access to this Module.", vbOKOnly, "Access Denied"
    End If
     GoTo DCNormalExit
DCTestExit:
    If Err = 53 Then
        frmModuleInfo.Label4 = "Vehicle Decals"
        frmModuleInfo.Show
    Else
        MsgBox "Error Code Was " + Err.Description + Str$(Err)
    End If
DCNormalExit:
    On Error GoTo 0
End Sub

Private Sub cmdAcctPayMenu_Click()
    On Error GoTo APTestExit
    Mdl = AP
    Findsettings
    If Oktogo Then
        If Not Exist(UBPath$ + "GLSETUP.DAT") Or Not Exist(UBPath$ + "GLACCT.DAT") Then
           MsgBox "You Must First Enter GL Setup Information And Accounts.", vbOKOnly, "Missing Setup Info."
           GoTo APNormalExit
        End If
        DoEvents
        Shell UBPath$ + "AP.exe", vbMaximizedFocus
        frmMainMenu.Enabled = False
        Timer1.Enabled = True
    Else
        MsgBox "Your Password does not allow access to this Module.", vbOKOnly, "Access Denied"
    End If
    GoTo APNormalExit
APTestExit:
    If Err = 53 Then
        frmModuleInfo.Label4 = "Accounts Payable"
        frmModuleInfo.Show
    Else
        MsgBox "Error Code Was " + Err.Description + Str$(Err)
    End If
APNormalExit:
    On Error GoTo 0
End Sub

Private Sub cmdGenLedgMenu_Click()
    On Error GoTo GLTestExit
    Mdl = GL
    Findsettings
    If Oktogo Then
        DoEvents
        Shell UBPath$ + "GL.exe", vbMaximizedFocus
        frmMainMenu.Enabled = False
        Timer1.Enabled = True
        GoTo GLNormalExit
    Else
        MsgBox "Your Password does not allow access to this Module.", vbOKOnly, "Access Denied"
    End If
    GoTo GLNormalExit
GLTestExit:
    If Err = 53 Then
        frmModuleInfo.Label4 = "General Ledger"
        frmModuleInfo.Show
    Else
        MsgBox "Error Code Was " + Err.Description + Str$(Err)
    End If
GLNormalExit:
    On Error GoTo 0
End Sub

Private Sub cmdPayroll_Click()
    On Error GoTo PRTestExit
    Mdl = PR
    Findsettings
    If Oktogo Then
        DoEvents
        Shell UBPath$ + "Payroll.EXE", vbMaximizedFocus
        frmMainMenu.Enabled = False
        Timer1.Enabled = True
    Else
        MsgBox "Your Password does not allow access to this Module.", vbOKOnly, "Access Denied"
    End If
    GoTo PRNormalExit
PRTestExit:
    If Err = 53 Then
        frmModuleInfo.Label4 = "Payroll"
        frmModuleInfo.Show
    Else
        MsgBox "Error Code Was " + Err.Description + Str$(Err)
    End If
PRNormalExit:
    On Error GoTo 0
End Sub

Private Sub cmdPassLogin_Click()
  MainLog "In PassLogin Menu"
  ClearInUse PWcnt
  Load frmPassLogin
  DoEvents
  frmPassLogin.Show
  Unload Me
End Sub
Private Sub cmdReceiptPrinter_Click()
  Load frmReceiptSetup
  DoEvents
  frmReceiptSetup.Show
  Unload frmMainMenu
End Sub
Private Sub Form_Load()
  Dim cnt&, dl&
  BlockInput True
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
  disabletheone
  Oktogo = False
  Mdl = 0
  BlockInput False
  DoEvents
  If PWfromMdl > 0 Then
    Timer2.Enabled = True
  End If
  DoEvents
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
      cmdPassLogin_Click
      KeyCode = 0
      DoEvents
    Case Else:
  End Select
End Sub

Public Function Findsettings()
On Error GoTo Cancel
  Dim NumPassRecs As Integer, cnt As Integer, CitiPassFile As Integer
  Dim Okcnt As Integer, LookFor As String
  Dim Citipass As CitiPassType
  Okcnt = 0
10:  If PWcnt = 0 And PWUser$ = "Sosoft Support" Then
    Okcnt = 1
12:  ElseIf PWcnt > 0 Then
13:    If Len(Dir$("Citipass.dat")) Then
14:      SetAttr ("CitiPass.dat"), vbNormal
15:      OpenCitiPassFile CitiPassFile, NumPassRecs
16:      If Not CitiPassFile = -1 Then
17:        Get CitiPassFile, PWcnt, Citipass
18:        If Not Citipass.DelFlag Then
19:        'THIS IS JUST TO SEE IF HAVE ANY ACCESS TO PARTICULAR MODULE
20:            If Citipass.Module(Mdl).FullAccess = True Then
21:               Okcnt = Okcnt + 1
22:            ElseIf Citipass.Module(Mdl).ReportsOnly = True Then
23:               Okcnt = Okcnt + 1
24:            ElseIf Citipass.Module(Mdl).PaymentAccess = True Then
25:               Okcnt = Okcnt + 1
26:            End If
27:        End If
28:        Citipass.FlagMod = Mdl
29:        Put CitiPassFile, PWcnt, Citipass
30:      End If
31:      Close CitiPassFile
32:    End If
  End If
  If Okcnt > 0 Then
    Oktogo = True
  Else
    Oktogo = False
  End If
  Exit Function
  
Cancel:
  If Err > 0 Then
    MsgBox "Error Code Was " + Err.Description + Str$(Err) + " (CitipakMain FindSettings - Line:" & Erl & ")"
    End
  End If
End Function

Private Sub cmdExitCitipak_Click()
  ClearInUse PWcnt
  End
  'Unload frmMainMenu
End Sub

Private Sub Timer1_Timer()
   End
  'Unload Me
End Sub
Private Sub Timer2_Timer()
  cmdAcctPayMenu.Enabled = True
  cmdBusinessLicense.Enabled = True
  cmdCashManagement.Enabled = True
  cmdFixedAssets.Enabled = True
  cmdPropertyTaxes.Enabled = True
  cmdUtilityBilling.Enabled = True
  cmdVehicleDecals.Enabled = True
  cmdGenLedgMenu.Enabled = True
  cmdPayroll.Enabled = True
End Sub

Private Sub disabletheone()
  Select Case PWfromMdl
  Case 1:
     cmdBusinessLicense.Enabled = False
  Case 2:
     cmdAcctPayMenu.Enabled = False
  Case 3:
     cmdGenLedgMenu.Enabled = False
  Case 4:
     cmdPayroll.Enabled = False
  Case 5:
     cmdFixedAssets.Enabled = False
  Case 6:
     cmdPropertyTaxes.Enabled = False
  Case 7:
     '
  Case 8:
     cmdCashManagement.Enabled = False
  Case 9:
     cmdUtilityBilling.Enabled = False
  Case 10:
    cmdVehicleDecals.Enabled = False
  Case Else
    'nothing
  End Select
End Sub

