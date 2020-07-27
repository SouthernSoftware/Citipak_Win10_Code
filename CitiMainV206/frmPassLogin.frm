VERSION 5.00
Begin VB.Form frmPassLogin 
   AutoRedraw      =   -1  'True
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "v2.05 CitiPak Password Login"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   12225
   ClipControls    =   0   'False
   Icon            =   "frmPassLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPasswords 
      BackColor       =   &H00D0D0D0&
      Caption         =   "PASSWORD &MAINTENANCE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   3330
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4128
      Width           =   5556
   End
   Begin VB.CommandButton cmdExitCitipak 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "E&XIT CITIPAK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   3330
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4980
      Width           =   5556
   End
   Begin VB.CommandButton cmdLogin 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "CITIPAK &LOGIN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   3330
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3276
      Width           =   5556
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   4404
      Left            =   1800
      Top             =   2160
      Width           =   8652
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   1092
      Left            =   1800
      Top             =   1080
      Width           =   8652
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CITIPAK PASSWORD LOGIN"
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
      TabIndex        =   3
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
End
Attribute VB_Name = "frmPassLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class


Private Sub cmdLogin_Click()
  If Exist(PassP$) Then
    GetTemp
    frmWarning.Label1.Caption = "CitiPak Login Still Active"
    frmWarning.Label6.Caption = "For: " + PWUser$
    frmWarning.Label4.Caption = "Continue to Cancel Previous Session or contact Password Administrator."
    frmWarning.Show 1
    If frmWarning.nogo = True Then
      Exit Sub
    Else
      MainLog "Prev session killed," + PWUser$
    End If
  End If
  PWUser = ""
  PWcnt = 0
  MainLog "In Main Login"
  Load frmSecurityPass
  DoEvents
  frmSecurityPass.Show
  Unload Me
End Sub

Private Sub cmdExitCitipak_Click()
  Unload Me
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
  Load frmMainPassWord
  DoEvents
  frmMainPassWord.Show
  'Unload Me
End Sub

Private Sub Form_Load()
  If App.PrevInstance Then
    ActivatePrevInstance
  End If
  
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me

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

