VERSION 5.00
Begin VB.Form frmBLLicRegisterMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busines License Register Menu"
   ClientHeight    =   8868
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   11652
   Icon            =   "frmBLLicRegisterMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPost 
      BackColor       =   &H008F8265&
      Caption         =   "&Post Data To Customer File"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4032
      MaskColor       =   &H8000000F&
      TabIndex        =   2
      Top             =   4428
      Width           =   3612
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit to Prior Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4032
      TabIndex        =   1
      ToolTipText     =   "Click this button to return to the main Business License menu."
      Top             =   5136
      Width           =   3612
   End
   Begin VB.CommandButton cmdPrintForms 
      BackColor       =   &H008F8265&
      Caption         =   "Print License &Register"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4032
      MaskColor       =   &H8000000F&
      TabIndex        =   0
      Top             =   3732
      Width           =   3612
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LICENSE REGISTER MENU"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2820
      TabIndex        =   3
      Top             =   1248
      Width           =   6012
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      X1              =   2100
      X2              =   3060
      Y1              =   1963
      Y2              =   1963
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000014&
      X1              =   8580
      X2              =   9540
      Y1              =   1963
      Y2              =   1963
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   3060
      X2              =   3060
      Y1              =   2203
      Y2              =   2083
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   2100
      X2              =   2100
      Y1              =   2083
      Y2              =   2203
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   8590
      X2              =   8590
      Y1              =   2100
      Y2              =   2205
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   9550
      X2              =   9550
      Y1              =   2197
      Y2              =   2082
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   8700
      X2              =   8700
      Y1              =   2206
      Y2              =   8078
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   8592
      X2              =   9542
      Y1              =   2203
      Y2              =   2203
   End
   Begin VB.Line Line7 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2100
      X2              =   3060
      Y1              =   2203
      Y2              =   2203
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2220
      X2              =   2220
      Y1              =   2203
      Y2              =   8075
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2219
      X2              =   2929
      Y1              =   8090
      Y2              =   8090
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8703
      X2              =   9403
      Y1              =   8079
      Y2              =   8079
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   2100
      X2              =   3060
      Y1              =   2083
      Y2              =   2083
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   8590
      X2              =   9540
      Y1              =   2083
      Y2              =   2083
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   1092
      Index           =   1
      Left            =   1500
      Top             =   886
      Width           =   8652
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1500
      Top             =   766
      Width           =   8652
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2100
      Top             =   1966
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2220
      Top             =   2196
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   8592
      Top             =   1966
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   8700
      Top             =   2194
      Width           =   732
   End
End
Attribute VB_Name = "frmBLLicRegisterMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  frmBLIssueAppsLics.Show
  DoEvents
  Unload frmBLLicRegisterMenu
End Sub

Private Sub cmdPrintForms_Click()
  frmBLLicRegister.Show
  DoEvents
  Unload frmBLLicRegisterMenu
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsBLTextBoxOverrider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
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
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      Call cmdExit_Click
      SendKeys "%X"
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      ClearInUse PWcnt
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLLicRegisterMenu.")
      Call Terminate
      End
    End If
  End If
End Sub

