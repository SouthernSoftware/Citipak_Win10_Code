VERSION 5.00
Begin VB.Form frmFAMainMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "v 2.01 Fixed Assets Main Menu"
   ClientHeight    =   8868
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   11652
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSetUp 
      BackColor       =   &H008F8265&
      Caption         =   "&FA Setup Maintenance"
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
      TabIndex        =   6
      Top             =   5880
      Width           =   3612
   End
   Begin VB.CommandButton cmdItemMaint 
      BackColor       =   &H008F8265&
      Caption         =   "&Item Maintenance"
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
      Left            =   4005
      MaskColor       =   &H8000000F&
      TabIndex        =   4
      Top             =   2904
      Width           =   3612
   End
   Begin VB.CommandButton cmdReportsMenu 
      BackColor       =   &H008F8265&
      Caption         =   "&Reports Menu"
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
      Left            =   4005
      MaskColor       =   &H8000000F&
      TabIndex        =   3
      Top             =   3648
      Width           =   3612
   End
   Begin VB.CommandButton cmdYearEndProc 
      BackColor       =   &H008F8265&
      Caption         =   "&Year End Processing"
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
      Left            =   4005
      MaskColor       =   &H8000000F&
      TabIndex        =   2
      Top             =   4392
      Width           =   3612
   End
   Begin VB.CommandButton cmdAssetCodeMaint 
      BackColor       =   &H008F8265&
      Caption         =   "&Asset Code Maintenance"
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
      TabIndex        =   1
      Top             =   5136
      Width           =   3612
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit Fixed Assets"
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
      TabIndex        =   0
      Top             =   6624
      Width           =   3612
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FIXED ASSETS MAIN MENU"
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
      TabIndex        =   5
      Top             =   1246
      Width           =   6012
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      X1              =   2100
      X2              =   3060
      Y1              =   1912.53
      Y2              =   1912.53
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000014&
      X1              =   8580
      X2              =   9540
      Y1              =   1912.53
      Y2              =   1912.53
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   3060
      X2              =   3060
      Y1              =   2146.36
      Y2              =   2029.445
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   2100
      X2              =   2100
      Y1              =   2029.445
      Y2              =   2146.36
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   8590
      X2              =   8590
      Y1              =   2046.008
      Y2              =   2148.309
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   9550
      X2              =   9550
      Y1              =   2140.514
      Y2              =   2028.471
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   8700
      X2              =   8700
      Y1              =   2149.283
      Y2              =   7870.311
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   8592
      X2              =   9542
      Y1              =   2146.36
      Y2              =   2146.36
   End
   Begin VB.Line Line7 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2100
      X2              =   3060
      Y1              =   2146.36
      Y2              =   2146.36
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2220
      X2              =   2220
      Y1              =   2146.36
      Y2              =   7867.388
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2219
      X2              =   2929
      Y1              =   7882.003
      Y2              =   7882.003
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8703
      X2              =   9403
      Y1              =   7882.003
      Y2              =   7882.003
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   2100
      X2              =   3060
      Y1              =   2029.445
      Y2              =   2029.445
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   8590
      X2              =   9540
      Y1              =   2029.445
      Y2              =   2029.445
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
Attribute VB_Name = "frmFAMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdAssetCodeMaint_Click()
  frmFAAssetsCodesmenu.Show
  DoEvents
  Unload frmFAMainMenu
End Sub

Private Sub cmdExit_Click()
  Close
  MainLog ("FixedAssets.exe terminated via normal exit in Payroll Main Menu.")
  If Exist(QPTrim$(StartPath) + "\" + "Citipak.exe") Then
    Shell QPTrim$(StartPath) + "\" + "Citipak.exe", vbMaximizedFocus
  End If
  DoEvents
  Call UnloadAllFormsAndOpn
  DoEvents
  End

End Sub

Private Sub cmdItemMaint_Click()
  frmFAItemMaintMenu.Show
  DoEvents
  Unload frmFAMainMenu
End Sub

Private Sub cmdReportsMenu_Click()
  frmFAReportMenu.Show
  DoEvents
  Unload frmFAMainMenu
End Sub

Private Sub cmdSetUp_Click()
  frmFASystemSetup.Show
  DoEvents
  Unload frmFAMainMenu
End Sub

Private Sub cmdYearEndProc_Click()
  frmFAYearEndMenu.Show
  DoEvents
  Unload frmFAMainMenu
End Sub

Private Sub Form_Load()
  Dim FirstThru As Boolean
  Dim Cnt&, dl&
  
  Clipboard.Clear
  If App.PrevInstance Then
    ActivatePrevInstance 'don't want two payroll
    'programs open at once
  End If
  
  'the next series of code is used to get the
  'identity of the current clerk using payroll
  'and recorded anytime MainLog is accessed
  Cnt& = 199
  ComputerName$ = String$(200, 0)
  dl& = GetUserName(ComputerName$, Cnt)
  ComputerName$ = QPTrim$(ComputerName$)
  
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsFATextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  
  'this saves the current path
  StartPath = App.Path
  If Right$(StartPath, 1) = "\" Then
    StartPath = Mid$(StartPath, 1, Len(StartPath) - 1)
  End If

  RecNum = 0
  Call ExtractDepts
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
      Call UnloadAllFormsAndOpn
'      ClearInUse PWcnt
      MainLog ("Payroll.exe terminated via menu bar on frmPayrollMainMenu.")
      End
    End If
  End If
End Sub

