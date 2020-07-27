VERSION 5.00
Begin VB.Form frmBLTransJrnlMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Transaction Journal Menu"
   ClientHeight    =   11376
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   11652
   Icon            =   "frmBLTransJrnlMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11376
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCust 
      BackColor       =   &H008F8265&
      Caption         =   "Transactions by Category Code"
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
      Top             =   4392
      Width           =   3612
   End
   Begin VB.CommandButton cmdType 
      BackColor       =   &H008F8265&
      Caption         =   "Transactions By Payment Type"
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
      Top             =   3696
      Width           =   3612
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit Menu"
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
      ToolTipText     =   "Click this button to return to the main Business License menu."
      Top             =   5100
      Width           =   3612
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   1092
      Index           =   1
      Left            =   1500
      Top             =   888
      Width           =   8652
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   8592
      X2              =   9542
      Y1              =   2088
      Y2              =   2088
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   2100
      X2              =   3060
      Y1              =   2088
      Y2              =   2088
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8700
      X2              =   9400
      Y1              =   8088
      Y2              =   8088
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2220
      X2              =   2930
      Y1              =   8100
      Y2              =   8100
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2220
      X2              =   2220
      Y1              =   2208
      Y2              =   8080
   End
   Begin VB.Line Line7 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2100
      X2              =   3060
      Y1              =   2208
      Y2              =   2208
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   8592
      X2              =   9542
      Y1              =   2208
      Y2              =   2208
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   8700
      X2              =   8700
      Y1              =   2208
      Y2              =   8080
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   9552
      X2              =   9552
      Y1              =   2203
      Y2              =   2088
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   8592
      X2              =   8592
      Y1              =   2112
      Y2              =   2217
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   2100
      X2              =   2100
      Y1              =   2088
      Y2              =   2208
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   3060
      X2              =   3060
      Y1              =   2208
      Y2              =   2088
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000014&
      X1              =   8580
      X2              =   9540
      Y1              =   1968
      Y2              =   1968
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      X1              =   2100
      X2              =   3060
      Y1              =   1968
      Y2              =   1968
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TRANSACTION JOURNAL MENU"
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
      Top             =   1260
      Width           =   6012
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1500
      Top             =   768
      Width           =   8652
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   8592
      Top             =   1968
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   8700
      Top             =   2196
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2100
      Top             =   1968
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2220
      Top             =   2208
      Width           =   732
   End
End
Attribute VB_Name = "frmBLTransJrnlMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  frmBLCustReportsMenu.Show
  DoEvents
  Unload frmBLTransJrnlMenu
End Sub

Private Sub cmdCust_Click()
  frmBLTransJrnlByCat.Show
  DoEvents
  Unload frmBLTransJrnlMenu
End Sub

Private Sub cmdPrintForms_Click()
  frmBLDelinquentNotices.Show
  DoEvents
  Unload frmBLDelinquentMenu
End Sub
Private Sub cmdType_Click()
  frmBLTransJournal.Show
  DoEvents
  Unload frmBLTransJrnlMenu
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
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLTransJrnlMenu.")
      Call Terminate
      End
    End If
  End If
End Sub


