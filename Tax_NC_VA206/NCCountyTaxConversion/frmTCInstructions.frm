VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmTCInstructions 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Instructions"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11655
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTCInstructions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   510
      Left            =   4950
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "Press the 'Cancel' button to exit this screen and return to the main 'Business License Reports' menu."
      Top             =   8160
      Width           =   1740
      _Version        =   131072
      _ExtentX        =   3069
      _ExtentY        =   900
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmTCInstructions.frx":08CA
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000FFFF&
      Caption         =   $"frmTCInstructions.frx":0AA8
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   840
      TabIndex        =   9
      Top             =   7440
      Width           =   9975
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000009&
      X1              =   660
      X2              =   10985
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000FFFF&
      Caption         =   "Step Eight: If you intend to convert optional revenues then be absolutely sure you have rate codes already set up for them. "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   840
      TabIndex        =   8
      Top             =   6600
      Width           =   9975
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000009&
      X1              =   675
      X2              =   11000
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      X1              =   660
      X2              =   10980
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      X1              =   660
      X2              =   10980
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      X1              =   660
      X2              =   10980
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      X1              =   660
      X2              =   10980
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   660
      X2              =   10980
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   660
      X2              =   10980
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      FillColor       =   &H00FFFFFF&
      Height          =   8055
      Left            =   660
      Top             =   120
      Width           =   10335
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000FFFF&
      Caption         =   $"frmTCInstructions.frx":0B82
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1095
      Left            =   840
      TabIndex        =   7
      Top             =   5280
      Width           =   9975
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000FFFF&
      Caption         =   $"frmTCInstructions.frx":0CDA
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   840
      TabIndex        =   6
      Top             =   4440
      Width           =   9975
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000FFFF&
      Caption         =   $"frmTCInstructions.frx":0D9F
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1095
      Left            =   840
      TabIndex        =   5
      Top             =   3120
      Width           =   9975
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000FFFF&
      Caption         =   "Step Four:  Clear the existing spreadsheet (even if there isn't an existing spreadsheet) on the main menu."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   2640
      Width           =   9975
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FFFF&
      Caption         =   $"frmTCInstructions.frx":0F3F
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   1920
      Width           =   9975
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Caption         =   $"frmTCInstructions.frx":1008
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   840
      TabIndex        =   2
      Top             =   960
      Width           =   9975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   $"frmTCInstructions.frx":10FD
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   9975
   End
End
Attribute VB_Name = "frmTCInstructions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
Private Sub cmdExit_Click()
  frmTCMainMenuNew.Show
  DoEvents
  Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%C"
      Call cmdExit_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub
Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTCConvert.")
      End
    End If
  End If

End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If

End Sub

