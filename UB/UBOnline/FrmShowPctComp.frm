VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmShowPctComp 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2892
   ClientLeft      =   36
   ClientTop       =   108
   ClientWidth     =   5592
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawWidth       =   3
   Icon            =   "FrmShowPctComp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2892
   ScaleWidth      =   5592
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   192
      Top             =   288
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   372
      Left            =   888
      TabIndex        =   2
      Top             =   1632
      Width           =   3852
      _ExtentX        =   6795
      _ExtentY        =   656
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   11
      Scrolling       =   1
   End
   Begin VB.Label lblComplete 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Completed"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   960
      TabIndex        =   6
      Top             =   2256
      Visible         =   0   'False
      Width           =   3660
   End
   Begin VB.Label AutoClose 
      BackStyle       =   0  'Transparent
      Height          =   300
      Left            =   360
      TabIndex        =   5
      Top             =   2376
      Visible         =   0   'False
      Width           =   684
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   150
      TabIndex        =   4
      Top             =   528
      Width           =   5292
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "% Complete."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3132
      TabIndex        =   3
      Top             =   1188
      Width           =   2340
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Processing:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   300
      TabIndex        =   0
      Top             =   1176
      Width           =   2220
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " 00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2262
      TabIndex        =   1
      Top             =   1212
      Width           =   732
   End
End
Attribute VB_Name = "FrmShowPctComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Temp_Class As Resize_Class
'Dim Over As clsTextBoxOverRider
'Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim vWidth%, vHeight%, vTop%, vLeft%
'Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal _
'    hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal _
'    cx As Long, ByVal cy As Long, wFlags As Long) As Long
Public Out As Boolean
'Const HWND_TOPMOST = -1
'Const SWP_SHOWWINDOW = &H40
'Const SWP_DRAWFRAME = &H20
'Const SWP_FRAMECHANGED = &H20

Private Sub CmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
    Case vbKeyEscape:
      SendKeys "%C"
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Initialize()
  vLeft = (Screen.Width * 0.5)  ' Set width of form.
  vTop = (Screen.Height * 0.5) ' Set height of form.
  vWidth = 525 '(Screen.Width - vWidth) \ 2   ' Center form horizontally.
  vHeight = 280 '((Screen.Height - vHeight) \ 2) + 10  ' Center form vertically.
  'Out = False
End Sub

'Private Sub cmdCancel_Click()
'  'MakeWindowTopMost Me.hwnd, False
'  If MsgBox("Are You Sure You Want To Cancel?", vbYesNo + vbSystemModal, "Cancel Processing") = vbYes Then
'    FrmShowPctComp.Out = True
'    MakeWindowTopMost Me.hWnd, False
'    Unload FrmShowPctComp
'  Else
'    MakeWindowTopMost Me.hWnd, True
'  End If
'End Sub

Private Sub Form_Load()
  Dim RetVal As Long, winhand As Long
  'Set Temp_Class = New Resize_Class
  'Temp_Class.InitResizeClass Me
'  Set Over = New clsTextBoxOverRider
'  Over.OverRide Me
'  MakeWindowTopMost Me.hWnd, True
  ProgressBar1.Value = 0
  FrmShowPctComp.Out = 0
'  Me.CmdCancel.SetFocus
  Out = False
End Sub

Public Sub ShowPctComp(ByVal cnt As Long, ByVal TotalCnt As Long)
  Dim PctComp As Long
  On Local Error Resume Next
  If TotalCnt = 0 Then
    TotalCnt = 1
    cnt = 1
  End If
  PctComp = Int((cnt / TotalCnt) * 100)
  FrmShowPctComp.Label3 = PctComp
  ProgressBar1.Value = PctComp
  If (PctComp = 100) And (Len(FrmShowPctComp.AutoClose) = 0) Then
    'MakeWindowTopMost Me.hWnd, False
    lblComplete.Visible = True
    Timer1.Enabled = True
    DoEvents
  Else
    DoEvents
  End If
End Sub

'Private Sub Form_Resize()
'  If Me.CmdCancel.Enabled Then
'    Me.CmdCancel.SetFocus
'  End If
'End Sub
Private Sub Timer1_Timer()
  Unload FrmShowPctComp
End Sub
