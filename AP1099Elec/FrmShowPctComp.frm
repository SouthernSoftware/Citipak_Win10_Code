VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmShowPctComp 
   BackColor       =   &H00D0D0D0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2892
   ClientLeft      =   36
   ClientTop       =   108
   ClientWidth     =   5592
   ControlBox      =   0   'False
   DrawWidth       =   3
   Icon            =   "FrmShowPctComp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2892
   ScaleWidth      =   5592
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2214
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2208
      Width           =   1164
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   396
      Left            =   870
      TabIndex        =   3
      Top             =   1548
      Width           =   3852
      _ExtentX        =   6795
      _ExtentY        =   699
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
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
      Height          =   540
      Left            =   606
      TabIndex        =   4
      Top             =   588
      Width           =   4380
   End
   Begin VB.Label Label2 
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
      Left            =   918
      TabIndex        =   0
      Top             =   1212
      Width           =   1596
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
      Left            =   3126
      TabIndex        =   2
      Top             =   1212
      Width           =   1572
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
'Private Sub Form_LostFocus()
'  FrmShowPctComp.SetFocus
'End Sub
'
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'  If (UnloadMode <> 1) Then
'    Cancel = True
'  End If
'End Sub
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim vWidth%, vHeight%, vTop%, vLeft%
Dim RetValue As Integer
'Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal _
'    hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal _
'    cx As Long, ByVal cy As Long, wFlags As Long) As Long
Public Out As Boolean
'Const HWND_TOPMOST = -1
'Const SWP_SHOWWINDOW = &H40
'Const SWP_DRAWFRAME = &H20
'Const SWP_FRAMECHANGED = &H20

'


Private Sub Form_Initialize()
  vLeft = (Screen.Width * 0.5)  ' Set width of form.
  vTop = (Screen.Height * 0.5) ' Set height of form.
  vWidth = 525 '(Screen.Width - vWidth) \ 2   ' Center form horizontally.
  vHeight = 280 '((Screen.Height - vHeight) \ 2) + 10  ' Center form vertically.

End Sub

Private Sub cmdCancel_Click()
  'MakeWindowTopMost Me.hwnd, False
  'RetValue = sndPlaySound("cancel.wav", SND_ASYNC Or SND_NODEFAULT)
  If MsgBox("Are You Sure You Want To Cancel?", vbYesNo + vbSystemModal, "Cancel Processing") = vbYes Then
    FrmShowPctComp.Out = True
    MakeWindowTopMost Me.hwnd, False
    Unload FrmShowPctComp
  Else
    MakeWindowTopMost Me.hwnd, True
  End If
End Sub

Private Sub Form_Load()
Dim RetVal As Long, winhand As Long
  'Set Temp_Class = New Resize_Class
  'Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  MakeWindowTopMost Me.hwnd, True
   'retVal = SetWindowPos(Me.hwnd, HWND_TOPMOST, vLeft, vTop, vWidth, vHeight, SWP_DRAWFRAME Or SWP_SHOWWINDOW)
'  Me.Width = vWidth
'  Me.Height = vHeight
'  Me.Left = vLeft
'  Me.Top = vTop
  ProgressBar1.Value = 0
  Out = False

End Sub



'Private Sub Form_Resize()
''  If Me.Visible Then
'    Temp_Class.ResizeControls Me
'    DoEvents
''  End If
'End Sub

Public Sub ShowPctComp(ByVal cnt As Long, ByVal TotalCnt As Long)
  Dim PctComp As Long
  If TotalCnt = 0 Then
    TotalCnt = 1
    cnt = 1
  End If
  PctComp = Int((cnt / TotalCnt) * 100)
  FrmShowPctComp.Label3 = PctComp
  ProgressBar1.Value = PctComp
  If PctComp = 100 Then
    MakeWindowTopMost Me.hwnd, False
    Unload FrmShowPctComp
    DoEvents
  Else
    DoEvents
  End If
End Sub

''Public Sub ShowAcct(ByVal Show As String)
''  FrmShowPctComp.Label5 = Show
''  DoEvents
''End Sub
