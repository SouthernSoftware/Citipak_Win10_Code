VERSION 5.00
Begin VB.Form frmSortFund 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sort Funds"
   ClientHeight    =   2700
   ClientLeft      =   48
   ClientTop       =   96
   ClientWidth     =   5448
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   5448
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00D0D0D0&
      Caption         =   "E&XIT"
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
      Left            =   3000
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   1212
   End
   Begin VB.CommandButton cmdOKSort 
      BackColor       =   &H00D0D0D0&
      Caption         =   "&OK"
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
      Left            =   1200
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1680
      Width           =   1212
   End
   Begin VB.Image Image1 
      Height          =   384
      Left            =   480
      Picture         =   "frmSortFund.frx":0000
      Top             =   600
      Width           =   384
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ready to Sort Fund Index?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   372
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   3732
   End
End
Attribute VB_Name = "frmSortFund"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim vWidth%, vHeight%, vTop%, vLeft%
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyUp:
      SendKeys "+{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%X"
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub cmdExit_Click()
  Unload frmSortFund
End Sub

Private Sub cmdOKSort_Click()
  Call MainLog("Sort Fund Started - Menu Option.")
  SortFundIndex 'this is in main module
  Call MainLog("Sort Fund Complete - Menu Option.")
  MsgBox "Sort is Complete, Press OK to Continue", vbOKOnly, "Sort Completed"
  Call cmdExit_Click
End Sub

Private Sub Form_Initialize()
  vWidth = Screen.Width * 0.45    ' Set width of form.
  vHeight = Screen.Height * 0.33  ' Set height of form.
  vLeft = (Screen.Width - vWidth) \ 2   ' Center form horizontally.
  vTop = ((Screen.Height - vHeight) \ 2) + 10  ' Center form vertically.
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  Me.Width = vWidth
  Me.Height = vHeight
  Me.Left = vLeft
  Me.Top = vTop
  Me.HelpContextID = hlpFundIndex
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
