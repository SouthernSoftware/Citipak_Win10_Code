VERSION 5.00
Begin VB.Form frmVoidPrnMsg 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Warning!"
   ClientHeight    =   3516
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   7512
   ControlBox      =   0   'False
   Icon            =   "frmVoidPrnMsg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3516
   ScaleWidth      =   7512
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   375
      Left            =   288
      Top             =   2712
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Escape E&xit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1590
      TabIndex        =   1
      Top             =   2832
      Width           =   1764
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "F10 &Continue"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4182
      TabIndex        =   0
      Top             =   2832
      Width           =   1740
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      Caption         =   $"frmVoidPrnMsg.frx":08CA
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   588
      Left            =   438
      TabIndex        =   5
      Top             =   1320
      Width           =   6636
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "The Next Step Must Be POSTING."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   396
      Left            =   1800
      TabIndex        =   4
      Top             =   936
      Width           =   3924
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "The Void Printed Check Procedure Should Be Followed By Printing A Check Register."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   636
      Left            =   564
      TabIndex        =   3
      Top             =   240
      Width           =   6396
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " Are you sure you want to continue? Press (ESC) to Cancel, (F10) to Continue."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   660
      Left            =   1542
      TabIndex        =   2
      Top             =   2088
      Width           =   4428
   End
End
Attribute VB_Name = "frmVoidPrnMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer

Private Sub cmdContinue_Click()
  frmChkPrnCancel.Show
  Unload frmAPChkProcessMenu
  Unload frmVoidPrnMsg
End Sub

Private Sub cmdExit_Click()
  Unload frmVoidPrnMsg
End Sub
Private Sub Timer1_Timer()
 ' Label2.Visible = Not Label2.Visible
  '&H0080FFFF&
  Static tog As Boolean
  tog = Not tog
  If tog Then
    Me.BackColor = &HC0&
  Else
    Me.BackColor = &H80&
  End If
  
End Sub

Private Sub Form_Load()
Dim RetVal As Long, winhand As Long
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  MakeWindowTopMost Me.hwnd, True

End Sub

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
    Case vbKeyF10:
      SendKeys "%C"
      KeyCode = 0
    Case Else:
  End Select
End Sub


Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub


