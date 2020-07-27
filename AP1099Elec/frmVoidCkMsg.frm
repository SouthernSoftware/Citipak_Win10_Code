VERSION 5.00
Begin VB.Form frmVoidCkMsg 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Void A/P Check Information"
   ClientHeight    =   3048
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   7656
   ControlBox      =   0   'False
   Icon            =   "frmVoidCkMsg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3048
   ScaleWidth      =   7656
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   375
      Left            =   312
      Top             =   2112
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   3330
      TabIndex        =   3
      Top             =   2352
      Width           =   996
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1812
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   7092
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "To void the invoice, select Void An Open Invoice from the Invoice Processing Menu."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   636
      TabIndex        =   2
      Top             =   1344
      Width           =   6396
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "The invoice paid by this check was reset to open."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   468
      Left            =   444
      TabIndex        =   1
      Top             =   888
      Width           =   6780
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "This operation voided this check only."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   636
      TabIndex        =   0
      Top             =   456
      Width           =   6396
   End
End
Attribute VB_Name = "frmVoidCkMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer


Private Sub cmdOk_Click()
  Unload frmVoidCkMsg
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
      SendKeys "%O"
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



