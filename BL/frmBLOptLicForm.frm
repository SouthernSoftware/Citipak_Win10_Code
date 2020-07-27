VERSION 5.00
Begin VB.Form frmBLOptRegForm 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7056
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   7056
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox fptxtHide 
      Height          =   684
      Left            =   480
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1968
      Width           =   492
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "ESC  E&xit and Return to Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   1248
      TabIndex        =   2
      Top             =   2256
      Width           =   4716
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "F6  Go To License Register Processing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   1248
      TabIndex        =   1
      Top             =   1680
      Width           =   4716
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "License Registers Must Be Processed and License Forms Must Be Processed Before License Posting Can Take Place. "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   876
      Left            =   480
      TabIndex        =   0
      Top             =   528
      Width           =   6108
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      Height          =   2796
      Left            =   240
      Top             =   192
      Width           =   6588
   End
End
Attribute VB_Name = "frmBLOptRegForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
  'if review is chosen then the selection is scoReviewChanges
  fptxtHide.Text = "Exit"
  Me.Hide
End Sub

Private Sub cmdRegister_Click()
  'if save is chosen then the selection is scoSave
  fptxtHide.Text = "Register"
  Me.Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF6:
      Call cmdRegister_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  fptxtHide.Visible = False
End Sub



