VERSION 5.00
Begin VB.Form frmBLOptPenInProg 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6948
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   6948
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox fptxtHide 
      Height          =   684
      Left            =   420
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1938
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
      Left            =   1188
      TabIndex        =   1
      Top             =   2184
      Width           =   4716
   End
   Begin VB.CommandButton cmdPenPost 
      Caption         =   "F6  Go To Penalty Post"
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
      Left            =   1188
      TabIndex        =   0
      Top             =   1608
      Width           =   4716
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmBLOptPenInProg.frx":0000
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
      Height          =   1020
      Left            =   420
      TabIndex        =   3
      Top             =   432
      Width           =   6108
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      Height          =   2796
      Left            =   180
      Top             =   162
      Width           =   6588
   End
End
Attribute VB_Name = "frmBLOptPenInProg"
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

Private Sub cmdPenPost_Click()
  'if save is chosen then the selection is scoSave
  fptxtHide.Text = "PenPost"
  Me.Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF6:
      Call cmdPenPost_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  fptxtHide.Visible = False
End Sub





