VERSION 5.00
Begin VB.Form frmBLOptPrintLic 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3012
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6804
   LinkTopic       =   "Form1"
   ScaleHeight     =   3012
   ScaleWidth      =   6804
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLaser 
      Caption         =   "F5  Go To Print &Laser Forms"
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
      Left            =   1104
      TabIndex        =   3
      Top             =   1008
      Width           =   4716
   End
   Begin VB.CommandButton cmdFormFed 
      Caption         =   "F6  Go To Print Form Fed Forms"
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
      Left            =   1104
      TabIndex        =   2
      Top             =   1584
      Width           =   4716
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
      Left            =   1104
      TabIndex        =   1
      Top             =   2160
      Width           =   4716
   End
   Begin VB.TextBox fptxtHide 
      Height          =   684
      Left            =   336
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1872
      Width           =   492
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      Height          =   2796
      Left            =   96
      Top             =   96
      Width           =   6588
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please Print License Data first. Select one of the three options below."
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
      Height          =   684
      Left            =   336
      TabIndex        =   4
      Top             =   288
      Width           =   6108
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmBLOptPrintLic"
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

Private Sub cmdFormFed_Click()
  'if save is chosen then the selection is scoSave
  fptxtHide.Text = "FormFed"
  Me.Hide
End Sub

Private Sub cmdLaser_Click()
  fptxtHide.Text = "Laser"
  Me.Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF5:
      Call cmdLaser_Click
      KeyCode = 0
    Case vbKeyF6:
      Call cmdFormFed_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  fptxtHide.Visible = False
End Sub




