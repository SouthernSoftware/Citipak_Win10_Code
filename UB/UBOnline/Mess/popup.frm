VERSION 5.00
Begin VB.Form popup 
   BackColor       =   &H000000FF&
   Caption         =   "Alert Message"
   ClientHeight    =   3450
   ClientLeft      =   3240
   ClientTop       =   1545
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   4635
   Begin VB.Label msg3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1020
      Left            =   120
      TabIndex        =   2
      Top             =   2340
      Width           =   4455
   End
   Begin VB.Label msg2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1020
      Left            =   105
      TabIndex        =   1
      Top             =   1095
      Width           =   4455
   End
   Begin VB.Label msg1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1020
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   4455
   End
End
Attribute VB_Name = "popup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set popup = Nothing
End Sub

Private Sub msg1_Click()

End Sub

Private Sub msg2_Click()

End Sub
