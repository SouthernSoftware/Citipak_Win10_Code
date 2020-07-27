VERSION 5.00
Begin VB.Form Dialog 
   BackColor       =   &H00000080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Message"
   ClientHeight    =   3195
   ClientLeft      =   1830
   ClientTop       =   2655
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   480
      Top             =   2715
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Close"
      Height          =   375
      Left            =   2190
      TabIndex        =   0
      Top             =   2790
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   60
      TabIndex        =   2
      Top             =   2400
      Visible         =   0   'False
      Width           =   5925
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2190
      Left            =   30
      TabIndex        =   1
      Top             =   60
      Width           =   5955
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MsgDuration As Long

Option Explicit

Private Sub Form_Load()

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Dialog = Nothing
End Sub

Private Sub OKButton_Click()
Unload Dialog
End Sub

Private Sub Timer1_Timer()
Static lCount As Long
Timer1.Enabled = False
lCount = lCount + 1
If MsgDuration <> 0 Then
    Label2.Visible = True
    If lCount >= MsgDuration * 10 Then
        Unload Me
        Exit Sub
    Else

        Label2.Caption = "This message will close in " & MsgDuration - Round(lCount / 10, 0) & " seconds."
    End If
Else
    Label2.Visible = False
    
End If
Timer1.Enabled = True

End Sub
