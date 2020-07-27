VERSION 5.00
Begin VB.Form frmPRFedAllow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   5220
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "GO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1883
      TabIndex        =   1
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "CitipakPR Federal Allowances Adjuster."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   450
      TabIndex        =   0
      Top             =   360
      Width           =   4335
   End
End
Attribute VB_Name = "frmPRFedAllow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ok2Exit As Boolean
Dim DoneExit As Boolean

Private Sub Command1_Click()
  If Ok2Exit Then
    End
  End If
  Command1.Enabled = False
  Command1.Caption = "Wait"
  Call FixEmpData
End Sub

Private Sub Form_Load()
  Ok2Exit = False
  DoneExit = True
End Sub

Private Sub FixEmpData()
  DoneExit = False
  Dim Emp2Rec As EmpData2Type
  Dim Emp2Len As Integer
  Dim RecCnt As Integer
  Dim LopCnt As Integer
  Dim TCnt As Integer
  Emp2Len = Len(Emp2Rec)
  Open "prdata\PREmp2.Dat" For Random As #1 Len = Emp2Len
  RecCnt = LOF(1) / Emp2Len
  For LopCnt = 1 To RecCnt
    Get #1, LopCnt, Emp2Rec
    If Emp2Rec.EMPTDATE = 0 And Emp2Rec.Deleted = 0 Then
       Emp2Rec.EMPFEDA = 1
       Put #1, LopCnt, Emp2Rec
       TCnt = TCnt + 1
    End If
  Next
  Close
  Command1.Caption = "Done"
  Command1.Enabled = True
  Label2.Caption = "Adjusted" + Str(TCnt) + " active employees"
  Ok2Exit = True
  DoneExit = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If DoneExit = False Then
    Cancel = 1
  End If
End Sub
