VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "PRPurgeHist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   240
      Top             =   1680
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   1740
      TabIndex        =   0
      Top             =   1185
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Done"
      Height          =   495
      Left            =   1733
      TabIndex        =   2
      Top             =   1185
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Payroll Old History Purge"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   660
      TabIndex        =   1
      Top             =   390
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Public DoLbl As Boolean

Private Sub Command1_Click()
  Command1.Enabled = False
  Call FixEmUp
  Label1.Caption = "History Purge Complete."
  Command1.Visible = False
  'End
End Sub
Private Sub Command2_Click()
  End
End Sub

Private Sub FixEmUp()
  Dim TRLen As Integer
  Dim TRCnt As Long
  Dim Bogus As Integer
  Dim Cnt As Long
  Dim TRRec As TransRecType
  TRLen = Len(TRRec)
  
  'Print MakeRegDate$(7671)
  'GoTo allDone:
  Call KillOldBack
  
  Name "prdata\prtransh.dat" As "prdata\prtransh.WAS"
  
  Open "prdata\prtransh.was" For Random As #1 Len = TRLen
  Open "prdata\prtransh.dat" For Random As #2 Len = TRLen
  TRCnt = LOF(1) / TRLen
  For Cnt = 1 To TRCnt
    Get #1, Cnt, TRRec
    If TRRec.CheckDate < 14056 Then '12/31/2003
      Put #2, , TRRec
    Else
      Bogus = Bogus + 1
    End If
    'Print MakeRegDate$(TRRec.CheckDate)
    'sleep
  Next
  
  Close
allDone:

Print "Removed: " + CStr(Bogus)
End Sub

Private Sub Form_Load()
  DoLbl = True
End Sub

Private Sub KillOldBack()
On Error Resume Next
  Kill "prdata\prtransh.WAS"
On Error GoTo 0
End Sub


Private Sub Timer1_Timer()
  DoLbl = Not DoLbl
  If DoLbl Then
    Form1.Caption = "Citipak"
  Else
    Form1.Caption = "Citipak"
  End If
  
End Sub
