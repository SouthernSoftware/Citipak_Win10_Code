VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Southern Software Inc."
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnOKCan 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   1620
      TabIndex        =   2
      Top             =   1995
      Width           =   1440
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   510
      TabIndex        =   1
      Top             =   870
      Width           =   3705
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "CM Transaction Correction Utility."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   495
      TabIndex        =   0
      Top             =   375
      Width           =   3705
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Dim CMTrans As CMTransRecType
Dim DoneWithIT As Boolean
Dim Date2Find As Integer
Dim Date2Set As Integer
Dim BadTR As Long

Private Sub Form_Load()
  DoneWithIT = False
End Sub

Private Sub btnOKCan_Click()
  If DoneWithIT Then
    End
  Else
    btnOKCan.Enabled = False
    Call DoCMTransThing
    Call ShowDone
  End If
    
End Sub

Private Sub DoCMTransThing()
  Dim TRLen As Integer
  Dim TRCnt As Long
  Dim ThisTR As Long
  Date2Find = Date2Num("12/31/1979")
  Date2Set = Date2Num("05/27/2014")
  TRLen = Len(CMTrans)
  Open "CMTRANS.DAT" For Random As #1 Len = TRLen
  TRCnt = LOF(1) / TRLen
  For ThisTR = 1 To TRCnt
    Get #1, ThisTR, CMTrans
    If CMTrans.TransDate = Date2Find Then
      CMTrans.TransDate = Date2Set
      BadTR = BadTR + 1
      Put #1, ThisTR, CMTrans
    End If
    Label2.Caption = MakePctComp(ThisTR, TRCnt) + " Complete."
    DoEvents
  Next
  Close
  DoneWithIT = True
End Sub

Private Sub ShowDone()
  Label2.Caption = "Processing Complete."
  btnOKCan.Caption = "DONE"
  'btnOKCan.Caption = CStr(BadTR) '"DONE"
  btnOKCan.Enabled = True
  
End Sub

Public Function MakePctComp(ByVal Cnt As Long, ByVal TotalCnt As Long) As String
  Dim PctComp As Long
  Dim RetStr As String
  PctComp = Int((Cnt / TotalCnt) * 100)
  RetStr = CStr(PctComp) + "%"
  MakePctComp = RetStr
End Function

Public Function Date2Num(TheDate$) As Integer
 'useful function throughout program...
 'takes a string date and converts into a number based on 12/31/1979
  Date2Num% = DateDiff("d", "12/31/1979", (TheDate$))
End Function

