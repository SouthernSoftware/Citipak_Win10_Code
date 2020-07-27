VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2880
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   6510
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1148
      TabIndex        =   3
      Top             =   2280
      Width           =   4215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   550
      Width           =   3615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   1040
      Width           =   3615
   End
   Begin VB.Label Label1 
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
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   1480
      Width           =   3615
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LoadLoadYears()
  ReDim BillYears(1 To 50) As Integer
  Dim WorkYears(1 To 50) As Integer
  Dim TransRec As UBTransRecType
  Dim TRLen As Integer
  Dim TRRecCnt As Long
  Dim LopCnt As Long
  Dim XCnt As Integer
  Dim BYCnt As Integer
  Dim TrBlYr As Integer
  Dim GotOne As Boolean
  
  TRLen = Len(TransRec)
  Open "UBTrans.Dat" For Random As #1 Len = TRLen
  BYCnt = 0
  TRRecCnt = LOF(1) / TRLen
  
  For LopCnt = 1 To TRRecCnt
   Get #1, LopCnt, TransRec
   If TransRec.TransType = 1 Then
     GotOne = False
     If BYCnt < 1 Then
       BYCnt = BYCnt + 1
       TrBlYr = GetBillYear(TransRec.TransDate)
       BillYears(BYCnt) = TrBlYr
     Else
       TrBlYr = GetBillYear(TransRec.TransDate)
       For XCnt = 1 To BYCnt
         If TrBlYr = BillYears(XCnt) Then
           GotOne = True
           Exit For
         End If
       Next
       If Not GotOne Then
         BYCnt = BYCnt + 1
         BillYears(BYCnt) = TrBlYr
       End If
     End If
   End If
   If LopCnt Mod 1000 = 0 Then
     Label1.Caption = "Scanning billing years. ." + ShowPctComp(LopCnt, TRRecCnt)
     DoEvents
   End If
  Next
  Close
  ReDim Preserve BillYears(1 To BYCnt) As Integer
  YearsQSort BillYears(), 1, BYCnt
End Sub

Public Sub YearsQSort(IdxBuff() As Integer, lLBound, lUBound)
  Dim lngCurMid As Long, lngCurLow As Long, lngCurHigh As Long
  Dim Temp As Integer
  Dim Temp2 As Integer
  lngCurLow = 1
  lngCurHigh = UBound(IdxBuff())
  'this is to exit loop if high and low are equal
  'Stop
  
  If lUBound <= lLBound Then Exit Sub 'GoTo Proc_Exit
    'find the midpoint
    lngCurMid = (lUBound + lLBound) \ 2
    '
    Temp = IdxBuff(lngCurMid)
    Do While (lngCurLow <= lngCurHigh)
      Do While IdxBuff(lngCurLow) < Temp
        lngCurLow = lngCurLow + 1
        If lngCurLow = lUBound Then Exit Do
      Loop
      Do While Temp < IdxBuff(lngCurHigh)
        lngCurHigh = lngCurHigh - 1
        If lngCurHigh = lLBound Then Exit Do
      Loop
      'if low is <= high then swap
      If (lngCurLow <= lngCurHigh) Then
        Temp2 = IdxBuff(lngCurLow)
        IdxBuff(lngCurLow) = IdxBuff(lngCurHigh)
        IdxBuff(lngCurHigh) = Temp2
        lngCurLow = lngCurLow + 1
        lngCurHigh = lngCurHigh - 1
      End If
    Loop
  'recurse if necessary
    If lLBound < lngCurHigh Then
      YearsQSort IdxBuff(), lLBound, lngCurHigh
    End If
  'recurse if necessary
    If lngCurLow < lUBound Then
      YearsQSort IdxBuff(), lngCurLow, lUBound
    End If
End Sub


Private Sub Form_Paint()
  Label4.Caption = "Copyright © 2019 Southern Software Inc."
  Label3.Caption = "Citipak V2.06"
  Label2.Caption = "UB Sales by Year Export"
  Label1.Caption = "Scanning billing years. . ."
  DoEvents
  Call LoadLoadYears
  Form1.Show
End Sub

Private Sub lblCompanyProduct_Click()

End Sub

