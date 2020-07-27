VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2205
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   2205
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Fix It."
      Height          =   615
      Left            =   1680
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If Command1.Caption = "Done" Then
  End
End If
    Call FixEmUp
End Sub

Private Sub FixEmUp()
  Dim Reclen As Integer
  ', GJFile As Integer,
  Dim NumTran As Long
  Dim cnt As Long, NumRecs As Long, BadRecs As Long
  Dim APV As VendorRecType, VenLen As Integer
  VenLen = Len(APV)
  'Dim tstDate As Integer
  'Dim GJRec(1) As TrEditRecType
  'tstDate = Date2Num("07/01/2017")
  Open "apvendor.dat" For Random As #1 Len = VenLen
  Get #1, 48, APV
  APV.vnum = "48"
  Put #1, 48, APV
  Close
  Command1.Caption = "Done"
'  Dim ApLedger As APLedger81RecType
'  Reclen = Len(ApLedger)
'  Open "APLEDGER.DAT" For Random Shared As 1 Len = Reclen
'  NumTran& = LOF(1) \ Reclen
'  For cnt = 1 To NumTran
'    Get #1, cnt, ApLedger
'    If Abs(ApLedger.TRCode) = 4 Then
'      Print cnt
'      BadRecs = BadRecs + 1
'    End If
'  Next
'  Print BadRecs
'  Close
  
  
'  NumRecs = 0
'  GJReclen = Len(GJRec(1))
'  GJFile = FreeFile
'  Open "GJEDIT.DAT" For Random Access Read Write Shared As GJFile Len = GJReclen
'  NumEdTrans = LOF(GJFile) \ GJReclen
'  For cnt = 1 To NumEdTrans
'    Get GJFile, cnt, GJRec(1)
'
'    If GJRec(1).TRDATE < tstDate Then
'       GJRec(1).Deleted = -1
'       Put GJFile, cnt, GJRec(1)
'       BadRecs = BadRecs + 1
'    End If
'  Next
'  Close


End Sub

Public Function Date2Num%(txtDate$)
  On Error GoTo BadDate2Num
  If Len(QPTrim$(txtDate$)) = 10 Then
    Date2Num% = DateDiff("d", "12/31/1979", txtDate$)
  Else
    Date2Num% = -32767
  End If
  Exit Function

BadDate2Num:
  On Error GoTo 0
  Date2Num% = -32767
End Function

Public Function Num2Date$(intDate%)
  On Error GoTo BadNum2Date
  If intDate% = -32767 Then
    Num2Date$ = ""
  Else
    Num2Date$ = Format(DateAdd("d", (intDate%), "12-31-1979"), "mm/dd/yyyy")
  End If
  Exit Function
BadNum2Date:
  On Error GoTo 0
  Num2Date = ""
End Function


Public Function QPTrim$(Text As String)
  'Dim CPos As Long
  Dim StrLen As Long
  Dim cnt As Long
  Dim ThisChar As Integer
  StrLen = Len(Text)
  For cnt = 1 To StrLen
    ThisChar = Asc(Mid$(Text, cnt, 1))
    If ThisChar = 0 Then
      Mid$(Text$, cnt, 1) = " "
    End If
  Next
  QPTrim$ = Trim$(Text)
End Function

