VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arSubTaxPayEditJrnl2 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9015
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   _ExtentX        =   15901
   _ExtentY        =   7938
   SectionData     =   "arSubTaxPayEditJrnl2.dsx":0000
End
Attribute VB_Name = "arSubTaxPayEditJrnl2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Private hFile As Integer

Private Sub ActiveReport_DataInitialize()
  Dim x As Integer
  hFile = FreeFile
  Open StartPath & "\TAXRPTS\SubEdPay2.RPT" For Input As #hFile
  Fields.Add ("fldYear") '0)
  Fields.Add ("fldPrinc") '1)
  Fields.Add ("fldInt") '2)
  Fields.Add ("fldAdvCol") '3)
  Fields.Add ("fldLateList") '4)
  Fields.Add ("fldRev1") '5)
  Fields.Add ("fldRev2") '6)
  Fields.Add ("fldRev3") '7)
  Fields.Add ("fldDisc") '8)
  Fields.Add ("fldGPrinc") '9)
  Fields.Add ("fldGInt") '10)
  Fields.Add ("fldGAdvCol") '11)
  Fields.Add ("fldGLateList") '12)
  Fields.Add ("fldGRev1") '13)
  Fields.Add ("fldGRev2") '14)
  Fields.Add ("fldGRev3") '15)
  Fields.Add ("fldGDisc") '16)
  Fields.Add ("fldRevDesc1") '17)
  Fields.Add ("fldRevDesc2") '18)
  Fields.Add ("fldRevDesc3") '19)
  Fields.Add ("fldPayLessDisc")
  Fields.Add ("fldGPayLessDisc")
  Fields.Add ("fldOut") '20)
  Fields.Add ("fldOverPay") '21)
  Fields.Add ("fldGOverPay") '22)
End Sub

Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
    frmTaxMsg.Label1.Caption = "Error Number: " & Str(Number) & " " & Description
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    Unload Me
  End If
  CancelDisplay = True 'removes the error message

End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)
  Dim sLine As String
  Dim arr() As String
  Dim RevTruncate As String * 12
  If VBA.eof(hFile) Then
    eof = True
    Exit Sub
  Else
    eof = False
  End If
  Line Input #hFile, sLine
  arr = Split(sLine, "~")
  Fields("fldYear").Value = arr(0)
  Fields("fldPrinc").Value = arr(1)
  Fields("fldInt").Value = arr(2)
  Fields("fldAdvCol").Value = arr(3)
  Fields("fldLateList").Value = arr(4)
  Fields("fldRev1").Value = arr(5)
  Fields("fldRev2").Value = arr(6)
  Fields("fldRev3").Value = arr(7)
  Fields("fldDisc").Value = arr(8)
  Fields("fldGPrinc").Value = arr(9)
  Fields("fldGInt").Value = arr(10)
  Fields("fldGAdvCol").Value = arr(11)
  Fields("fldGLateList").Value = arr(12)
  Fields("fldGRev1").Value = arr(13)
  Fields("fldGRev2").Value = arr(14)
  Fields("fldGRev3").Value = arr(15)
  Fields("fldGDisc").Value = arr(16)
  RevTruncate = arr(17)
  Fields("fldRevDesc1").Value = QPTrim$(RevTruncate)
  RevTruncate = arr(18)
  Fields("fldRevDesc2").Value = QPTrim$(RevTruncate)
  RevTruncate = arr(19)
  Fields("fldRevDesc3").Value = QPTrim$(RevTruncate)
  Fields("fldOut").Value = arr(20)
  If arr(20) = "True" Then
    Line3.Visible = False
  Else
    Line3.Visible = True
  End If
  Fields("fldOverPay").Value = arr(21)
  Fields("fldGOverPay").Value = arr(22)
  Fields("fldPayLessDisc").Value = OldRound(CDbl(arr(21)) + CDbl(arr(1)) + CDbl(arr(2)) + CDbl(arr(3)) + CDbl(arr(4)) + CDbl(arr(5)) + CDbl(arr(6)) + CDbl(arr(7)) - CDbl(arr(8)))
  Fields("fldGPayLessDisc").Value = OldRound(CDbl(arr(22)) + CDbl(arr(9)) + CDbl(arr(10)) + CDbl(arr(11)) + CDbl(arr(12)) + CDbl(arr(13)) + CDbl(arr(14)) + CDbl(arr(15)) - CDbl(arr(16)))
  
End Sub

Private Sub ActiveReport_ReportEnd()
  If hFile <> 0 Then
    Close #hFile
  End If
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If
End Sub
