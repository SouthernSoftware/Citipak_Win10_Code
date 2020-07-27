VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arVASubTaxPayEditJrnl2 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5208
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10575
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   _ExtentX        =   18653
   _ExtentY        =   9181
   SectionData     =   "arVASubTaxPayEditJrnl2.dsx":0000
End
Attribute VB_Name = "arVASubTaxPayEditJrnl2"
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
  Fields.Add ("fldType") '23)
  Fields.Add ("fldPenAmt") '24)
  Fields.Add ("fldGPenAmt") '25)
  Fields.Add ("fldOpt1") '26)
  Fields.Add ("fldOpt2") '27)
  Fields.Add ("fldOpt3") '28)
  Fields.Add ("fldGOpt1") '29)
  Fields.Add ("fldGOpt2") '30)
  Fields.Add ("fldGOpt3") '31)
  Fields.Add ("fldPOpt1Desc") '32)
  Fields.Add ("fldPOpt2Desc") '33)
  Fields.Add ("fldPOpt3Desc") '34)
End Sub

Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
    frmVATaxMsg.Label1.Caption = "Error Number: " & Str(Number) & " " & Description
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
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
  If arr(23) = "P" Then
    Fields("fldPayLessDisc").Value = OldRound(CDbl(arr(26)) + CDbl(arr(27)) + CDbl(arr(28)) + CDbl(arr(21)) + CDbl(arr(1)) + CDbl(arr(2)) + CDbl(arr(3)) + CDbl(arr(4)) + CDbl(arr(5)) + CDbl(arr(6)) + CDbl(arr(7)) + CDbl(arr(24)) - CDbl(arr(8)))
    Fields("fldGPayLessDisc").Value = OldRound(CDbl(arr(29)) + CDbl(arr(30)) + CDbl(arr(31)) + CDbl(arr(22)) + CDbl(arr(9)) + CDbl(arr(10)) + CDbl(arr(11)) + CDbl(arr(12)) + CDbl(arr(13)) + CDbl(arr(14)) + CDbl(arr(15)) + CDbl(arr(25)) - CDbl(arr(16)))
  ElseIf arr(23) = "R" Then
    Fields("fldPayLessDisc").Value = OldRound(CDbl(arr(21)) + CDbl(arr(1)) + CDbl(arr(2)) + CDbl(arr(3)) + CDbl(arr(4)) + CDbl(arr(5)) + CDbl(arr(6)) + CDbl(arr(7)) + CDbl(arr(24)) - CDbl(arr(8)))
    Fields("fldGPayLessDisc").Value = OldRound(CDbl(arr(22)) + CDbl(arr(9)) + CDbl(arr(10)) + CDbl(arr(11)) + CDbl(arr(12)) + CDbl(arr(13)) + CDbl(arr(14)) + CDbl(arr(15)) + CDbl(arr(25)) - CDbl(arr(16)))
  End If
  Fields("fldType").Value = arr(23)
  Fields("fldPenAmt").Value = arr(24)
  Fields("fldGPenAmt").Value = arr(25)
  Fields("fldOpt1").Value = arr(26)
  Fields("fldOpt2").Value = arr(27)
  Fields("fldOpt3").Value = arr(28)
  Fields("fldGOpt1").Value = arr(29)
  Fields("fldGOpt2").Value = arr(30)
  Fields("fldGOpt3").Value = arr(31)
  Fields("fldPOpt1Desc").Value = arr(32)
  Fields("fldPOpt2Desc").Value = arr(33)
  Fields("fldPOpt3Desc").Value = arr(34)
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

Private Sub Detail_Format()
  
  If QPTrim$(Fields("fldPOpt1Desc").Value) = "" And QPTrim$(Fields("fldPOpt2Desc").Value) = "" And QPTrim$(Fields("fldPOpt3Desc").Value) = "" Then
    Detail.Height = 636
    Line3.Y1 = 540
    Line3.Y2 = 540
    Field42.Visible = False
    Field43.Visible = False
    Field44.Visible = False
 
  ElseIf QPTrim$(Fields("fldPOpt1Desc").Value) <> "" Or QPTrim$(Fields("fldPOpt2Desc").Value) <> "" Or QPTrim$(Fields("fldPOpt3Desc").Value) <> "" Then
    Detail.Height = 924
    Line1.Y1 = 860
    Line1.Y2 = 860
  End If
  
  If Fields("fldType").Value = "P" Then
    If QPTrim$(Fields("fldPOpt1Desc").Value) <> "" Then
      Field42.Visible = True
    End If
    If QPTrim$(Fields("fldPOpt2Desc").Value) <> "" Then
      Field43.Visible = True
    End If
    If QPTrim$(Fields("fldPOpt3Desc").Value) <> "" Then
      Field44.Visible = True
    End If
  End If
  
End Sub

Private Sub GroupFooter1_Format()
  Field45.Visible = False
  Field46.Visible = False
  Field47.Visible = False
  If Fields("fldType").Value = "P" Then
    If QPTrim$(Fields("fldPOpt1Desc").Value) <> "" Or QPTrim$(Fields("fldPOpt2Desc").Value) <> "" Or QPTrim$(Fields("fldPOpt3Desc").Value) <> "" Then
      If QPTrim$(Fields("fldPOpt1Desc").Value) <> "" Then
        Field45.Visible = True
      End If
      If QPTrim$(Fields("fldPOpt2Desc").Value) <> "" Then
        Field46.Visible = True
      End If
      If QPTrim$(Fields("fldPOpt3Desc").Value) <> "" Then
        Field47.Visible = True
      End If
    End If
  End If
End Sub

Private Sub GroupHeader1_Format()
  Dim ThisDesc As String * 12
  GroupHeader1.Height = 948
  Line1.Y1 = 900
  Line1.Y2 = 900
  Label34.Visible = False
  Label35.Visible = False
  Label36.Visible = False
  If Fields("fldType").Value = "R" Then
    Label32.Caption = "Real Property"
  ElseIf Fields("fldType").Value = "P" Then
    Label32.Caption = "Personal Property"
    Label24.Caption = "Personal"
    Label25.Caption = "Mach Tools"
    Label26.Caption = "Merch Cap"
    Label27.Caption = "Farm Equip"
    If QPTrim$(Fields("fldPOpt1Desc").Value) <> "" Or QPTrim$(Fields("fldPOpt2Desc").Value) <> "" Or QPTrim$(Fields("fldPOpt3Desc").Value) <> "" Then
      GroupHeader1.Height = 1152
      Line1.Y1 = 1100
      Line1.Y2 = 1100
    End If
    If QPTrim$(Fields("fldPOpt1Desc").Value) <> "" Then
      RSet ThisDesc = QPTrim$(Fields("fldPOpt1Desc").Value)
      Label34.Caption = ThisDesc
      Label34.Visible = True
    End If
    If QPTrim$(Fields("fldPOpt2Desc").Value) <> "" Then
      RSet ThisDesc = QPTrim$(Fields("fldPOpt2Desc").Value)
      Label35.Caption = ThisDesc
      Label35.Visible = True
    End If
    If QPTrim$(Fields("fldPOpt3Desc").Value) <> "" Then
      RSet ThisDesc = QPTrim$(Fields("fldPOpt3Desc").Value)
      Label36.Caption = ThisDesc
      Label36.Visible = True
    End If
  End If
End Sub
