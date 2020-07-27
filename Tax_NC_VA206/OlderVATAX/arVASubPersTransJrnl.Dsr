VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arVASubPersTransJrnl 
   Caption         =   "ActiveReport1"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9495
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   16748
   _ExtentY        =   10001
   SectionData     =   "arVASubPersTransJrnl.dsx":0000
End
Attribute VB_Name = "arVASubPersTransJrnl"
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
  Open StartPath & "\TAXRPTS\SUBTAXJRNLP.RPT" For Input As #hFile
  Fields.Add ("fldDesc") '0)
  Fields.Add ("fldYear") '1)
  Fields.Add ("fldAmt") '2)
  Fields.Add ("fldCnt") '3)
  Fields.Add ("fldPers") '4)
  Fields.Add ("fldInt") '5)
  Fields.Add ("fldAdv") '6)
  Fields.Add ("fldLateList") '7)
  Fields.Add ("fldOpt1") '8)
  Fields.Add ("fldOpt2") '9)
  Fields.Add ("fldOpt3") '10)
  Fields.Add ("fldType") '11)
  Fields.Add ("fldOpt1Desc") '12)
  Fields.Add ("fldOpt2Desc") '13)
  Fields.Add ("fldOpt3Desc") '14)
  Fields.Add ("fldPen") '15)
  Fields.Add ("fldMT") '16)
  Fields.Add ("fldMC") '17)
  Fields.Add ("fldFE") '18)
  Fields.Add ("fldMH") '19)
  
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

  If VBA.eof(hFile) Then
    eof = True
    Exit Sub
  Else
    eof = False
  End If
  Line Input #hFile, sLine
  arr = Split(sLine, "~")
  ' Here we set the values of the fields that we define as unbound
  ' or user defined.
  Fields("fldDesc").Value = arr(0)
  Fields("fldYear").Value = arr(1)
  Fields("fldAmt").Value = arr(2)
  Fields("fldCnt").Value = arr(3)
  Fields("fldPers").Value = arr(4)
  Fields("fldInt").Value = arr(5)
  Fields("fldAdv").Value = arr(6)
  Fields("fldLateList").Value = arr(7)
  Fields("fldOpt1").Value = arr(8)
  Fields("fldOpt2").Value = arr(9)
  Fields("fldOpt3").Value = arr(10)
  Fields("fldType").Value = arr(11)
  Fields("fldOpt1Desc").Value = arr(12)
  Fields("fldOpt2Desc").Value = arr(13)
  Fields("fldOpt3Desc").Value = arr(14)
  Fields("fldPen").Value = arr(15)
  Fields("fldMT").Value = arr(16)
  Fields("fldMC").Value = arr(17)
  Fields("fldFE").Value = arr(18)
  Fields("fldMH").Value = arr(19)
End Sub

Private Sub ActiveReport_ReportEnd()
  If hFile <> 0 Then
    Close #hFile
  End If
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  'Set Temp_Class = New Resize_Class
  'Temp_Class.InitResizeClass Me

End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    'Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If
End Sub

Private Sub Detail_Format()
  Dim Opt1 As Boolean
  Dim Opt2 As Boolean
  Dim Opt3 As Boolean
  
  Field12.Visible = False 'opt 1
  Field9.Visible = False 'opt 1
  Field13.Visible = False 'opt 2
  Field10.Visible = False 'opt 2
  Field14.Visible = False 'opt 3
  Field11.Visible = False 'opt 3

  If CInt(Fields("fldType").Value) = 44 Then
    Detail.Height = 0
    Label1.Visible = False 'Pers
    Field18.Visible = False 'Pers
    Label39.Visible = False 'MT
    Field22.Visible = False 'MT
    Label40.Visible = False 'MC
    Field23.Visible = False 'MC
    Label41.Visible = False 'FE
    Field24.Visible = False 'FE
    Label42.Visible = False 'MH
    Field25.Visible = False 'MH
    Label2.Visible = False 'Int
    Field19.Visible = False 'Int
    Label3.Visible = False 'Adv
    Field20.Visible = False 'Adv
    Label4.Visible = False 'Late List
    Field8.Visible = False 'Late List
    Label38.Visible = False 'Penalty
    Field21.Visible = False 'Penalty
    Exit Sub
  End If

  Select Case CInt(Fields("fldType").Value)
    Case 1:
      Detail.Height = 2706
    Case 2:
      Detail.Height = 3246
    Case 3:
      Detail.Height = 3246
    Case 4:
      Detail.Height = 10
      Label1.Visible = False 'pers
      Field18.Visible = False 'pers
      Label2.Visible = False 'Int
      Field19.Visible = False 'Int
      Label3.Visible = False 'Adv
      Field20.Visible = False 'Adv
      Label4.Visible = False 'LL
      Field8.Visible = False 'LL
      Label39.Visible = False 'MT
      Field22.Visible = False 'MT
      Label40.Visible = False 'MC
      Field23.Visible = False 'MC
      Label41.Visible = False 'FE
      Field24.Visible = False 'FE
      Label42.Visible = False 'MH
      Field25.Visible = False 'MH
      Exit Sub
    Case 5:
      Detail.Height = 0
      Label1.Visible = False 'pers
      Field18.Visible = False 'pers
      Label2.Visible = False 'Int
      Field19.Visible = False 'Int
      Label3.Visible = False 'Adv
      Field20.Visible = False 'Adv
      Label4.Visible = False 'LL
      Field8.Visible = False 'LL
      Label39.Visible = False 'MT
      Field22.Visible = False 'MT
      Label40.Visible = False 'MC
      Field23.Visible = False 'MC
      Label41.Visible = False 'FE
      Field24.Visible = False 'FE
      Label42.Visible = False 'MH
      Field25.Visible = False 'MH
      Exit Sub
    Case 6:
      Detail.Height = 0
      Label1.Visible = False 'pers
      Field18.Visible = False 'pers
      Label2.Visible = False 'Int
      Field19.Visible = False 'Int
      Label3.Visible = False 'Adv
      Field20.Visible = False 'Adv
      Label4.Visible = False 'LL
      Field8.Visible = False 'LL
      Label39.Visible = False 'MT
      Field22.Visible = False 'MT
      Label40.Visible = False 'MC
      Field23.Visible = False 'MC
      Label41.Visible = False 'FE
      Field24.Visible = False 'FE
      Label42.Visible = False 'MH
      Field25.Visible = False 'MH
      Exit Sub
    Case 7:
      Detail.Height = 3246
    Case 8:
      Detail.Height = 0
      Label1.Visible = False 'pers
      Field18.Visible = False 'pers
      Label2.Visible = False 'Int
      Field19.Visible = False 'Int
      Label3.Visible = False 'Adv
      Field20.Visible = False 'Adv
      Label4.Visible = False 'LL
      Field8.Visible = False 'LL
      Label39.Visible = False 'MT
      Field22.Visible = False 'MT
      Label40.Visible = False 'MC
      Field23.Visible = False 'MC
      Label41.Visible = False 'FE
      Field24.Visible = False 'FE
      Label42.Visible = False 'MH
      Field25.Visible = False 'MH
      Exit Sub
    Case 9:
      Detail.Height = 3246
    Case 10:
      Detail.Height = 3246
    Case 11:
      Detail.Height = 3246
    Case 12:
      Detail.Height = 0
      Label1.Visible = False 'pers
      Field18.Visible = False 'pers
      Label2.Visible = False 'Int
      Field19.Visible = False 'Int
      Label3.Visible = False 'Adv
      Field20.Visible = False 'Adv
      Label4.Visible = False 'LL
      Field8.Visible = False 'LL
      Label39.Visible = False 'MT
      Field22.Visible = False 'MT
      Label40.Visible = False 'MC
      Field23.Visible = False 'MC
      Label41.Visible = False 'FE
      Field24.Visible = False 'FE
      Label42.Visible = False 'MH
      Field25.Visible = False 'MH
      Exit Sub
    Case 13:
      Detail.Height = 3246
    Case 14:
      Detail.Height = 3246
    Case 15:
      Detail.Height = 0
      Label1.Visible = False 'pers
      Field18.Visible = False 'pers
      Label2.Visible = False 'Int
      Field19.Visible = False 'Int
      Label3.Visible = False 'Adv
      Field20.Visible = False 'Adv
      Label4.Visible = False 'LL
      Field8.Visible = False 'LL
      Label39.Visible = False 'MT
      Field22.Visible = False 'MT
      Label40.Visible = False 'MC
      Field23.Visible = False 'MC
      Label41.Visible = False 'FE
      Field24.Visible = False 'FE
      Label42.Visible = False 'MH
      Field25.Visible = False 'MH
      Exit Sub
    Case 16:
      Detail.Height = 0
      Label1.Visible = False 'pers
      Field18.Visible = False 'pers
      Label2.Visible = False 'Int
      Field19.Visible = False 'Int
      Label3.Visible = False 'Adv
      Field20.Visible = False 'Adv
      Label4.Visible = False 'LL
      Field8.Visible = False 'LL
      Label39.Visible = False 'MT
      Field22.Visible = False 'MT
      Label40.Visible = False 'MC
      Field23.Visible = False 'MC
      Label41.Visible = False 'FE
      Field24.Visible = False 'FE
      Label42.Visible = False 'MH
      Field25.Visible = False 'MH
      Exit Sub
    Case 17:
      Detail.Height = 0
      Label1.Visible = False 'pers
      Field18.Visible = False 'pers
      Label2.Visible = False 'Int
      Field19.Visible = False 'Int
      Label3.Visible = False 'Adv
      Field20.Visible = False 'Adv
      Label4.Visible = False 'LL
      Field8.Visible = False 'LL
      Label39.Visible = False 'MT
      Field22.Visible = False 'MT
      Label40.Visible = False 'MC
      Field23.Visible = False 'MC
      Label41.Visible = False 'FE
      Field24.Visible = False 'FE
      Label42.Visible = False 'MH
      Field25.Visible = False 'MH
      Exit Sub
    Case Else:
      Detail.Height = 0
      Label1.Visible = False 'pers
      Field18.Visible = False 'pers
      Label2.Visible = False 'Int
      Field19.Visible = False 'Int
      Label3.Visible = False 'Adv
      Field20.Visible = False 'Adv
      Label4.Visible = False 'LL
      Field8.Visible = False 'LL
      Label39.Visible = False 'MT
      Field22.Visible = False 'MT
      Label40.Visible = False 'MC
      Field23.Visible = False 'MC
      Label41.Visible = False 'FE
      Field24.Visible = False 'FE
      Label42.Visible = False 'MH
      Field25.Visible = False 'MH
      Exit Sub
  End Select

  Opt1 = False
  Opt2 = False
  Opt3 = False

  If QPTrim$(Fields("fldOpt1Desc").Value) <> "" Then Opt1 = True
  If QPTrim$(Fields("fldOpt2Desc").Value) <> "" Then Opt2 = True
  If QPTrim$(Fields("fldOpt3Desc").Value) <> "" Then Opt3 = True

  If Opt1 = True And Opt2 = True And Opt3 = True Then
    Field12.Visible = True
    Field9.Visible = True
    Field13.Visible = True
    Field10.Visible = True
    Field14.Visible = True
    Field11.Visible = True
    Detail.Height = 3240
  ElseIf Opt1 = True And Opt2 = False And Opt3 = False Then
    Field9.Visible = True
    Field12.Visible = True
    Detail.Height = 2700
  ElseIf Opt1 = True And Opt2 = True And Opt3 = False Then
    Field9.Visible = True
    Field12.Visible = True
    Field10.Visible = True
    Field13.Visible = True
    Detail.Height = 2970
  ElseIf Opt1 = True And Opt2 = False And Opt3 = True Then
    Field12.Visible = True
    Field9.Visible = True
    Field14.Visible = True
    Field11.Visible = True
    Field11.Top = 2700
    Field14.Top = 2700
    Detail.Height = 2970
  ElseIf Opt1 = False And Opt2 = True And Opt3 = False Then
    Field13.Visible = True
    Field10.Visible = True
    Field10.Top = 2430
    Field13.Top = 2430
    Detail.Height = 2700
  ElseIf Opt1 = False And Opt2 = True And Opt3 = True Then
    Field13.Visible = True
    Field10.Visible = True
    Field14.Visible = True
    Field11.Visible = True
    Field10.Top = 2430
    Field13.Top = 2430
    Field11.Top = 2700
    Field14.Top = 2700
    Detail.Height = 2970
  ElseIf Opt1 = False And Opt2 = False And Opt3 = True Then
    Field14.Visible = True
    Field11.Visible = True
    Field11.Top = 2430
    Field14.Top = 2430
    Detail.Height = 2700
  ElseIf Opt1 = False And Opt2 = False And Opt3 = False Then
    Detail.Height = 2430
  End If
  
End Sub

Private Sub GroupHeader1_Format()
  If CInt(Fields("fldType").Value) = 44 Then
    Field5.Visible = False
    Field7.Visible = False
    Field6.Visible = False
  End If
End Sub

Private Sub GroupHeader2_Format()
  If CInt(Fields("fldType").Value) <> 44 Then
    GroupHeader2.Height = 0
    Field15.Visible = False
    Field16.Visible = False
    Field17.Visible = False
  Else
   GroupHeader2.Height = 300
   Field15.Visible = True
   Field16.Visible = True
   Field17.Visible = True
  End If

End Sub


