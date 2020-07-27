VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arVASubTransJrnl 
   BorderStyle     =   0  'None
   Caption         =   "ActiveReport1"
   ClientHeight    =   6150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   16907
   _ExtentY        =   10848
   SectionData     =   "arVASubTransJrnl.dsx":0000
End
Attribute VB_Name = "arVASubTransJrnl"
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
  Open StartPath & "\TAXRPTS\SUBTAXJRNL.RPT" For Input As #hFile
  Fields.Add ("fldDesc") '0)
  Fields.Add ("fldYear") '1)
  Fields.Add ("fldAmt") '2)
  Fields.Add ("fldCnt") '3)
  Fields.Add ("fldPrinc") '4)
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
  Fields("fldPrinc").Value = arr(4)
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
  Dim Opt1 As Boolean
  Dim Opt2 As Boolean
  Dim Opt3 As Boolean
  
  Field9.Visible = False 'opt1
  Field12.Visible = False 'opt1
  Field10.Visible = False 'opt2
  Field13.Visible = False 'opt2
  Field11.Visible = False 'opt3
  Field14.Visible = False 'opt3
  
  If CInt(Fields("fldType").Value) = 44 Then
    Detail.Height = 0
    Label1.Visible = False 'principle
    Field18.Visible = False 'principle
    Label2.Visible = False 'interest
    Field19.Visible = False 'interest
    Label3.Visible = False 'advertising
    Field20.Visible = False 'advertising
    Label4.Visible = False 'late list
    Field8.Visible = False 'late list
    Label38.Visible = False 'penalty
    Field21.Visible = False 'penalty
    Exit Sub
  End If
  
  Select Case CInt(Fields("fldType").Value)
    Case 1: 'billing does not affect interest or advertising billing
'      Detail.Height = 1620
      Detail.Height = 2220 '7/11/06 put int, adv and pen fields back to accommodate manual bill revs
    Case 2:
      Detail.Height = 2220
    Case 3:
      Detail.Height = 2220
    Case 4:
      Detail.Height = 10
      Label1.Visible = False 'principle
      Field18.Visible = False 'principle
      Label2.Visible = False 'interest
      Field19.Visible = False 'interest
      Label3.Visible = False 'advertising
      Field20.Visible = False 'advertising
      Label4.Visible = False 'late list
      Field8.Visible = False 'late list
      Label38.Visible = False 'penalty
      Field21.Visible = False 'penalty
      Exit Sub
    Case 5:
      Detail.Height = 0
      Label1.Visible = False 'principle
      Field18.Visible = False 'principle
      Label2.Visible = False 'interest
      Field19.Visible = False 'interest
      Label3.Visible = False 'advertising
      Field20.Visible = False 'advertising
      Label4.Visible = False 'late list
      Field8.Visible = False 'late list
      Label38.Visible = False 'penalty
      Field21.Visible = False 'penalty
      Exit Sub
    Case 6:
      Detail.Height = 0
      Label1.Visible = False 'principle
      Field18.Visible = False 'principle
      Label2.Visible = False 'interest
      Field19.Visible = False 'interest
      Label3.Visible = False 'advertising
      Field20.Visible = False 'advertising
      Label4.Visible = False 'late list
      Field8.Visible = False 'late list
      Label38.Visible = False 'penalty
      Field21.Visible = False 'penalty
      Exit Sub
    Case 7:
      Detail.Height = 2220
    Case 8:
      Detail.Height = 0
      Label1.Visible = False 'principle
      Field18.Visible = False 'principle
      Label2.Visible = False 'interest
      Field19.Visible = False 'interest
      Label3.Visible = False 'advertising
      Field20.Visible = False 'advertising
      Label4.Visible = False 'late list
      Field8.Visible = False 'late list
      Label38.Visible = False 'penalty
      Field21.Visible = False 'penalty
      Exit Sub
    Case 9:
      Detail.Height = 2220
    Case 10:
      Detail.Height = 2220
    Case 11:
      Detail.Height = 2220
    Case 12:
      Detail.Height = 0
      Label1.Visible = False 'principle
      Field18.Visible = False 'principle
      Label2.Visible = False 'interest
      Field19.Visible = False 'interest
      Label3.Visible = False 'advertising
      Field20.Visible = False 'advertising
      Label4.Visible = False 'late list
      Field8.Visible = False 'late list
      Label38.Visible = False 'penalty
      Field21.Visible = False 'penalty
      Exit Sub
    Case 13:
      Detail.Height = 2220
    Case 14:
      Detail.Height = 2220
    Case 15:
      Detail.Height = 0
      Label1.Visible = False 'principle
      Field18.Visible = False 'principle
      Label2.Visible = False 'interest
      Field19.Visible = False 'interest
      Label3.Visible = False 'advertising
      Field20.Visible = False 'advertising
      Label4.Visible = False 'late list
      Field8.Visible = False 'late list
      Label38.Visible = False 'penalty
      Field21.Visible = False 'penalty
      Exit Sub
    Case 16:
      Detail.Height = 0
      Label1.Visible = False 'principle
      Field18.Visible = False 'principle
      Label2.Visible = False 'interest
      Field19.Visible = False 'interest
      Label3.Visible = False 'advertising
      Field20.Visible = False 'advertising
      Label4.Visible = False 'late list
      Field8.Visible = False 'late list
      Label38.Visible = False 'penalty
      Field21.Visible = False 'penalty
      Exit Sub
    Case 17:
      Detail.Height = 0
      Label1.Visible = False 'principle
      Field18.Visible = False 'principle
      Label2.Visible = False 'interest
      Field19.Visible = False 'interest
      Label3.Visible = False 'advertising
      Field20.Visible = False 'advertising
      Label4.Visible = False 'late list
      Field8.Visible = False 'late list
      Label38.Visible = False 'penalty
      Field21.Visible = False 'penalty
      Exit Sub
    Case Else:
      Detail.Height = 0
      Label1.Visible = False 'principle
      Field18.Visible = False 'principle
      Label2.Visible = False 'interest
      Field19.Visible = False 'interest
      Label3.Visible = False 'advertising
      Field20.Visible = False 'advertising
      Label4.Visible = False 'late list
      Field8.Visible = False 'late list
      Label38.Visible = False 'penalty
      Field21.Visible = False 'penalty
      Exit Sub
  End Select
  
  Opt1 = False
  Opt2 = False
  Opt3 = False
  
  If QPTrim$(Fields("fldOpt1Desc").Value) <> "" Then Opt1 = True
  If QPTrim$(Fields("fldOpt2Desc").Value) <> "" Then Opt2 = True
  If QPTrim$(Fields("fldOpt3Desc").Value) <> "" Then Opt3 = True
      
  If Opt1 = True And Opt2 = True And Opt3 = True Then
    Field9.Visible = True
    Field12.Visible = True
    Field10.Visible = True
    Field13.Visible = True
    Field11.Visible = True
    Field14.Visible = True
    Detail.Height = 2220
  ElseIf Opt1 = True And Opt2 = False And Opt3 = False Then
    Field9.Visible = True
    Field12.Visible = True
    Detail.Height = 1620
  ElseIf Opt1 = True And Opt2 = True And Opt3 = False Then
    Field9.Visible = True
    Field12.Visible = True
    Field10.Visible = True
    Field13.Visible = True
    Detail.Height = 1890
  ElseIf Opt1 = True And Opt2 = False And Opt3 = True Then
    Field9.Visible = True
    Field12.Visible = True
    Field11.Visible = True
    Field14.Visible = True
    Field11.Top = 1620
    Field14.Top = 1620
    Detail.Height = 1890
  ElseIf Opt1 = False And Opt2 = True And Opt3 = False Then
    Field10.Visible = True
    Field13.Visible = True
    Field10.Top = 1350
    Field13.Top = 1350
    Detail.Height = 1350
  ElseIf Opt1 = False And Opt2 = True And Opt3 = True Then
    Field10.Visible = True
    Field13.Visible = True
    Field11.Visible = True
    Field14.Visible = True
    Field10.Top = 1350
    Field13.Top = 1350
    Field11.Top = 1620
    Field14.Top = 1620
    Detail.Height = 1890
  ElseIf Opt1 = False And Opt2 = False And Opt3 = True Then
    Field11.Visible = True
    Field14.Visible = True
    Field11.Top = 1350
    Field14.Top = 1350
    Detail.Height = 1620
  ElseIf Opt1 = False And Opt2 = False And Opt3 = False Then
    Detail.Height = 1350
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

