VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arVASubTaxJrnlDetAll 
   BorderStyle     =   0  'None
   ClientHeight    =   6195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   14023
   _ExtentY        =   10927
   SectionData     =   "arVASubTaxJrnlDetAll.dsx":0000
End
Attribute VB_Name = "arVASubTaxJrnlDetAll"
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
  Open StartPath & "\TAXRPTS\SUBTXJRLDET.RPT" For Input As #hFile
  Fields.Add ("fldDesc") '0)
  Fields.Add ("fldYear") '1)
  Fields.Add ("fldAmt") '2)
  Fields.Add ("fldTransCnt") '3)
  Fields.Add ("fldTotByYrAndPers") '4)
  Fields.Add ("fldTotByYrAndInt") '5)
  Fields.Add ("fldTotByYrAndAdv") '6)
  Fields.Add ("fldTotByYrAndLateList") '7)
  Fields.Add ("fldTotByYrAndOpt1") '8)
  Fields.Add ("fldTotByYrAndOpt2") '9)
  Fields.Add ("fldTotByYrAndOpt3") '10)
  Fields.Add ("fldBillType") '11)
  Fields.Add ("fldOpt1Desc") '12)
  Fields.Add ("fldOpt2Desc") '13)
  Fields.Add ("fldOpt3Desc") '14)
  Fields.Add ("fldTotByYrAndMT") '15)
  Fields.Add ("fldTotByYrAndMC") '16)
  Fields.Add ("fldTotByYrAndFE") '17)
  Fields.Add ("fldTotByYrAndMH") '18)
  Fields.Add ("fldTotByYrAndPen") '19)
  Fields.Add ("fldYearCnt") '20)
  Fields.Add ("fldYearAmt") '21)
  Fields.Add ("fldAllYN") '22)
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
  Fields("fldTransCnt").Value = arr(3)
  Fields("fldTotByYrAndPers").Value = arr(4)
  Fields("fldTotByYrAndInt").Value = arr(5)
  Fields("fldTotByYrAndAdv").Value = arr(6)
  Fields("fldTotByYrAndLateList").Value = arr(7)
  Fields("fldTotByYrAndOpt1").Value = arr(8)
  Fields("fldTotByYrAndOpt2").Value = arr(9)
  Fields("fldTotByYrAndOpt3").Value = arr(10)
  Fields("fldBillType").Value = arr(11)
  Fields("fldOpt1Desc").Value = arr(12)
  Fields("fldOpt2Desc").Value = arr(13)
  Fields("fldOpt3Desc").Value = arr(14)
  Fields("fldTotByYrAndMT").Value = arr(15)
  Fields("fldTotByYrAndMC").Value = arr(16)
  Fields("fldTotByYrAndFE").Value = arr(17)
  Fields("fldTotByYrAndMH").Value = arr(18)
  Fields("fldTotByYrAndPen").Value = arr(19)
  Fields("fldYearCnt").Value = arr(20)
  Fields("fldYearAmt").Value = arr(21)
  Fields("fldAllYN").Value = arr(22)
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
  Dim Bill As Boolean
  
  Bill = False
  Field13.Visible = False 'opt1
  Field10.Visible = False
  Field14.Visible = False 'opt2
  Field11.Visible = False
  Field15.Visible = False 'opt3
  Field12.Visible = False
'  Field13.Top = 2430 'opt1
'  Field10.Top = 2430
'  Field14.Top = 2700 'opt2
'  Field11.Top = 2700
'  Field15.Top = 2970 'opt3
'  Field12.Top = 2970
  
  If CInt(Fields("fldBillType").Value) = 44 Then
    Detail.Height = 0
    Label2.Visible = False 'pers
    Field19.Visible = False
    Label38.Visible = False 'mt
    Field23.Visible = False
    Label39.Visible = False 'mc
    Field24.Visible = False
    Label40.Visible = False 'fe
    Field25.Visible = False
    Label41.Visible = False 'mh
    Field26.Visible = False
    Label3.Visible = False 'int
    Field20.Visible = False
    Label4.Visible = False 'adv
    Field21.Visible = False
    Label5.Visible = False 'latelist
    Field22.Visible = False
    Label42.Visible = False 'penalty
    Field27.Visible = False
    Exit Sub
  End If
  
  Select Case CInt(Fields("fldBillType").Value)
    Case 1: 'billing
      Detail.Height = 3360 '7/11/06 put back int, adv and pen fields to allow for manual bill revs
'      Detail.Height = 2240
'      Label3.Visible = False 'int
'      Field20.Visible = False
'      Label4.Visible = False 'adv
'      Field21.Visible = False
'      Label42.Visible = False 'pen
'      Field27.Visible = False
'      Label5.Visible = False 'll
'      Field22.Visible = False
'      Field13.Top = 1350 'opt1
'      Field10.Top = 1350
'      Field14.Top = 1620 'opt2
'      Field11.Top = 1620
'      Field15.Top = 1890 'opt3
'      Field12.Top = 1890
'      Bill = True
    Case 2:
      Detail.Height = 3360
    Case 3:
      Detail.Height = 3360
    Case 4:
      Detail.Height = 10
      Label2.Visible = False 'pers
      Field19.Visible = False
      Label38.Visible = False 'mt
      Field23.Visible = False
      Label39.Visible = False 'mc
      Field24.Visible = False
      Label40.Visible = False 'fe
      Field25.Visible = False
      Label41.Visible = False 'mh
      Field26.Visible = False
      Label3.Visible = False 'int
      Field20.Visible = False
      Label4.Visible = False 'adv
      Field21.Visible = False
      Label5.Visible = False 'll
      Field22.Visible = False
      Label42.Visible = False 'pen
      Field27.Visible = False
      Exit Sub
    Case 5:
      Detail.Height = 0
      Label2.Visible = False 'pers
      Field19.Visible = False
      Label38.Visible = False 'mt
      Field23.Visible = False
      Label39.Visible = False 'mc
      Field24.Visible = False
      Label40.Visible = False 'fe
      Field25.Visible = False
      Label41.Visible = False 'mh
      Field26.Visible = False
      Label3.Visible = False 'int
      Field20.Visible = False
      Label4.Visible = False 'adv
      Field21.Visible = False
      Label5.Visible = False 'll
      Field22.Visible = False
      Label42.Visible = False 'pen
      Field27.Visible = False
      Exit Sub
    Case 6:
      Detail.Height = 0
      Label2.Visible = False 'pers
      Field19.Visible = False
      Label38.Visible = False 'mt
      Field23.Visible = False
      Label39.Visible = False 'mc
      Field24.Visible = False
      Label40.Visible = False 'fe
      Field25.Visible = False
      Label41.Visible = False 'mh
      Field26.Visible = False
      Label3.Visible = False 'int
      Field20.Visible = False
      Label4.Visible = False 'adv
      Field21.Visible = False
      Label5.Visible = False 'll
      Field22.Visible = False
      Label42.Visible = False 'pen
      Field27.Visible = False
      Exit Sub
    Case 7:
      Detail.Height = 3360
    Case 8:
      Detail.Height = 0
      Label2.Visible = False 'pers
      Field19.Visible = False
      Label38.Visible = False 'mt
      Field23.Visible = False
      Label39.Visible = False 'mc
      Field24.Visible = False
      Label40.Visible = False 'fe
      Field25.Visible = False
      Label41.Visible = False 'mh
      Field26.Visible = False
      Label3.Visible = False 'int
      Field20.Visible = False
      Label4.Visible = False 'adv
      Field21.Visible = False
      Label5.Visible = False 'll
      Field22.Visible = False
      Label42.Visible = False 'pen
      Field27.Visible = False
      Exit Sub
    Case 9:
      Detail.Height = 3360
    Case 10:
      Detail.Height = 3360
    Case 11:
      Detail.Height = 3360
    Case 12:
      Detail.Height = 0
      Label2.Visible = False 'pers
      Field19.Visible = False
      Label38.Visible = False 'mt
      Field23.Visible = False
      Label39.Visible = False 'mc
      Field24.Visible = False
      Label40.Visible = False 'fe
      Field25.Visible = False
      Label41.Visible = False 'mh
      Field26.Visible = False
      Label3.Visible = False 'int
      Field20.Visible = False
      Label4.Visible = False 'adv
      Field21.Visible = False
      Label5.Visible = False 'll
      Field22.Visible = False
      Label42.Visible = False 'pen
      Field27.Visible = False
      Exit Sub
    Case 13:
      Detail.Height = 3360
    Case 14:
      Detail.Height = 3360
    Case 15:
      Detail.Height = 0
      Label2.Visible = False 'pers
      Field19.Visible = False
      Label38.Visible = False 'mt
      Field23.Visible = False
      Label39.Visible = False 'mc
      Field24.Visible = False
      Label40.Visible = False 'fe
      Field25.Visible = False
      Label41.Visible = False 'mh
      Field26.Visible = False
      Label3.Visible = False 'int
      Field20.Visible = False
      Label4.Visible = False 'adv
      Field21.Visible = False
      Label5.Visible = False 'll
      Field22.Visible = False
      Label42.Visible = False 'pen
      Field27.Visible = False
      Exit Sub
    Case 16:
      Detail.Height = 0
      Label2.Visible = False 'pers
      Field19.Visible = False
      Label38.Visible = False 'mt
      Field23.Visible = False
      Label39.Visible = False 'mc
      Field24.Visible = False
      Label40.Visible = False 'fe
      Field25.Visible = False
      Label41.Visible = False 'mh
      Field26.Visible = False
      Label3.Visible = False 'int
      Field20.Visible = False
      Label4.Visible = False 'adv
      Field21.Visible = False
      Label5.Visible = False 'll
      Field22.Visible = False
      Label42.Visible = False 'pen
      Field27.Visible = False
      Exit Sub
    Case 17:
      Detail.Height = 0
      Label2.Visible = False 'pers
      Field19.Visible = False
      Label38.Visible = False 'mt
      Field23.Visible = False
      Label39.Visible = False 'mc
      Field24.Visible = False
      Label40.Visible = False 'fe
      Field25.Visible = False
      Label41.Visible = False 'mh
      Field26.Visible = False
      Label3.Visible = False 'int
      Field20.Visible = False
      Label4.Visible = False 'adv
      Field21.Visible = False
      Label5.Visible = False 'll
      Field22.Visible = False
      Label42.Visible = False 'pen
      Field27.Visible = False
      Exit Sub
    Case Else:
      Detail.Height = 0
      Label2.Visible = False 'pers
      Field19.Visible = False
      Label38.Visible = False 'mt
      Field23.Visible = False
      Label39.Visible = False 'mc
      Field24.Visible = False
      Label40.Visible = False 'fe
      Field25.Visible = False
      Label41.Visible = False 'mh
      Field26.Visible = False
      Label3.Visible = False 'int
      Field20.Visible = False
      Label4.Visible = False 'adv
      Field21.Visible = False
      Label5.Visible = False 'll
      Field22.Visible = False
      Label42.Visible = False 'pen
      Field27.Visible = False
      Exit Sub
  End Select
  
  Opt1 = False
  Opt2 = False
  Opt3 = False
  
  If QPTrim$(Fields("fldOpt1Desc").Value) <> "" Then Opt1 = True
  If QPTrim$(Fields("fldOpt2Desc").Value) <> "" Then Opt2 = True
  If QPTrim$(Fields("fldOpt3Desc").Value) <> "" Then Opt3 = True
      
  If Opt1 = True And Opt2 = True And Opt3 = True Then
    Field13.Visible = True 'opt1
    Field10.Visible = True
    Field14.Visible = True 'opt2
    Field11.Visible = True
    Field15.Visible = True 'opt3
    Field12.Visible = True
    If Bill = True Then
      Field13.Top = 1350
      Field10.Top = 1350
      Field14.Top = 1620
      Field11.Top = 1620
      Field15.Top = 1890
      Field12.Top = 1890
      Detail.Height = 2280
    Else
      Detail.Height = 3360
    End If
  ElseIf Opt1 = True And Opt2 = False And Opt3 = False Then
    Field13.Visible = True 'opt1
    Field10.Visible = True
    If Bill = False Then
      Detail.Height = 2820
    Else
      Field13.Top = 1350
      Field10.Top = 1350
      Detail.Height = 1740
    End If
  ElseIf Opt1 = True And Opt2 = True And Opt3 = False Then
    Field13.Visible = True
    Field10.Visible = True
    Field14.Visible = True
    Field11.Visible = True
    If Bill = False Then
      Detail.Height = 3090
    Else
      Field13.Top = 1350
      Field10.Top = 1350
      Field14.Top = 1620
      Field11.Top = 1620
      Detail.Height = 2010
    End If
  ElseIf Opt1 = True And Opt2 = False And Opt3 = True Then
    Field13.Visible = True
    Field10.Visible = True
    Field15.Visible = True
    Field12.Visible = True
    If Bill = False Then
      Field15.Top = 2700
      Field12.Top = 2700
      Detail.Height = 3090
    Else
      Field13.Top = 1350
      Field10.Top = 1350
      Field15.Top = 1620
      Field12.Top = 1620
      Detail.Height = 2010
    End If
  ElseIf Opt1 = False And Opt2 = True And Opt3 = False Then
    Field14.Visible = True
    Field11.Visible = True
    If Bill = False Then
      Field14.Top = 2430
      Field11.Top = 2430
      Detail.Height = 2820
    Else
      Field14.Top = 1350
      Field11.Top = 1350
      Detail.Height = 1740
    End If
  ElseIf Opt1 = False And Opt2 = True And Opt3 = True Then
    Field11.Visible = True
    Field14.Visible = True
    Field12.Visible = True
    Field15.Visible = True
    If Bill = False Then
      Field11.Top = 2430
      Field14.Top = 2430
      Field12.Top = 2700
      Field15.Top = 2700
      Detail.Height = 3090
    Else
      Field11.Top = 1350
      Field14.Top = 1350
      Field12.Top = 1620
      Field15.Top = 1620
      Detail.Height = 1970
    End If
  ElseIf Opt1 = False And Opt2 = False And Opt3 = True Then
    Field12.Visible = True
    Field15.Visible = True
    If Bill = False Then
      Field12.Top = 2430
      Field15.Top = 2430
      Detail.Height = 2820
    Else
      Field12.Top = 1350
      Field15.Top = 1350
      Detail.Height = 1740
    End If
  ElseIf Opt1 = False And Opt2 = False And Opt3 = False Then
    If Bill = False Then
      Detail.Height = 2550
    Else
      Detail.Height = 1740
    End If
  End If

End Sub

Private Sub GroupHeader1_Format()
  If CInt(Fields("fldBillType").Value) = 44 Then
    Field30.Visible = False
    Field31.Visible = False
    Field32.Visible = False
  End If

End Sub

'Private Sub GroupFooter1_Format()
'  If Fields("fldAllYN").Value = "Y" Then
'    GroupFooter1.Height = 500
'    Label43.Visible = True
'    Field28.Visible = True
'    Field29.Visible = True
'    Line2.Y1 = 270
'    Line2.Y2 = 270
'  Else
'    GroupFooter1.Height = 20
'    Label43.Visible = False
'    Field28.Visible = False
'    Field29.Visible = False
'    Line2.Y1 = 10
'    Line2.Y2 = 10
'  End If
'End Sub

Private Sub GroupHeader2_Format()
  If Fields("fldBillType").Value <> 44 Then
    GroupHeader2.Height = 0
    Field16.Visible = False
    Field17.Visible = False
    Field18.Visible = False
  Else
    GroupHeader2.Height = 300
    Field16.Visible = True
    Field17.Visible = True
    Field18.Visible = True
  End If

End Sub
