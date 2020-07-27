VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arSub2TaxJrnlDetAll 
   BorderStyle     =   0  'None
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   17198
   _ExtentY        =   10213
   SectionData     =   "arVASub2TaxJrnlDetAll.dsx":0000
End
Attribute VB_Name = "arSub2TaxJrnlDetAll"
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
  Open StartPath & "\TAXRPTS\SUB2TXJRLDET.RPT" For Input As #hFile
  Fields.Add ("fldTotAmt") '0)
  Fields.Add ("fldTotCnt") '1)
  Fields.Add ("fldGTotAmt") '2)
  Fields.Add ("fldGTotCnt") '3)
  Fields.Add ("fldGPersTot") '4)
  Fields.Add ("fldGIntTot") '5)
  Fields.Add ("fldGAdvTot") '6)
  Fields.Add ("fldGLateListTot") '7)
  Fields.Add ("fldGOpt1Tot") '8)
  Fields.Add ("fldGOpt2Tot") '9)
  Fields.Add ("fldGOpt3Tot") '10)
  Fields.Add ("fldOpt1Desc") '11)
  Fields.Add ("fldOpt2Desc") '12)
  Fields.Add ("fldOpt3Desc") '13)
  Fields.Add ("fldGMTTot") '14)
  Fields.Add ("fldGMCTot") '15)
  Fields.Add ("fldGFETot") '16)
  Fields.Add ("fldGMHTot") '17)
  Fields.Add ("fldGPenTot") '18)
  Fields.Add ("fldRange") '19)
  Fields.Add ("fldType") '20)
  Fields.Add ("fldShow") '21)
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
  Fields("fldTotAmt").Value = arr(0)
  Fields("fldTotCnt").Value = arr(1)
  Fields("fldGTotAmt").Value = arr(2)
  Fields("fldGTotCnt").Value = arr(3)
  Fields("fldGPersTot").Value = arr(4)
  Fields("fldGIntTot").Value = arr(5)
  Fields("fldGAdvTot").Value = arr(6)
  Fields("fldGLateListTot").Value = arr(7)
  Fields("fldGOpt1Tot").Value = arr(8)
  Fields("fldGOpt2Tot").Value = arr(9)
  Fields("fldGOpt3Tot").Value = arr(10)
  Fields("fldOpt1Desc").Value = arr(11)
  Fields("fldOpt2Desc").Value = arr(12)
  Fields("fldOpt3Desc").Value = arr(13)
  Fields("fldGMTTot").Value = arr(14)
  Fields("fldGMCTot").Value = arr(15)
  Fields("fldGFETot").Value = arr(16)
  Fields("fldGMHTot").Value = arr(17)
  Fields("fldGPenTot").Value = arr(18)
  Fields("fldRange").Value = arr(19)
  Fields("fldType").Value = arr(20)
  Fields("fldShow").Value = arr(21)
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
Private Sub GroupFooter1_Format()
  Dim Opt1 As Boolean
  Dim Opt2 As Boolean
  Dim Opt3 As Boolean
  Dim Bill As Boolean
  
  Bill = False
  Label11.Top = 2160 'penalty
  Field21.Top = 2160
  Field7.Top = 2430 'opt1
  Field14.Top = 2430
  Field8.Top = 2700 'opt2
  Field15.Top = 2700
  Field9.Top = 2970 'opt3
  Field16.Top = 2970
  Line2.Y1 = 3240
  Line2.Y2 = 3240
  Label1.Top = 3330 'grand totals
  Field6.Top = 3330
  Field5.Top = 3330
  GroupFooter1.Height = 3744
  
  If CInt(Fields("fldRange").Value) = 2 Or CInt(Fields("fldShow").Value) = 44 Then
    Label2.Visible = False 'personal
    Field10.Visible = False
    Label7.Visible = False 'mt
    Field17.Visible = False
    Label8.Visible = False 'mc
    Field18.Visible = False
    Label9.Visible = False 'fe
    Field19.Visible = False
    Label10.Visible = False 'mh
    Field20.Visible = False
    Label3.Visible = False 'int
    Field11.Visible = False
    Label4.Visible = False 'adv
    Field12.Visible = False
    Label5.Visible = False 'latelist
    Field13.Visible = False
    Label11.Visible = False 'pen
    Field21.Visible = False
    Field7.Visible = False 'opt1
    Field14.Visible = False
    Field8.Visible = False 'opt2
    Field15.Visible = False
    Field9.Visible = False 'opt3
    Field16.Visible = False
    Line1.Visible = False
    GroupFooter1.Height = 624
    Line2.Y1 = 0
    Line2.Y2 = 0
    Label1.Top = 90
    Field6.Top = 90
    Field5.Top = 90
    Exit Sub
    '7/11/06 took out bills to allow revenues for manual billings
'  ElseIf Fields("fldType").Value = "Billing" Then
'    Bill = True
'    Label3.Visible = False 'int
'    Field11.Visible = False
'    Label4.Visible = False 'adv
'    Field12.Visible = False
'    Label11.Visible = False 'pen
'    Field21.Visible = False
'    Label5.Visible = False 'll
'    Field13.Visible = False
'    GroupFooter1.Height = 2664
'    Line2.Y1 = 2298
'    Line2.Y2 = 2298
'    Label1.Top = 2388
'    Field6.Top = 2388
'    Field5.Top = 2388
'    Field7.Top = 1350
'    Field14.Top = 1350
'    Field8.Top = 1620
'    Field15.Top = 1620
'    Field9.Top = 1890
'    Field16.Top = 1890
  End If
    
  Opt1 = False
  Opt2 = False
  Opt3 = False
  
  If QPTrim$(Fields("fldOpt1Desc").Value) <> "" Then
    Opt1 = True
  End If
  
  If QPTrim$(Fields("fldOpt2Desc").Value) <> "" Then
    Opt2 = True
  End If
  
  If QPTrim$(Fields("fldOpt3Desc").Value) <> "" Then
    Opt3 = True
  End If
  
  Field7.Visible = True 'opt1
  Field14.Visible = True
  Field8.Visible = True 'opt2
  Field15.Visible = True
  Field9.Visible = True 'opt3
  Field16.Visible = True
  If Bill = False Then
    Line2.Y1 = 3240
    Line2.Y2 = 3240
    Label1.Top = 3330
    Field6.Top = 3330
    Field5.Top = 3330
  Else
    Line2.Y1 = 2430
    Line2.Y2 = 2430
    Label1.Top = 2520
    Field6.Top = 2520
    Field5.Top = 2520
  End If
  
  If Opt1 = True And Opt2 = False And Opt3 = False Then
    Field8.Visible = False
    Field15.Visible = False
    Field9.Visible = False
    Field16.Visible = False
    If Bill = False Then
      Line2.Y1 = 2700
      Line2.Y2 = 2700
      Label1.Top = 2790
      Field6.Top = 2790
      Field5.Top = 2790
      GroupFooter1.Height = 3204
    Else
      Field7.Top = 1350
      Field14.Top = 1350
      Line2.Y1 = 1620
      Line2.Y2 = 1620
      Label1.Top = 1710
      Field6.Top = 1710
      Field5.Top = 1710
      GroupFooter1.Height = 1980
    End If
  ElseIf Opt1 = True And Opt2 = True And Opt3 = False Then
    Field9.Visible = False
    Field16.Visible = False
    If Bill = False Then
      Line2.Y1 = 2970
      Line2.Y2 = 2970
      Label1.Top = 3060
      Field6.Top = 3060
      Field5.Top = 3060
      GroupFooter1.Height = 3474
    Else
      Field7.Top = 1350
      Field14.Top = 1350
      Field8.Top = 1620
      Field15.Top = 1620
      Line2.Y1 = 1890
      Line2.Y2 = 1890
      Label1.Top = 1980
      Field6.Top = 1980
      Field5.Top = 1980
      GroupFooter1.Height = 2330
    End If
  ElseIf Opt1 = True And Opt2 = False And Opt3 = True Then
    Field8.Visible = False
    Field15.Visible = False
    If Bill = False Then
      Field9.Top = 2700
      Field16.Top = 2700
      Line2.Y1 = 2970
      Line2.Y2 = 2970
      Label1.Top = 3060
      Field6.Top = 3060
      Field5.Top = 3060
      GroupFooter1.Height = 3474
    Else
      Field7.Top = 1350
      Field14.Top = 1350
      Field9.Top = 1620
      Field16.Top = 1620
      Line2.Y1 = 1890
      Line2.Y2 = 1890
      Label1.Top = 1980
      Field6.Top = 1980
      Field5.Top = 1980
      GroupFooter1.Height = 2330
    End If
  ElseIf Opt1 = False And Opt2 = True And Opt3 = False Then
    Field7.Visible = False
    Field14.Visible = False
    Field9.Visible = False
    Field16.Visible = False
    If Bill = False Then
      Field8.Top = 2430
      Field15.Top = 2430
      Line2.Y1 = 2700
      Line2.Y2 = 2700
      Label1.Top = 2790
      Field6.Top = 2790
      Field5.Top = 2790
      GroupFooter1.Height = 3204
    Else
      Field8.Top = 1350
      Field15.Top = 1350
      Line2.Y1 = 1620
      Line2.Y2 = 1620
      Label1.Top = 1710
      Field6.Top = 1710
      Field5.Top = 1710
      GroupFooter1.Height = 2060
    End If
  ElseIf Opt1 = False And Opt2 = True And Opt3 = True Then
    Field7.Visible = False
    Field14.Visible = False
    If Bill = False Then
      Field8.Top = 2430
      Field15.Top = 2430
      Field9.Top = 2700
      Field16.Top = 2700
      Line2.Y1 = 2970
      Line2.Y2 = 2970
      Label1.Top = 3060
      Field6.Top = 3060
      Field5.Top = 3060
      GroupFooter1.Height = 3474
    Else
      Field8.Top = 1350
      Field15.Top = 1350
      Field9.Top = 1620
      Field16.Top = 1620
      Line2.Y1 = 1890
      Line2.Y2 = 1890
      Label1.Top = 1980
      Field6.Top = 1980
      Field5.Top = 1980
      GroupFooter1.Height = 2330
    End If
  ElseIf Opt1 = False And Opt2 = False And Opt3 = True Then
    Field7.Visible = False
    Field14.Visible = False
    Field8.Visible = False
    Field15.Visible = False
    If Bill = False Then
      Field9.Top = 2430
      Field16.Top = 2430
      Line2.Y1 = 2700
      Line2.Y2 = 2700
      Label1.Top = 2790
      Field6.Top = 2790
      Field5.Top = 2790
      GroupFooter1.Height = 3204
    Else
      Field9.Top = 1350
      Field16.Top = 1350
      Line2.Y1 = 1620
      Line2.Y2 = 1620
      Label1.Top = 1710
      Field6.Top = 1710
      Field5.Top = 1710
      GroupFooter1.Height = 2060
    End If
  ElseIf Opt1 = False And Opt2 = False And Opt3 = False Then
    Field7.Visible = False
    Field14.Visible = False
    Field8.Visible = False
    Field15.Visible = False
    Field9.Visible = False
    Field16.Visible = False
    If Bill = False Then
      Line2.Y1 = 2430
      Line2.Y2 = 2430
      Label1.Top = 2520
      Field6.Top = 2520
      Field5.Top = 2520
      GroupFooter1.Height = 2934
    Else
      Line2.Y1 = 1350
      Line2.Y2 = 1350
      Label1.Top = 1440
      Field6.Top = 1440
      Field5.Top = 1440
      GroupFooter1.Height = 1790
    End If
  End If

End Sub
