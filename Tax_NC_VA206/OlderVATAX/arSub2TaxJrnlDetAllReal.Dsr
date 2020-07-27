VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arSub2TaxJrnlDetAllReal 
   BorderStyle     =   0  'None
   ClientHeight    =   5835
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   12039
   _ExtentY        =   10292
   SectionData     =   "arSub2TaxJrnlDetAllReal.dsx":0000
End
Attribute VB_Name = "arSub2TaxJrnlDetAllReal"
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
  Open StartPath & "\TAXRPTS\SUB2TXJRLDETREAL.RPT" For Input As #hFile
  Fields.Add ("fldTotAmt") '0)
  Fields.Add ("fldTotCnt") '1)
  Fields.Add ("fldGTotAmt") '2)
  Fields.Add ("fldGTotCnt") '3)
  Fields.Add ("fldGPrincTot") '4)
  Fields.Add ("fldGIntTot") '5)
  Fields.Add ("fldGAdvTot") '6)
  Fields.Add ("fldGLateListTot") '7)
  Fields.Add ("fldGOpt1Tot") '8)
  Fields.Add ("fldGOpt2Tot") '9)
  Fields.Add ("fldGOpt3Tot") '10)
  Fields.Add ("fldOpt1Desc") '11)
  Fields.Add ("fldOpt2Desc") '12)
  Fields.Add ("fldOpt3Desc") '13)
  Fields.Add ("fldGPenTot") '14)
  Fields.Add ("fldRange") '15)
  Fields.Add ("fldType") '16)
  Fields.Add ("fldShow") '17)
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
  Fields("fldGPrincTot").Value = arr(4)
  Fields("fldGIntTot").Value = arr(5)
  Fields("fldGAdvTot").Value = arr(6)
  Fields("fldGLateListTot").Value = arr(7)
  Fields("fldGOpt1Tot").Value = arr(8)
  Fields("fldGOpt2Tot").Value = arr(9)
  Fields("fldGOpt3Tot").Value = arr(10)
  Fields("fldOpt1Desc").Value = arr(11)
  Fields("fldOpt2Desc").Value = arr(12)
  Fields("fldOpt3Desc").Value = arr(13)
  Fields("fldGPenTot").Value = arr(14)
  Fields("fldRange").Value = arr(15)
  Fields("fldType").Value = arr(16)
  Fields("fldShow").Value = arr(17)
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
  Label11.Top = 1080 'penalty
  Field21.Top = 1080
  Field7.Top = 1350 'opt1
  Field14.Top = 1350
  Field8.Top = 1620 'opt2
  Field15.Top = 1620
  Field9.Top = 1890 'opt3
  Field16.Top = 1890
  Line2.Y1 = 2160
  Line2.Y2 = 2160
  Label1.Top = 2250 'grand totals
  Field6.Top = 2250
  Field5.Top = 2250
  GroupFooter1.Height = 2600

  If CInt(Fields("fldRange").Value) = 2 Or CInt(Fields("fldShow").Value) = 44 Then
    Label2.Visible = False 'principle
    Field10.Visible = False
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
    '7/11/06 took out the bill filter to allow revenues for manual billing
'  ElseIf Fields("fldType").Value = "Billing" Then
'    Bill = True
'    Label3.Visible = False 'int
'    Field11.Visible = False
'    Label4.Visible = False 'adv
'    Field12.Visible = False
'    Label11.Visible = False 'pen
'    Field21.Visible = False
'    Label5.Top = 270 'll
'    Field13.Top = 270
'    GroupFooter1.Height = 1790
'    Line2.Y1 = 1350
'    Line2.Y2 = 1350
'    Label1.Top = 1350
'    Field6.Top = 1350
'    Field5.Top = 1350
'    Field7.Top = 810
'    Field14.Top = 810
'    Field8.Top = 1080
'    Field15.Top = 1080
'    Field9.Top = 1350
'    Field16.Top = 1350
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
    Line2.Y1 = 2160
    Line2.Y2 = 2160
    Label1.Top = 2250
    Field6.Top = 2250
    Field5.Top = 2250
  Else
    Line2.Y1 = 1350
    Line2.Y2 = 1350
    Label1.Top = 1710
    Field6.Top = 1710
    Field5.Top = 1710
  End If

  If Opt1 = True And Opt2 = False And Opt3 = False Then
    Field8.Visible = False
    Field15.Visible = False
    Field9.Visible = False
    Field16.Visible = False
    If Bill = False Then
      Line2.Y1 = 1620
      Line2.Y2 = 1620
      Label1.Top = 1710 'grand total (-int, -adv, -pen)
      Field6.Top = 1710
      Field5.Top = 1710
      GroupFooter1.Height = 2060
    Else 'bill = true
      Label5.Top = 270 'll
      Field13.Top = 270
      Field7.Top = 540 'opt1
      Field14.Top = 540
      Line2.Y1 = 810
      Line2.Y2 = 810
      Label1.Top = 900
      Field6.Top = 900
      Field5.Top = 900
      GroupFooter1.Height = 1250
    End If
  ElseIf Opt1 = True And Opt2 = True And Opt3 = False Then
    Field9.Visible = False
    Field16.Visible = False
    If Bill = False Then
      Line2.Y1 = 1890
      Line2.Y2 = 1890
      Label1.Top = 1980
      Field6.Top = 1980
      Field5.Top = 1980
      GroupFooter1.Height = 2330
    Else ' bill = true
      Field7.Top = 540
      Field14.Top = 540
      Field8.Top = 810
      Field15.Top = 810
      Line2.Y1 = 1080
      Line2.Y2 = 1080
      Label1.Top = 1170
      Field6.Top = 1170
      Field5.Top = 1170
      GroupFooter1.Height = 1520
    End If
  ElseIf Opt1 = True And Opt2 = False And Opt3 = True Then
    Field8.Visible = False
    Field15.Visible = False
    If Bill = False Then
      Field9.Top = 1620
      Field16.Top = 1620
      Line2.Y1 = 1890
      Line2.Y2 = 1890
      Label1.Top = 1980
      Field6.Top = 1980
      Field5.Top = 1980
      GroupFooter1.Height = 2330
    Else
      Field7.Top = 540
      Field14.Top = 540
      Field9.Top = 810
      Field16.Top = 810
      Line2.Y1 = 1080
      Line2.Y2 = 1080
      Label1.Top = 1170
      Field6.Top = 1170
      Field5.Top = 1170
      GroupFooter1.Height = 1520
    End If
  ElseIf Opt1 = False And Opt2 = True And Opt3 = False Then
    Field7.Visible = False
    Field14.Visible = False
    Field9.Visible = False
    Field16.Visible = False
    If Bill = False Then
      Field8.Top = 1350
      Field15.Top = 1350
      Line2.Y1 = 1620
      Line2.Y2 = 1620
      Label1.Top = 1710
      Field6.Top = 1710
      Field5.Top = 1710
      GroupFooter1.Height = 2060
    Else
      Field8.Top = 540
      Field15.Top = 540
      Line2.Y1 = 810
      Line2.Y2 = 810
      Label1.Top = 900
      Field6.Top = 900
      Field5.Top = 900
      GroupFooter1.Height = 1250
    End If
  ElseIf Opt1 = False And Opt2 = True And Opt3 = True Then
    Field7.Visible = False
    Field14.Visible = False
    If Bill = False Then
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
    Else
      Field8.Top = 540
      Field15.Top = 540
      Field9.Top = 810
      Field16.Top = 810
      Line2.Y1 = 1080
      Line2.Y2 = 1080
      Label1.Top = 1170
      Field6.Top = 1170
      Field5.Top = 1170
      GroupFooter1.Height = 1520
    End If
  ElseIf Opt1 = False And Opt2 = False And Opt3 = True Then
    Field7.Visible = False
    Field14.Visible = False
    Field8.Visible = False
    Field15.Visible = False
    If Bill = False Then
      Field9.Top = 1350
      Field16.Top = 1350
      Line2.Y1 = 1620
      Line2.Y2 = 1620
      Label1.Top = 1710
      Field6.Top = 1710
      Field5.Top = 1710
      GroupFooter1.Height = 2060
    Else
      Field9.Top = 540
      Field16.Top = 540
      Line2.Y1 = 810
      Line2.Y2 = 810
      Label1.Top = 900
      Field6.Top = 900
      Field5.Top = 900
      GroupFooter1.Height = 1250
    End If
  ElseIf Opt1 = False And Opt2 = False And Opt3 = False Then
    Field7.Visible = False
    Field14.Visible = False
    Field8.Visible = False
    Field15.Visible = False
    Field9.Visible = False
    Field16.Visible = False
    If Bill = False Then
      Line2.Y1 = 1350
      Line2.Y2 = 1350
      Label1.Top = 1440
      Field6.Top = 1440
      Field5.Top = 1440
      GroupFooter1.Height = 1790
    Else
      Line2.Y1 = 540
      Line2.Y2 = 540
      Label1.Top = 630
      Field6.Top = 630
      Field5.Top = 630
      GroupFooter1.Height = 980
    End If
  End If

End Sub

