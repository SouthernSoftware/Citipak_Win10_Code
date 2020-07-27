VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arSub2TransJrnl 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5790
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   9615
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   16960
   _ExtentY        =   10213
   SectionData     =   "arSub2TransJrnl.dsx":0000
End
Attribute VB_Name = "arSub2TransJrnl"
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
  Open StartPath & "\TAXRPTS\SUB2TAXJRNL.RPT" For Input As #hFile
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
  Fields.Add ("fldRange") '14)
  Fields.Add ("fldType") '15)
  Fields.Add ("fldShow") '16)
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
  Fields("fldRange").Value = arr(14)
  Fields("fldType").Value = arr(15)
  Fields("fldShow").Value = arr(16)
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
Private Sub GroupFooter1_Format()
  Dim Opt1 As Boolean
  Dim Opt2 As Boolean
  Dim Opt3 As Boolean
  Dim Bill As Boolean
  
  Bill = False
  Label5.Top = 1170
  Field13.Top = 1170
  Field7.Top = 1440
  Field14.Top = 1440
  Field8.Top = 1710
  Field15.Top = 1710
  Field9.Top = 1980
  Field16.Top = 1980
  Line2.Y1 = 2250
  Line2.Y2 = 2250
  Label1.Top = 2340
  Field6.Top = 2340
  Field5.Top = 2340
  GroupFooter1.Height = 2610
  '7/10/06 added back advertising and interest to bills to accommodate manuals
  
  If CInt(Fields("fldRange").Value) = 2 Or CInt(Fields("fldShow").Value) = 44 Then
    Label2.Visible = False
    Field10.Visible = False
    Label3.Visible = False
    Field11.Visible = False
    Label4.Visible = False
    Field12.Visible = False
    Label5.Visible = False
    Field13.Visible = False
    Field7.Visible = False
    Field14.Visible = False
    Field8.Visible = False
    Field15.Visible = False
    Field9.Visible = False
    Field16.Visible = False
    Label6.Visible = False
    Line1.Visible = False
    GroupFooter1.Height = 624
    Line2.Y1 = 0
    Line2.Y2 = 0
    Label1.Top = 90
    Field6.Top = 90
    Field5.Top = 90
    Exit Sub
'  ElseIf Fields("fldType").Value = "Billing" Then
'    Label3.Visible = False
'    Field11.Visible = False
'    Label4.Visible = False
'    Field12.Visible = False
'    Label5.Top = 630
'    Field13.Top = 630
'    Bill = True
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
  
  Field7.Visible = True
  Field14.Visible = True
  Field8.Visible = True
  Field15.Visible = True
  Field9.Visible = True
  Field16.Visible = True
  If Bill = False Then
    Line2.Y1 = 2250
    Line2.Y2 = 2250
    Label1.Top = 2340
    Field6.Top = 2340
    Field5.Top = 2340
  Else
    Field7.Top = 900
    Field14.Top = 900
    Field8.Top = 1170
    Field15.Top = 1170
    Field9.Top = 1440
    Field16.Top = 1440
    Line2.Y1 = 1710
    Line2.Y2 = 1710
    Label1.Top = 1800
    Field6.Top = 1800
    Field5.Top = 1800
  End If
  
  If Opt1 = True And Opt2 = False And Opt3 = False Then
    Field8.Visible = False
    Field15.Visible = False
    Field9.Visible = False
    Field16.Visible = False
    If Bill = False Then
      Line2.Y1 = 1710
      Line2.Y2 = 1710
      Label1.Top = 1800
      Field6.Top = 1800
      Field5.Top = 1800
      GroupFooter1.Height = 2076
    Else
      Field7.Top = 900
      Field14.Top = 900
      Line2.Y1 = 1170
      Line2.Y2 = 1170
      Label1.Top = 1260
      Field6.Top = 1260
      Field5.Top = 1260
      GroupFooter1.Height = 1536
    End If
  ElseIf Opt1 = True And Opt2 = True And Opt3 = False Then
    Field9.Visible = False
    Field16.Visible = False
    If Bill = False Then
      Line2.Y1 = 1980
      Line2.Y2 = 1980
      Label1.Top = 2070
      Field6.Top = 2070
      Field5.Top = 2070
      GroupFooter1.Height = 2346
    Else
      Field7.Top = 900
      Field14.Top = 900
      Field8.Top = 1170
      Field15.Top = 1170
      Line2.Y1 = 1440
      Line2.Y2 = 1440
      Label1.Top = 1530
      Field6.Top = 1530
      Field5.Top = 1530
      GroupFooter1.Height = 1806
    End If
  ElseIf Opt1 = True And Opt2 = False And Opt3 = True Then
    Field8.Visible = False
    Field15.Visible = False
    If Bill = False Then
      Field9.Top = 1710
      Field16.Top = 1710
      Line2.Y1 = 1980
      Line2.Y2 = 1980
      Label1.Top = 2070
      Field6.Top = 2070
      Field5.Top = 2070
      GroupFooter1.Height = 2346
    Else
      Field7.Top = 900
      Field14.Top = 900
      Field9.Top = 1170
      Field16.Top = 1170
      Line2.Y1 = 1440
      Line2.Y2 = 1440
      Label1.Top = 1530
      Field6.Top = 1530
      Field5.Top = 1530
      GroupFooter1.Height = 1806
    End If
  ElseIf Opt1 = False And Opt2 = True And Opt3 = False Then
    Field7.Visible = False
    Field14.Visible = False
    Field9.Visible = False
    Field16.Visible = False
    If Bill = False Then
      Field8.Top = 1440
      Field15.Top = 1440
      Line2.Y1 = 1710
      Line2.Y2 = 1710
      Label1.Top = 1800
      Field6.Top = 1800
      Field5.Top = 1800
      GroupFooter1.Height = 2076
    Else
      Field8.Top = 900
      Field15.Top = 900
      Line2.Y1 = 1170
      Line2.Y2 = 1170
      Label1.Top = 1260
      Field6.Top = 1260
      Field5.Top = 1260
      GroupFooter1.Height = 1536
    End If
  ElseIf Opt1 = False And Opt2 = True And Opt3 = True Then
    Field7.Visible = False
    Field14.Visible = False
    If Bill = False Then
      Field8.Top = 1440
      Field15.Top = 1440
      Field9.Top = 1710
      Field16.Top = 1710
      Line2.Y1 = 1980
      Line2.Y2 = 1980
      Label1.Top = 2070
      Field6.Top = 2070
      Field5.Top = 2070
      GroupFooter1.Height = 2346
    Else
      Field8.Top = 900
      Field15.Top = 900
      Field9.Top = 1170
      Field16.Top = 1170
      Line2.Y1 = 1440
      Line2.Y2 = 1440
      Label1.Top = 1530
      Field6.Top = 1530
      Field5.Top = 1530
      GroupFooter1.Height = 1806
    End If
  ElseIf Opt1 = False And Opt2 = False And Opt3 = True Then
    Field7.Visible = False
    Field14.Visible = False
    Field8.Visible = False
    Field15.Visible = False
    If Bill = False Then
      Field9.Top = 1440
      Field16.Top = 1440
      Line2.Y1 = 1710
      Line2.Y2 = 1710
      Label1.Top = 1800
      Field6.Top = 1800
      Field5.Top = 1800
      GroupFooter1.Height = 2076
    Else
      Field9.Top = 900
      Field16.Top = 900
      Line2.Y1 = 1170
      Line2.Y2 = 1170
      Label1.Top = 1260
      Field6.Top = 1260
      Field5.Top = 1260
      GroupFooter1.Height = 1536
    End If
  ElseIf Opt1 = False And Opt2 = False And Opt3 = False Then
    Field7.Visible = False
    Field14.Visible = False
    Field8.Visible = False
    Field15.Visible = False
    Field9.Visible = False
    Field16.Visible = False
    If Bill = False Then
      Line2.Y1 = 1440
      Line2.Y2 = 1440
      Label1.Top = 1530
      Field6.Top = 1530
      Field5.Top = 1530
      GroupFooter1.Height = 1806
    Else
      Line2.Y1 = 900
      Line2.Y2 = 900
      Label1.Top = 990
      Field6.Top = 990
      Field5.Top = 990
      GroupFooter1.Height = 1266
    End If
  End If
    
End Sub
