VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arVASub2PersTransJrnl 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5784
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   9615
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   16960
   _ExtentY        =   10213
   SectionData     =   "arVASub2PersTransJrnl.dsx":0000
End
Attribute VB_Name = "arVASub2PersTransJrnl"
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
  Open StartPath & "\TAXRPTS\SUB2TAXJRNLP.RPT" For Input As #hFile
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
  Fields.Add ("fldGPenTot") '14)
  Fields.Add ("fldGMTTot") '15)
  Fields.Add ("fldGMCTot") '16)
  Fields.Add ("fldGFETot") '17)
  Fields.Add ("fldGMHTot") '18)
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
  Fields("fldGPenTot").Value = arr(14)
  Fields("fldGMTTot").Value = arr(15)
  Fields("fldGMCTot").Value = arr(16)
  Fields("fldGFETot").Value = arr(17)
  Fields("fldGMHTot").Value = arr(18)
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

  Label5.Top = 2250
  Field13.Top = 2250
  Label7.Top = 2520
  Field17.Top = 2520
  Field7.Top = 2790 'opt 1
  Field14.Top = 2790
  Field8.Top = 3060 'opt 2
  Field15.Top = 3060
  Field9.Top = 3330 'opt 3
  Field16.Top = 3330
  Line2.Y1 = 3600
  Line2.Y2 = 3600
  Label1.Top = 3690
  Field6.Top = 3690
  Field5.Top = 3690
  GroupFooter1.Height = 4056

  If CInt(Fields("fldRange").Value) = 2 Or CInt(Fields("fldShow").Value) = 44 Then
    Label2.Visible = False 'Pers
    Field10.Visible = False 'Pers
    Label8.Visible = False 'MT
    Field18.Visible = False 'MT
    Label9.Visible = False 'MC
    Field19.Visible = False 'MC
    Label10.Visible = False 'FE
    Field20.Visible = False 'FE
    Label11.Visible = False 'MH
    Field21.Visible = False 'MH
    Label3.Visible = False 'Int
    Field11.Visible = False 'Int
    Label4.Visible = False 'Adv
    Field12.Visible = False 'Adv
    Label5.Visible = False 'Late List
    Field13.Visible = False 'Late List
    Label7.Visible = False 'Pen
    Field17.Visible = False 'Pen
    Field7.Visible = False 'Opt1
    Field14.Visible = False 'Opt1
    Field8.Visible = False 'Opt2
    Field15.Visible = False 'Opt2
    Field9.Visible = False 'Opt3
    Field16.Visible = False 'Opt3
    Label6.Visible = False 'REVENUE
    Line1.Visible = False 'top line
    GroupFooter1.Height = 624
    Line2.Y1 = 0 'bottom line
    Line2.Y2 = 0
    Label1.Top = 90 'GRAND TOTALS
    Field6.Top = 90 'Cnt
    Field5.Top = 90 'Amt
    Exit Sub
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
  Line2.Y1 = 3600
  Line2.Y2 = 3600
  Label1.Top = 3690
  Field6.Top = 3690
  Field5.Top = 3690

  If Opt1 = True And Opt2 = False And Opt3 = False Then
    Field8.Visible = False
    Field15.Visible = False
    Field9.Visible = False
    Field16.Visible = False
    Line2.Y1 = 3060
    Line2.Y2 = 3060
    Label1.Top = 3150
    Field6.Top = 3150
    Field5.Top = 3150
    GroupFooter1.Height = 3516
  ElseIf Opt1 = True And Opt2 = True And Opt3 = False Then
    Field9.Visible = False
    Field16.Visible = False
    Line2.Y1 = 3330
    Line2.Y2 = 3330
    Label1.Top = 3420
    Field6.Top = 3420
    Field5.Top = 3420
    GroupFooter1.Height = 3786
  ElseIf Opt1 = True And Opt2 = False And Opt3 = True Then
    Field8.Visible = False
    Field15.Visible = False
    Field9.Top = 3060
    Field16.Top = 3060
    Line2.Y1 = 3330
    Line2.Y2 = 3330
    Label1.Top = 3420
    Field6.Top = 3420
    Field5.Top = 3420
    GroupFooter1.Height = 3786
  ElseIf Opt1 = False And Opt2 = True And Opt3 = False Then
    Field7.Visible = False
    Field14.Visible = False
    Field9.Visible = False
    Field16.Visible = False
    Field8.Top = 2790
    Field15.Top = 2790
    Line2.Y1 = 3060
    Line2.Y2 = 3060
    Label1.Top = 3150
    Field6.Top = 3150
    Field5.Top = 3150
    GroupFooter1.Height = 3516
  ElseIf Opt1 = False And Opt2 = True And Opt3 = True Then
    Field7.Visible = False
    Field14.Visible = False
    Field8.Top = 2790
    Field15.Top = 2790
    Field9.Top = 3060
    Field16.Top = 3060
    Line2.Y1 = 3330
    Line2.Y2 = 3330
    Label1.Top = 3420
    Field6.Top = 3420
    Field5.Top = 3420
    GroupFooter1.Height = 3786
  ElseIf Opt1 = False And Opt2 = False And Opt3 = True Then
    Field7.Visible = False
    Field14.Visible = False
    Field8.Visible = False
    Field15.Visible = False
    Field9.Top = 2790
    Field16.Top = 2790
    Line2.Y1 = 3060
    Line2.Y2 = 3060
    Label1.Top = 3150
    Field6.Top = 3150
    Field5.Top = 3150
    GroupFooter1.Height = 3516
  ElseIf Opt1 = False And Opt2 = False And Opt3 = False Then
    Field7.Visible = False
    Field14.Visible = False
    Field8.Visible = False
    Field15.Visible = False
    Field9.Visible = False
    Field16.Visible = False
    Line2.Y1 = 2790
    Line2.Y2 = 2790
    Label1.Top = 2880
    Field6.Top = 2880
    Field5.Top = 2880
    GroupFooter1.Height = 3246
  End If
    
End Sub



