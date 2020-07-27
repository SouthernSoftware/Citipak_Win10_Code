VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arVASubTaxPayEditJrnl 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8265
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   _ExtentX        =   14579
   _ExtentY        =   12991
   SectionData     =   "arVASubTaxPayEditJrnl.dsx":0000
End
Attribute VB_Name = "arVASubTaxPayEditJrnl"
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
  Open StartPath & "\TAXRPTS\SubEdPay1.RPT" For Input As #hFile
  Fields.Add ("fldGPrincDisc") '0)
  Fields.Add ("fldGInt") '1)
  Fields.Add ("fldGAdvCol") '2)
  Fields.Add ("fldGLateList") '3)
  Fields.Add ("fldGRev1") '4)
  Fields.Add ("fldGRev2") '5)
  Fields.Add ("fldGRev3") '6)
  Fields.Add ("fldGTot") '7)
  Fields.Add ("fldGPrinc") '8)
  Fields.Add ("fldGDisc") '9)
  Fields.Add ("fldGOverPay") '10)
  Fields.Add ("fldRevDesc1") '11)
  Fields.Add ("fldRevDesc2") '12)
  Fields.Add ("fldRevDesc3") '13)
  Fields.Add ("fldType") '14)
  Fields.Add ("fldGPenAmt") '15)
  Fields.Add ("fldOpt1") '16)
  Fields.Add ("fldOpt2") '17)
  Fields.Add ("fldOpt3") '18)
  
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
  ' Here we set the values of the fields that we defines as unbound
  ' or user defined.
  Fields("fldGPrincDisc").Value = arr(0) 'also GPers
  Fields("fldGInt").Value = arr(1) 'also GMachTools
  Fields("fldGAdvCol").Value = arr(2) 'also GMerchCap
  Fields("fldGLateList").Value = arr(3) 'also FarmEquip
  Fields("fldGTot").Value = arr(7)
  Fields("fldGPrinc").Value = arr(8) 'also GPers(dup)
  Fields("fldGDisc").Value = arr(9)
  Fields("fldGOverPay").Value = arr(10)
  Fields("fldRevDesc1").Value = arr(11)
  Fields("fldRevDesc2").Value = arr(12)
  Fields("fldRevDesc3").Value = arr(13)
  Fields("fldType").Value = arr(14)
  Fields("fldGPenAmt").Value = arr(15)
  Fields("fldOpt1").Value = arr(16)
  Fields("fldOpt2").Value = arr(17)
  Fields("fldOpt3").Value = arr(18)
  If arr(14) = "R" Then
    Fields("fldGRev1").Value = arr(4) 'also MobHomes
    Fields("fldGRev2").Value = arr(5) 'also PGInt
    Fields("fldGRev3").Value = arr(6) 'also GPPenalty
  ElseIf arr(14) = "P" Then
    Fields("fldGPenAmt").Value = arr(4)
    Fields("fldGRev1").Value = arr(5)
    Fields("fldGRev2").Value = arr(15)
  End If
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

'Private Sub Form_Resize()
'  If Me.WindowState <> vbMinimized Then
'    Me.Visible = False
'    'Temp_Class.ResizeControls Me
'    Me.Visible = True
'    Me.SetFocus
'    DoEvents
'  End If
'End Sub

Private Sub Detail_Format()
  Field29.Visible = True
  Field22.Visible = True
  Field34.Visible = True
  Field31.Visible = True
  Field35.Visible = True
  Field32.Visible = True
  Field20.Visible = True
  Field21.Visible = True
  Field22.Visible = True
  If Fields("fldType").Value = "P" Then
    Label23.Visible = False 'Princ + Disc label
    Field16.Visible = False
    Line2.Visible = False
    
    Label28.Top = 90
    Label28.Caption = "Personal"
    Field24.Top = 90
    
    Label29.Top = 360 'discount
    Field25.Top = 360
    
    Label24.Top = 630
    Label24.Caption = "Machine Tools"
    Field17.Top = 630
    
    Label25.Top = 900
    Label25.Caption = "Merchant Capital"
    Field18.Top = 900
    
    Label26.Top = 1170
    Label26.Caption = "Farm Equipment"
    Field19.Top = 1170
    
    Label34.Top = 1440
    Label34.Caption = "Mobile Homes"
    Field30.Top = 1440
    
    Field27.Top = 1710
    Field27.Text = "Interest"
    Field20.Top = 1710
    
    Field28.Top = 1980
    Field28.Text = "Penalty"
    Field21.Top = 1980
    
    If QPTrim$(Fields("fldRevDesc1").Value) <> "" Then
      Field29.Text = QPTrim$(Fields("fldRevDesc1").Value)
    Else
      Field29.Text = ""
      Field22.Text = ""
    End If
    Field29.Top = 2250
    Field22.Top = 2250
    
    If QPTrim$(Fields("fldRevDesc2").Value) <> "" Then
      Field34.Text = QPTrim$(Fields("fldRevDesc2").Value)
    Else
      Field34.Text = ""
      Field31.Text = ""
    End If
    Field34.Top = 2520
    Field31.Top = 2520
    
    If QPTrim$(Fields("fldRevDesc3").Value) <> "" Then
      Field35.Text = QPTrim$(Fields("fldRevDesc3").Value)
    Else
      Field35.Text = ""
      Field32.Text = ""
    End If
    Field35.Top = 2790
    Field32.Top = 2790
    
    Label30.Top = 3060
    Field26.Top = 3060
    
    Line1.Y1 = 3380
    Line1.Y2 = 3380
    
    Label27.Top = 3430
    Field23.Top = 3430
  ElseIf Fields("fldType").Value = "R" Then
    Field34.Visible = False
    Field31.Visible = False
    Field35.Visible = False
    Field32.Visible = False
    Label30.Top = 2970
    Field26.Top = 2970
    Line1.Y1 = 3290
    Line1.Y2 = 3290
    Label27.Top = 3340
    Field23.Top = 3340
    Detail.Height = 3400
    If QPTrim$(Fields("fldRevDesc1").Value) = "" Then
      Field20.Visible = False
    End If
    If QPTrim$(Fields("fldRevDesc2").Value) = "" Then
      Field21.Visible = False
    End If
    If QPTrim$(Fields("fldRevDesc3").Value) = "" Then
      Field22.Visible = False
    End If
  End If
End Sub

Private Sub GroupHeader1_Format()
  If Fields("fldType").Value = "R" Then
    Label33.Caption = "REAL PROPERTY"
  ElseIf Fields("fldType").Value = "P" Then
    Label33.Caption = "PERSONAL PROPERTY"
  End If
End Sub
