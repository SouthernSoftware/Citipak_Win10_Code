VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arVASubTaxPMastBalDet 
   Caption         =   "ActiveReport1"
   ClientHeight    =   4332
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9330
   StartUpPosition =   2  'CenterScreen
   _ExtentX        =   16457
   _ExtentY        =   7646
   SectionData     =   "arVASubTaxPMastBalDet.dsx":0000
End
Attribute VB_Name = "arVASubTaxPMastBalDet"
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
  Open StartPath & "\TAXRPTS\TXPMSTBALSUBDET.RPT" For Input As #hFile
  Fields.Add ("fldYear") '0)
  Fields.Add ("fldYearBal") '1)
  Fields.Add ("fldOverage") '2)
  Fields.Add ("fldLastOne") '3)
  Fields.Add ("fldPersTot") '4)
  Fields.Add ("fldIntTot") '5)
  Fields.Add ("fldMTTot") '6)
  Fields.Add ("fldMCTot") '7)
  Fields.Add ("fldOpt1Tot") '8)
  Fields.Add ("fldOpt2Tot") '9)
  Fields.Add ("fldOpt3Tot") '10)
  Fields.Add ("fldOpt1Desc") '11)
  Fields.Add ("fldOpt2Desc") '12)
  Fields.Add ("fldOpt3Desc") '13)
  Fields.Add ("fldFETot") '14)
  Fields.Add ("fldMHTot") '15)
  Fields.Add ("fldPenTot") '16)
  
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
  Fields("fldYear").Value = arr(0)
  Fields("fldYearBal").Value = arr(1)
  Fields("fldOverage").Value = arr(2)
  Fields("fldLastOne").Value = arr(3)
'  If CInt(Fields("fldLastOne").Value) = 1 Then
'    Field4.Visible = True
'  Else
'    Field4.Visible = False
'  End If
  Fields("fldPersTot").Value = arr(4)
  Fields("fldIntTot").Value = arr(5)
  Fields("fldMTTot").Value = arr(6)
  Fields("fldMCTot").Value = arr(7)
  Fields("fldOpt1Tot").Value = arr(8)
  Fields("fldOpt2Tot").Value = arr(9)
  Fields("fldOpt3Tot").Value = arr(10)
  Fields("fldOpt1Desc").Value = arr(11)
  Fields("fldOpt2Desc").Value = arr(12)
  Fields("fldOpt3Desc").Value = arr(13)
  Fields("fldFETot").Value = arr(14)
  Fields("fldMHTot").Value = arr(15)
  Fields("fldPenTot").Value = arr(16)
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
  
  Opt1 = True
  Opt2 = True
  Opt3 = True
  
  If QPTrim$(Fields("fldOpt1Desc").Value) = "" Then
    Opt1 = False
  End If
  
  If QPTrim$(Fields("fldOpt2Desc").Value) = "" Then
    Opt2 = False
  End If
  
  If QPTrim$(Fields("fldOpt3Desc").Value) = "" Then
    Opt3 = False
  End If
  
  If Opt1 = True And Opt2 = True And Opt3 = True Then
    Exit Sub
  End If
  
  If Opt1 = True And Opt2 = False And Opt3 = False Then
    Field11.Visible = False
    Field12.Visible = False
    Field13.Visible = False
    Field14.Visible = False
  End If
  
  If Opt1 = True And Opt2 = True And Opt3 = False Then
    Field13.Visible = False
    Field14.Visible = False
  End If
  
  If Opt1 = True And Opt2 = False And Opt3 = True Then
    Field11.Visible = False
    Field12.Visible = False
    Field13.Top = 630
    Field14.Top = 630
  End If
  
  If Opt1 = False And Opt2 = True And Opt3 = True Then
    Field9.Visible = False
    Field10.Visible = False
    Field11.Top = 360
    Field12.Top = 360
    Field13.Top = 630
    Field14.Top = 630
  End If
  
  If Opt1 = False And Opt2 = True And Opt3 = False Then
    Field9.Visible = False
    Field10.Visible = False
    Field13.Visible = False
    Field14.Visible = False
    Field11.Top = 360
    Field12.Top = 360
  End If
  
  If Opt1 = False And Opt2 = False And Opt3 = True Then
    Field9.Visible = False
    Field10.Visible = False
    Field11.Visible = False
    Field12.Visible = False
    Field13.Top = 360
    Field14.Top = 360
  End If
  
  If Opt1 = False And Opt2 = False And Opt3 = False Then
    Field9.Visible = False
    Field10.Visible = False
    Field11.Visible = False
    Field12.Visible = False
    Field13.Visible = False
    Field14.Visible = False
  End If
  
'  If QPTrim$(Fields("fldOpt1Tot").Value) = "" Then
'    Field9.Visible = False
'    Field10.Visible = False
'  End If
'
'  If QPTrim$(Fields("fldOpt2Tot").Value) = "" Then
'    Field11.Visible = False
'    Field12.Visible = False
'  End If
'
'  If QPTrim$(Fields("fldOpt3Tot").Value) = "" Then
'    Field13.Visible = False
'    Field14.Visible = False
'  End If
  

End Sub

