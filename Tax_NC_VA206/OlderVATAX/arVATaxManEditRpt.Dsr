VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arVATaxManEditRpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manual Tax Bill Edit Report"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "arVATaxManEditRpt.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20532
   _ExtentY        =   15399
   SectionData     =   "arVATaxManEditRpt.dsx":08CA
End
Attribute VB_Name = "arVATaxManEditRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private hFile As Integer
Private Temp_Class As Resize_Class
Dim ReportYN As Boolean

Private Sub ActiveReport_DataInitialize()
  Dim x As Integer
  hFile = FreeFile
  Open StartPath & "\TAXRPTS\TXMANEDT.RPT" For Input As #hFile
  Fields.Add ("fldTown") '0)
  Fields.Add ("fldAcctNum") '1)
  Fields.Add ("fldCustName") '2)
  Fields.Add ("fldBillNum") '3)
  Fields.Add ("fldClass") '4)
  Fields.Add ("fldOpt1Desc") '5)
  Fields.Add ("fldOpt2Desc") '6)
  Fields.Add ("fldOpt3Desc") '7)
  Fields.Add ("fldPrinc") '8)
  Fields.Add ("fldInt") '9)
  Fields.Add ("fldAdvCol") '10)
  Fields.Add ("fldLateList") '11)
  Fields.Add ("fldOpt1") '12)
  Fields.Add ("fldOpt2") '13)
  Fields.Add ("fldOpt3") '14)
  Fields.Add ("fldTaxYear") '15)
  Fields.Add ("fldTransDate") '16)
  Fields.Add ("fldPin") '17)
  Fields.Add ("fldTotCustAmt") '18)
  Fields.Add ("fldGPrinc") '19)
  Fields.Add ("fldGInt") '20)
  Fields.Add ("fldGAdvCol") '21)
  Fields.Add ("fldGLateList") '22)
  Fields.Add ("fldGOpt1") '23)
  Fields.Add ("fldGOpt2") '24)
  Fields.Add ("fldGOpt3") '25)
  Fields.Add ("fldGTotal") '26)
  Fields.Add ("fldBillCnt") '27)
  Fields.Add ("fldFE") '28)
  Fields.Add ("fldTotFE") '29)
  Fields.Add ("fldMH") '30)
  Fields.Add ("fldTotMH") '31)
  Fields.Add ("fldPGTotal") '32)
  Fields.Add ("fldRealCnt") '33)
  Fields.Add ("fldPersCnt") '34)
  Fields.Add ("fldPGOpt1") '35)
  Fields.Add ("fldPGOpt2") '36)
  Fields.Add ("fldPGOpt3") '37)
  Fields.Add ("fldGPers") '38)
  Fields.Add ("fldTotPInt") '39)
  Fields.Add ("fldGMT") '40)
  Fields.Add ("fldPGPen") '41)
  Fields.Add ("fldRGPen") '42)
  Fields.Add ("fldRGTotal") '43)
  Fields.Add ("fldPOpt1Desc") '44)
  Fields.Add ("fldPOpt2Desc") '45)
  Fields.Add ("fldPOpt3DEsc") '46)
  Fields.Add ("fldTotMC") '47)
  Fields.Add ("fldPen") '48)
 
  Fields.Add ("fldROpt1Desc")
  Fields.Add ("fldROpt2Desc")
  Fields.Add ("fldROpt3Desc")
  Fields.Add ("fldGPOpt1Desc")
  Fields.Add ("fldGPOpt2Desc")
  Fields.Add ("fldGPOpt3Desc")
  
End Sub

Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
'    Unload frmLoadReport
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
  Fields("fldTown").Value = arr(0)
  Fields("fldAcctNum").Value = arr(1)
  Fields("fldCustName").Value = arr(2)
  Fields("fldBillNum").Value = arr(3)
  Fields("fldClass").Value = arr(4)
  Fields("fldOpt1Desc").Value = arr(5)
  Fields("fldOpt2Desc").Value = arr(6)
  Fields("fldOpt3Desc").Value = arr(7)
  Fields("fldPrinc").Value = arr(8)
  Fields("fldInt").Value = arr(9)
  Fields("fldAdvCol").Value = arr(10)
  Fields("fldLateList").Value = arr(11)
  Fields("fldOpt1").Value = arr(12)
  Fields("fldOpt2").Value = arr(13)
  Fields("fldOpt3").Value = arr(14)
  Fields("fldTaxYear").Value = arr(15)
  Fields("fldTransDate").Value = arr(16)
  Fields("fldPin").Value = arr(17)
  If QPTrim$(arr(5)) = "" Then
    Field4.Visible = False
    Field14.Visible = False
    Field23.Visible = False
    Field26.Visible = False
  Else
    Field4.Visible = True
    Field14.Visible = True
    Field23.Visible = True
    Field26.Visible = True
  End If
  If QPTrim$(arr(6)) = "" Then
    Field5.Visible = False
    Field15.Visible = False
    Field24.Visible = False
    Field27.Visible = False
  Else
    Field5.Visible = True
    Field15.Visible = True
    Field24.Visible = True
    Field27.Visible = True
  End If
  If QPTrim$(arr(7)) = "" Then
    Field6.Visible = False
    Field16.Visible = False
    Field25.Visible = False
    Field28.Visible = False
  Else
    Field6.Visible = True
    Field16.Visible = True
    Field25.Visible = True
    Field28.Visible = True
  End If
  Fields("fldTotCustAmt").Value = arr(18)
  Fields("fldGPrinc").Value = arr(19)
  Fields("fldGInt").Value = arr(20)
  Fields("fldGAdvCol").Value = arr(21)
  Fields("fldGLateList").Value = arr(22)
  Fields("fldGOpt1").Value = arr(23)
  Fields("fldGOpt2").Value = arr(24)
  Fields("fldGOpt3").Value = arr(25)
  Fields("fldGTotal").Value = arr(26)
  Fields("fldBillCnt").Value = arr(27)
  Fields("fldFE").Value = arr(28)
  Fields("fldTotFE").Value = arr(29)
  Fields("fldMH").Value = arr(30)
  Fields("fldTotMH").Value = arr(31)
  Fields("fldPGTotal").Value = arr(32)
  Fields("fldRealCnt").Value = arr(33)
  Fields("fldPersCnt").Value = arr(34)
  Fields("fldPGOpt1").Value = arr(35)
  Fields("fldPGOpt2").Value = arr(36)
  Fields("fldPGOpt3").Value = arr(37)
  Fields("fldGPers").Value = arr(38)
  Fields("fldTotPInt").Value = arr(39)
  Fields("fldGMT").Value = arr(40)
  Fields("fldPGPen").Value = arr(41)
  Fields("fldRGPen").Value = arr(42)
  Fields("fldRGTotal").Value = arr(43)
  Fields("fldPOpt1Desc").Value = arr(44)
  If QPTrim$(arr(44)) = "" Then
    Field34.Visible = False
    Field37.Visible = False
  End If
  Fields("fldPOpt2Desc").Value = arr(45)
  If QPTrim$(arr(45)) = "" Then
    Field35.Visible = False
    Field38.Visible = False
  End If
  Fields("fldPOpt3DEsc").Value = arr(46)
  If QPTrim$(arr(46)) = "" Then
    Field36.Visible = False
    Field39.Visible = False
  End If
  Fields("fldTotMC").Value = arr(47)
  Fields("fldPen").Value = arr(48)
  
  Fields("fldROpt1Desc") = "Total " + QPTrim$(arr(5))
  Fields("fldROpt2Desc") = "Total " + QPTrim$(arr(6))
  Fields("fldROpt3Desc") = "Total " + QPTrim$(arr(7))
  Fields("fldGPOpt1Desc") = "Total " + QPTrim$(arr(44))
  Fields("fldGPOpt2Desc") = "Total " + QPTrim$(arr(45))
  Fields("fldGPOpt3Desc") = "Total " + QPTrim$(arr(46))
End Sub

Private Sub ActiveReport_Initialize()
  Me.Toolbar.Tools.Add "&Close"
  Me.Toolbar.Tools.Add "Save/&Excel"
  Me.Toolbar.Tools.Add "&Text"
End Sub

Private Sub ActiveReport_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
  If KeyCode = vbKeyEscape Then
    Unload Me
    KeyCode = 0
  End If
  If Shift = 4 Then
    If KeyCode = vbKeyC Then
      Unload Me
      KeyCode = 0
    ElseIf KeyCode = vbKeyE Then
      Screen.MousePointer = vbHourglass
      ExportReport 1
      Screen.MousePointer = vbDefault
      DoEvents
      frmVATaxMsg.Label1.Caption = "File - TaxManEditRpt.xls, created in the Citipak Directory."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show vbModal
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      frmVATaxMsg.Label1.Caption = "File - TaxManEditRpt.txt, created in the Citipak Directory."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show vbModal
      KeyCode = 0
    End If
  End If
End Sub

Private Sub ActiveReport_Terminate()
  Close
End Sub

Private Sub ActiveReport_ToolbarClick(ByVal Tool As DDActiveReports2.DDTool)
  If Tool = "&Close" Then
    Close
    Unload Me
  End If
  If Tool = "Save/&Excel" Then
    Screen.MousePointer = vbHourglass
    ExportReport 1
    Screen.MousePointer = vbDefault
    DoEvents
    frmVATaxMsg.Label1.Caption = "File - TaxManEditRpt.xls, created in the Citipak Directory."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    frmVATaxMsg.Label1.Caption = "File - TaxManEditRpt.txt, created in the Citipak Directory."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
  End If
End Sub
Private Sub ExportReport(x As Integer)
  Dim oEXL As ActiveReportsExcelExport.ARExportExcel
  Dim oTXT As ActiveReportsTextExport.ARExportText
  Dim outfile As String
  If Right$(StartPath, 1) = ":" Then
    outfile = StartPath
  Else
    outfile = StartPath & "\"
  End If
  Select Case x
    Case 1   '"Excel"
        Set oEXL = New ActiveReportsExcelExport.ARExportExcel
        oEXL.FileName = outfile & "TaxManEditRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "TaxManEditRpt.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
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
'  End If
'End Sub

Private Sub ActiveReport_ReportStart()
  Me.fldTimeDate.Text = Date
  Me.Zoom = -1
  ReportYN = False
End Sub

Private Sub Detail_Format()
  If Fields("fldClass").Value = "REAL" Then
    Field50.Visible = False
    Field51.Visible = False
    If QPTrim$(Fields("fldOpt1Desc").Value) <> "" Then
      Field14.Visible = True
    End If
    If QPTrim$(Fields("fldOpt2Desc").Value) <> "" Then
      Field15.Visible = True
    End If
    If QPTrim$(Fields("fldOpt3Desc").Value) <> "" Then
      Field16.Visible = True
    End If
  ElseIf Fields("fldClass").Value = "PERSONAL" Then
    Field50.Visible = True
    Field51.Visible = True
    If QPTrim$(Fields("fldPOpt1Desc").Value) <> "" Then
      Field14.Visible = True
    End If
    If QPTrim$(Fields("fldPOpt2Desc").Value) <> "" Then
      Field15.Visible = True
    End If
    If QPTrim$(Fields("fldPOpt3Desc").Value) <> "" Then
      Field16.Visible = True
    End If
  End If
End Sub

Private Sub GroupHeader1_Format()
  Label56.Visible = True
  Label57.Visible = True
  Field4.Visible = False
  Field5.Visible = False
  Field6.Visible = False
  If Fields("fldClass").Value = "REAL" Then
    Label19.Caption = "Principle"
    Label21.Caption = "Adv/Col"
    Label27.Caption = "Late List"
    Label56.Visible = False 'Farm Equipment
    Label57.Visible = False 'Mobile Homes
    If QPTrim$(Fields("fldOpt1Desc").Value) <> "" Then
      Field4.Visible = True
    End If
    If QPTrim$(Fields("fldOpt2Desc").Value) <> "" Then
      Field5.Visible = True
    End If
    If QPTrim$(Fields("fldOpt3Desc").Value) <> "" Then
      Field6.Visible = True
    End If
  ElseIf Fields("fldClass").Value = "PERSONAL" Then
    Label19.Caption = "Personal"
    Label21.Caption = "Machine Tools"
    Label27.Caption = "Merchant Capital"
    If QPTrim$(Fields("fldPOpt1Desc").Value) <> "" Then
      Field4.Visible = True
    End If
    If QPTrim$(Fields("fldPOpt2Desc").Value) <> "" Then
      Field5.Visible = True
    End If
    If QPTrim$(Fields("fldPOpt3Desc").Value) <> "" Then
      Field6.Visible = True
    End If
  End If
End Sub

'Private Sub PageHeader_Format()
'  If ReportYN = True Then
'    PageHeader.Height = 1260
'    Line1.Y1 = 1250
'    Line1.Y2 = 1250
'  End If
'End Sub

Private Sub ReportFooter_Format()
  ReportYN = True
  Set SubReport1 = New arVASubTaxManEdit
'  Field4.Visible = False
'  Field5.Visible = False
'  Field6.Visible = False
'  Label16.Visible = False
'  Label23.Visible = False
'  Label17.Visible = False
'  Label28.Visible = False
'  Label29.Visible = False
'  Label30.Visible = False
'  Label19.Visible = False
'  Label20.Visible = False
'  Label21.Visible = False
'  Label27.Visible = False
'  Label31.Visible = False
'  Label40.Visible = False
  
End Sub
