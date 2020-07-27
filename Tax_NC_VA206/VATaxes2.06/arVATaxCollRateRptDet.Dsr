VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arVATaxCollRateRptDet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Collection Rate Report"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11655
   Icon            =   "arVATaxCollRateRptDet.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20558
   _ExtentY        =   15452
   SectionData     =   "arVATaxCollRateRptDet.dsx":08CA
End
Attribute VB_Name = "arVATaxCollRateRptDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim Over As clsTextBoxOverRider
  Private hFile As Integer
  Private Temp_Class As Resize_Class
Private Sub ActiveReport_DataInitialize()
  Dim x As Integer
  hFile = FreeFile
  Open StartPath & "\TAXRPTS\COLLECTRTD.RPT" For Input As #hFile
  Fields.Add ("fldTown") '0)
  Fields.Add ("fldYear") '1)
  Fields.Add ("fldChrgs") '2)
  Fields.Add ("fldPaid") '3)
  Fields.Add ("fldPct") '4)
  Fields.Add ("fldGTotChrgs") '5)
  Fields.Add ("fldGTotPaid") '6)
  Fields.Add ("fldGTotPct") '7)
  Fields.Add ("fldBillType") '8)
  Fields.Add ("fldPrincChrgs") '9)
  Fields.Add ("fldPrincPaid") '10)
  Fields.Add ("fldPrincPct") '11)
  Fields.Add ("fldRIntChrgs") '12)
  Fields.Add ("fldRIntPaid") '13)
  Fields.Add ("fldRIntPct") '14)
  Fields.Add ("fldAdvChrgs") '15)
  Fields.Add ("fldAdvPaid") '16)
  Fields.Add ("fldAdvPct") '17)
  Fields.Add ("fldLateListChrgs") '18)
  Fields.Add ("fldLateListPaid") '19)
  Fields.Add ("fldLateListPct") '20)
  Fields.Add ("fldRPenChrgs") '21)
  Fields.Add ("fldRPenPaid") '22)
  Fields.Add ("fldRPenPct") '23)
  Fields.Add ("fldROpt1Chrgs") '24)
  Fields.Add ("fldROpt1Paid") '25)
  Fields.Add ("fldROPt1Pct") '26)
  Fields.Add ("fldROpt2Chrgs") '27)
  Fields.Add ("fldROpt2Paid") '28)
  Fields.Add ("fldROPt2Pct") '29)
  Fields.Add ("fldROpt3Chrgs") '30)
  Fields.Add ("fldROpt3Paid") '31)
  Fields.Add ("fldROPt3Pct") '32)
  Fields.Add ("fldPersChrgs") '33)
  Fields.Add ("fldPersPaid") '34)
  Fields.Add ("fldPersPct") '35)
  Fields.Add ("fldMTChrgs") '36)
  Fields.Add ("fldMTPaid") '37)
  Fields.Add ("fldMTPct") '38)
  Fields.Add ("fldMCChrgs") '39)
  Fields.Add ("fldMCPaid") '40)
  Fields.Add ("fldMCPct") '41)
  Fields.Add ("fldFEChrgs") '42)
  Fields.Add ("fldFEPaid") '43)
  Fields.Add ("fldFEPct") '44)
  Fields.Add ("fldMHChrgs") '45)
  Fields.Add ("fldMHPaid") '46)
  Fields.Add ("fldMHPct") '47)
  Fields.Add ("fldPIntChrgs") '48)
  Fields.Add ("fldPIntPaid") '49)
  Fields.Add ("fldPIntPct") '50)
  Fields.Add ("fldPPenChrgs") '51)
  Fields.Add ("fldPPenPaid") '52)
  Fields.Add ("fldPPenPct") '53)
  Fields.Add ("fldPOpt1Chrgs") '54)
  Fields.Add ("fldPOPt1Paid") '55)
  Fields.Add ("fldPOpt1Pct") '56)
  Fields.Add ("fldPOpt2Chrgs") '57)
  Fields.Add ("fldPOPt2Paid") '58)
  Fields.Add ("fldPOpt2Pct") '59)
  Fields.Add ("fldPOpt3Chrgs") '60)
  Fields.Add ("fldPOPt3Paid") '61)
  Fields.Add ("fldPOpt3Pct") '62)
  Fields.Add ("fldGTPrincChrgs") '63)
  Fields.Add ("fldGTPrincPaid") '64)
  Fields.Add ("fldGTPrincPct") '65)
  Fields.Add ("fldGTRIntChrgs") '66)
  Fields.Add ("fldGTRIntPaid") '67)
  Fields.Add ("fldGTRIntPct") '68)
  Fields.Add ("fldGTAdvChrgs") '69)
  Fields.Add ("fldGTAdvPaid") '70)
  Fields.Add ("fldGTAdvPct") '71)
  Fields.Add ("fldGTLateListChrgs") '72)
  Fields.Add ("fldGTLateListPaid") '73)
  Fields.Add ("fldGTLateListPct") '74)
  Fields.Add ("fldGTRPenChrgs") '75)
  Fields.Add ("fldGTRPenPaid") '76)
  Fields.Add ("fldGTRPenPct") '77)
  Fields.Add ("fldGTROpt1Chrgs") '78)
  Fields.Add ("fldGTROpt1Paid") '79)
  Fields.Add ("fldGTROpt1Pct") '80)
  Fields.Add ("fldGTROpt2Chrgs") '81)
  Fields.Add ("fldGTROpt2Paid") '82)
  Fields.Add ("fldGTROpt2Pct") '83)
  Fields.Add ("fldGTROpt3Chrgs") '84)
  Fields.Add ("fldGTROpt3Paid") '85)
  Fields.Add ("fldGTROpt3Pct") '86)
  Fields.Add ("fldGTPersChrgs") '87)
  Fields.Add ("fldGTPersPaid") '88)
  Fields.Add ("fldGTPersPct") '89)
  Fields.Add ("fldGTMTChrgs") '90)
  Fields.Add ("fldGTMTPaid") '91)
  Fields.Add ("fldGTMTPct") '92)
  Fields.Add ("fldGTMCChrgs") '93)
  Fields.Add ("fldGTMCPaid") '94)
  Fields.Add ("fldGTMCPct") '95)
  Fields.Add ("fldGTFEChrgs") '96)
  Fields.Add ("fldGTFEPaid") '97)
  Fields.Add ("fldGTFEPct") '98)
  Fields.Add ("fldGTMHChrgs") '99)
  Fields.Add ("fldGTMHPaid") '100)
  Fields.Add ("fldGTMHPct") '101)
  Fields.Add ("fldGTPIntChrgs") '102)
  Fields.Add ("fldGTPIntPaid") '103)
  Fields.Add ("fldGTPIntPct") '104)
  Fields.Add ("fldGTPPenChrgs") '105)
  Fields.Add ("fldGTPPenPaid") '106)
  Fields.Add ("fldGTPPenPct") '107)
  Fields.Add ("fldGTPOpt1Chrgs") '108)
  Fields.Add ("fldGTPOpt1Paid") '109)
  Fields.Add ("fldGTPOpt1Pct") '110)
  Fields.Add ("fldGTPOpt2Chrgs") '111)
  Fields.Add ("fldGTPOpt2Paid") '112)
  Fields.Add ("fldGTPOpt2Pct") '113)
  Fields.Add ("fldGTPOpt3Chrgs") '114)
  Fields.Add ("fldGTPOpt3Paid") '115)
  Fields.Add ("fldGTPOpt3Pct") '116)
  
End Sub
Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
    Unload frmVATaxLoadReport
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
  Fields("fldYear").Value = arr(1)
  Fields("fldChrgs").Value = arr(2)
  Fields("fldPaid").Value = arr(3)
  Fields("fldPct").Value = arr(4)
  Fields("fldGTotChrgs").Value = arr(5)
  Fields("fldGTotPaid").Value = arr(6)
  Fields("fldGTotPct").Value = arr(7)
  Fields("fldBillType").Value = arr(8)
  Fields("fldPrincChrgs").Value = arr(9)
  Fields("fldPrincPaid").Value = arr(10)
  Fields("fldPrincPct").Value = arr(11)
  Fields("fldRIntChrgs").Value = arr(12)
  Fields("fldRIntPaid").Value = arr(13)
  Fields("fldRIntPct").Value = arr(14)
  Fields("fldAdvChrgs").Value = arr(15)
  Fields("fldAdvPaid").Value = arr(16)
  Fields("fldAdvPct").Value = arr(17)
  Fields("fldLateListChrgs").Value = arr(18)
  Fields("fldLateListPaid").Value = arr(19)
  Fields("fldLateListPct").Value = arr(20)
  Fields("fldRPenChrgs").Value = arr(21)
  Fields("fldRPenPaid").Value = arr(22)
  Fields("fldRPenPct").Value = arr(23)
  Fields("fldROpt1Chrgs").Value = arr(24)
  Fields("fldROpt1Paid").Value = arr(25)
  Fields("fldROPt1Pct").Value = arr(26)
  Fields("fldROpt2Chrgs").Value = arr(27)
  Fields("fldROpt2Paid").Value = arr(28)
  Fields("fldROPt2Pct").Value = arr(29)
  Fields("fldROpt3Chrgs").Value = arr(30)
  Fields("fldROpt3Paid").Value = arr(31)
  Fields("fldROPt3Pct").Value = arr(32)
  Fields("fldPersChrgs").Value = arr(33)
  Fields("fldPersPaid").Value = arr(34)
  Fields("fldPersPct").Value = arr(35)
  Fields("fldMTChrgs").Value = arr(36)
  Fields("fldMTPaid").Value = arr(37)
  Fields("fldMTPct").Value = arr(38)
  Fields("fldMCChrgs").Value = arr(39)
  Fields("fldMCPaid").Value = arr(40)
  Fields("fldMCPct").Value = arr(41)
  Fields("fldFEChrgs").Value = arr(42)
  Fields("fldFEPaid").Value = arr(43)
  Fields("fldFEPct").Value = arr(44)
  Fields("fldMHChrgs").Value = arr(45)
  Fields("fldMHPaid").Value = arr(46)
  Fields("fldMHPct").Value = arr(47)
  Fields("fldPIntChrgs").Value = arr(48)
  Fields("fldPIntPaid").Value = arr(49)
  Fields("fldPIntPct").Value = arr(50)
  Fields("fldPPenChrgs").Value = arr(51)
  Fields("fldPPenPaid").Value = arr(52)
  Fields("fldPPenPct").Value = arr(53)
  Fields("fldPOpt1Chrgs").Value = arr(54)
  Fields("fldPOPt1Paid").Value = arr(55)
  Fields("fldPOpt1Pct").Value = arr(56)
  Fields("fldPOpt2Chrgs").Value = arr(57)
  Fields("fldPOPt2Paid").Value = arr(58)
  Fields("fldPOpt2Pct").Value = arr(59)
  Fields("fldPOpt3Chrgs").Value = arr(60)
  Fields("fldPOPt3Paid").Value = arr(61)
  Fields("fldPOpt3Pct").Value = arr(62)
  Fields("fldGTPrincChrgs").Value = arr(63)
  Fields("fldGTPrincPaid").Value = arr(64)
  Fields("fldGTPrincPct").Value = arr(65)
  Fields("fldGTRIntChrgs").Value = arr(66)
  Fields("fldGTRIntPaid").Value = arr(67)
  Fields("fldGTRIntPct").Value = arr(68)
  Fields("fldGTAdvChrgs").Value = arr(69)
  Fields("fldGTAdvPaid").Value = arr(70)
  Fields("fldGTAdvPct").Value = arr(71)
  Fields("fldGTLateListChrgs").Value = arr(72)
  Fields("fldGTLateListPaid").Value = arr(73)
  Fields("fldGTLateListPct").Value = arr(74)
  Fields("fldGTRPenChrgs").Value = arr(75)
  Fields("fldGTRPenPaid").Value = arr(76)
  Fields("fldGTRPenPct").Value = arr(77)
  Fields("fldGTROpt1Chrgs").Value = arr(78)
  Fields("fldGTROpt1Paid").Value = arr(79)
  Fields("fldGTROpt1Pct").Value = arr(80)
  Fields("fldGTROpt2Chrgs").Value = arr(81)
  Fields("fldGTROpt2Paid").Value = arr(82)
  Fields("fldGTROpt2Pct").Value = arr(83)
  Fields("fldGTROpt3Chrgs").Value = arr(84)
  Fields("fldGTROpt3Paid").Value = arr(85)
  Fields("fldGTROpt3Pct").Value = arr(86)
  Fields("fldGTPersChrgs").Value = arr(87)
  Fields("fldGTPersPaid").Value = arr(88)
  Fields("fldGTPersPct").Value = arr(89)
  Fields("fldGTMTChrgs").Value = arr(90)
  Fields("fldGTMTPaid").Value = arr(91)
  Fields("fldGTMTPct").Value = arr(92)
  Fields("fldGTMCChrgs").Value = arr(93)
  Fields("fldGTMCPaid").Value = arr(94)
  Fields("fldGTMCPct").Value = arr(95)
  Fields("fldGTFEChrgs").Value = arr(96)
  Fields("fldGTFEPaid").Value = arr(97)
  Fields("fldGTFEPct").Value = arr(98)
  Fields("fldGTMHChrgs").Value = arr(99)
  Fields("fldGTMHPaid").Value = arr(100)
  Fields("fldGTMHPct").Value = arr(101)
  Fields("fldGTPIntChrgs").Value = arr(102)
  Fields("fldGTPIntPaid").Value = arr(103)
  Fields("fldGTPIntPct").Value = arr(104)
  Fields("fldGTPPenChrgs").Value = arr(105)
  Fields("fldGTPPenPaid").Value = arr(106)
  Fields("fldGTPPenPct").Value = arr(107)
  Fields("fldGTPOpt1Chrgs").Value = arr(108)
  Fields("fldGTPOpt1Paid").Value = arr(109)
  Fields("fldGTPOpt1Pct").Value = arr(110)
  Fields("fldGTPOpt2Chrgs").Value = arr(111)
  Fields("fldGTPOpt2Paid").Value = arr(112)
  Fields("fldGTPOpt2Pct").Value = arr(113)
  Fields("fldGTPOpt3Chrgs").Value = arr(114)
  Fields("fldGTPOpt3Paid").Value = arr(115)
  Fields("fldGTPOpt3Pct").Value = arr(116)
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
      frmVATaxMsg.Label1.Caption = "File - CollRateDetRpt.xls, created in the Citipak Directory."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show vbModal
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      frmVATaxMsg.Label1.Caption = "File - CollRateDetRpt.txt, created in the Citipak Directory."
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
    frmVATaxMsg.Label1.Caption = "File - CollRateDetRpt.xls, created in the Citipak Directory."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    frmVATaxMsg.Label1.Caption = "File - CollRateDetRpt.txt, created in the Citipak Directory."
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
        oEXL.FileName = outfile & "CollRateDetRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "CollRateDetRpt.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub ActiveReport_ReportEnd()
  Unload frmVATaxLoadReport
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
  End If
End Sub

Private Sub ActiveReport_ReportStart()
  Me.fldTimeDate.Text = Date
  Me.Zoom = -1
End Sub


