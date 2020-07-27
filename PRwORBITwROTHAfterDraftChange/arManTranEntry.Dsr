VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arManTranEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manual Transaction Register"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "arManTranEntry.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20558
   _ExtentY        =   15637
   SectionData     =   "arManTranEntry.dsx":08CA
End
Attribute VB_Name = "arManTranEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Private hFile As Integer
Dim EndReport As Boolean
Dim DedCnt As Integer
Private Sub ActiveReport_Initialize()
  Me.ToolBar.Tools.Add "&Close"
  Me.ToolBar.Tools.Add "Save/&Excel"
  Me.ToolBar.Tools.Add "&Text"
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
      MsgBox "File - ManualRegisterRpt.xls, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - ManualRegisterRpt.txt, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    End If
  End If
End Sub

Private Sub ActiveReport_Terminate()
  Close
End Sub

Private Sub ActiveReport_ToolbarClick(ByVal Tool As DDActiveReports2.DDTool)
  If Tool = "&Close" Then
    Unload Me
  End If
  If Tool = "Save/&Excel" Then
    Screen.MousePointer = vbHourglass
    ExportReport 1
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - ManualRegisterRpt.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - ManualRegisterRpt.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "ManualRegisterRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "ManualRegisterRpt.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  EndReport = False
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
    Unload frmLoadingRpt
    MsgBox "Error Number: " & Str(Number) & " " & Description, vbOKOnly, "Printer Error"
    Unload Me
  End If
  CancelDisplay = True 'removes the error message
End Sub
Private Sub ActiveReport_DataInitialize()
  hFile = FreeFile
  Open StartPath & "\PRRPTS\MANREGISG.RPT" For Input As #hFile
  Fields.Add "fldEmployer" '(0)
  Fields.Add "fldDate1" '(1)
  Fields.Add "fldEmpNum" '(2)
  Fields.Add "fldEmpName" '(3)
  Fields.Add "fldDate2" '(4)
  Fields.Add "fldDate3" '(5)
  Fields.Add "fldDate4" '(6)
  Fields.Add "fldCheckNum" '(7)
  Fields.Add "fldBaseRate" '(8)
  Fields.Add "fldOTRate" '(9)
  Fields.Add "fldRegHrs" '(10)
  Fields.Add "fldVacHrs" '(11)
  Fields.Add "fldSickHrs" '(12)
  Fields.Add "fldHolHrs" '(13)
  Fields.Add "fldCompHrs" '(14)
  Fields.Add "fldPersHrs" '(15)
  Fields.Add "fldTotHrs" '(16)
  Fields.Add "fldOTHrs" '(17)
  Fields.Add "fldOTPaid" '(18)
  Fields.Add "fldOTComp" '(19)
  Fields.Add "fldRegEarn" '(20)
  Fields.Add "fldOTEarn" '(21)
  Fields.Add "fldGrsPay" '(22)
  Fields.Add "fldEIC" '(23)
  Fields.Add "fldSSTax" '(24)
  Fields.Add "fldMedTax" '(25)
  Fields.Add "fldFWT" '(26)
  Fields.Add "fldSWT" '(27)
  Fields.Add "fldRetTax" '(28)
  Fields.Add "fldNetPay" '(29)
  Fields.Add "fldDedVal1Det" '(30)
  Fields.Add "fldDedVal2Det" '(31)
  Fields.Add "fldDedVal3Det" '(32)
  Fields.Add "fldDedVal4Det" '(33)
  Fields.Add "fldDedVal5Det" '(34)
  Fields.Add "fldDedVal6Det" '(35)
  Fields.Add "fldDedVal7Det" '(36)
  Fields.Add "fldDedVal8Det" '(37)
  Fields.Add "fldDedVal9Det" '(38)
  Fields.Add "fldDedVal10Det" '(39)
  Fields.Add "fldDedVal11Det" '(40)
  Fields.Add "fldDedVal12Det" '(41)
  Fields.Add "fldDedVal13Det" '(42)
  Fields.Add "fldDedVal14Det" '(43)
  Fields.Add "fldDedVal15Det" '(44)
  Fields.Add "fldDedVal16Det" '(45)
  Fields.Add "fldDedVal17Det" '(46)
  Fields.Add "fldDedVal18Det" '(47)
  Fields.Add "fldDedVal19Det" '(48)
  Fields.Add "fldDedVal20Det" '(49)
  Fields.Add "fldDedVal21Det" '(50)
  Fields.Add "fldDedVal22Det" '(51)
  Fields.Add "fldDedVal23Det" '(52)
  Fields.Add "fldDedVal24Det" '(53)
  Fields.Add "fldDedVal25Det" '(54)
  Fields.Add "fldDedVal26Det" '(55)
  Fields.Add "fldDedVal27Det" '(56)
  Fields.Add "fldDedVal28Det" '(57)
  Fields.Add "fldDedVal29Det" '(58)
  Fields.Add "fldDedVal30Det" '(59)
  Fields.Add "fldDedVal31Det" '(60)
  Fields.Add "fldDedVal32Det" '(61)
  Fields.Add "fldDedVal33Det" '(62)
  Fields.Add "fldDedVal34Det" '(63)
  Fields.Add "fldDedVal35Det" '(64)
  Fields.Add "fldDedVal36Det" '(65)
  Fields.Add "fldDedVal37Det" '(66)
  Fields.Add "fldDedVal38Det" '(67)
  Fields.Add "fldDedVal39Det" '(68)
  Fields.Add "fldDedVal40Det" '(69)
  Fields.Add "fldDedVal41Det" '(70)
  Fields.Add "fldDedVal42Det" '(71)
  Fields.Add "fldDedVal43Det" '(72)
  Fields.Add "fldDedVal44Det" '(73)
  Fields.Add "fldDedVal45Det" '(74)
  Fields.Add "fldDedVal46Det" '(75)
  Fields.Add "fldDedVal47Det" '(76)
  Fields.Add "fldDedVal48Det" '(77)
  Fields.Add "fldDedVal49Det" '(78)
  Fields.Add "fldDedVal50Det" '(79)
  Fields.Add "fldDAcct1" '(80)
  Fields.Add "fldDAcct2" '(81)
  Fields.Add "fldDAcct3" '(82)
  Fields.Add "fldDAcct4" '(83)
  Fields.Add "fldFedGrs" '(84)
  Fields.Add "fldStateGrs" '(85)
  Fields.Add "fldSocGrs" '(86)
  Fields.Add "fldMedGrs" '(87)
  Fields.Add "fldDAmts1" '(88)
  Fields.Add "fldDAmts2" '(89)
  Fields.Add "fldDAmts3" '(90)
  Fields.Add "fldDAmts4" '(91)
  Fields.Add "fldRetGrs" '(92)
  Fields.Add "fldSalNum" '(93)
  Fields.Add "fldHrNum" '(94)
  Fields.Add "fldRegHrsttl" '(95)
  Fields.Add "fldVacHrsttl" '(96)
  Fields.Add "fldSickHrsttl" '(97)
  Fields.Add "fldHolHrsttl" '(98)
  Fields.Add "fldCompHrsttl" '(99)
  Fields.Add "fldPersHrsttl" '(100)
  Fields.Add "fldTotHrsttl" '(101)
  Fields.Add "fldOTHrsttl" '(102)
  Fields.Add "fldOTPaidttl" '(103)
  Fields.Add "fldOTCompttl" '(104)
  Fields.Add "fldRegEarnttl" '(105)
  Fields.Add "fldOTEarnttl" '(106)
  Fields.Add "fldGrsPayttl" '(107)
  Fields.Add "fldEICttl" '(108)
  Fields.Add "fldSSTaxttl" '(109)
  Fields.Add "fldMedTaxttl" '(110)
  Fields.Add "fldFedTaxttl" '(111)
  Fields.Add "fldStateTaxttl" '(112)
  Fields.Add "fldRetTaxttl" '(113)
  Fields.Add "fldNetPayttl" '(114)
  
  Fields.Add "fldDedVal1ttl" '(115)
  Fields.Add "fldDedVal2ttl" '(116)
  Fields.Add "fldDedVal3ttl" '(117)
  Fields.Add "fldDedVal4ttl" '(118)
  Fields.Add "fldDedVal5ttl" '(119)
  Fields.Add "fldDedVal6ttl" '(120)
  Fields.Add "fldDedVal7ttl" '(121)
  Fields.Add "fldDedVal8ttl" '(122)
  Fields.Add "fldDedVal9ttl" '(123)
  Fields.Add "fldDedVal10ttl" '(124)
  Fields.Add "fldDedVal11ttl" '(125)
  Fields.Add "fldDedVal12ttl" '(126)
  Fields.Add "fldDedVal13ttl" '(127)
  Fields.Add "fldDedVal14ttl" '(128)
  Fields.Add "fldDedVal15ttl" '(129)
  Fields.Add "fldDedVal16ttl" '(130)
  Fields.Add "fldDedVal17ttl" '(131)
  Fields.Add "fldDedVal18ttl" '(132)
  Fields.Add "fldDedVal19ttl" '(133)
  Fields.Add "fldDedVal20ttl" '(134)
  Fields.Add "fldDedVal21ttl" '(135)
  Fields.Add "fldDedVal22ttl" '(136)
  Fields.Add "fldDedVal23ttl" '(137)
  Fields.Add "fldDedVal24ttl" '(138)
  Fields.Add "fldDedVal25ttl" '(139)
  Fields.Add "fldDedVal26ttl" '(140)
  Fields.Add "fldDedVal27ttl" '(141)
  Fields.Add "fldDedVal28ttl" '(142)
  Fields.Add "fldDedVal29ttl" '(143)
  Fields.Add "fldDedVal30ttl" '(144)
  Fields.Add "fldDedVal31ttl" '(145)
  Fields.Add "fldDedVal32ttl" '(146)
  Fields.Add "fldDedVal33ttl" '(147)
  Fields.Add "fldDedVal34ttl" '(148)
  Fields.Add "fldDedVal35ttl" '(149)
  Fields.Add "fldDedVal36ttl" '(150)
  Fields.Add "fldDedVal37ttl" '(151)
  Fields.Add "fldDedVal38ttl" '(152)
  Fields.Add "fldDedVal39ttl" '(153)
  Fields.Add "fldDedVal40ttl" '(154)
  Fields.Add "fldDedVal41ttl" '(155)
  Fields.Add "fldDedVal42ttl" '(156)
  Fields.Add "fldDedVal43ttl" '(157)
  Fields.Add "fldDedVal44ttl" '(158)
  Fields.Add "fldDedVal45ttl" '(159)
  Fields.Add "fldDedVal46ttl" '(160)
  Fields.Add "fldDedVal47ttl" '(161)
  Fields.Add "fldDedVal48ttl" '(162)
  Fields.Add "fldDedVal49ttl" '(163)
  Fields.Add "fldDedVal50ttl" '(164)
  Fields.Add "fldFedGrsttl" '(165)
  Fields.Add "fldStateGrsttl" '(166)
  Fields.Add "fldMedGrsttl" '(167)
  Fields.Add "fldSocGrsttl" '(168)
  Fields.Add "fldRetGrsttl" '(169)
  Fields.Add "fldEmpty" '(170)
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
  Fields("fldEmployer").Value = arr(0)
  Fields("fldDate1").Value = arr(1)
  Fields("fldEmpNum").Value = arr(2)
  Fields("fldEmpName").Value = arr(3)
  Fields("fldDate2").Value = arr(4)
  Fields("fldDate3").Value = arr(5)
  Fields("fldDate4").Value = arr(6)
  Fields("fldCheckNum").Value = arr(7)
  Fields("fldBaseRate").Value = arr(8)
  Fields("fldOTRate").Value = arr(9)
  Fields("fldRegHrs").Value = arr(10)
  Fields("fldVacHrs").Value = arr(11)
  Fields("fldSickHrs").Value = arr(12)
  Fields("fldHolHrs").Value = arr(13)
  Fields("fldCompHrs").Value = arr(14)
  Fields("fldPersHrs").Value = arr(15)
  Fields("fldTotHrs").Value = arr(16)
  Fields("fldOTHrs").Value = arr(17)
  Fields("fldOTPaid").Value = arr(18)
  Fields("fldOTComp").Value = arr(19)
  Fields("fldRegEarn").Value = arr(20)
  Fields("fldOTEarn").Value = arr(21)
  Fields("fldGrsPay").Value = arr(22)
  Fields("fldEIC").Value = arr(23)
  Fields("fldSSTax").Value = arr(24)
  Fields("fldMedTax").Value = arr(25)
  Fields("fldFWT").Value = arr(26)
  Fields("fldSWT").Value = arr(27)
  Fields("fldRetTax").Value = arr(28)
  Fields("fldNetPay").Value = arr(29)
  Fields("fldDedVal1Det").Value = arr(30)
  Fields("fldDedVal2Det").Value = arr(31)
  Fields("fldDedVal3Det").Value = arr(32)
  Fields("fldDedVal4Det").Value = arr(33)
  Fields("fldDedVal5Det").Value = arr(34)
  Fields("fldDedVal6Det").Value = arr(35)
  Fields("fldDedVal7Det").Value = arr(36)
  Fields("fldDedVal8Det").Value = arr(37)
  Fields("fldDedVal9Det").Value = arr(38)
  Fields("fldDedVal10Det").Value = arr(39)
  Fields("fldDedVal11Det").Value = arr(40)
  Fields("fldDedVal12Det").Value = arr(41)
  Fields("fldDedVal13Det").Value = arr(42)
  Fields("fldDedVal14Det").Value = arr(43)
  Fields("fldDedVal15Det").Value = arr(44)
  Fields("fldDedVal16Det").Value = arr(45)
  Fields("fldDedVal17Det").Value = arr(46)
  Fields("fldDedVal18Det").Value = arr(47)
  Fields("fldDedVal19Det").Value = arr(48)
  Fields("fldDedVal20Det").Value = arr(49)
  Fields("fldDedVal21Det").Value = arr(50)
  Fields("fldDedVal22Det").Value = arr(51)
  Fields("fldDedVal23Det").Value = arr(52)
  Fields("fldDedVal24Det").Value = arr(53)
  Fields("fldDedVal25Det").Value = arr(54)
  Fields("fldDedVal26Det").Value = arr(55)
  Fields("fldDedVal27Det").Value = arr(56)
  Fields("fldDedVal28Det").Value = arr(57)
  Fields("fldDedVal29Det").Value = arr(58)
  Fields("fldDedVal30Det").Value = arr(59)
  Fields("fldDedVal31Det").Value = arr(60)
  Fields("fldDedVal32Det").Value = arr(61)
  Fields("fldDedVal33Det").Value = arr(62)
  Fields("fldDedVal34Det").Value = arr(63)
  Fields("fldDedVal35Det").Value = arr(64)
  Fields("fldDedVal36Det").Value = arr(65)
  Fields("fldDedVal37Det").Value = arr(66)
  Fields("fldDedVal38Det").Value = arr(67)
  Fields("fldDedVal39Det").Value = arr(68)
  Fields("fldDedVal40Det").Value = arr(69)
  Fields("fldDedVal41Det").Value = arr(70)
  Fields("fldDedVal42Det").Value = arr(71)
  Fields("fldDedVal43Det").Value = arr(72)
  Fields("fldDedVal44Det").Value = arr(73)
  Fields("fldDedVal45Det").Value = arr(74)
  Fields("fldDedVal46Det").Value = arr(75)
  Fields("fldDedVal47Det").Value = arr(76)
  Fields("fldDedVal48Det").Value = arr(77)
  Fields("fldDedVal49Det").Value = arr(78)
  Fields("fldDedVal50Det").Value = arr(79)
  Fields("fldDAcct1").Value = QPTrim$(arr(80))
  Fields("fldDAcct2").Value = QPTrim$(arr(81))
  Fields("fldDAcct3").Value = QPTrim$(arr(82))
  Fields("fldDAcct4").Value = QPTrim$(arr(83))
  Fields("fldFedGrs").Value = arr(84)
  Fields("fldStateGrs").Value = arr(85)
  Fields("fldSocGrs").Value = arr(86)
  Fields("fldMedGrs").Value = arr(87)
  If QPTrim$(arr(80)) = "" Then arr(88) = ""
  Fields("fldDAmts1").Value = arr(88)
  If QPTrim$(arr(81)) = "" Then arr(89) = ""
  Fields("fldDAmts2").Value = arr(89)
  If QPTrim$(arr(82)) = "" Then arr(90) = ""
  Fields("fldDAmts3").Value = arr(90)
  If QPTrim$(arr(83)) = "" Then arr(91) = ""
  Fields("fldDAmts4").Value = arr(91)
  Fields("fldRetGrs").Value = arr(92)
  Fields("fldSalNum").Value = arr(93)
  Fields("fldHrNum").Value = arr(94)
  Fields("fldRegHrsttl").Value = arr(95)
  Fields("fldVacHrsttl").Value = arr(96)
  Fields("fldSickHrsttl").Value = arr(97)
  Fields("fldHolHrsttl").Value = arr(98)
  Fields("fldCompHrsttl").Value = arr(99)
  Fields("fldPersHrsttl").Value = arr(100)
  Fields("fldTotHrsttl").Value = arr(101)
  Fields("fldOTHrsttl").Value = arr(102)
  Fields("fldOTPaidttl").Value = arr(103)
  Fields("fldOTCompttl").Value = arr(104)
  Fields("fldRegEarnttl").Value = arr(105)
  Fields("fldOTEarnttl").Value = arr(106)
  Fields("fldGrsPayttl").Value = arr(107)
  Fields("fldEICttl").Value = arr(108)
  Fields("fldSSTaxttl").Value = arr(109)
  Fields("fldMedTaxttl").Value = arr(110)
  Fields("fldFedTaxttl").Value = arr(111)
  Fields("fldStateTaxttl").Value = arr(112)
  Fields("fldRetTaxttl").Value = arr(113)
  Fields("fldNetPayttl").Value = arr(114)
  
  Fields("fldDedVal1ttl").Value = arr(115)
  Fields("fldDedVal2ttl").Value = arr(116)
  Fields("fldDedVal3ttl").Value = arr(117)
  Fields("fldDedVal4ttl").Value = arr(118)
  Fields("fldDedVal5ttl").Value = arr(119)
  Fields("fldDedVal6ttl").Value = arr(120)
  Fields("fldDedVal7ttl").Value = arr(121)
  Fields("fldDedVal8ttl").Value = arr(122)
  Fields("fldDedVal9ttl").Value = arr(123)
  Fields("fldDedVal10ttl").Value = arr(124)
  Fields("fldDedVal11ttl").Value = arr(125)
  Fields("fldDedVal12ttl").Value = arr(126)
  Fields("fldDedVal13ttl").Value = arr(127)
  Fields("fldDedVal14ttl").Value = arr(128)
  Fields("fldDedVal15ttl").Value = arr(129)
  Fields("fldDedVal16ttl").Value = arr(130)
  Fields("fldDedVal17ttl").Value = arr(131)
  Fields("fldDedVal18ttl").Value = arr(132)
  Fields("fldDedVal19ttl").Value = arr(133)
  Fields("fldDedVal20ttl").Value = arr(134)
  Fields("fldDedVal21ttl").Value = arr(135)
  Fields("fldDedVal22ttl").Value = arr(136)
  Fields("fldDedVal23ttl").Value = arr(137)
  Fields("fldDedVal24ttl").Value = arr(138)
  Fields("fldDedVal25ttl").Value = arr(139)
  Fields("fldDedVal26ttl").Value = arr(140)
  Fields("fldDedVal27ttl").Value = arr(141)
  Fields("fldDedVal28ttl").Value = arr(142)
  Fields("fldDedVal29ttl").Value = arr(143)
  Fields("fldDedVal30ttl").Value = arr(144)
  Fields("fldDedVal31ttl").Value = arr(145)
  Fields("fldDedVal32ttl").Value = arr(146)
  Fields("fldDedVal33ttl").Value = arr(147)
  Fields("fldDedVal34ttl").Value = arr(148)
  Fields("fldDedVal35ttl").Value = arr(149)
  Fields("fldDedVal36ttl").Value = arr(150)
  Fields("fldDedVal37ttl").Value = arr(151)
  Fields("fldDedVal38ttl").Value = arr(152)
  Fields("fldDedVal39ttl").Value = arr(153)
  Fields("fldDedVal40ttl").Value = arr(154)
  Fields("fldDedVal41ttl").Value = arr(155)
  Fields("fldDedVal42ttl").Value = arr(156)
  Fields("fldDedVal43ttl").Value = arr(157)
  Fields("fldDedVal44ttl").Value = arr(158)
  Fields("fldDedVal45ttl").Value = arr(159)
  Fields("fldDedVal46ttl").Value = arr(160)
  Fields("fldDedVal47ttl").Value = arr(161)
  Fields("fldDedVal48ttl").Value = arr(162)
  Fields("fldDedVal49ttl").Value = arr(163)
  Fields("fldDedVal50ttl").Value = arr(164)
  Fields("fldFedGrsttl").Value = arr(165)
  Fields("fldStateGrsttl").Value = arr(166)
  Fields("fldMedGrsttl").Value = arr(167)
  Fields("fldSocGrsttl").Value = arr(168)
  Fields("fldRetGrsttl").Value = arr(169)
  Fields("fldEmpty").Value = arr(170)
End Sub
Private Sub ActiveReport_ReportEnd()
  Unload frmLoadingRpt
  If hFile <> 0 Then
    Close #hFile
  End If
End Sub

Private Sub ActiveReport_ReportStart()
  Dim DedRec As DedCodeRecType
  Dim DHandle As Integer
  Dim x As Integer
  Me.Zoom = -1
  
  OpenDedCodeFile DHandle
  DedCnt = LOF(DHandle) / Len(DedRec)
  Close DHandle
  'for all but the last page
  Select Case DedCnt
    Case 1 To 10
'      PageHeader.Height = 3100
      Detail.Height = 1800
    Case 11 To 20
'      PageHeader.Height = 3100
      Detail.Height = 1800
    Case 21 To 30
'      PageHeader.Height = 2600
      Detail.Height = 2250
    Case 31 To 40
'      PageHeader.Height = 2850
      Detail.Height = 2500
      SubReport1.Height = 1100
    Case 41 To 50
'      PageHeader.Height = 5000
'      Detail.Height = 2700
      SubReport1.Height = 1400
    Case Else
'      Detail.Height = 2700
      SubReport1.Height = 1400
  End Select
  Me.PageSettings.RightMargin = 300
  Me.PageSettings.LeftMargin = 200
  Me.fldTimeDate.Text = Now
  GroupFooter1.Height = 0
  PageFooter.Height = 0
End Sub

Private Sub Detail_Format()
  Dim ctrl As Control
  Dim sec As Section
  Dim y As Integer
  Dim x As Integer
  Set sec = arManTranEntry.Sections("Detail")
  x = 1
  
  For y = 0 To sec.Controls.Count - 1
    sec.Controls(y).Visible = True
    If y <= 16 Or y >= 67 Then
      GoTo NotNow
    End If
    If x > DedCnt Then
      sec.Controls(y).Visible = False
      x = x + 1
    Else
      sec.Controls(y).Visible = True
      x = x + 1
    End If
NotNow:
  Next y

End Sub

Private Sub PageHeader_Format()
  Set SubReport1.object = New arManDedDesc
  If EndReport = True Then
    PageHeader.Visible = False
  End If
End Sub

Private Sub ReportFooter_Format()
  Set SubReport2.object = New arManDedDesc
  Me.fldTimeDaterf.Text = Now
  
  EndReport = True
  Dim ctrl As Control
  Dim sec As Section
  Dim y As Integer
  Dim x As Integer
  Set sec = arManTranEntry.Sections("ReportFooter")
  x = 1
  
  For y = 0 To sec.Controls.Count - 1
    sec.Controls(y).Visible = True
    If y <= 16 Or y >= 67 Then
      GoTo NotNow
    End If
    If x > DedCnt Then
      sec.Controls(y).Visible = False
      x = x + 1
    Else
      sec.Controls(y).Visible = True
      x = x + 1
    End If
NotNow:
  Next y

End Sub



