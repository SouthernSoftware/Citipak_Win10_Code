VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arPayRollRegisterNS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Register"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "arPRRegisterNS.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20558
   _ExtentY        =   15637
   SectionData     =   "arPRRegisterNS.dsx":08CA
End
Attribute VB_Name = "arPayRollRegisterNS"
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
Dim NoEscape As Boolean
Private Sub ActiveReport_Initialize()
  Me.ToolBar.Tools.Add "&Close"
  Me.ToolBar.Tools.Add "Save/&Excel"
  Me.ToolBar.Tools.Add "&Text"
End Sub

Private Sub ActiveReport_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
  If KeyCode = vbKeyEscape And NoEscape = False Then
    NoEscape = True
    DoEvents
    Unload Me
    DoEvents
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
      MsgBox "File - EarningsRegisterNSRpt.xls, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - EarningsRegisterNSRpt.txt, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    End If
  End If
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
    MsgBox "File - EarningsRegisterNSRpt.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - EarningsRegisterNSRpt.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "EarningsRegisterRptNS.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "EarningsRegisterRptNS.txt"
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
  NoEscape = False
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
  Open StartPath & "\PRRPTS\REGISTERNSG.RPT" For Input As #hFile
  Fields.Add ("fldEmployer") '(0)
  Fields.Add ("fldDates") '(1)
  Fields.Add ("fldEmpNum") '(2)
  Fields.Add ("fldEmpName") '(3)
  
  Fields.Add ("fldBaseRate") '(4)
  Fields.Add ("fldOTRate") '(5)
  Fields.Add ("fldTaxFr") '(6)
  Fields.Add ("fldRegHrsDet") '(7)
  Fields.Add ("fldVacDet") '(8)
  Fields.Add ("fldSickDet") '(9)
  Fields.Add ("fldHolDet") '(10)
  Fields.Add ("fldCompDet") '(11)
  Fields.Add ("fldPersDet") '(12)
  Fields.Add ("fldTotHrsDet") '(13)
  Fields.Add ("fldOTPaidDet") '(14)
  Fields.Add ("fldOTComp") '(15)
  Fields.Add ("fldRegEarnDet") '(16)
  Fields.Add ("fldOTEarnDet") '(17)
  Fields.Add ("fldEarn3Det") '(18)
  Fields.Add ("fldEarn2Det") '(19)
  Fields.Add ("fldEarn1Det") '(20)
  Fields.Add ("fldEarnDsc3") '(21)
  Fields.Add ("fldEarnDsc2") '(22)
  Fields.Add ("fldEarnDsc1") '(23)
  
  Fields.Add ("fldGrossPayDet") '(24)
  Fields.Add ("fldSocSecDet") '(25)
  Fields.Add ("fldMedDet") '(26)
  Fields.Add ("fldFWTDet") '(27)
  Fields.Add ("fldSWTDet") '(28)
  Fields.Add ("fldRetDet") '(29)
  Fields.Add ("fldNetPayDet") '(30)
  Fields.Add ("fldEICDet") '(31)
  
  Fields.Add ("fldDedDsc1") '(32)
  Fields.Add ("fldDedDsc2") '(33)
  Fields.Add ("fldDedDsc3") '(34)
  Fields.Add ("fldDedDsc4") '(35)
  Fields.Add ("fldDedDsc5") '(36)
  Fields.Add ("fldDedDsc6") '(37)
  Fields.Add ("fldDedDsc7") '(38)
  Fields.Add ("fldDedDsc8") '(39)
  Fields.Add ("fldDedDsc9") '(40)
  Fields.Add ("fldDedDsc10") '(41)
  Fields.Add ("fldDedDsc11") '(42)
  Fields.Add ("fldDedDsc12") '(43)
  Fields.Add ("fldDedDsc13") '(44)
  Fields.Add ("fldDedDsc14") '(45)
  Fields.Add ("fldDedDsc15") '(46)
  Fields.Add ("fldDedDsc16") '(47)
  Fields.Add ("fldDedDsc17") '(48)
  Fields.Add ("fldDedDsc18") '(49)
  Fields.Add ("fldDedDsc19") '(50)
  Fields.Add ("fldDedDsc20") '(51)
  Fields.Add ("fldDedDsc21") '(52)
  Fields.Add ("fldDedDsc22") '(53)
  Fields.Add ("fldDedDsc23") '(54)
  Fields.Add ("fldDedDsc24") '(55)
  Fields.Add ("fldDedDsc25") '(56)
  Fields.Add ("fldDedDsc26") '(57)
  Fields.Add ("fldDedDsc27") '(58)
  Fields.Add ("fldDedDsc28") '(59)
  Fields.Add ("fldDedDsc29") '(60)
  Fields.Add ("fldDedDsc30") '(61)
  Fields.Add ("fldDedDsc31") '(62)
  Fields.Add ("fldDedDsc32") '(63)
  Fields.Add ("fldDedDsc33") '(64)
  Fields.Add ("fldDedDsc34") '(65)
  Fields.Add ("fldDedDsc35") '(66)
  Fields.Add ("fldDedDsc36") '(67)
  Fields.Add ("fldDedDsc37") '(68)
  Fields.Add ("fldDedDsc38") '(69)
  Fields.Add ("fldDedDsc39") '(70)
  Fields.Add ("fldDedDsc40") '(71)
  Fields.Add ("fldDedDsc41") '(72)
  Fields.Add ("fldDedDsc42") '(73)
  Fields.Add ("fldDedDsc43") '(74)
  Fields.Add ("fldDedDsc44") '(75)
  Fields.Add ("fldDedDsc45") '(76)
  Fields.Add ("fldDedDsc46") '(77)
  Fields.Add ("fldDedDsc47") '(78)
  Fields.Add ("fldDedDsc48") '(79)
  Fields.Add ("fldDedDsc49") '(80)
  Fields.Add ("fldDedDsc50") '(81)
  Fields.Add ("fldDedVal1Det") '(82)
  Fields.Add ("fldDedVal2Det") '(83)
  Fields.Add ("fldDedVal3Det") '(84)
  Fields.Add ("fldDedVal4Det") '(85)
  Fields.Add ("fldDedVal5Det") '(86)
  Fields.Add ("fldDedVal6Det") '(87)
  Fields.Add ("fldDedVal7Det") '(88)
  Fields.Add ("fldDedVal8Det") '(89)
  Fields.Add ("fldDedVal9Det") '(90)
  Fields.Add ("fldDedVal10Det") '(91)
  Fields.Add ("fldDedVal11Det") '(92)
  Fields.Add ("fldDedVal12Det") '(93)
  Fields.Add ("fldDedVal13Det") '(94)
  Fields.Add ("fldDedVal14Det") '(95)
  Fields.Add ("fldDedVal15Det") '(96)
  Fields.Add ("fldDedVal16Det") '(97)
  Fields.Add ("fldDedVal17Det") '(98)
  Fields.Add ("fldDedVal18Det") '(99)
  Fields.Add ("fldDedVal19Det") '(100)
  Fields.Add ("fldDedVal20Det") '(101)
  Fields.Add ("fldDedVal21Det") '(102)
  Fields.Add ("fldDedVal22Det") '(103)
  Fields.Add ("fldDedVal23Det") '(104)
  Fields.Add ("fldDedVal24Det") '(105)
  Fields.Add ("fldDedVal25Det") '(106)
  Fields.Add ("fldDedVal26Det") '(107)
  Fields.Add ("fldDedVal27Det") '(108)
  Fields.Add ("fldDedVal28Det") '(109)
  Fields.Add ("fldDedVal29Det") '(110)
  Fields.Add ("fldDedVal30Det") '(111)
  Fields.Add ("fldDedVal31Det") '(112)
  Fields.Add ("fldDedVal32Det") '(113)
  Fields.Add ("fldDedVal33Det") '(114)
  Fields.Add ("fldDedVal34Det") '(115)
  Fields.Add ("fldDedVal35Det") '(116)
  Fields.Add ("fldDedVal36Det") '(117)
  Fields.Add ("fldDedVal37Det") '(118)
  Fields.Add ("fldDedVal38Det") '(119)
  Fields.Add ("fldDedVal39Det") '(120)
  Fields.Add ("fldDedVal40Det") '(121)
  Fields.Add ("fldDedVal41Det") '(122)
  Fields.Add ("fldDedVal42Det") '(123)
  Fields.Add ("fldDedVal43Det") '(124)
  Fields.Add ("fldDedVal44Det") '(125)
  Fields.Add ("fldDedVal45Det") '(126)
  Fields.Add ("fldDedVal46Det") '(127)
  Fields.Add ("fldDedVal47Det") '(128)
  Fields.Add ("fldDedVal48Det") '(129)
  Fields.Add ("fldDedVal49Det") '(130)
  Fields.Add ("fldDedVal50Det") '(131)
  Fields.Add ("fldNumOfDeds") '(132)
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
  Fields("fldDates").Value = arr(1)
  Fields("fldEmpNum").Value = arr(2)
  Fields("fldEmpName").Value = arr(3)
  
  Fields("fldBaseRate").Value = arr(4)
  Fields("fldOTRate").Value = arr(5)
  Fields("fldTaxFr").Value = arr(6)
  Fields("fldRegHrsDet").Value = arr(7)
  Fields("fldVacDet").Value = arr(8)
  Fields("fldSickDet").Value = arr(9)
  Fields("fldHolDet").Value = arr(10)
  Fields("fldCompDet").Value = arr(11)
  Fields("fldPersDet").Value = arr(12)
  Fields("fldTotHrsDet").Value = arr(13)
  Fields("fldOTPaidDet").Value = arr(14)
  Fields("fldOTComp").Value = arr(15)
  Fields("fldRegEarnDet").Value = arr(16)
  Fields("fldOTEarnDet").Value = arr(17)
  Fields("fldEarn3Det").Value = arr(18)
  Fields("fldEarn2Det").Value = arr(19)
  Fields("fldEarn1Det").Value = arr(20)
  Fields("fldEarnDsc3").Value = arr(21)
  Fields("fldEarnDsc2").Value = arr(22)
  Fields("fldEarnDsc1").Value = arr(23)
  
  Fields("fldGrossPayDet").Value = arr(24)
  Fields("fldSocSecDet").Value = arr(25)
  Fields("fldMedDet").Value = arr(26)
  Fields("fldFWTDet").Value = arr(27)
  Fields("fldSWTDet").Value = arr(28)
  Fields("fldRetDet").Value = arr(29)
  Fields("fldNetPayDet").Value = arr(30)
  Fields("fldEICDet").Value = arr(31)
  
  Fields("fldDedDsc1").Value = arr(32)
  Fields("fldDedDsc2").Value = arr(33)
  Fields("fldDedDsc3").Value = arr(34)
  Fields("fldDedDsc4").Value = arr(35)
  Fields("fldDedDsc5").Value = arr(36)
  Fields("fldDedDsc6").Value = arr(37)
  Fields("fldDedDsc7").Value = arr(38)
  Fields("fldDedDsc8").Value = arr(39)
  Fields("fldDedDsc9").Value = arr(40)
  Fields("fldDedDsc10").Value = arr(41)
  Fields("fldDedDsc11").Value = arr(42)
  Fields("fldDedDsc12").Value = arr(43)
  Fields("fldDedDsc13").Value = arr(44)
  Fields("fldDedDsc14").Value = arr(45)
  Fields("fldDedDsc15").Value = arr(46)
  Fields("fldDedDsc16").Value = arr(47)
  Fields("fldDedDsc17").Value = arr(48)
  Fields("fldDedDsc18").Value = arr(49)
  Fields("fldDedDsc19").Value = arr(50)
  Fields("fldDedDsc20").Value = arr(51)
  Fields("fldDedDsc21").Value = arr(52)
  Fields("fldDedDsc22").Value = arr(53)
  Fields("fldDedDsc23").Value = arr(54)
  Fields("fldDedDsc24").Value = arr(55)
  Fields("fldDedDsc25").Value = arr(56)
  Fields("fldDedDsc26").Value = arr(57)
  Fields("fldDedDsc27").Value = arr(58)
  Fields("fldDedDsc28").Value = arr(59)
  Fields("fldDedDsc29").Value = arr(60)
  Fields("fldDedDsc30").Value = arr(61)
  Fields("fldDedDsc31").Value = arr(62)
  Fields("fldDedDsc32").Value = arr(63)
  Fields("fldDedDsc33").Value = arr(64)
  Fields("fldDedDsc34").Value = arr(65)
  Fields("fldDedDsc35").Value = arr(66)
  Fields("fldDedDsc36").Value = arr(67)
  Fields("fldDedDsc37").Value = arr(68)
  Fields("fldDedDsc38").Value = arr(69)
  Fields("fldDedDsc39").Value = arr(70)
  Fields("fldDedDsc40").Value = arr(71)
  Fields("fldDedDsc41").Value = arr(72)
  Fields("fldDedDsc42").Value = arr(73)
  Fields("fldDedDsc43").Value = arr(74)
  Fields("fldDedDsc44").Value = arr(75)
  Fields("fldDedDsc45").Value = arr(76)
  Fields("fldDedDsc46").Value = arr(77)
  Fields("fldDedDsc47").Value = arr(78)
  Fields("fldDedDsc48").Value = arr(79)
  Fields("fldDedDsc49").Value = arr(80)
  Fields("fldDedDsc50").Value = arr(81)
  Fields("fldDedVal1Det").Value = arr(82)
  Fields("fldDedVal2Det").Value = arr(83)
  Fields("fldDedVal3Det").Value = arr(84)
  Fields("fldDedVal4Det").Value = arr(85)
  Fields("fldDedVal5Det").Value = arr(86)
  Fields("fldDedVal6Det").Value = arr(87)
  Fields("fldDedVal7Det").Value = arr(88)
  Fields("fldDedVal8Det").Value = arr(89)
  Fields("fldDedVal9Det").Value = arr(90)
  Fields("fldDedVal10Det").Value = arr(91)
  Fields("fldDedVal11Det").Value = arr(92)
  Fields("fldDedVal12Det").Value = arr(93)
  Fields("fldDedVal13Det").Value = arr(94)
  Fields("fldDedVal14Det").Value = arr(95)
  Fields("fldDedVal15Det").Value = arr(96)
  Fields("fldDedVal16Det").Value = arr(97)
  Fields("fldDedVal17Det").Value = arr(98)
  Fields("fldDedVal18Det").Value = arr(99)
  Fields("fldDedVal19Det").Value = arr(100)
  Fields("fldDedVal20Det").Value = arr(101)
  Fields("fldDedVal21Det").Value = arr(102)
  Fields("fldDedVal22Det").Value = arr(103)
  Fields("fldDedVal23Det").Value = arr(104)
  Fields("fldDedVal24Det").Value = arr(105)
  Fields("fldDedVal25Det").Value = arr(106)
  Fields("fldDedVal26Det").Value = arr(107)
  Fields("fldDedVal27Det").Value = arr(108)
  Fields("fldDedVal28Det").Value = arr(109)
  Fields("fldDedVal29Det").Value = arr(110)
  Fields("fldDedVal30Det").Value = arr(111)
  Fields("fldDedVal31Det").Value = arr(112)
  Fields("fldDedVal32Det").Value = arr(113)
  Fields("fldDedVal33Det").Value = arr(114)
  Fields("fldDedVal34Det").Value = arr(115)
  Fields("fldDedVal35Det").Value = arr(116)
  Fields("fldDedVal36Det").Value = arr(117)
  Fields("fldDedVal37Det").Value = arr(118)
  Fields("fldDedVal38Det").Value = arr(119)
  Fields("fldDedVal39Det").Value = arr(120)
  Fields("fldDedVal40Det").Value = arr(121)
  Fields("fldDedVal41Det").Value = arr(122)
  Fields("fldDedVal42Det").Value = arr(123)
  Fields("fldDedVal43Det").Value = arr(124)
  Fields("fldDedVal44Det").Value = arr(125)
  Fields("fldDedVal45Det").Value = arr(126)
  Fields("fldDedVal46Det").Value = arr(127)
  Fields("fldDedVal47Det").Value = arr(128)
  Fields("fldDedVal48Det").Value = arr(129)
  Fields("fldDedVal49Det").Value = arr(130)
  Fields("fldDedVal50Det").Value = arr(131)
  Fields("fldNumOfDeds").Value = arr(132)
'  DedCnt = arr(132)
  
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
  Me.Zoom = -1
  
  Label47.Visible = False 'Summary

  OpenDedCodeFile DHandle
  DedCnt = LOF(DHandle) / Len(DedRec)
  Close DHandle
  
  'for all but the last page
  Select Case DedCnt
    Case 1 To 5
      PageHeader.Height = 2100
      Detail.Height = 460
      Line4.Y1 = 2200
      Line4.Y2 = 2200
    Case 6 To 15
      PageHeader.Height = 2350
      Detail.Height = 650
      Line4.Y1 = 2500
      Line4.Y2 = 2500
    Case 16 To 25
      PageHeader.Height = 2600
      Detail.Height = 875
      Line4.Y1 = 2750
      Line4.Y2 = 2750
    Case 26 To 35
      PageHeader.Height = 2850
      Detail.Height = 1000
      Line4.Y1 = 3000
      Line4.Y2 = 3000
    Case 36 To 45
      PageHeader.Height = 3100
      Detail.Height = 1125
      Line4.Y1 = 3300
      Line4.Y2 = 3300
    Case 46 To 50
      PageHeader.Height = 3350
      Detail.Height = 1350
      Line4.Y1 = 3600
      Line4.Y2 = 3600
    Case Else
      PageHeader.Height = 3350
      Detail.Height = 1350
      Line4.Y1 = 3600
      Line4.Y2 = 3600
  End Select
  Me.PageSettings.RightMargin = 300
  Me.PageSettings.LeftMargin = 200
  Me.fldTimeDate.Text = Now
  GroupFooter1.Height = 0
  PageFooter.Height = 0
End Sub

Private Sub PageHeader_Format()
  If EndReport = True Then
    Label47.Visible = True
    PageHeader.Height = 1350
  End If
End Sub

Private Sub ReportFooter_Format()
  Set SubReport1.object = New arSubPRRegNSTotals
  SubReport1.Height = 5000
  EndReport = True
  Detail.Height = 0
  GroupHeader1.Height = 0
End Sub

Private Sub ReportHeader_Format()
  ReportHeader.Height = 0
End Sub


