VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arSubPRRegNSTotals 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ActiveReport1"
   ClientHeight    =   4356
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   8340
   Icon            =   "arSubPRRegNSTotals.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   14711
   _ExtentY        =   7684
   SectionData     =   "arSubPRRegNSTotals.dsx":08CA
End
Attribute VB_Name = "arSubPRRegNSTotals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Private hFile As Integer
Dim DedCnt As Integer
Private Sub ActiveReport_Initialize()
  Me.ToolBar.Tools.Add "Save/&Excel"
  Me.ToolBar.Tools.Add "&Text"
  Me.ToolBar.Tools.Add "&Close"
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
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
    MsgBox "Error Number: " & Str(Number) & " " & Description, vbOKOnly, "Printer Error"
    Unload Me
  End If
  CancelDisplay = True 'removes the error message
End Sub
Private Sub ActiveReport_DataInitialize()
  hFile = FreeFile
  Open StartPath & "\PRRPTS\REGISTERNSTOTAL.RPT" For Input As #hFile
  
  Fields.Add "fldEmployer" '(0)
  Fields.Add "fldDate" '(1)
  Fields.Add "fldSalNum" '(2)
  Fields.Add "fldHrNum" '(3)
  Fields.Add "fldTaxFrngttl" '(4)
  Fields.Add "fldRegHrsttl" '(5)
  Fields.Add "fldVacttl" '(6)
  Fields.Add "fldSickttl" '(7)
  Fields.Add "fldHolttl" '(8)
  Fields.Add "fldCompttl" '(9)
  Fields.Add "fldPersttl" '(10)
  Fields.Add "fldTotHrsttl" '(11)
  Fields.Add "fldOTPaidttl" '(12)
  Fields.Add "fldTOTComp" '(13)
  Fields.Add "fldRegEarnttl" '(14)
  Fields.Add "fldOTEarnttl" '(15)
  Fields.Add "fldEarn3ttl" '(16)
  Fields.Add "fldEarn2ttl" '(17)
  Fields.Add "fldEarn1ttl" '(18)
  Fields.Add "fldGrossPayttl" '(19)
  Fields.Add "fldSocSecttl" '(20)
  Fields.Add "fldMedttl" '(21)
  Fields.Add "fldFWTttl" '(22)
  Fields.Add "fldSWTttl" '(23)
  Fields.Add "fldRetttl" '(24)
  Fields.Add "fldNetPayttl" '(25)

  Fields.Add "fldDedVal1ttl" '(26)
  Fields.Add "fldDedVal2ttl" '(27)
  Fields.Add "fldDedVal3ttl" '(28)
  Fields.Add "fldDedVal4ttl" '(29)
  Fields.Add "fldDedVal5ttl" '(30)
  Fields.Add "fldDedVal6ttl" '(31)
  Fields.Add "fldDedVal7ttl" '(32)
  Fields.Add "fldDedVal8ttl" '(33)
  Fields.Add "fldDedVal9ttl" '(34)
  Fields.Add "fldDedVal10ttl" '(35)
  Fields.Add "fldDedVal11ttl" '(36)
  Fields.Add "fldDedVal12ttl" '(37)
  Fields.Add "fldDedVal13ttl" '(38)
  Fields.Add "fldDedVal14ttl" '(39)
  Fields.Add "fldDedVal15ttl" '(40)
  Fields.Add "fldDedVal16ttl" '(41)
  Fields.Add "fldDedVal17ttl" '(42)
  Fields.Add "fldDedVal18ttl" '(43)
  Fields.Add "fldDedVal19ttl" '(44)
  Fields.Add "fldDedVal20ttl" '(45)
  Fields.Add "fldDedVal21ttl" '(46)
  Fields.Add "fldDedVal22ttl" '(47)
  Fields.Add "fldDedVal23ttl" '(48)
  Fields.Add "fldDedVal24ttl" '(49)
  Fields.Add "fldDedVal25ttl" '(50)
  Fields.Add "fldDedVal26ttl" '(51)
  Fields.Add "fldDedVal27ttl" '(52)
  Fields.Add "fldDedVal28ttl" '(53)
  Fields.Add "fldDedVal29ttl" '(54)
  Fields.Add "fldDedVal30ttl" '(55)
  Fields.Add "fldDedVal31ttl" '(56)
  Fields.Add "fldDedVal32ttl" '(57)
  Fields.Add "fldDedVal33ttl" '(58)
  Fields.Add "fldDedVal34ttl" '(59)
  Fields.Add "fldDedVal35ttl" '(60)
  Fields.Add "fldDedVal36ttl" '(61)
  Fields.Add "fldDedVal37ttl" '(62)
  Fields.Add "fldDedVal38ttl" '(63)
  Fields.Add "fldDedVal39ttl" '(64)
  Fields.Add "fldDedVal40ttl" '(65)
  Fields.Add "fldDedVal41ttl" '(66)
  Fields.Add "fldDedVal42ttl" '(67)
  Fields.Add "fldDedVal43ttl" '(68)
  Fields.Add "fldDedVal44ttl" '(69)
  Fields.Add "fldDedVal45ttl" '(70)
  Fields.Add "fldDedVal46ttl" '(71)
  Fields.Add "fldDedVal47ttl" '(72)
  Fields.Add "fldDedVal48ttl" '(73)
  Fields.Add "fldDedVal49ttl" '(74)
  Fields.Add "fldDedVal50ttl" '(75)
  
  Fields.Add "fldFedGrsttl" '(76)
  Fields.Add "fldStaGrsttl" '(77)
  Fields.Add "fldMedGrsttl" '(78)
  Fields.Add "fldSocGrsttl" '(79)
  Fields.Add "fldRetGrsttl" '(80)
  Fields.Add "fldEICttl" '(81)
  
  Fields.Add ("fldDedDsc1") '(82)
  Fields.Add ("fldDedDsc2") '(83)
  Fields.Add ("fldDedDsc3") '(84)
  Fields.Add ("fldDedDsc4") '(85)
  Fields.Add ("fldDedDsc5") '(86)
  Fields.Add ("fldDedDsc6") '(87)
  Fields.Add ("fldDedDsc7") '(88)
  Fields.Add ("fldDedDsc8") '(89)
  Fields.Add ("fldDedDsc9") '(90)
  Fields.Add ("fldDedDsc10") '(91)
  Fields.Add ("fldDedDsc11") '(92)
  Fields.Add ("fldDedDsc12") '(93)
  Fields.Add ("fldDedDsc13") '(94)
  Fields.Add ("fldDedDsc14") '(95)
  Fields.Add ("fldDedDsc15") '(96)
  Fields.Add ("fldDedDsc16") '(97)
  Fields.Add ("fldDedDsc17") '(98)
  Fields.Add ("fldDedDsc18") '(99)
  Fields.Add ("fldDedDsc19") '(100)
  Fields.Add ("fldDedDsc20") '(101)
  Fields.Add ("fldDedDsc21") '(102)
  Fields.Add ("fldDedDsc22") '(103)
  Fields.Add ("fldDedDsc23") '(104)
  Fields.Add ("fldDedDsc24") '(105)
  Fields.Add ("fldDedDsc25") '(106)
  Fields.Add ("fldDedDsc26") '(107)
  Fields.Add ("fldDedDsc27") '(108)
  Fields.Add ("fldDedDsc28") '(109)
  Fields.Add ("fldDedDsc29") '(110)
  Fields.Add ("fldDedDsc30") '(111)
  Fields.Add ("fldDedDsc31") '(112)
  Fields.Add ("fldDedDsc32") '(113)
  Fields.Add ("fldDedDsc33") '(114)
  Fields.Add ("fldDedDsc34") '(115)
  Fields.Add ("fldDedDsc35") '(116)
  Fields.Add ("fldDedDsc36") '(117)
  Fields.Add ("fldDedDsc37") '(118)
  Fields.Add ("fldDedDsc38") '(119)
  Fields.Add ("fldDedDsc39") '(120)
  Fields.Add ("fldDedDsc40") '(121)
  Fields.Add ("fldDedDsc41") '(122)
  Fields.Add ("fldDedDsc42") '(123)
  Fields.Add ("fldDedDsc43") '(124)
  Fields.Add ("fldDedDsc44") '(125)
  Fields.Add ("fldDedDsc45") '(126)
  Fields.Add ("fldDedDsc46") '(127)
  Fields.Add ("fldDedDsc47") '(128)
  Fields.Add ("fldDedDsc48") '(129)
  Fields.Add ("fldDedDsc49") '(130)
  Fields.Add ("fldDedDsc50") '(131)
  
  Fields.Add ("fldNumOfDeds") '(132)
  Fields.Add ("fldEarnDsc3") '133) '8/5/05
  Fields.Add ("fldEarnDsc2") '134) '8/5/05
  Fields.Add ("fldEarnDsc1") '135) '8/5/05
  
  End Sub
Private Sub ActiveReport_FetchData(eof As Boolean)
  Dim sLine As String
  Dim arr() As String
  Dim x As Integer
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
  Fields("fldDate").Value = arr(1)
  Fields("fldSalNum").Value = arr(2)
  Fields("fldHrNum").Value = arr(3)
  Fields("fldTaxFrngttl").Value = arr(4)
  Fields("fldRegHrsttl").Value = arr(5)
  Fields("fldVacttl").Value = arr(6)
  Fields("fldSickttl").Value = arr(7)
  Fields("fldHolttl").Value = arr(8)
  Fields("fldCompttl").Value = arr(9)
  Fields("fldPersttl").Value = arr(10)
  Fields("fldTotHrsttl").Value = arr(11)
  Fields("fldOTPaidttl").Value = arr(12)
  Fields("fldTOTComp").Value = arr(13)
  Fields("fldRegEarnttl").Value = arr(14)
  Fields("fldOTEarnttl").Value = arr(15)
  Fields("fldEarn3ttl").Value = arr(16)
  Fields("fldEarn2ttl").Value = arr(17)
  Fields("fldEarn1ttl").Value = arr(18)
  Fields("fldGrossPayttl").Value = arr(19)
  Fields("fldSocSecttl").Value = arr(20)
  Fields("fldMedttl").Value = arr(21)
  Fields("fldFWTttl").Value = arr(22)
  Fields("fldSWTttl").Value = arr(23)
  Fields("fldRetttl").Value = arr(24)
  Fields("fldNetPayttl").Value = arr(25)

  For x = 26 To 75
    If x >= 26 + DedCnt Then
      arr(x) = ""
    End If
  Next x
  
  Fields("fldDedVal1ttl").Value = arr(26)
  Fields("fldDedVal2ttl").Value = arr(27)
  Fields("fldDedVal3ttl").Value = arr(28)
  Fields("fldDedVal4ttl").Value = arr(29)
  Fields("fldDedVal5ttl").Value = arr(30)
  Fields("fldDedVal6ttl").Value = arr(31)
  Fields("fldDedVal7ttl").Value = arr(32)
  Fields("fldDedVal8ttl").Value = arr(33)
  Fields("fldDedVal9ttl").Value = arr(34)
  Fields("fldDedVal10ttl").Value = arr(35)
  Fields("fldDedVal11ttl").Value = arr(36)
  Fields("fldDedVal12ttl").Value = arr(37)
  Fields("fldDedVal13ttl").Value = arr(38)
  Fields("fldDedVal14ttl").Value = arr(39)
  Fields("fldDedVal15ttl").Value = arr(40)
  Fields("fldDedVal16ttl").Value = arr(41)
  Fields("fldDedVal17ttl").Value = arr(42)
  Fields("fldDedVal18ttl").Value = arr(43)
  Fields("fldDedVal19ttl").Value = arr(44)
  Fields("fldDedVal20ttl").Value = arr(45)
  Fields("fldDedVal21ttl").Value = arr(46)
  Fields("fldDedVal22ttl").Value = arr(47)
  Fields("fldDedVal23ttl").Value = arr(48)
  Fields("fldDedVal24ttl").Value = arr(49)
  Fields("fldDedVal25ttl").Value = arr(50)
  Fields("fldDedVal26ttl").Value = arr(51)
  Fields("fldDedVal27ttl").Value = arr(52)
  Fields("fldDedVal28ttl").Value = arr(53)
  Fields("fldDedVal29ttl").Value = arr(54)
  Fields("fldDedVal30ttl").Value = arr(55)
  Fields("fldDedVal31ttl").Value = arr(56)
  Fields("fldDedVal32ttl").Value = arr(57)
  Fields("fldDedVal33ttl").Value = arr(58)
  Fields("fldDedVal34ttl").Value = arr(59)
  Fields("fldDedVal35ttl").Value = arr(60)
  Fields("fldDedVal36ttl").Value = arr(61)
  Fields("fldDedVal37ttl").Value = arr(62)
  Fields("fldDedVal38ttl").Value = arr(63)
  Fields("fldDedVal39ttl").Value = arr(64)
  Fields("fldDedVal40ttl").Value = arr(65)
  Fields("fldDedVal41ttl").Value = arr(66)
  Fields("fldDedVal42ttl").Value = arr(67)
  Fields("fldDedVal43ttl").Value = arr(68)
  Fields("fldDedVal44ttl").Value = arr(69)
  Fields("fldDedVal45ttl").Value = arr(70)
  Fields("fldDedVal46ttl").Value = arr(71)
  Fields("fldDedVal47ttl").Value = arr(72)
  Fields("fldDedVal48ttl").Value = arr(73)
  Fields("fldDedVal49ttl").Value = arr(74)
  Fields("fldDedVal50ttl").Value = arr(75)
  
  Fields("fldFedGrsttl").Value = arr(76)
  Fields("fldStaGrsttl").Value = arr(77)
  Fields("fldMedGrsttl").Value = arr(78)
  Fields("fldSocGrsttl").Value = arr(79)
  Fields("fldRetGrsttl").Value = arr(80)
  Fields("fldEICttl").Value = arr(81)
  
  Fields("fldDedDsc1").Value = arr(82)
  Fields("fldDedDsc2").Value = arr(83)
  Fields("fldDedDsc3").Value = arr(84)
  Fields("fldDedDsc4").Value = arr(85)
  Fields("fldDedDsc5").Value = arr(86)
  Fields("fldDedDsc6").Value = arr(87)
  Fields("fldDedDsc7").Value = arr(88)
  Fields("fldDedDsc8").Value = arr(89)
  Fields("fldDedDsc9").Value = arr(90)
  Fields("fldDedDsc10").Value = arr(91)
  Fields("fldDedDsc11").Value = arr(92)
  Fields("fldDedDsc12").Value = arr(93)
  Fields("fldDedDsc13").Value = arr(94)
  Fields("fldDedDsc14").Value = arr(95)
  Fields("fldDedDsc15").Value = arr(96)
  Fields("fldDedDsc16").Value = arr(97)
  Fields("fldDedDsc17").Value = arr(98)
  Fields("fldDedDsc18").Value = arr(99)
  Fields("fldDedDsc19").Value = arr(100)
  Fields("fldDedDsc20").Value = arr(101)
  Fields("fldDedDsc21").Value = arr(102)
  Fields("fldDedDsc22").Value = arr(103)
  Fields("fldDedDsc23").Value = arr(104)
  Fields("fldDedDsc24").Value = arr(105)
  Fields("fldDedDsc25").Value = arr(106)
  Fields("fldDedDsc26").Value = arr(107)
  Fields("fldDedDsc27").Value = arr(108)
  Fields("fldDedDsc28").Value = arr(109)
  Fields("fldDedDsc29").Value = arr(110)
  Fields("fldDedDsc30").Value = arr(111)
  Fields("fldDedDsc31").Value = arr(112)
  Fields("fldDedDsc32").Value = arr(113)
  Fields("fldDedDsc33").Value = arr(114)
  Fields("fldDedDsc34").Value = arr(115)
  Fields("fldDedDsc35").Value = arr(116)
  Fields("fldDedDsc36").Value = arr(117)
  Fields("fldDedDsc37").Value = arr(118)
  Fields("fldDedDsc38").Value = arr(119)
  Fields("fldDedDsc39").Value = arr(120)
  Fields("fldDedDsc40").Value = arr(121)
  Fields("fldDedDsc41").Value = arr(122)
  Fields("fldDedDsc42").Value = arr(123)
  Fields("fldDedDsc43").Value = arr(124)
  Fields("fldDedDsc44").Value = arr(125)
  Fields("fldDedDsc45").Value = arr(126)
  Fields("fldDedDsc46").Value = arr(127)
  Fields("fldDedDsc47").Value = arr(128)
  Fields("fldDedDsc48").Value = arr(129)
  Fields("fldDedDsc49").Value = arr(130)
  Fields("fldDedDsc50").Value = arr(131)
  Fields("fldNumOfDeds").Value = arr(132)
  Fields("fldEarnDsc3").Value = arr(133)  '8/5/05
  Fields("fldEarnDsc2").Value = arr(134) '8/5/05
  Fields("fldEarnDsc1").Value = arr(135) '8/5/05
  
End Sub
Private Sub ActiveReport_ReportEnd()
  If hFile <> 0 Then
    Close #hFile
  End If
End Sub

Private Sub ActiveReport_ReportStart()
  
  Dim DedRec As DedCodeRecType
  Dim DHandle As Integer
  OpenDedCodeFile DHandle
  DedCnt = LOF(DHandle) / Len(DedRec)
  Close DHandle
  
End Sub

