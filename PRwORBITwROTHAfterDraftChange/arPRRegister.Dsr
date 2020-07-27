VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arPRRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Register"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   13035
   Icon            =   "arPRRegister.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   22992
   _ExtentY        =   15637
   SectionData     =   "arPRRegister.dsx":08CA
End
Attribute VB_Name = "arPRRegister"
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
      MsgBox "File - EarningsRegisterRpt.xls, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - EarningsRegisterRpt.txt, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    End If
  End If
End Sub

Private Sub ActiveReport_Terminate()
  Close hFile
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
    MsgBox "File - EarningsRegisterRpt.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - EarningsRegisterRpt.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "EarningsRegisterRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "EarningsRegisterRpt.txt"
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
  Open StartPath & "\PRRPTS\REGISTERG.RPT" For Input As #hFile
  Fields.Add "fldEmployer" '0
  Fields.Add "fldDates" '1
  Fields.Add "fldEmpNum" '2
  Fields.Add "fldEmpName" '3
  Fields.Add "fldEarn1Det" '4
  Fields.Add "fldEarnDsc1" '5
  Fields.Add "fldEarn2Det" '6
  Fields.Add "fldEarnDsc2" '7
  Fields.Add "fldEarn3Det" '8
  Fields.Add "fldEarnDsc3" '9
  Fields.Add "fldDedDsc1" '10
  Fields.Add "fldDedDsc2" '11
  Fields.Add "fldDedDsc3" '12
  Fields.Add "fldDedDsc4" '13
  Fields.Add "fldDedDsc5" '14
  Fields.Add "fldDedDsc6" '15
  Fields.Add "fldDedDsc7" '16
  Fields.Add "fldDedDsc8" '17
  Fields.Add "fldDedDsc9" '18
  Fields.Add "fldDedDsc10" '19
  Fields.Add "fldDedDsc11" '20
  Fields.Add "fldDedDsc12" '21
  Fields.Add "fldDedDsc13" '22
  Fields.Add "fldDedDsc14" '23
  Fields.Add "fldDedDsc15" '24
  Fields.Add "fldDedDsc16" '25
  Fields.Add "fldDedDsc17" '26
  Fields.Add "fldDedDsc18" '27
  Fields.Add "fldDedDsc19" '28
  Fields.Add "fldDedDsc20" '29
  Fields.Add "fldDedDsc21" '30
  Fields.Add "fldDedDsc22" '31
  Fields.Add "fldDedDsc23" '32
  Fields.Add "fldDedDsc24" '33
  Fields.Add "fldDedDsc25" '34
  Fields.Add "fldDedDsc26" '35
  Fields.Add "fldDedDsc27" '36
  Fields.Add "fldDedDsc28" '37
  Fields.Add "fldDedDsc29" '38
  Fields.Add "fldDedDsc30" '39
  Fields.Add "fldDedDsc31" '40
  Fields.Add "fldDedDsc32" '41
  Fields.Add "fldDedDsc33" '42
  Fields.Add "fldDedDsc34" '43
  Fields.Add "fldDedDsc35" '44
  Fields.Add "fldDedDsc36" '45
  Fields.Add "fldDedDsc37" '46
  Fields.Add "fldDedDsc38" '47
  Fields.Add "fldDedDsc39" '48
  Fields.Add "fldDedDsc40" '49
  Fields.Add "fldDedDsc41" '50
  Fields.Add "fldDedDsc42" '51
  Fields.Add "fldDedDsc43" '52
  Fields.Add "fldDedDsc44" '53
  Fields.Add "fldDedDsc45" '54
  Fields.Add "fldDedDsc46" '55
  Fields.Add "fldDedDsc47" '56
  Fields.Add "fldDedDsc48" '57
  Fields.Add "fldDedDsc49" '58
  Fields.Add "fldDedDsc50" '59
  Fields.Add "fldBaseRate" '60
  Fields.Add "fldOTRate" '61
  Fields.Add "fldTaxFr" '62
  Fields.Add "fldRegHrsDet" '63
  Fields.Add "fldVacDet" '64
  Fields.Add "fldSickDet" '65
  Fields.Add "fldHolDet" '66
  Fields.Add "fldCompDet" '67
  Fields.Add "fldPersDet" '68
  Fields.Add "fldTotHrsDet" '69
  Fields.Add "fldOTPaidDet" '70
  Fields.Add "fldOTComp" '71
  Fields.Add "fldEICDet" '72
  Fields.Add "fldRegEarnDet" '73
  Fields.Add "fldOTEarnDet" '74
  Fields.Add "fldEarn1ttl" '75
  Fields.Add "fldEarn2ttl" '76
  Fields.Add "fldEarn3ttl" '77
  Fields.Add "fldGrossPayDet" '78
  Fields.Add "fldSocSecDet" '79
  Fields.Add "fldMedDet" '80
  Fields.Add "fldFWTDet" '81
  Fields.Add "fldSWTDet" '82
  Fields.Add "fldRetDet" '83
  Fields.Add "fldNetPayDet" '84
  Fields.Add "fldDedVal1Det" '85
  Fields.Add "fldDedVal2Det" '86
  Fields.Add "fldDedVal3Det" '87
  Fields.Add "fldDedVal4Det" '88
  Fields.Add "fldDedVal5Det" '89
  Fields.Add "fldDedVal6Det" '90
  Fields.Add "fldDedVal7Det" '91
  Fields.Add "fldDedVal8Det" '92
  Fields.Add "fldDedVal9Det" '93
  Fields.Add "fldDedVal10Det" '94
  Fields.Add "fldDedVal11Det" '95
  Fields.Add "fldDedVal12Det" '96
  Fields.Add "fldDedVal13Det" '97
  Fields.Add "fldDedVal14Det" '98
  Fields.Add "fldDedVal15Det" '99
  Fields.Add "fldDedVal16Det" '100
  Fields.Add "fldDedVal17Det" '101
  Fields.Add "fldDedVal18Det" '102
  Fields.Add "fldDedVal19Det" '103
  Fields.Add "fldDedVal20Det" '104
  Fields.Add "fldDedVal21Det" '105
  Fields.Add "fldDedVal22Det" '106
  Fields.Add "fldDedVal23Det" '107
  Fields.Add "fldDedVal24Det" '108
  Fields.Add "fldDedVal25Det" '109
  Fields.Add "fldDedVal26Det" '110
  Fields.Add "fldDedVal27Det" '111
  Fields.Add "fldDedVal28Det" '112
  Fields.Add "fldDedVal29Det" '113
  Fields.Add "fldDedVal30Det" '114
  Fields.Add "fldDedVal31Det" '115
  Fields.Add "fldDedVal32Det" '116
  Fields.Add "fldDedVal33Det" '117
  Fields.Add "fldDedVal34Det" '118
  Fields.Add "fldDedVal35Det" '119
  Fields.Add "fldDedVal36Det" '120
  Fields.Add "fldDedVal37Det" '121
  Fields.Add "fldDedVal38Det" '122
  Fields.Add "fldDedVal39Det" '123
  Fields.Add "fldDedVal40Det" '124
  Fields.Add "fldDedVal41Det" '125
  Fields.Add "fldDedVal42Det" '126
  Fields.Add "fldDedVal43Det" '127
  Fields.Add "fldDedVal44Det" '128
  Fields.Add "fldDedVal45Det" '129
  Fields.Add "fldDedVal46Det" '130
  Fields.Add "fldDedVal47Det" '131
  Fields.Add "fldDedVal48Det" '132
  Fields.Add "fldDedVal49Det" '133
  Fields.Add "fldDedVal50Det" '134
  Fields.Add "fldSalNum" '135
  Fields.Add "fldHrNum" '136
  Fields.Add "fldFedGrsttl" '137
  Fields.Add "fldStaGrsttl" '138
  Fields.Add "fldSocGrsttl" '139
  Fields.Add "fldRetGrsttl" '140
  Fields.Add "fldMedGrsttl" '141
  
  Fields.Add "fldDedVal1ttl" '142
  Fields.Add "fldDedVal2ttl" '143
  Fields.Add "fldDedVal3ttl" '144
  Fields.Add "fldDedVal4ttl" '145
  Fields.Add "fldDedVal5ttl" '146
  Fields.Add "fldDedVal6ttl" '147
  Fields.Add "fldDedVal7ttl" '148
  Fields.Add "fldDedVal8ttl" '149
  Fields.Add "fldDedVal9ttl" '150
  Fields.Add "fldDedVal10ttl" '151
  Fields.Add "fldDedVal11ttl" '152
  Fields.Add "fldDedVal12ttl" '153
  Fields.Add "fldDedVal13ttl" '154
  Fields.Add "fldDedVal14ttl" '155
  Fields.Add "fldDedVal15ttl" '156
  Fields.Add "fldDedVal16ttl" '157
  Fields.Add "fldDedVal17ttl" '158
  Fields.Add "fldDedVal18ttl" '159
  Fields.Add "fldDedVal19ttl" '160
  Fields.Add "fldDedVal20ttl" '161
  Fields.Add "fldDedVal21ttl" '162
  Fields.Add "fldDedVal22ttl" '163
  Fields.Add "fldDedVal23ttl" '164
  Fields.Add "fldDedVal24ttl" '165
  Fields.Add "fldDedVal25ttl" '166
  Fields.Add "fldDedVal26ttl" '167
  Fields.Add "fldDedVal27ttl" '168
  Fields.Add "fldDedVal28ttl" '169
  Fields.Add "fldDedVal29ttl" '170
  Fields.Add "fldDedVal30ttl" '171
  Fields.Add "fldDedVal31ttl" '172
  Fields.Add "fldDedVal32ttl" '173
  Fields.Add "fldDedVal33ttl" '174
  Fields.Add "fldDedVal34ttl" '175
  Fields.Add "fldDedVal35ttl" '176
  Fields.Add "fldDedVal36ttl" '177
  Fields.Add "fldDedVal37ttl" '178
  Fields.Add "fldDedVal38ttl" '179
  Fields.Add "fldDedVal39ttl" '180
  Fields.Add "fldDedVal40ttl" '181
  Fields.Add "fldDedVal41ttl" '182
  Fields.Add "fldDedVal42ttl" '183
  Fields.Add "fldDedVal43ttl" '184
  Fields.Add "fldDedVal44ttl" '185
  Fields.Add "fldDedVal45ttl" '186
  Fields.Add "fldDedVal46ttl" '187
  Fields.Add "fldDedVal47ttl" '188
  Fields.Add "fldDedVal48ttl" '189
  Fields.Add "fldDedVal49ttl" '190
  Fields.Add "fldDedVal50ttl" '191
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
  Fields("fldEarn1Det").Value = arr(4)
  Fields("fldEarnDsc1").Value = arr(5)
  Fields("fldEarn2Det").Value = arr(6)
  Fields("fldEarnDsc2").Value = arr(7)
  Fields("fldEarn3Det").Value = arr(8)
  Fields("fldEarnDsc3").Value = arr(9)
  Fields("fldDedDsc1").Value = arr(10)
  Fields("fldDedDsc2").Value = arr(11)
  Fields("fldDedDsc3").Value = arr(12)
  Fields("fldDedDsc4").Value = arr(13)
  Fields("fldDedDsc5").Value = arr(14)
  Fields("fldDedDsc6").Value = arr(15)
  Fields("fldDedDsc7").Value = arr(16)
  Fields("fldDedDsc8").Value = arr(17)
  Fields("fldDedDsc9").Value = arr(18)
  Fields("fldDedDsc10").Value = arr(19)
  Fields("fldDedDsc11").Value = arr(20)
  Fields("fldDedDsc12").Value = arr(21)
  Fields("fldDedDsc13").Value = arr(22)
  Fields("fldDedDsc14").Value = arr(23)
  Fields("fldDedDsc15").Value = arr(24)
  Fields("fldDedDsc16").Value = arr(25)
  Fields("fldDedDsc17").Value = arr(26)
  Fields("fldDedDsc18").Value = arr(27)
  Fields("fldDedDsc19").Value = arr(28)
  Fields("fldDedDsc20").Value = arr(29)
  Fields("fldDedDsc21").Value = arr(30)
  Fields("fldDedDsc22").Value = arr(31)
  Fields("fldDedDsc23").Value = arr(32)
  Fields("fldDedDsc24").Value = arr(33)
  Fields("fldDedDsc25").Value = arr(34)
  Fields("fldDedDsc26").Value = arr(35)
  Fields("fldDedDsc27").Value = arr(36)
  Fields("fldDedDsc28").Value = arr(37)
  Fields("fldDedDsc29").Value = arr(38)
  Fields("fldDedDsc30").Value = arr(39)
  Fields("fldDedDsc31").Value = arr(40)
  Fields("fldDedDsc32").Value = arr(41)
  Fields("fldDedDsc33").Value = arr(42)
  Fields("fldDedDsc34").Value = arr(43)
  Fields("fldDedDsc35").Value = arr(44)
  Fields("fldDedDsc36").Value = arr(45)
  Fields("fldDedDsc37").Value = arr(46)
  Fields("fldDedDsc38").Value = arr(47)
  Fields("fldDedDsc39").Value = arr(48)
  Fields("fldDedDsc40").Value = arr(49)
  Fields("fldDedDsc41").Value = arr(50)
  Fields("fldDedDsc42").Value = arr(51)
  Fields("fldDedDsc43").Value = arr(52)
  Fields("fldDedDsc44").Value = arr(53)
  Fields("fldDedDsc45").Value = arr(54)
  Fields("fldDedDsc46").Value = arr(55)
  Fields("fldDedDsc47").Value = arr(56)
  Fields("fldDedDsc48").Value = arr(57)
  Fields("fldDedDsc49").Value = arr(58)
  Fields("fldDedDsc50").Value = arr(59)
  Fields("fldBaseRate").Value = arr(60)
  Fields("fldOTRate").Value = arr(61)
  Fields("fldTaxFr").Value = arr(62)
  Fields("fldRegHrsDet").Value = arr(63)
  Fields("fldVacDet").Value = arr(64)
  Fields("fldSickDet").Value = arr(65)
  Fields("fldHolDet").Value = arr(66)
  Fields("fldCompDet").Value = arr(67)
  Fields("fldPersDet").Value = arr(68)
  Fields("fldTotHrsDet").Value = arr(69)
  Fields("fldOTPaidDet").Value = arr(70)
  Fields("fldOTComp").Value = arr(71)
  Fields("fldEICDet").Value = arr(72)
  Fields("fldRegEarnDet").Value = arr(73)
  Fields("fldOTEarnDet").Value = arr(74)
  Fields("fldEarn1ttl").Value = arr(75)
  Fields("fldEarn2ttl").Value = arr(76)
  Fields("fldEarn3ttl").Value = arr(77)
  Fields("fldGrossPayDet").Value = arr(78)
  Fields("fldSocSecDet").Value = arr(79)
  Fields("fldMedDet").Value = arr(80)
  Fields("fldFWTDet").Value = arr(81)
  Fields("fldSWTDet").Value = arr(82)
  Fields("fldRetDet").Value = arr(83)
  Fields("fldNetPayDet").Value = arr(84)
  Fields("fldDedVal1Det").Value = arr(85)
  Fields("fldDedVal2Det").Value = arr(86)
  Fields("fldDedVal3Det").Value = arr(87)
  Fields("fldDedVal4Det").Value = arr(88)
  Fields("fldDedVal5Det").Value = arr(89)
  Fields("fldDedVal6Det").Value = arr(90)
  Fields("fldDedVal7Det").Value = arr(91)
  Fields("fldDedVal8Det").Value = arr(92)
  Fields("fldDedVal9Det").Value = arr(93)
  Fields("fldDedVal10Det").Value = arr(94)
  Fields("fldDedVal11Det").Value = arr(95)
  Fields("fldDedVal12Det").Value = arr(96)
  Fields("fldDedVal13Det").Value = arr(97)
  Fields("fldDedVal14Det").Value = arr(98)
  Fields("fldDedVal15Det").Value = arr(99)
  Fields("fldDedVal16Det").Value = arr(100)
  Fields("fldDedVal17Det").Value = arr(101)
  Fields("fldDedVal18Det").Value = arr(102)
  Fields("fldDedVal19Det").Value = arr(103)
  Fields("fldDedVal20Det").Value = arr(104)
  Fields("fldDedVal21Det").Value = arr(105)
  Fields("fldDedVal22Det").Value = arr(106)
  Fields("fldDedVal23Det").Value = arr(107)
  Fields("fldDedVal24Det").Value = arr(108)
  Fields("fldDedVal25Det").Value = arr(109)
  Fields("fldDedVal26Det").Value = arr(110)
  Fields("fldDedVal27Det").Value = arr(111)
  Fields("fldDedVal28Det").Value = arr(112)
  Fields("fldDedVal29Det").Value = arr(113)
  Fields("fldDedVal30Det").Value = arr(114)
  Fields("fldDedVal31Det").Value = arr(115)
  Fields("fldDedVal32Det").Value = arr(116)
  Fields("fldDedVal33Det").Value = arr(117)
  Fields("fldDedVal34Det").Value = arr(118)
  Fields("fldDedVal35Det").Value = arr(119)
  Fields("fldDedVal36Det").Value = arr(120)
  Fields("fldDedVal37Det").Value = arr(121)
  Fields("fldDedVal38Det").Value = arr(122)
  Fields("fldDedVal39Det").Value = arr(123)
  Fields("fldDedVal40Det").Value = arr(124)
  Fields("fldDedVal41Det").Value = arr(125)
  Fields("fldDedVal42Det").Value = arr(126)
  Fields("fldDedVal43Det").Value = arr(127)
  Fields("fldDedVal44Det").Value = arr(128)
  Fields("fldDedVal45Det").Value = arr(129)
  Fields("fldDedVal46Det").Value = arr(130)
  Fields("fldDedVal47Det").Value = arr(131)
  Fields("fldDedVal48Det").Value = arr(132)
  Fields("fldDedVal49Det").Value = arr(133)
  Fields("fldDedVal50Det").Value = arr(134)
  Fields("fldSalNum").Value = arr(135)
  Fields("fldHrNum").Value = arr(136)
  Fields("fldFedGrsttl").Value = arr(137)
  Fields("fldStaGrsttl").Value = arr(138)
  Fields("fldSocGrsttl").Value = arr(139)
  Fields("fldRetGrsttl").Value = arr(140)
  Fields("fldMedGrsttl").Value = arr(141)
  
  Fields("fldDedVal1ttl").Value = arr(142)
  Fields("fldDedVal2ttl").Value = arr(143)
  Fields("fldDedVal3ttl").Value = arr(144)
  Fields("fldDedVal4ttl").Value = arr(145)
  Fields("fldDedVal5ttl").Value = arr(146)
  Fields("fldDedVal6ttl").Value = arr(147)
  Fields("fldDedVal7ttl").Value = arr(148)
  Fields("fldDedVal8ttl").Value = arr(149)
  Fields("fldDedVal9ttl").Value = arr(150)
  Fields("fldDedVal10ttl").Value = arr(151)
  Fields("fldDedVal11ttl").Value = arr(152)
  Fields("fldDedVal12ttl").Value = arr(153)
  Fields("fldDedVal13ttl").Value = arr(154)
  Fields("fldDedVal14ttl").Value = arr(155)
  Fields("fldDedVal15ttl").Value = arr(156)
  Fields("fldDedVal16ttl").Value = arr(157)
  Fields("fldDedVal17ttl").Value = arr(158)
  Fields("fldDedVal18ttl").Value = arr(159)
  Fields("fldDedVal19ttl").Value = arr(160)
  Fields("fldDedVal20ttl").Value = arr(161)
  Fields("fldDedVal21ttl").Value = arr(162)
  Fields("fldDedVal22ttl").Value = arr(163)
  Fields("fldDedVal23ttl").Value = arr(164)
  Fields("fldDedVal24ttl").Value = arr(165)
  Fields("fldDedVal25ttl").Value = arr(166)
  Fields("fldDedVal26ttl").Value = arr(167)
  Fields("fldDedVal27ttl").Value = arr(168)
  Fields("fldDedVal28ttl").Value = arr(169)
  Fields("fldDedVal29ttl").Value = arr(170)
  Fields("fldDedVal30ttl").Value = arr(171)
  Fields("fldDedVal31ttl").Value = arr(172)
  Fields("fldDedVal32ttl").Value = arr(173)
  Fields("fldDedVal33ttl").Value = arr(174)
  Fields("fldDedVal34ttl").Value = arr(175)
  Fields("fldDedVal35ttl").Value = arr(176)
  Fields("fldDedVal36ttl").Value = arr(177)
  Fields("fldDedVal37ttl").Value = arr(178)
  Fields("fldDedVal38ttl").Value = arr(179)
  Fields("fldDedVal39ttl").Value = arr(180)
  Fields("fldDedVal40ttl").Value = arr(181)
  Fields("fldDedVal41ttl").Value = arr(182)
  Fields("fldDedVal42ttl").Value = arr(183)
  Fields("fldDedVal43ttl").Value = arr(184)
  Fields("fldDedVal44ttl").Value = arr(185)
  Fields("fldDedVal45ttl").Value = arr(186)
  Fields("fldDedVal46ttl").Value = arr(187)
  Fields("fldDedVal47ttl").Value = arr(188)
  Fields("fldDedVal48ttl").Value = arr(189)
  Fields("fldDedVal49ttl").Value = arr(190)
  Fields("fldDedVal50ttl").Value = arr(191)
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
  
  Label47.Visible = False 'Summary

  OpenDedCodeFile DHandle
  DedCnt = LOF(DHandle) / Len(DedRec)
  Close DHandle
  'for report footer only
  Select Case DedCnt
    Case 1 To 5
      ReportFooter.Height = 4250
    Case 6 To 10
      ReportFooter.Height = 4800
    Case 11 To 15
      ReportFooter.Height = 5500
    Case 16 To 20
      ReportFooter.Height = 6200
    Case 21 To 25
      ReportFooter.Height = 6900
    Case 26 To 30
      ReportFooter.Height = 7600
    Case 31 To 35
      ReportFooter.Height = 8000
    Case 36 To 40
      ReportFooter.Height = 8625
    Case 41 To 45
      ReportFooter.Height = 9250
    Case 46 To 50
      ReportFooter.Height = 9900
    Case Else
      ReportFooter.Height = 9900
  End Select
  
  'for all but the last page
  Select Case DedCnt
    Case 1 To 5
      PageHeader.Height = 2100
      Detail.Height = 460
      ReportFooter.Height = 4250
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

Private Sub Detail_Format()
  If Fields("fldEarnDsc1").Value = "" Then
    Fields("fldEarn1ttl").Value = ""
  End If
  If Fields("fldEarnDsc2").Value = "" Then
    Fields("fldEarn2ttl").Value = ""
  End If
  If Fields("fldEarnDsc3").Value = "" Then
    Fields("fldEarn3ttl").Value = ""
  End If

End Sub

Private Sub PageHeader_Format()
  If EndReport = True Then
    Label47.Visible = True
    PageHeader.Height = 1350
  End If
End Sub

Private Sub ReportFooter_Format()
  EndReport = True
  Detail.Height = 0
  GroupHeader1.Height = 0
End Sub

Private Sub ReportHeader_Format()
  ReportHeader.Height = 0
End Sub

