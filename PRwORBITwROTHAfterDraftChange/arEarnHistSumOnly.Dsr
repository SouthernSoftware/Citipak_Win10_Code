VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arEarnHistSumOnly 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Earning History Sum Only"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "arEarnHistSumOnly.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20558
   _ExtentY        =   15637
   SectionData     =   "arEarnHistSumOnly.dsx":08CA
End
Attribute VB_Name = "arEarnHistSumOnly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Private hFile As Integer
Dim EndReport As Boolean
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
      MsgBox "File - EarnHistSumOnly.xls, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - EarnHistSumOnly.txt, created in the Citipak Directory.", vbOKOnly
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
    MsgBox "File - EarnHistSumOnly.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - EarnHistSumOnly.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "EarnHistSumOnly.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "EarnHistSumOnly.txt"
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
  Open StartPath & "\PRRPTS\EMPHISTSUMG.RPT" For Input As #hFile
  Fields.Add "ghGrpHdr1" '0 'group totals on this value
  Fields.Add "fldEmployer" '1
  Fields.Add "fldStartDate" '2
  Fields.Add "fldEndDate" '3
  Fields.Add "fldEmpNum" '4
  Fields.Add "fldEmpName" '5
  Fields.Add "fldEarnDsc1" '6
  Fields.Add "fldEarnDsc2" '7
  Fields.Add "fldEarnDsc3" '8
  Fields.Add "fldDedDsc1" '9
  Fields.Add "fldDedDsc2" '10
  Fields.Add "fldDedDsc3" '11
  Fields.Add "fldDedDsc4" '12
  Fields.Add "fldDedDsc5" '13
  Fields.Add "fldDedDsc6" '14
  Fields.Add "fldDedDsc7" '15
  Fields.Add "fldDedDsc8" '16
  Fields.Add "fldDedDsc9" '17
  Fields.Add "fldDedDsc10" '18
  Fields.Add "fldDedDsc11" '19
  Fields.Add "fldDedDsc12" '20
  Fields.Add "fldDedDsc13" '21
  Fields.Add "fldDedDsc14" '22
  Fields.Add "fldDedDsc15" '23
  Fields.Add "fldDedDsc16" '24
  Fields.Add "fldDedDsc17" '25
  Fields.Add "fldDedDsc18" '26
  Fields.Add "fldDedDsc19" '27
  Fields.Add "fldDedDsc20" '28
  Fields.Add "fldDedDsc21" '29
  Fields.Add "fldDedDsc22" '30
  Fields.Add "fldDedDsc23" '31
  Fields.Add "fldDedDsc24" '32
  Fields.Add "fldDedDsc25" '33
  Fields.Add "fldDedDsc26" '34
  Fields.Add "fldDedDsc27" '35
  Fields.Add "fldDedDsc28" '36
  Fields.Add "fldDedDsc29" '37
  Fields.Add "fldDedDsc30" '38
  Fields.Add "fldDedDsc31" '39
  Fields.Add "fldDedDsc32" '40
  Fields.Add "fldDedDsc33" '41
  Fields.Add "fldDedDsc34" '42
  Fields.Add "fldDedDsc35" '43
  Fields.Add "fldDedDsc36" '44
  Fields.Add "fldDedDsc37" '45
  Fields.Add "fldDedDsc38" '46
  Fields.Add "fldDedDsc39" '47
  Fields.Add "fldDedDsc40" '48
  Fields.Add "fldDedDsc41" '49
  Fields.Add "fldDedDsc42" '50
  Fields.Add "fldDedDsc43" '51
  Fields.Add "fldDedDsc44" '52
  Fields.Add "fldDedDsc45" '53
  Fields.Add "fldDedDsc46" '54
  Fields.Add "fldDedDsc47" '55
  Fields.Add "fldDedDsc48" '56
  Fields.Add "fldDedDsc49" '57
  Fields.Add "fldDedDsc50" '58
  Fields.Add "fldTaxFr" '59
  Fields.Add "fldRegHrs" '60
  Fields.Add "fldVac" '61
  Fields.Add "fldSick" '62
  Fields.Add "fldHol" '63
  Fields.Add "fldComp" '64
  Fields.Add "fldPers" '65
  Fields.Add "fldTotHrs" '66
  Fields.Add "fldOTPaid" '67
  Fields.Add "fldEIC" '68
  Fields.Add "fldRegEarn" '69
  Fields.Add "fldOTEarn" '70
  Fields.Add "fldEarn1" '71
  Fields.Add "fldEarn2" '72
  Fields.Add "fldEarn3" '73
  Fields.Add "fldGrossPay" '74
  Fields.Add "fldSocSec" '75
  Fields.Add "fldMed" '76
  Fields.Add "fldFWT" '77
  Fields.Add "fldSWT" '78
  Fields.Add "fldRet" '79
  Fields.Add "fldNetPay" '80
  Fields.Add "fldDedVal1" '81
  Fields.Add "fldDedVal2" '82
  Fields.Add "fldDedVal3" '83
  Fields.Add "fldDedVal4" '84
  Fields.Add "fldDedVal5" '85
  Fields.Add "fldDedVal6" '86
  Fields.Add "fldDedVal7" '87
  Fields.Add "fldDedVal8" '88
  Fields.Add "fldDedVal9" '89
  Fields.Add "fldDedVal10" '90
  Fields.Add "fldDedVal11" '91
  Fields.Add "fldDedVal12" '92
  Fields.Add "fldDedVal13" '93
  Fields.Add "fldDedVal14" '94
  Fields.Add "fldDedVal15" '95
  Fields.Add "fldDedVal16" '96
  Fields.Add "fldDedVal17" '97
  Fields.Add "fldDedVal18" '98
  Fields.Add "fldDedVal19" '99
  Fields.Add "fldDedVal20" '100
  Fields.Add "fldDedVal21" '101
  Fields.Add "fldDedVal22" '102
  Fields.Add "fldDedVal23" '103
  Fields.Add "fldDedVal24" '104
  Fields.Add "fldDedVal25" '105
  Fields.Add "fldDedVal26" '106
  Fields.Add "fldDedVal27" '107
  Fields.Add "fldDedVal28" '108
  Fields.Add "fldDedVal29" '109
  Fields.Add "fldDedVal30" '110
  Fields.Add "fldDedVal31" '111
  Fields.Add "fldDedVal32" '112
  Fields.Add "fldDedVal33" '113
  Fields.Add "fldDedVal34" '114
  Fields.Add "fldDedVal35" '115
  Fields.Add "fldDedVal36" '116
  Fields.Add "fldDedVal37" '117
  Fields.Add "fldDedVal38" '118
  Fields.Add "fldDedVal39" '119
  Fields.Add "fldDedVal40" '120
  Fields.Add "fldDedVal41" '121
  Fields.Add "fldDedVal42" '122
  Fields.Add "fldDedVal43" '123
  Fields.Add "fldDedVal44" '124
  Fields.Add "fldDedVal45" '125
  Fields.Add "fldDedVal46" '126
  Fields.Add "fldDedVal47" '127
  Fields.Add "fldDedVal48" '128
  Fields.Add "fldDedVal49" '129
  Fields.Add "fldDedVal50" '130
  Fields.Add "fldFedGrs" '131
  Fields.Add "fldStaGrs" '132
  Fields.Add "fldSocGrs" '133
  Fields.Add "fldMedGrs" '134
  Fields.Add "fldRetGrs" '135
  Fields.Add "fldTaxFRttl" '136
  Fields.Add "fldRegHrsttl" '137
  Fields.Add "fldVacttl" '138
  Fields.Add "fldSickttl" '139
  Fields.Add "fldHolttl" '140
  Fields.Add "fldCompttl" '141
  Fields.Add "fldPersttl" '142
  Fields.Add "fldTotHrsttl" '143
  Fields.Add "fldEarnDsc1ttl" '6
  Fields.Add "fldEarnDsc2ttl" '7
  Fields.Add "fldEarnDsc3ttl" '8
  Fields.Add "fldOTPaidttl" '144
  Fields.Add "fldEICttl" '145
  Fields.Add "fldRegEarnttl" '146
  Fields.Add "fldOTEarnttl" '147
  Fields.Add "fldEarn1ttl" '148
  Fields.Add "fldEarn2ttl" '149
  Fields.Add "fldEarn3ttl" '150
  Fields.Add "fldGrossPayttl" '151
  Fields.Add "fldSocSecttl" '152
  Fields.Add "fldMedttl" '153
  Fields.Add "fldDedDsc1ttl" '9
  Fields.Add "fldDedDsc2ttl" '10
  Fields.Add "fldDedDsc3ttl" '11
  Fields.Add "fldDedDsc4ttl" '12
  Fields.Add "fldDedDsc5ttl" '13
  Fields.Add "fldFWTttl" '154
  Fields.Add "fldSWTttl" '155
  Fields.Add "fldRetttl" '156
  Fields.Add "fldNetPayttl" '157
  Fields.Add "fldDedVal1ttl" '158
  Fields.Add "fldDedVal2ttl" '159
  Fields.Add "fldDedVal3ttl" '160
  Fields.Add "fldDedVal4ttl" '161
  Fields.Add "fldDedVal5ttl" '162
  Fields.Add "fldDedDsc6ttl" '14
  Fields.Add "fldDedDsc7ttl" '15
  Fields.Add "fldDedDsc8ttl" '16
  Fields.Add "fldDedDsc9ttl" '17
  Fields.Add "fldDedDsc10ttl" '18
  Fields.Add "fldDedDsc11ttl" '19
  Fields.Add "fldDedDsc12ttl" '20
  Fields.Add "fldDedDsc13ttl" '21
  Fields.Add "fldDedDsc14ttl" '22
  Fields.Add "fldDedDsc15ttl" '23
  Fields.Add "fldDedVal6ttl" '163
  Fields.Add "fldDedVal7ttl" '164
  Fields.Add "fldDedVal8ttl" '165
  Fields.Add "fldDedVal9ttl" '166
  Fields.Add "fldDedVal10ttl" '167
  Fields.Add "fldDedVal11ttl" '168
  Fields.Add "fldDedVal12ttl" '169
  Fields.Add "fldDedVal13ttl" '170
  Fields.Add "fldDedVal14ttl" '171
  Fields.Add "fldDedVal15ttl" '172
  Fields.Add "fldDedDsc16ttl" '24
  Fields.Add "fldDedDsc17ttl" '25
  Fields.Add "fldDedDsc18ttl" '26
  Fields.Add "fldDedDsc19ttl" '27
  Fields.Add "fldDedDsc20ttl" '28
  Fields.Add "fldDedDsc21ttl" '29
  Fields.Add "fldDedDsc22ttl" '30
  Fields.Add "fldDedDsc23ttl" '31
  Fields.Add "fldDedDsc24ttl" '32
  Fields.Add "fldDedDsc25ttl" '33
  Fields.Add "fldDedVal16ttl" '173
  Fields.Add "fldDedVal17ttl" '174
  Fields.Add "fldDedVal18ttl" '175
  Fields.Add "fldDedVal19ttl" '176
  Fields.Add "fldDedVal20ttl" '177
  Fields.Add "fldDedVal21ttl" '178
  Fields.Add "fldDedVal22ttl" '179
  Fields.Add "fldDedVal23ttl" '180
  Fields.Add "fldDedVal24ttl" '181
  Fields.Add "fldDedVal25ttl" '182
  Fields.Add "fldDedDsc26ttl" '34
  Fields.Add "fldDedDsc27ttl" '35
  Fields.Add "fldDedDsc28ttl" '36
  Fields.Add "fldDedDsc29ttl" '37
  Fields.Add "fldDedDsc30ttl" '38
  Fields.Add "fldDedDsc31ttl" '39
  Fields.Add "fldDedDsc32ttl" '40
  Fields.Add "fldDedDsc33ttl" '41
  Fields.Add "fldDedDsc34ttl" '42
  Fields.Add "fldDedDsc35ttl" '43
  Fields.Add "fldDedVal26ttl" '183
  Fields.Add "fldDedVal27ttl" '184
  Fields.Add "fldDedVal28ttl" '185
  Fields.Add "fldDedVal29ttl" '186
  Fields.Add "fldDedVal30ttl" '187
  Fields.Add "fldDedVal31ttl" '188
  Fields.Add "fldDedVal32ttl" '189
  Fields.Add "fldDedVal33ttl" '190
  Fields.Add "fldDedVal34ttl" '191
  Fields.Add "fldDedVal35ttl" '192
  Fields.Add "fldDedDsc36ttl" '44
  Fields.Add "fldDedDsc37ttl" '45
  Fields.Add "fldDedDsc38ttl" '46
  Fields.Add "fldDedDsc39ttl" '47
  Fields.Add "fldDedDsc40ttl" '48
  Fields.Add "fldDedDsc41ttl" '49
  Fields.Add "fldDedDsc42ttl" '50
  Fields.Add "fldDedDsc43ttl" '51
  Fields.Add "fldDedDsc44ttl" '52
  Fields.Add "fldDedDsc45ttl" '53
  Fields.Add "fldDedVal36ttl" '193
  Fields.Add "fldDedVal37ttl" '194
  Fields.Add "fldDedVal38ttl" '195
  Fields.Add "fldDedVal39ttl" '196
  Fields.Add "fldDedVal40ttl" '197
  Fields.Add "fldDedVal41ttl" '198
  Fields.Add "fldDedVal42ttl" '199
  Fields.Add "fldDedVal43ttl" '200
  Fields.Add "fldDedVal44ttl" '201
  Fields.Add "fldDedVal45ttl" '202
  Fields.Add "fldDedDsc46ttl" '54
  Fields.Add "fldDedDsc47ttl" '55
  Fields.Add "fldDedDsc48ttl" '56
  Fields.Add "fldDedDsc49ttl" '57
  Fields.Add "fldDedDsc50ttl" '58
  Fields.Add "fldDedVal46ttl" '203
  Fields.Add "fldDedVal47ttl" '204
  Fields.Add "fldDedVal48ttl" '205
  Fields.Add "fldDedVal49ttl" '206
  Fields.Add "fldDedVal50ttl" '207
  Fields.Add "fldFedGrsttl" '208
  Fields.Add "fldStaGrsttl" '209
  Fields.Add "fldSocGrsttl" '210
  Fields.Add "fldMedGrsttl" '211
  Fields.Add "fldRetGrsttl" '212
  Fields.Add "fldEmpNo" '213
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
  Fields("ghGrpHdr1").Value = arr(0) 'group totals on this value
  Fields("fldEmployer").Value = arr(1)
  Fields("fldStartDate").Value = arr(2)
  Fields("fldEndDate").Value = arr(3)
  Fields("fldEmpNum").Value = arr(4)
  Fields("fldEmpName").Value = arr(5)
  Fields("fldEarnDsc1").Value = arr(6)
  Fields("fldEarnDsc2").Value = arr(7)
  Fields("fldEarnDsc3").Value = arr(8)
  Fields("fldDedDsc1").Value = arr(9)
  Fields("fldDedDsc2").Value = arr(10)
  Fields("fldDedDsc3").Value = arr(11)
  Fields("fldDedDsc4").Value = arr(12)
  Fields("fldDedDsc5").Value = arr(13)
  Fields("fldDedDsc6").Value = arr(14)
  Fields("fldDedDsc7").Value = arr(15)
  Fields("fldDedDsc8").Value = arr(16)
  Fields("fldDedDsc9").Value = arr(17)
  Fields("fldDedDsc10").Value = arr(18)
  Fields("fldDedDsc11").Value = arr(19)
  Fields("fldDedDsc12").Value = arr(20)
  Fields("fldDedDsc13").Value = arr(21)
  Fields("fldDedDsc14").Value = arr(22)
  Fields("fldDedDsc15").Value = arr(23)
  Fields("fldDedDsc16").Value = arr(24)
  Fields("fldDedDsc17").Value = arr(25)
  Fields("fldDedDsc18").Value = arr(26)
  Fields("fldDedDsc19").Value = arr(27)
  Fields("fldDedDsc20").Value = arr(28)
  Fields("fldDedDsc21").Value = arr(29)
  Fields("fldDedDsc22").Value = arr(30)
  Fields("fldDedDsc23").Value = arr(31)
  Fields("fldDedDsc24").Value = arr(32)
  Fields("fldDedDsc25").Value = arr(33)
  Fields("fldDedDsc26").Value = arr(34)
  Fields("fldDedDsc27").Value = arr(35)
  Fields("fldDedDsc28").Value = arr(36)
  Fields("fldDedDsc29").Value = arr(37)
  Fields("fldDedDsc30").Value = arr(38)
  Fields("fldDedDsc31").Value = arr(39)
  Fields("fldDedDsc32").Value = arr(40)
  Fields("fldDedDsc33").Value = arr(41)
  Fields("fldDedDsc34").Value = arr(42)
  Fields("fldDedDsc35").Value = arr(43)
  Fields("fldDedDsc36").Value = arr(44)
  Fields("fldDedDsc37").Value = arr(45)
  Fields("fldDedDsc38").Value = arr(46)
  Fields("fldDedDsc39").Value = arr(47)
  Fields("fldDedDsc40").Value = arr(48)
  Fields("fldDedDsc41").Value = arr(49)
  Fields("fldDedDsc42").Value = arr(50)
  Fields("fldDedDsc43").Value = arr(51)
  Fields("fldDedDsc44").Value = arr(52)
  Fields("fldDedDsc45").Value = arr(53)
  Fields("fldDedDsc46").Value = arr(54)
  Fields("fldDedDsc47").Value = arr(55)
  Fields("fldDedDsc48").Value = arr(56)
  Fields("fldDedDsc49").Value = arr(57)
  Fields("fldDedDsc50").Value = arr(58)
  Fields("fldTaxFr").Value = arr(59)
  Fields("fldRegHrs").Value = arr(60)
  Fields("fldVac").Value = arr(61)
  Fields("fldSick").Value = arr(62)
  Fields("fldHol").Value = arr(63)
  Fields("fldComp").Value = arr(64)
  Fields("fldPers").Value = arr(65)
  Fields("fldTotHrs").Value = arr(66)
  Fields("fldOTPaid").Value = arr(67)
  Fields("fldEIC").Value = arr(68)
  Fields("fldRegEarn").Value = arr(69)
  Fields("fldOTEarn").Value = arr(70)
  Fields("fldEarn1").Value = arr(71)
  Fields("fldEarn2").Value = arr(72)
  Fields("fldEarn3").Value = arr(73)
  Fields("fldGrossPay").Value = arr(74)
  Fields("fldSocSec").Value = arr(75)
  Fields("fldMed").Value = arr(76)
  Fields("fldFWT").Value = arr(77)
  Fields("fldSWT").Value = arr(78)
  Fields("fldRet").Value = arr(79)
  Fields("fldNetPay").Value = arr(80)
  Fields("fldDedVal1").Value = arr(81)
  Fields("fldDedVal2").Value = arr(82)
  Fields("fldDedVal3").Value = arr(83)
  Fields("fldDedVal4").Value = arr(84)
  Fields("fldDedVal5").Value = arr(85)
  Fields("fldDedVal6").Value = arr(86)
  Fields("fldDedVal7").Value = arr(87)
  Fields("fldDedVal8").Value = arr(88)
  Fields("fldDedVal9").Value = arr(89)
  Fields("fldDedVal10").Value = arr(90)
  Fields("fldDedVal11").Value = arr(91)
  Fields("fldDedVal12").Value = arr(92)
  Fields("fldDedVal13").Value = arr(93)
  Fields("fldDedVal14").Value = arr(94)
  Fields("fldDedVal15").Value = arr(95)
  Fields("fldDedVal16").Value = arr(96)
  Fields("fldDedVal17").Value = arr(97)
  Fields("fldDedVal18").Value = arr(98)
  Fields("fldDedVal19").Value = arr(99)
  Fields("fldDedVal20").Value = arr(100)
  Fields("fldDedVal21").Value = arr(101)
  Fields("fldDedVal22").Value = arr(102)
  Fields("fldDedVal23").Value = arr(103)
  Fields("fldDedVal24").Value = arr(104)
  Fields("fldDedVal25").Value = arr(105)
  Fields("fldDedVal26").Value = arr(106)
  Fields("fldDedVal27").Value = arr(107)
  Fields("fldDedVal28").Value = arr(108)
  Fields("fldDedVal29").Value = arr(109)
  Fields("fldDedVal30").Value = arr(110)
  Fields("fldDedVal31").Value = arr(111)
  Fields("fldDedVal32").Value = arr(112)
  Fields("fldDedVal33").Value = arr(113)
  Fields("fldDedVal34").Value = arr(114)
  Fields("fldDedVal35").Value = arr(115)
  Fields("fldDedVal36").Value = arr(116)
  Fields("fldDedVal37").Value = arr(117)
  Fields("fldDedVal38").Value = arr(118)
  Fields("fldDedVal39").Value = arr(119)
  Fields("fldDedVal40").Value = arr(120)
  Fields("fldDedVal41").Value = arr(121)
  Fields("fldDedVal42").Value = arr(122)
  Fields("fldDedVal43").Value = arr(123)
  Fields("fldDedVal44").Value = arr(124)
  Fields("fldDedVal45").Value = arr(125)
  Fields("fldDedVal46").Value = arr(126)
  Fields("fldDedVal47").Value = arr(127)
  Fields("fldDedVal48").Value = arr(128)
  Fields("fldDedVal49").Value = arr(129)
  Fields("fldDedVal50").Value = arr(130)
  Fields("fldFedGrs").Value = arr(131)
  Fields("fldStaGrs").Value = arr(132)
  Fields("fldSocGrs").Value = arr(133)
  Fields("fldMedGrs").Value = arr(134)
  Fields("fldRetGrs").Value = arr(135)
'  Fields("fldEmpNamettl").Value = arr(136)
  Fields("fldTaxFRttl").Value = arr(136)
  Fields("fldRegHrsttl").Value = arr(137)
  Fields("fldVacttl").Value = arr(138)
  Fields("fldSickttl").Value = arr(139)
  Fields("fldHolttl").Value = arr(140)
  Fields("fldCompttl").Value = arr(141)
  Fields("fldPersttl").Value = arr(142)
  Fields("fldTotHrsttl").Value = arr(143)
  Fields("fldEarnDsc1ttl").Value = arr(6)
  Fields("fldEarnDsc2ttl").Value = arr(7)
  Fields("fldEarnDsc3ttl").Value = arr(8)
  Fields("fldOTPaidttl").Value = arr(144)
  Fields("fldEICttl").Value = arr(145)
  Fields("fldRegEarnttl").Value = arr(146)
  Fields("fldOTEarnttl").Value = arr(147)
  Fields("fldEarn1ttl").Value = arr(148)
  Fields("fldEarn2ttl").Value = arr(149)
  Fields("fldEarn3ttl").Value = arr(150)
  Fields("fldGrossPayttl").Value = arr(151)
  Fields("fldSocSecttl").Value = arr(152)
  Fields("fldMedttl").Value = arr(153)
  Fields("fldDedDsc1ttl").Value = arr(9)
  Fields("fldDedDsc2ttl").Value = arr(10)
  Fields("fldDedDsc3ttl").Value = arr(11)
  Fields("fldDedDsc4ttl").Value = arr(12)
  Fields("fldDedDsc5ttl").Value = arr(13)
  Fields("fldFWTttl").Value = arr(154)
  Fields("fldSWTttl").Value = arr(155)
  Fields("fldRetttl").Value = arr(156)
  Fields("fldNetPayttl").Value = arr(157)
  Fields("fldDedVal1ttl").Value = arr(158)
  Fields("fldDedVal2ttl").Value = arr(159)
  Fields("fldDedVal3ttl").Value = arr(160)
  Fields("fldDedVal4ttl").Value = arr(161)
  Fields("fldDedVal5ttl").Value = arr(162)
  Fields("fldDedDsc6ttl").Value = arr(14)
  Fields("fldDedDsc7ttl").Value = arr(15)
  Fields("fldDedDsc8ttl").Value = arr(16)
  Fields("fldDedDsc9ttl").Value = arr(17)
  Fields("fldDedDsc10ttl").Value = arr(18)
  Fields("fldDedDsc11ttl").Value = arr(19)
  Fields("fldDedDsc12ttl").Value = arr(20)
  Fields("fldDedDsc13ttl").Value = arr(21)
  Fields("fldDedDsc14ttl").Value = arr(22)
  Fields("fldDedDsc15ttl").Value = arr(23)
  Fields("fldDedVal6ttl").Value = arr(163)
  Fields("fldDedVal7ttl").Value = arr(164)
  Fields("fldDedVal8ttl").Value = arr(165)
  Fields("fldDedVal9ttl").Value = arr(166)
  Fields("fldDedVal10ttl").Value = arr(167)
  Fields("fldDedVal11ttl").Value = arr(168)
  Fields("fldDedVal12ttl").Value = arr(169)
  Fields("fldDedVal13ttl").Value = arr(170)
  Fields("fldDedVal14ttl").Value = arr(171)
  Fields("fldDedVal15ttl").Value = arr(172)
  Fields("fldDedDsc16ttl").Value = arr(24)
  Fields("fldDedDsc17ttl").Value = arr(25)
  Fields("fldDedDsc18ttl").Value = arr(26)
  Fields("fldDedDsc19ttl").Value = arr(27)
  Fields("fldDedDsc20ttl").Value = arr(28)
  Fields("fldDedDsc21ttl").Value = arr(29)
  Fields("fldDedDsc22ttl").Value = arr(30)
  Fields("fldDedDsc23ttl").Value = arr(31)
  Fields("fldDedDsc24ttl").Value = arr(32)
  Fields("fldDedDsc25ttl").Value = arr(33)
  Fields("fldDedVal16ttl").Value = arr(173)
  Fields("fldDedVal17ttl").Value = arr(174)
  Fields("fldDedVal18ttl").Value = arr(175)
  Fields("fldDedVal19ttl").Value = arr(176)
  Fields("fldDedVal20ttl").Value = arr(177)
  Fields("fldDedVal21ttl").Value = arr(178)
  Fields("fldDedVal22ttl").Value = arr(179)
  Fields("fldDedVal23ttl").Value = arr(180)
  Fields("fldDedVal24ttl").Value = arr(181)
  Fields("fldDedVal25ttl").Value = arr(182)
  Fields("fldDedDsc26ttl").Value = arr(34)
  Fields("fldDedDsc27ttl").Value = arr(35)
  Fields("fldDedDsc28ttl").Value = arr(36)
  Fields("fldDedDsc29ttl").Value = arr(37)
  Fields("fldDedDsc30ttl").Value = arr(38)
  Fields("fldDedDsc31ttl").Value = arr(39)
  Fields("fldDedDsc32ttl").Value = arr(40)
  Fields("fldDedDsc33ttl").Value = arr(41)
  Fields("fldDedDsc34ttl").Value = arr(42)
  Fields("fldDedDsc35ttl").Value = arr(43)
  Fields("fldDedVal26ttl").Value = arr(183)
  Fields("fldDedVal27ttl").Value = arr(184)
  Fields("fldDedVal28ttl").Value = arr(185)
  Fields("fldDedVal29ttl").Value = arr(186)
  Fields("fldDedVal30ttl").Value = arr(187)
  Fields("fldDedVal31ttl").Value = arr(188)
  Fields("fldDedVal32ttl").Value = arr(189)
  Fields("fldDedVal33ttl").Value = arr(190)
  Fields("fldDedVal34ttl").Value = arr(191)
  Fields("fldDedVal35ttl").Value = arr(192)
  Fields("fldDedDsc36ttl").Value = arr(44)
  Fields("fldDedDsc37ttl").Value = arr(45)
  Fields("fldDedDsc38ttl").Value = arr(46)
  Fields("fldDedDsc39ttl").Value = arr(47)
  Fields("fldDedDsc40ttl").Value = arr(48)
  Fields("fldDedDsc41ttl").Value = arr(49)
  Fields("fldDedDsc42ttl").Value = arr(50)
  Fields("fldDedDsc43ttl").Value = arr(51)
  Fields("fldDedDsc44ttl").Value = arr(52)
  Fields("fldDedDsc45ttl").Value = arr(53)
  Fields("fldDedVal36ttl").Value = arr(193)
  Fields("fldDedVal37ttl").Value = arr(194)
  Fields("fldDedVal38ttl").Value = arr(195)
  Fields("fldDedVal39ttl").Value = arr(196)
  Fields("fldDedVal40ttl").Value = arr(197)
  Fields("fldDedVal41ttl").Value = arr(198)
  Fields("fldDedVal42ttl").Value = arr(199)
  Fields("fldDedVal43ttl").Value = arr(200)
  Fields("fldDedVal44ttl").Value = arr(201)
  Fields("fldDedVal45ttl").Value = arr(202)
  Fields("fldDedDsc46ttl").Value = arr(54)
  Fields("fldDedDsc47ttl").Value = arr(55)
  Fields("fldDedDsc48ttl").Value = arr(56)
  Fields("fldDedDsc49ttl").Value = arr(57)
  Fields("fldDedDsc50ttl").Value = arr(58)
  Fields("fldDedVal46ttl").Value = arr(203)
  Fields("fldDedVal47ttl").Value = arr(204)
  Fields("fldDedVal48ttl").Value = arr(205)
  Fields("fldDedVal49ttl").Value = arr(206)
  Fields("fldDedVal50ttl").Value = arr(207)
  Fields("fldFedGrsttl").Value = arr(208)
  Fields("fldStaGrsttl").Value = arr(209)
  Fields("fldSocGrsttl").Value = arr(210)
  Fields("fldMedGrsttl").Value = arr(211)
  Fields("fldRetGrsttl").Value = arr(212)
  Fields("fldEmpNo").Value = arr(213)
  If Len(Fields("fldEarnDsc1").Value) = 0 Then
    Fields("fldEarn1").Value = ""
  End If
  If Len(Fields("fldEarnDsc2").Value) = 0 Then
    Fields("fldEarn2").Value = ""
  End If
  If Len(Fields("fldEarnDsc3").Value) = 0 Then
    Fields("fldEarn3").Value = ""
  End If
End Sub
Private Sub ActiveReport_ReportEnd()
  Unload frmLoadingRpt
  If hFile <> 0 Then
    Close #hFile
  End If
End Sub

Private Sub ActiveReport_ReportStart()
  Dim DedCnt As Integer
  Dim DedRec As DedCodeRecType
  Dim DHandle As Integer
  Dim x As Integer
  Me.Zoom = -1
  
  OpenDedCodeFile DHandle
  DedCnt = LOF(DHandle) / Len(DedRec)
  Close DHandle
  Select Case DedCnt
    Case 1 To 5
      Detail.Height = 1900
    Case 6 To 15
      Detail.Height = 2500
    Case 16 To 25
      Detail.Height = 2900
    Case 26 To 35
      Detail.Height = 3400
    Case 36 To 45
      Detail.Height = 3800
    Case 46 To 50
      Detail.Height = 4200
    Case Else
      Detail.Height = 4200
  End Select
  Me.PageSettings.RightMargin = 300
  Me.PageSettings.LeftMargin = 200
  Me.fldTimeDate.Text = Now
End Sub

Private Sub Detail_Format()
  GroupHeader1.GroupValue = Fields("ghGrpHdr1").Value
  If GroupHeader1.GroupValue = "" Then
    GroupFooter1.NewPage = 0
  End If
  If Len(Fields("fldEarnDsc3ttl").Value) = 0 Then
    Fields("fldEarn3ttl").Value = ""
  End If
  If Len(Fields("fldEarnDsc2ttl").Value) = 0 Then
    Fields("fldEarn2ttl").Value = ""
  End If
  If Len(Fields("fldEarnDsc1ttl").Value) = 0 Then
    Fields("fldEarn1ttl").Value = ""
  End If
End Sub

Private Sub ReportHeader_Format()
  ReportHeader.Height = 0
End Sub

