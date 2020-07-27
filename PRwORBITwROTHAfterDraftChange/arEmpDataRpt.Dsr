VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arEmpDataRpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Data Report"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "arEmpDataRpt.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20558
   _ExtentY        =   15637
   SectionData     =   "arEmpDataRpt.dsx":08CA
End
Attribute VB_Name = "arEmpDataRpt"
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
  Me.ToolBar.Tools.Add "&Close"
  Me.ToolBar.Tools.Add "Save/&Excel"
  Me.ToolBar.Tools.Add "&Text"
End Sub

Private Sub ActiveReport_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
  If KeyCode = vbKeyEscape Then
    Unload Me
    DoEvents
'    frmEmpDataPrint.Show
  End If
  If Shift = 4 Then
    If KeyCode = vbKeyC Then
      Unload Me
      DoEvents
'      frmEmpDataPrint.Show
    ElseIf KeyCode = vbKeyE Then
      Screen.MousePointer = vbHourglass
      ExportReport 1
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - EmpDataRpt.xls, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - EmpDataRpt.txt, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    End If
  End If
End Sub

Private Sub ActiveReport_Terminate()
  Close
End Sub

Private Sub ActiveReport_ToolbarClick(ByVal Tool As DDActiveReports2.DDTool)
  If Tool.Caption = "&Close" Then
    Unload Me
    DoEvents
'    frmEmpDataPrint.Show
  End If
  If Tool = "Save/&Excel" Then
    Screen.MousePointer = vbHourglass
    ExportReport 1
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - EmpDataRpt.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - EmpDataRpt.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "EmpDataRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "EmpDataRpt.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub ActiveReport_DataInitialize()
  hFile = FreeFile
  Open StartPath & "\PRRPTS\EMPDATAG.RPT" For Input As #hFile
  Fields.Add "fldEmpNum" '0
  Fields.Add "fldEmpLName" '1
  Fields.Add "fldSSN" '2
  Fields.Add "fldAddress1" '3
  Fields.Add "fldEmpFName" '4
  Fields.Add "fldAddress2" '5
  Fields.Add "fldEmpCity" '6
  Fields.Add "fldEmpState" '7
  Fields.Add "fldZip" '8
  Fields.Add "fldEmpBDay" '9
  Fields.Add "fldGender" '10
  Fields.Add "fldRace" '11
  Fields.Add "fldEmpRetNum" '12
  Fields.Add "fldRetType" '13
  Fields.Add "fldDraftCode" '14
  Fields.Add "fldBankAcctNo" '15
  Fields.Add "fldPreNoted" '16
  Fields.Add "fldBankName" '17
  Fields.Add "fldBankLoc" '18
  Fields.Add "fldBankTransNo" '19
  Fields.Add "fldJobTitle" '20
  Fields.Add "fldWCCode" '21
  Fields.Add "fldStatus" '22
  Fields.Add "fldBenePct" '23
  Fields.Add "fldPayType" '24
  Fields.Add "fldFreq" '25
  Fields.Add "fldRate" '26
  Fields.Add "fldOTRate" '27
  Fields.Add "fldHDate" '28
  Fields.Add "fldRDate" '29
  Fields.Add "fldTDate" '30
  Fields.Add "fldFedX" '31
  Fields.Add "fldFedAP" '32
  Fields.Add "fldFedFig" '33
  Fields.Add "fldFedSts" '34
  Fields.Add "fldFedAll" '35
  Fields.Add "fldFedAddAll" '36
  Fields.Add "fldStaX" '37*
  Fields.Add "fldStaAP" '38
  Fields.Add "fldStaFig" '39
  Fields.Add "fldStaSts" '40
  Fields.Add "fldStaAll" '41
  Fields.Add "fldStaAddAll" '42
  Fields.Add "fldSSX" '43
  Fields.Add "fldMedX" '44
  Fields.Add "fldEICCode" '45
  
  Fields.Add "fldDedDesc1" '46
  Fields.Add "fldDedAP1" '47
  Fields.Add "fldDedFig1" '48
  Fields.Add "fldDedOT1" '49
  Fields.Add "fldDedDesc2" '50
  Fields.Add "fldDedAP2" '51
  Fields.Add "fldDedFig2" '52
  Fields.Add "fldDedOT2" '53
  Fields.Add "fldDedDesc3" '54
  Fields.Add "fldDedAP3" '55
  Fields.Add "fldDedFig3" '56
  Fields.Add "fldDedOT3" '57
  Fields.Add "fldDedDesc4" '58
  Fields.Add "fldDedAP4" '59
  Fields.Add "fldDedFig4" '60
  Fields.Add "fldDedOT4" '61
  Fields.Add "fldDedDesc5" '62
  Fields.Add "fldDedAP5" '63
  Fields.Add "fldDedFig5" '64
  Fields.Add "fldDedOT5" '65
  Fields.Add "fldDedDesc6" '66
  Fields.Add "fldDedAP6" '67
  Fields.Add "fldDedFig6" '68
  Fields.Add "fldDedOT6" '69
  Fields.Add "fldDedDesc7" '70
  Fields.Add "fldDedAP7" '71
  Fields.Add "fldDedFig7" '72
  Fields.Add "fldDedOT7" '73
  Fields.Add "fldDedDesc8" '74
  Fields.Add "fldDedAP8" '75
  Fields.Add "fldDedFig8" '76
  Fields.Add "fldDedOT8" '77
  Fields.Add "fldDedDesc9" '78
  Fields.Add "fldDedAP9" '79
  Fields.Add "fldDedFig9" '80
  Fields.Add "fldDedOT9" '81
  Fields.Add "fldDedDesc10" '82
  Fields.Add "fldDedAP10" '83
  Fields.Add "fldDedFig10" '84
  Fields.Add "fldDedOT10" '85
  Fields.Add "fldDedDesc11" '86
  Fields.Add "fldDedAP11" '87
  Fields.Add "fldDedFig11" '88
  Fields.Add "fldDedOT11" '89
  Fields.Add "fldDedDesc12" '90
  Fields.Add "fldDedAP12" '91
  Fields.Add "fldDedFig12" '92
  Fields.Add "fldDedOT12" '93
  Fields.Add "fldDedDesc13" '94
  Fields.Add "fldDedAP13" '95
  Fields.Add "fldDedFig13" '96
  Fields.Add "fldDedOT13" '97
  Fields.Add "fldDedDesc14" '98
  Fields.Add "fldDedAP14" '99
  Fields.Add "fldDedFig14" '100
  Fields.Add "fldDedOT14" '101
  Fields.Add "fldDedDesc15" '102
  Fields.Add "fldDedAP15" '103
  Fields.Add "fldDedFig15" '104
  Fields.Add "fldDedOT15" '105
  Fields.Add "fldDedDesc16" '106
  Fields.Add "fldDedAP16" '107
  Fields.Add "fldDedFig16" '108
  Fields.Add "fldDedOT16" '109
  Fields.Add "fldDedDesc17" '110
  Fields.Add "fldDedAP17" '111
  Fields.Add "fldDedFig17" '112
  Fields.Add "fldDedOT17" '113
  Fields.Add "fldDedDesc18" '114
  Fields.Add "fldDedAP18" '115
  Fields.Add "fldDedFig18" '116
  Fields.Add "fldDedOT18" '117
  Fields.Add "fldDedDesc19" '118
  Fields.Add "fldDedAP19" '119
  Fields.Add "fldDedFig19" '120
  Fields.Add "fldDedOT19" '121
  Fields.Add "fldDedDesc20" '122
  Fields.Add "fldDedAP20" '123
  Fields.Add "fldDedFig20" '124
  Fields.Add "fldDedOT20" '125
  Fields.Add "fldDedDesc21" '126
  Fields.Add "fldDedAP21" '127
  Fields.Add "fldDedFig21" '128
  Fields.Add "fldDedOT21" '129
  Fields.Add "fldDedDesc22" '130
  Fields.Add "fldDedAP22" '131
  Fields.Add "fldDedFig22" '132
  Fields.Add "fldDedOT22" '133
  Fields.Add "fldDedDesc23" '134
  Fields.Add "fldDedAP23" '135
  Fields.Add "fldDedFig23" '136
  Fields.Add "fldDedOT23" '137
  Fields.Add "fldDedDesc24" '138
  Fields.Add "fldDedAP24" '139
  Fields.Add "fldDedFig24" '140
  Fields.Add "fldDedOT24" '141
  Fields.Add "fldDedDesc25" '142
  Fields.Add "fldDedAP25" '143
  Fields.Add "fldDedFig25" '144
  Fields.Add "fldDedOT25" '145
  Fields.Add "fldDedDesc26" '146
  Fields.Add "fldDedAP26" '147
  Fields.Add "fldDedFig26" '148
  Fields.Add "fldDedOT26" '149
  Fields.Add "fldDedDesc27" '150
  Fields.Add "fldDedAP27" '151
  Fields.Add "fldDedFig27" '152
  Fields.Add "fldDedOT27" '153
  Fields.Add "fldDedDesc28" '154
  Fields.Add "fldDedAP28" '155
  Fields.Add "fldDedFig28" '156
  Fields.Add "fldDedOT28" '157
  Fields.Add "fldDedDesc29" '158
  Fields.Add "fldDedAP29" '159
  Fields.Add "fldDedFig29" '160
  Fields.Add "fldDedOT29" '161
  Fields.Add "fldDedDesc30" '162
  Fields.Add "fldDedAP30" '163
  Fields.Add "fldDedFig30" '164
  Fields.Add "fldDedOT30" '165
  Fields.Add "fldDedDesc31" '166
  Fields.Add "fldDedAP31" '167
  Fields.Add "fldDedFig31" '168
  Fields.Add "fldDedOT31" '169
  Fields.Add "fldDedDesc32" '170
  Fields.Add "fldDedAP32" '171
  Fields.Add "fldDedFig32" '172
  Fields.Add "fldDedOT32" '173
  Fields.Add "fldDedDesc33" '174
  Fields.Add "fldDedAP33" '175
  Fields.Add "fldDedFig33" '176
  Fields.Add "fldDedOT33" '177
  Fields.Add "fldDedDesc34" '178
  Fields.Add "fldDedAP34" '179
  Fields.Add "fldDedFig34" '180
  Fields.Add "fldDedOT34" '181
  Fields.Add "fldDedDesc35" '182
  Fields.Add "fldDedAP35" '183
  Fields.Add "fldDedFig35" '184
  Fields.Add "fldDedOT35" '185
  Fields.Add "fldDedDesc36" '186
  Fields.Add "fldDedAP36" '187
  Fields.Add "fldDedFig36" '188
  Fields.Add "fldDedOT36" '189
  Fields.Add "fldDedDesc37" '190
  Fields.Add "fldDedAP37" '191
  Fields.Add "fldDedFig37" '192
  Fields.Add "fldDedOT37" '193
  Fields.Add "fldDedDesc38" '194
  Fields.Add "fldDedAP38" '195
  Fields.Add "fldDedFig38" '196
  Fields.Add "fldDedOT38" '197
  Fields.Add "fldDedDesc39" '198
  Fields.Add "fldDedAP39" '199
  Fields.Add "fldDedFig39" '200
  Fields.Add "fldDedOT39" '201
  Fields.Add "fldDedDesc40" '202
  Fields.Add "fldDedAP40" '203
  Fields.Add "fldDedFig40" '204
  Fields.Add "fldDedOT40" '205
  Fields.Add "fldDedDesc41" '206
  Fields.Add "fldDedAP41" '207
  Fields.Add "fldDedFig41" '208
  Fields.Add "fldDedOT41" '209
  Fields.Add "fldDedDesc42" '210
  Fields.Add "fldDedAP42" '211
  Fields.Add "fldDedFig42" '212
  Fields.Add "fldDedOT42" '213
  Fields.Add "fldDedDesc43" '214
  Fields.Add "fldDedAP43" '215
  Fields.Add "fldDedFig43" '216
  Fields.Add "fldDedOT43" '217
  Fields.Add "fldDedDesc44" '218
  Fields.Add "fldDedAP44" '219
  Fields.Add "fldDedFig44" '220
  Fields.Add "fldDedOT44" '221
  Fields.Add "fldDedDesc45" '222
  Fields.Add "fldDedAP45" '223
  Fields.Add "fldDedFig45" '224
  Fields.Add "fldDedOT45" '225
  Fields.Add "fldDedDesc46" '226
  Fields.Add "fldDedAP46" '227
  Fields.Add "fldDedFig46" '228
  Fields.Add "fldDedOT46" '229
  Fields.Add "fldDedDesc47" '230
  Fields.Add "fldDedAP47" '231
  Fields.Add "fldDedFig47" '232
  Fields.Add "fldDedOT47" '233
  Fields.Add "fldDedDesc48" '234
  Fields.Add "fldDedAP48" '235
  Fields.Add "fldDedFig48" '236
  Fields.Add "fldDedOT48" '237
  Fields.Add "fldDedDesc49" '238
  Fields.Add "fldDedAP49" '239
  Fields.Add "fldDedFig49" '240
  Fields.Add "fldDedOT49" '241
  Fields.Add "fldDedDesc50" '242
  Fields.Add "fldDedAP50" '243
  Fields.Add "fldDedFig50" '244
  Fields.Add "fldDedOT50" '245
  
  Fields.Add "fldEarnDesc1" '246
  Fields.Add "fldEarnNum1" '247
  Fields.Add "fldEarnAmt1" '248
  Fields.Add "fldEarnDesc2" '249
  Fields.Add "fldEarnNum2" '250
  Fields.Add "fldEarnAmt2" '251
  Fields.Add "fldEarnDesc3" '252
  Fields.Add "fldEarnNum3" '253
  Fields.Add "fldEarnAmt3" '254
  
  Fields.Add "fldWANums1" '255
  Fields.Add "fldWADD1" '256
  Fields.Add "fldWANums2" '257
  Fields.Add "fldWADD2" '258
  Fields.Add "fldWANums3" '259
  Fields.Add "fldWADD3" '260
  Fields.Add "fldWANums4" '261
  Fields.Add "fldWADD4" '262
  Fields.Add "fldWANums5" '263
  Fields.Add "fldWADD5" '264
  Fields.Add "fldWANums6" '265
  Fields.Add "fldWADD6" '266
  Fields.Add "fldWANums7" '267
  Fields.Add "fldWADD7" '268
  Fields.Add "fldWANums8" '269
  Fields.Add "fldWADD8" '270
 
  Fields.Add "fldVacE" '271
  Fields.Add "fldVacU" '272
  Fields.Add "fldVacB" '273
  Fields.Add "fldSLE" '274
  Fields.Add "fldSLU" '275
  Fields.Add "fldSLB" '276
  Fields.Add "fldCTE" '277
  Fields.Add "fldCTU" '278
  Fields.Add "fldCTB" '279
  Fields.Add "fldPE" '280
  Fields.Add "fldPU" '281
  Fields.Add "fldPB" '282
  Fields.Add "fldHE" '283
  Fields.Add "fldHU" '284
  Fields.Add "fldHB" '285
  Fields.Add "fldLvTbl" '286
  Fields.Add "fldXESC" '287
  Fields.Add "fldEmployer" '288
  Fields.Add "fld401K" '289
  Fields.Add "Comment" '290
End Sub

Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
    Unload frmLoadingRpt
    MsgBox "Error Number: " & Str(Number) & " " & Description, vbOKOnly, "Printer Error"
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
  Fields("fldEmpNum").Value = arr(0) 'change group value
  Fields("fldEmpLName").Value = arr(1)
  If Len(QPTrim$(arr(2))) = 9 Then
    Fields("fldSSN").Value = AddDashToSSN(arr(2))
  End If
  Fields("fldAddress1").Value = arr(3)
  Fields("fldEmpFName").Value = arr(4)
  Fields("fldAddress2").Value = arr(5)
  Fields("fldEmpCity").Value = arr(6)
  Fields("fldEmpState").Value = arr(7)
  Fields("fldZip").Value = arr(8)
  Fields("fldEmpBDay").Value = arr(9)
  Fields("fldGender").Value = arr(10)
  Fields("fldRace").Value = arr(11) & " "
  Fields("fldEmpRetNum").Value = arr(12)
  Fields("fldRetType").Value = arr(13)
  Fields("fldDraftCode").Value = arr(14)
  Fields("fldBankAcctNo").Value = arr(15)
  Fields("fldPreNoted").Value = arr(16)
  Fields("fldBankName").Value = arr(17)
  Fields("fldBankLoc").Value = arr(18)
  Fields("fldBankTransNo").Value = arr(19)
  Fields("fldJobTitle").Value = arr(20)
  Fields("fldWCCode").Value = arr(21)
  Fields("fldStatus").Value = arr(22)
  Fields("fldBenePct").Value = arr(23)
  Fields("fldPayType").Value = arr(24)
  Fields("fldFreq").Value = arr(25)
  Fields("fldRate").Value = arr(26)
  Fields("fldOTRate").Value = arr(27)
  Fields("fldHDate").Value = arr(28)
  Fields("fldRDate").Value = arr(29)
  Fields("fldTDate").Value = arr(30)
  Fields("fldFedX").Value = arr(31)
  Fields("fldFedAP").Value = arr(32)
  Fields("fldFedFig").Value = arr(33)
  Fields("fldFedSts").Value = arr(34)
  Fields("fldFedAll").Value = arr(35)
  Fields("fldFedAddAll").Value = arr(36)
  Fields("fldStaX").Value = arr(37)
  Fields("fldStaAP").Value = arr(38)
  Fields("fldStaFig").Value = arr(39)
  Fields("fldStaSts").Value = arr(40)
  Fields("fldStaAll").Value = arr(41)
  Fields("fldStaAddAll").Value = arr(42)
  Fields("fldSSX").Value = arr(43)
  Fields("fldMedX").Value = arr(44)
  Fields("fldEICCode").Value = arr(45)
  
  Fields("fldDedDesc1").Value = arr(46)
  Fields("fldDedAP1").Value = arr(47)
  Fields("fldDedFig1").Value = arr(48)
  Fields("fldDedOT1").Value = arr(49)
  Fields("fldDedDesc2").Value = arr(50)
  Fields("fldDedAP2").Value = arr(51)
  Fields("fldDedFig2").Value = arr(52)
  Fields("fldDedOT2").Value = arr(53)
  Fields("fldDedDesc3").Value = arr(54)
  Fields("fldDedAP3").Value = arr(55)
  Fields("fldDedFig3").Value = arr(56)
  Fields("fldDedOT3").Value = arr(57)
  Fields("fldDedDesc4").Value = arr(58)
  Fields("fldDedAP4").Value = arr(59)
  Fields("fldDedFig4").Value = arr(60)
  Fields("fldDedOT4").Value = arr(61)
  Fields("fldDedDesc5").Value = arr(62)
  Fields("fldDedAP5").Value = arr(63)
  Fields("fldDedFig5").Value = arr(64)
  Fields("fldDedOT5").Value = arr(65)
  Fields("fldDedDesc6").Value = arr(66)
  Fields("fldDedAP6").Value = arr(67)
  Fields("fldDedFig6").Value = arr(68)
  Fields("fldDedOT6").Value = arr(69)
  Fields("fldDedDesc7").Value = arr(70)
  Fields("fldDedAP7").Value = arr(71)
  Fields("fldDedFig7").Value = arr(72)
  Fields("fldDedOT7").Value = arr(73)
  Fields("fldDedDesc8").Value = arr(74)
  Fields("fldDedAP8").Value = arr(75)
  Fields("fldDedFig8").Value = arr(76)
  Fields("fldDedOT8").Value = arr(77)
  Fields("fldDedDesc9").Value = arr(78)
  Fields("fldDedAP9").Value = arr(79)
  Fields("fldDedFig9").Value = arr(80)
  Fields("fldDedOT9").Value = arr(81)
  Fields("fldDedDesc10").Value = arr(82)
  Fields("fldDedAP10").Value = arr(83)
  Fields("fldDedFig10").Value = arr(84)
  Fields("fldDedOT10").Value = arr(85)
  Fields("fldDedDesc11").Value = arr(86)
  Fields("fldDedAP11").Value = arr(87)
  Fields("fldDedFig11").Value = arr(88)
  Fields("fldDedOT11").Value = arr(89)
  Fields("fldDedDesc12").Value = arr(90)
  Fields("fldDedAP12").Value = arr(91)
  Fields("fldDedFig12").Value = arr(92)
  Fields("fldDedOT12").Value = arr(93)
  Fields("fldDedDesc13").Value = arr(94)
  Fields("fldDedAP13").Value = arr(95)
  Fields("fldDedFig13").Value = arr(96)
  Fields("fldDedOT13").Value = arr(97)
  Fields("fldDedDesc14").Value = arr(98)
  Fields("fldDedAP14").Value = arr(99)
  Fields("fldDedFig14").Value = arr(100)
  Fields("fldDedOT14").Value = arr(101)
  Fields("fldDedDesc15").Value = arr(102)
  Fields("fldDedAP15").Value = arr(103)
  Fields("fldDedFig15").Value = arr(104)
  Fields("fldDedOT15").Value = arr(105)
  Fields("fldDedDesc16").Value = arr(106)
  Fields("fldDedAP16").Value = arr(107)
  Fields("fldDedFig16").Value = arr(108)
  Fields("fldDedOT16").Value = arr(109)
  Fields("fldDedDesc17").Value = arr(110)
  Fields("fldDedAP17").Value = arr(111)
  Fields("fldDedFig17").Value = arr(112)
  Fields("fldDedOT17").Value = arr(113)
  Fields("fldDedDesc18").Value = arr(114)
  Fields("fldDedAP18").Value = arr(115)
  Fields("fldDedFig18").Value = arr(116)
  Fields("fldDedOT18").Value = arr(117)
  Fields("fldDedDesc19").Value = arr(118)
  Fields("fldDedAP19").Value = arr(119)
  Fields("fldDedFig19").Value = arr(120)
  Fields("fldDedOT19").Value = arr(121)
  Fields("fldDedDesc20").Value = arr(122)
  Fields("fldDedAP20").Value = arr(123)
  Fields("fldDedFig20").Value = arr(124)
  Fields("fldDedOT20").Value = arr(125)
  Fields("fldDedDesc21").Value = arr(126)
  Fields("fldDedAP21").Value = arr(127)
  Fields("fldDedFig21").Value = arr(128)
  Fields("fldDedOT21").Value = arr(129)
  Fields("fldDedDesc22").Value = arr(130)
  Fields("fldDedAP22").Value = arr(131)
  Fields("fldDedFig22").Value = arr(132)
  Fields("fldDedOT22").Value = arr(133)
  Fields("fldDedDesc23").Value = arr(134)
  Fields("fldDedAP23").Value = arr(135)
  Fields("fldDedFig23").Value = arr(136)
  Fields("fldDedOT23").Value = arr(137)
  Fields("fldDedDesc24").Value = arr(138)
  Fields("fldDedAP24").Value = arr(139)
  Fields("fldDedFig24").Value = arr(140)
  Fields("fldDedOT24").Value = arr(141)
  Fields("fldDedDesc25").Value = arr(142)
  Fields("fldDedAP25").Value = arr(143)
  Fields("fldDedFig25").Value = arr(144)
  Fields("fldDedOT25").Value = arr(145)
  Fields("fldDedDesc26").Value = arr(146)
  Fields("fldDedAP26").Value = arr(147)
  Fields("fldDedFig26").Value = arr(148)
  Fields("fldDedOT26").Value = arr(149)
  Fields("fldDedDesc27").Value = arr(150)
  Fields("fldDedAP27").Value = arr(151)
  Fields("fldDedFig27").Value = arr(152)
  Fields("fldDedOT27").Value = arr(153)
  Fields("fldDedDesc28").Value = arr(154)
  Fields("fldDedAP28").Value = arr(155)
  Fields("fldDedFig28").Value = arr(156)
  Fields("fldDedOT28").Value = arr(157)
  Fields("fldDedDesc29").Value = arr(158)
  Fields("fldDedAP29").Value = arr(159)
  Fields("fldDedFig29").Value = arr(160)
  Fields("fldDedOT29").Value = arr(161)
  Fields("fldDedDesc30").Value = arr(162)
  Fields("fldDedAP30").Value = arr(163)
  Fields("fldDedFig30").Value = arr(164)
  Fields("fldDedOT30").Value = arr(165)
  Fields("fldDedDesc31").Value = arr(166)
  Fields("fldDedAP31").Value = arr(167)
  Fields("fldDedFig31").Value = arr(168)
  Fields("fldDedOT31").Value = arr(169)
  Fields("fldDedDesc32").Value = arr(170)
  Fields("fldDedAP32").Value = arr(171)
  Fields("fldDedFig32").Value = arr(172)
  Fields("fldDedOT32").Value = arr(173)
  Fields("fldDedDesc33").Value = arr(174)
  Fields("fldDedAP33").Value = arr(175)
  Fields("fldDedFig33").Value = arr(176)
  Fields("fldDedOT33").Value = arr(177)
  Fields("fldDedDesc34").Value = arr(178)
  Fields("fldDedAP34").Value = arr(179)
  Fields("fldDedFig34").Value = arr(180)
  Fields("fldDedOT34").Value = arr(181)
  Fields("fldDedDesc35").Value = arr(182)
  Fields("fldDedAP35").Value = arr(183)
  Fields("fldDedFig35").Value = arr(184)
  Fields("fldDedOT35").Value = arr(185)
  Fields("fldDedDesc36").Value = arr(186)
  Fields("fldDedAP36").Value = arr(187)
  Fields("fldDedFig36").Value = arr(188)
  Fields("fldDedOT36").Value = arr(189)
  Fields("fldDedDesc37").Value = arr(190)
  Fields("fldDedAP37").Value = arr(191)
  Fields("fldDedFig37").Value = arr(192)
  Fields("fldDedOT37").Value = arr(193)
  Fields("fldDedDesc38").Value = arr(194)
  Fields("fldDedAP38").Value = arr(195)
  Fields("fldDedFig38").Value = arr(196)
  Fields("fldDedOT38").Value = arr(197)
  Fields("fldDedDesc39").Value = arr(198)
  Fields("fldDedAP39").Value = arr(199)
  Fields("fldDedFig39").Value = arr(200)
  Fields("fldDedOT39").Value = arr(201)
  Fields("fldDedDesc40").Value = arr(202)
  Fields("fldDedAP40").Value = arr(203)
  Fields("fldDedFig40").Value = arr(204)
  Fields("fldDedOT40").Value = arr(205)
  Fields("fldDedDesc41").Value = arr(206)
  Fields("fldDedAP41").Value = arr(207)
  Fields("fldDedFig41").Value = arr(208)
  Fields("fldDedOT41").Value = arr(209)
  Fields("fldDedDesc42").Value = arr(210)
  Fields("fldDedAP42").Value = arr(211)
  Fields("fldDedFig42").Value = arr(212)
  Fields("fldDedOT42").Value = arr(213)
  Fields("fldDedDesc43").Value = arr(214)
  Fields("fldDedAP43").Value = arr(215)
  Fields("fldDedFig43").Value = arr(216)
  Fields("fldDedOT43").Value = arr(217)
  Fields("fldDedDesc44").Value = arr(218)
  Fields("fldDedAP44").Value = arr(219)
  Fields("fldDedFig44").Value = arr(220)
  Fields("fldDedOT44").Value = arr(221)
  Fields("fldDedDesc45").Value = arr(222)
  Fields("fldDedAP45").Value = arr(223)
  Fields("fldDedFig45").Value = arr(224)
  Fields("fldDedOT45").Value = arr(225)
  Fields("fldDedDesc46").Value = arr(226)
  Fields("fldDedAP46").Value = arr(227)
  Fields("fldDedFig46").Value = arr(228)
  Fields("fldDedOT46").Value = arr(229)
  Fields("fldDedDesc47").Value = arr(230)
  Fields("fldDedAP47").Value = arr(231)
  Fields("fldDedFig47").Value = arr(232)
  Fields("fldDedOT47").Value = arr(233)
  Fields("fldDedDesc48").Value = arr(234)
  Fields("fldDedAP48").Value = arr(235)
  Fields("fldDedFig48").Value = arr(236)
  Fields("fldDedOT48").Value = arr(237)
  Fields("fldDedDesc49").Value = arr(238)
  Fields("fldDedAP49").Value = arr(239)
  Fields("fldDedFig49").Value = arr(240)
  Fields("fldDedOT49").Value = arr(241)
  Fields("fldDedDesc50").Value = arr(242)
  Fields("fldDedAP50").Value = arr(243)
  Fields("fldDedFig50").Value = arr(244)
  Fields("fldDedOT50").Value = arr(245)
 
  Fields("fldEarnDesc1").Value = arr(246)
  Fields("fldEarnNum1").Value = arr(247)
  Fields("fldEarnAmt1").Value = arr(248)
  Fields("fldEarnDesc2").Value = arr(249)
  Fields("fldEarnNum2").Value = arr(250)
  Fields("fldEarnAmt2").Value = arr(251)
  Fields("fldEarnDesc3").Value = arr(252)
  Fields("fldEarnNum3").Value = arr(253)
  Fields("fldEarnAmt3").Value = arr(254)
  
  Fields("fldWANums1").Value = arr(255)
  Fields("fldWADD1").Value = arr(256)
  Fields("fldWANums2").Value = arr(257)
  Fields("fldWADD2").Value = arr(258)
  Fields("fldWANums3").Value = arr(259)
  Fields("fldWADD3").Value = arr(260)
  Fields("fldWANums4").Value = arr(261)
  Fields("fldWADD4").Value = arr(262)
  Fields("fldWANums5").Value = arr(263)
  Fields("fldWADD5").Value = arr(264)
  Fields("fldWANums6").Value = arr(265)
  Fields("fldWADD6").Value = arr(266)
  Fields("fldWANums7").Value = arr(267)
  Fields("fldWADD7").Value = arr(268)
  Fields("fldWANums8").Value = arr(269)
  Fields("fldWADD8").Value = arr(270)
  
  
  Fields("fldVacE").Value = arr(271)
  Fields("fldVacU").Value = arr(272)
  Fields("fldVacB").Value = arr(273)
  Fields("fldSLE").Value = arr(274)
  Fields("fldSLU").Value = arr(275)
  Fields("fldSLB").Value = arr(276)
  Fields("fldCTE").Value = arr(277)
  Fields("fldCTU").Value = arr(278)
  Fields("fldCTB").Value = arr(279)
  Fields("fldPE").Value = arr(280)
  Fields("fldPU").Value = arr(281)
  Fields("fldPB").Value = arr(282)
  Fields("fldHE").Value = arr(283)
  Fields("fldHU").Value = arr(284)
  Fields("fldHB").Value = arr(285)
  Fields("fldLvTbl").Value = arr(286)
  Fields("fldXESC").Value = arr(287)
  Fields("fldEmployer").Value = arr(288)
  Fields("fld401K").Value = arr(289)
  Fields("Comment").Value = arr(290)
End Sub

Private Sub ActiveReport_ReportEnd()
  Unload frmLoadingRpt
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
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub ActiveReport_ReportStart()
  Dim DedRec As DedCodeRecType
  Dim DedHandle As Integer
  Dim DedCnt As Integer
  Me.Zoom = -1
  OpenDedCodeFile DedHandle
  DedCnt = LOF(DedHandle) / Len(DedRec)
  Close DedHandle
  Select Case DedCnt:
    Case 1 To 2
      If DedCnt = 1 Then
        fldDedDesc2.Visible = False
        fldDedAP2.Visible = False
        fldDedFig2.Visible = False
        fldDedOT2.Visible = False
        Label76.Visible = False
      End If
      Detail.Height = 11000
      Line12.Y2 = 11775
    Case 3 To 4
      If DedCnt = 3 Then
        fldDedDesc4.Visible = False
        fldDedAP4.Visible = False
        fldDedFig4.Visible = False
        fldDedOT4.Visible = False
        Label78.Visible = False
      End If
      Detail.Height = 11000
      Line12.Y2 = 12025
    Case 5 To 6
      If DedCnt = 5 Then
        fldDedDesc6.Visible = False
        fldDedAP6.Visible = False
        fldDedFig6.Visible = False
        fldDedOT6.Visible = False
        Label80.Visible = False
      End If
      Detail.Height = 12300
      Line12.Y2 = 12325
    Case 7 To 8
      If DedCnt = 7 Then
        fldDedDesc8.Visible = False
        fldDedAP8.Visible = False
        fldDedFig8.Visible = False
        fldDedOT8.Visible = False
        Label82.Visible = False
      End If
      Detail.Height = 12500
      Line12.Y2 = 12600
    Case 9 To 10
      If DedCnt = 9 Then
        fldDedDesc10.Visible = False
        fldDedAP10.Visible = False
        fldDedFig10.Visible = False
        fldDedOT10.Visible = False
        Label84.Visible = False
      End If
      Detail.Height = 12700
      Line12.Y2 = 12850
    Case 11 To 12
      If DedCnt = 11 Then
        fldDedDesc12.Visible = False
        fldDedAP12.Visible = False
        fldDedFig12.Visible = False
        fldDedOT12.Visible = False
        Label86.Visible = False
      End If
      Detail.Height = 12900
      Line12.Y2 = 13125
    Case 13 To 14
      If DedCnt = 13 Then
        fldDedDesc14.Visible = False
        fldDedAP14.Visible = False
        fldDedFig14.Visible = False
        fldDedOT14.Visible = False
        Label88.Visible = False
      End If
      Detail.Height = 13100
      Line12.Y2 = 13400
    Case 15 To 16
      If DedCnt = 15 Then
        fldDedDesc16.Visible = False
        fldDedAP16.Visible = False
        fldDedFig16.Visible = False
        fldDedOT16.Visible = False
        Label90.Visible = False
      End If
      Detail.Height = 13300
      Line12.Y2 = 13675
    Case 17 To 18
      If DedCnt = 17 Then
        fldDedDesc18.Visible = False
        fldDedAP18.Visible = False
        fldDedFig18.Visible = False
        fldDedOT18.Visible = False
        Label92.Visible = False
      End If
      Detail.Height = 13500
      Line12.Y2 = 13950
    Case 19 To 20
      If DedCnt = 19 Then
        fldDedDesc20.Visible = False
        fldDedAP20.Visible = False
        fldDedFig20.Visible = False
        fldDedOT20.Visible = False
        Label94.Visible = False
      End If
      Detail.Height = 13700
      Line12.Y2 = 14225
    Case 21 To 22
      If DedCnt = 21 Then
        fldDedDesc22.Visible = False
        fldDedAP22.Visible = False
        fldDedFig22.Visible = False
        fldDedOT22.Visible = False
        Label96.Visible = False
      End If
      Detail.Height = 13900
      Line12.Y2 = 14500
    Case 23 To 24
      If DedCnt = 23 Then
        fldDedDesc24.Visible = False
        fldDedAP24.Visible = False
        fldDedFig24.Visible = False
        fldDedOT24.Visible = False
        Label98.Visible = False
      End If
      Detail.Height = 14100
      Line12.Y2 = 14775
    Case 25 To 26
      If DedCnt = 25 Then
        fldDedDesc26.Visible = False
        fldDedAP26.Visible = False
        fldDedFig26.Visible = False
        fldDedOT26.Visible = False
        Label100.Visible = False
      End If
      Detail.Height = 14300
      Line12.Y2 = 15050
    Case 27 To 28
      If DedCnt = 27 Then
        fldDedDesc28.Visible = False
        fldDedAP28.Visible = False
        fldDedFig28.Visible = False
        fldDedOT28.Visible = False
        Label114.Visible = False
      End If
      Detail.Height = 14500
      Line12.Y2 = 15325
    Case 29 To 30
      If DedCnt = 29 Then
        fldDedDesc30.Visible = False
        fldDedAP30.Visible = False
        fldDedFig30.Visible = False
        fldDedOT30.Visible = False
        Label115.Visible = False
      End If
      Detail.Height = 14700
      Line12.Y2 = 15575
    Case 31 To 32
      If DedCnt = 31 Then
        fldDedDesc32.Visible = False
        fldDedAP32.Visible = False
        fldDedFig32.Visible = False
        fldDedOT32.Visible = False
        Label116.Visible = False
      End If
      Detail.Height = 14900
      Line12.Y2 = 15850
    Case 33 To 34
      If DedCnt = 33 Then
        fldDedDesc34.Visible = False
        fldDedAP34.Visible = False
        fldDedFig34.Visible = False
        fldDedOT34.Visible = False
        Label117.Visible = False
      End If
      Detail.Height = 15100
      Line12.Y2 = 16125
    Case 35 To 36
      If DedCnt = 35 Then
        fldDedDesc36.Visible = False
        fldDedAP36.Visible = False
        fldDedFig36.Visible = False
        fldDedOT36.Visible = False
        Label118.Visible = False
      End If
      Detail.Height = 15300
      Line12.Y2 = 16400
    Case 37 To 38
      If DedCnt = 37 Then
        fldDedDesc38.Visible = False
        fldDedAP38.Visible = False
        fldDedFig38.Visible = False
        fldDedOT38.Visible = False
        Label119.Visible = False
      End If
      Detail.Height = 15500
      Line12.Y2 = 16675
    Case 39 To 40
      If DedCnt = 39 Then
        fldDedDesc40.Visible = False
        fldDedAP40.Visible = False
        fldDedFig40.Visible = False
        fldDedOT40.Visible = False
        Label120.Visible = False
      End If
      Detail.Height = 15700
      Line12.Y2 = 16925
    Case 41 To 42
      If DedCnt = 41 Then
        fldDedDesc42.Visible = False
        fldDedAP42.Visible = False
        fldDedFig42.Visible = False
        fldDedOT42.Visible = False
        Label121.Visible = False
      End If
      Detail.Height = 15900
      Line12.Y2 = 17200
    Case 43 To 44
      If DedCnt = 43 Then
        fldDedDesc44.Visible = False
        fldDedAP44.Visible = False
        fldDedFig44.Visible = False
        fldDedOT44.Visible = False
        Label122.Visible = False
      End If
      Detail.Height = 16100
      Line12.Y2 = 17475
    Case 45 To 46
      If DedCnt = 45 Then
        fldDedDesc46.Visible = False
        fldDedAP46.Visible = False
        fldDedFig46.Visible = False
        fldDedOT46.Visible = False
        Label123.Visible = False
      End If
      Detail.Height = 16300
      Line12.Y2 = 17750
    Case 47 To 48
      If DedCnt = 47 Then
        fldDedDesc48.Visible = False
        fldDedAP48.Visible = False
        fldDedFig48.Visible = False
        fldDedOT48.Visible = False
        Label124.Visible = False
      End If
      Detail.Height = 16500
      Line12.Y2 = 18000
    Case 49 To 50
      If DedCnt = 49 Then
        fldDedDesc50.Visible = False
        fldDedAP50.Visible = False
        fldDedFig50.Visible = False
        fldDedOT50.Visible = False
        Label125.Visible = False
      End If
      Detail.Height = 16700
      Line12.Y2 = 18275
    Case Else
      Detail.Height = 16700
      Line12.Y2 = 18275
  End Select
  Me.fldTimeDate.Text = Now

End Sub

Private Sub Detail_Format()
  GroupHeader1.GroupValue = Fields("fldEmpNum").Value

End Sub


