VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arEarningsHistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Earnings History"
   ClientHeight    =   8868
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   12225
   Icon            =   "arEarningsHistory.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   21537
   _ExtentY        =   15637
   SectionData     =   "arEarningsHistory.dsx":08CA
End
Attribute VB_Name = "arEarningsHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Private HFile As Integer
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
      MsgBox "File - EarnHist.xls, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - EarnHist.txt, created in the Citipak Directory.", vbOKOnly
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
    MsgBox "File - EarnHist.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - EarnHist.txt, created in the Citipak Directory.", vbOKOnly
  End If
End Sub
Private Sub ExportReport(X As Integer)
  Dim oEXL As ActiveReportsExcelExport.ARExportExcel
  Dim oTXT As ActiveReportsTextExport.ARExportText
  Dim outfile As String
  If Right$(StartPath, 1) = ":" Then
    outfile = StartPath
  Else
    outfile = StartPath & "\"
  End If
  Select Case X
    Case 1   '"Excel"
        Set oEXL = New ActiveReportsExcelExport.ARExportExcel
        oEXL.FileName = outfile & "EarnHist.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "EarnHist.txt"
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
    Me.Visible = False
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
  HFile = FreeFile
  Open StartPath & "\PRRPTS\EMPHISTG.RPT" For Input As #HFile
  Fields.Add "ghGrpHdr1" '0
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
  Fields.Add "fldTransDateDet" '59
  Fields.Add "fldCheckNo" '60
  Fields.Add "fldTaxFr" '61
  Fields.Add "fldRegHrsDet" '62
  Fields.Add "fldVacDet" '63
  Fields.Add "fldSickDet" '64
  Fields.Add "fldHolDet" '65
  Fields.Add "fldCompDet" '66
  Fields.Add "fldPersDet" '67
  Fields.Add "fldTotHrsDet" '68
  Fields.Add "fldOTPaidDet" '69
  Fields.Add "fldEICDet" '70
  Fields.Add "fldRegEarnDet" '71
  Fields.Add "fldOTEarnDet" '72
  Fields.Add "fldEarn1Det" '73
  Fields.Add "fldEarn2Det" '74
  Fields.Add "fldEarn3Det" '75
  Fields.Add "fldGrossPayDet" '76
  Fields.Add "fldSocSecDet" '77
  Fields.Add "fldMedDet" '78
  Fields.Add "fldFWTDet" '79
  Fields.Add "fldSWTDet" '80
  Fields.Add "fldRetDet" '81
  Fields.Add "fldNetPayDet" '82
  Fields.Add "fldDedVal1Det" '83
  Fields.Add "fldDedVal2Det" '84
  Fields.Add "fldDedVal3Det" '85
  Fields.Add "fldDedVal4Det" '86
  Fields.Add "fldDedVal5Det" '87
  Fields.Add "fldDedVal6Det" '88
  Fields.Add "fldDedVal7Det" '89
  Fields.Add "fldDedVal8Det" '90
  Fields.Add "fldDedVal9Det" '91
  Fields.Add "fldDedVal10Det" '92
  Fields.Add "fldDedVal11Det" '93
  Fields.Add "fldDedVal12Det" '94
  Fields.Add "fldDedVal13Det" '95
  Fields.Add "fldDedVal14Det" '96
  Fields.Add "fldDedVal15Det" '97
  Fields.Add "fldDedVal16Det" '98
  Fields.Add "fldDedVal17Det" '99
  Fields.Add "fldDedVal18Det" '100
  Fields.Add "fldDedVal19Det" '101
  Fields.Add "fldDedVal20Det" '102
  Fields.Add "fldDedVal21Det" '103
  Fields.Add "fldDedVal22Det" '104
  Fields.Add "fldDedVal23Det" '105
  Fields.Add "fldDedVal24Det" '106
  Fields.Add "fldDedVal25Det" '107
  Fields.Add "fldDedVal26Det" '108
  Fields.Add "fldDedVal27Det" '109
  Fields.Add "fldDedVal28Det" '110
  Fields.Add "fldDedVal29Det" '111
  Fields.Add "fldDedVal30Det" '112
  Fields.Add "fldDedVal31Det" '113
  Fields.Add "fldDedVal32Det" '114
  Fields.Add "fldDedVal33Det" '115
  Fields.Add "fldDedVal34Det" '116
  Fields.Add "fldDedVal35Det" '117
  Fields.Add "fldDedVal36Det" '118
  Fields.Add "fldDedVal37Det" '119
  Fields.Add "fldDedVal38Det" '120
  Fields.Add "fldDedVal39Det" '121
  Fields.Add "fldDedVal40Det" '122
  Fields.Add "fldDedVal41Det" '123
  Fields.Add "fldDedVal42Det" '124
  Fields.Add "fldDedVal43Det" '125
  Fields.Add "fldDedVal44Det" '126
  Fields.Add "fldDedVal45Det" '127
  Fields.Add "fldDedVal46Det" '128
  Fields.Add "fldDedVal47Det" '129
  Fields.Add "fldDedVal48Det" '130
  Fields.Add "fldDedVal49Det" '131
  Fields.Add "fldDedVal50Det" '132
  
  Fields.Add "fldTaxFRftr" '133
  Fields.Add "fldRegHrsftr" '134
  Fields.Add "fldVacftr" '135
  Fields.Add "fldSickftr" '136
  Fields.Add "fldHolftr" '137
  
  Fields.Add "fldCompftr" '138
  Fields.Add "fldPersftr" '139
  Fields.Add "fldTotHrsftr" '140
  Fields.Add "fldEarnDsc3ftr" '141
  Fields.Add "fldEarnDsc2ftr" '142
  
  Fields.Add "fldEarnDsc1ftr" '143
  Fields.Add "fldOTPaidftr" '144
  Fields.Add "fldEICftr" '145
  Fields.Add "fldRegEarnftr" '146
  Fields.Add "fldOTEarnftr" '147
  
  Fields.Add "fldEarn3ftr" '148
  Fields.Add "fldEarn2ftr" '149
  Fields.Add "fldEarn1ftr" '150
  Fields.Add "fldGrossPayftr" '151
  Fields.Add "fldSocSecftr" '152
  Fields.Add "fldMedftr" '153
  Fields.Add "fldDedDsc1ftr" '154
  Fields.Add "fldDedDsc2ftr" '155
  Fields.Add "fldDedDsc3ftr" '156
  Fields.Add "fldDedDsc4ftr" '157
  Fields.Add "fldDedDsc5ftr" '158
  Fields.Add "fldFWTftr" '159
  Fields.Add "fldSWTftr" '160
  Fields.Add "fldRetftr" '161
  Fields.Add "fldNetPayftr" '162
  Fields.Add "fldDedVal1ftr" '163
  Fields.Add "fldDedVal2ftr" '164
  Fields.Add "fldDedVal3ftr" '165
  Fields.Add "fldDedVal4ftr" '166
  Fields.Add "fldDedVal5ftr" '167
  
  Fields.Add "fldDedDsc6ftr" '168
  Fields.Add "fldDedDsc7ftr" '169
  Fields.Add "fldDedDsc8ftr" '170
  Fields.Add "fldDedDsc9ftr" '171
  Fields.Add "fldDedDsc10ftr" '172
  Fields.Add "fldDedDsc11ftr" '173
  Fields.Add "fldDedDsc12ftr" '174
  Fields.Add "fldDedDsc13ftr" '175
  Fields.Add "fldDedDsc14ftr" '176
  Fields.Add "fldDedDsc15ftr" '177
  Fields.Add "fldDedVal6ftr" '178
  Fields.Add "fldDedVal7ftr" '179
  Fields.Add "fldDedVal8ftr" '180
  Fields.Add "fldDedVal9ftr" '181
  Fields.Add "fldDedVal10ftr" '182
  Fields.Add "fldDedVal11ftr" '183
  Fields.Add "fldDedVal12ftr" '184
  Fields.Add "fldDedVal13ftr" '185
  Fields.Add "fldDedVal14ftr" '186
  Fields.Add "fldDedVal15ftr" '187
  
  Fields.Add "fldDedDsc16ftr" '188
  Fields.Add "fldDedDsc17ftr" '189
  Fields.Add "fldDedDsc18ftr" '190
  Fields.Add "fldDedDsc19ftr" '191
  Fields.Add "fldDedDsc20ftr" '192
  Fields.Add "fldDedDsc21ftr" '193
  Fields.Add "fldDedDsc22ftr" '194
  Fields.Add "fldDedDsc23ftr" '195
  Fields.Add "fldDedDsc24ftr" '196
  Fields.Add "fldDedDsc25ftr" '197
  Fields.Add "fldDedVal16ftr" '198
  Fields.Add "fldDedVal17ftr" '199
  Fields.Add "fldDedVal18ftr" '200
  Fields.Add "fldDedVal19ftr" '201
  Fields.Add "fldDedVal20ftr" '202
  Fields.Add "fldDedVal21ftr" '203
  Fields.Add "fldDedVal22ftr" '204
  Fields.Add "fldDedVal23ftr" '205
  Fields.Add "fldDedVal24ftr" '206
  Fields.Add "fldDedVal25ftr" '207
  
  Fields.Add "fldDedDsc26ftr" '208
  Fields.Add "fldDedDsc27ftr" '209
  Fields.Add "fldDedDsc28ftr" '210
  Fields.Add "fldDedDsc29ftr" '211
  Fields.Add "fldDedDsc30ftr" '212
  Fields.Add "fldDedDsc31ftr" '213
  Fields.Add "fldDedDsc32ftr" '214
  Fields.Add "fldDedDsc33ftr" '215
  Fields.Add "fldDedDsc34ftr" '216
  Fields.Add "fldDedDsc35ftr" '217
  Fields.Add "fldDedVal26ftr" '218
  Fields.Add "fldDedVal27ftr" '219
  Fields.Add "fldDedVal28ftr" '220
  Fields.Add "fldDedVal29ftr" '221
  Fields.Add "fldDedVal30ftr" '222
  Fields.Add "fldDedVal31ftr" '223
  Fields.Add "fldDedVal32ftr" '224
  Fields.Add "fldDedVal33ftr" '225
  Fields.Add "fldDedVal34ftr" '226
  Fields.Add "fldDedVal35ftr" '227
  
  Fields.Add "fldDedDsc36ftr" '228
  Fields.Add "fldDedDsc37ftr" '229
  Fields.Add "fldDedDsc38ftr" '230
  Fields.Add "fldDedDsc39ftr" '231
  Fields.Add "fldDedDsc40ftr" '232
  Fields.Add "fldDedDsc41ftr" '233
  Fields.Add "fldDedDsc42ftr" '234
  Fields.Add "fldDedDsc43ftr" '235
  Fields.Add "fldDedDsc44ftr" '236
  Fields.Add "fldDedDsc45ftr" '237
  Fields.Add "fldDedVal36ftr" '238
  Fields.Add "fldDedVal37ftr" '239
  Fields.Add "fldDedVal38ftr" '240
  Fields.Add "fldDedVal39ftr" '241
  Fields.Add "fldDedVal40ftr" '242
  Fields.Add "fldDedVal41ftr" '243
  Fields.Add "fldDedVal42ftr" '244
  Fields.Add "fldDedVal43ftr" '245
  Fields.Add "fldDedVal44ftr" '246
  Fields.Add "fldDedVal45ftr" '247
  
  Fields.Add "fldDedDsc46ftr" '248
  Fields.Add "fldDedDsc47ftr" '249
  Fields.Add "fldDedDsc48ftr" '250
  Fields.Add "fldDedDsc49ftr" '251
  Fields.Add "fldDedDsc50ftr" '252
  
  Fields.Add "fldDedVal46ftr" '253
  Fields.Add "fldDedVal47ftr" '254
  Fields.Add "fldDedVal48ftr" '255
  Fields.Add "fldDedVal49ftr" '256
  Fields.Add "fldDedVal50ftr" '257
  
  Fields.Add "fldFedGrsftr" '258
  Fields.Add "fldStaGrsftr" '259
  Fields.Add "fldSocGrsftr" '260
  Fields.Add "fldMedGrsftr" '261
  Fields.Add "fldRetGrsftr" '262
  Fields.Add "fldEmpFNameftr" '263
  Fields.Add "fldEmpLNameftr" '264
  Fields.Add "fldTaxFRttl" '265
  Fields.Add "fldRegHrsttl" '266
  Fields.Add "fldVacttl" '267
  Fields.Add "fldSickttl" '268
  Fields.Add "fldHolttl" '269
  Fields.Add "fldCompttl" '270
  Fields.Add "fldPersttl" '271
  Fields.Add "fldTotHrsttl" '272
  Fields.Add "fldEarnDsc1ttl" '273
  Fields.Add "fldEarnDsc2ttl" '274
  Fields.Add "fldEarnDsc3ttl" '275
  Fields.Add "fldOTPaidttl" '276
  Fields.Add "fldEICttl" '277
  Fields.Add "fldRegEarnttl" '278
  Fields.Add "fldOTEarnttl" '279
  Fields.Add "fldEarn1ttl" '280
  Fields.Add "fldEarn2ttl" '281
  Fields.Add "fldEarn3ttl" '282
  Fields.Add "fldGrossPayttl" '283
  Fields.Add "fldSocSecttl" '284
  Fields.Add "fldMedttl" '285
  Fields.Add "fldDedDsc1ttl" '286
  Fields.Add "fldDedDsc2ttl" '287
  Fields.Add "fldDedDsc3ttl" '288
  Fields.Add "fldDedDsc4ttl" '289
  Fields.Add "fldDedDsc5ttl" '290
  Fields.Add "fldFWTttl" '291
  Fields.Add "fldSWTttl" '292
  Fields.Add "fldRetttl" '293
  Fields.Add "fldNetPayttl" '294
  Fields.Add "fldDedVal1ttl" '295
  Fields.Add "fldDedVal2ttl" '296
  Fields.Add "fldDedVal3ttl" '297
  Fields.Add "fldDedVal4ttl" '298
  Fields.Add "fldDedVal5ttl" '299
  Fields.Add "fldDedDsc6ttl" '300
  Fields.Add "fldDedDsc7ttl" '301
  Fields.Add "fldDedDsc8ttl" '302
  Fields.Add "fldDedDsc9ttl" '303
  Fields.Add "fldDedDsc10ttl" '304
  Fields.Add "fldDedDsc11ttl" '305
  Fields.Add "fldDedDsc12ttl" '306
  Fields.Add "fldDedDsc13ttl" '307
  Fields.Add "fldDedDsc14ttl" '308
  Fields.Add "fldDedDsc15ttl" '309
  Fields.Add "fldDedVal6ttl" '310
  Fields.Add "fldDedVal7ttl" '311
  Fields.Add "fldDedVal8ttl" '312
  Fields.Add "fldDedVal9ttl" '313
  Fields.Add "fldDedVal10ttl" '314
  Fields.Add "fldDedVal11ttl" '315
  Fields.Add "fldDedVal12ttl" '316
  Fields.Add "fldDedVal13ttl" '317
  Fields.Add "fldDedVal14ttl" '318
  Fields.Add "fldDedVal15ttl" '319
  Fields.Add "fldDedDsc16ttl" '320
  Fields.Add "fldDedDsc17ttl" '321
  Fields.Add "fldDedDsc18ttl" '322
  Fields.Add "fldDedDsc19ttl" '323
  Fields.Add "fldDedDsc20ttl" '324
  Fields.Add "fldDedDsc21ttl" '325
  Fields.Add "fldDedDsc22ttl" '326
  Fields.Add "fldDedDsc23ttl" '327
  Fields.Add "fldDedDsc24ttl" '328
  Fields.Add "fldDedDsc25ttl" '329
  Fields.Add "fldDedVal16ttl" '330
  Fields.Add "fldDedVal17ttl" '331
  Fields.Add "fldDedVal18ttl" '332
  Fields.Add "fldDedVal19ttl" '333
  Fields.Add "fldDedVal20ttl" '334
  Fields.Add "fldDedVal21ttl" '335
  Fields.Add "fldDedVal22ttl" '336
  Fields.Add "fldDedVal23ttl" '337
  Fields.Add "fldDedVal24ttl" '338
  Fields.Add "fldDedVal25ttl" '339
  Fields.Add "fldDedDsc26ttl" '340
  Fields.Add "fldDedDsc27ttl" '341
  Fields.Add "fldDedDsc28ttl" '342
  Fields.Add "fldDedDsc29ttl" '343
  Fields.Add "fldDedDsc30ttl" '344
  Fields.Add "fldDedDsc31ttl" '345
  Fields.Add "fldDedDsc32ttl" '346
  Fields.Add "fldDedDsc33ttl" '347
  Fields.Add "fldDedDsc34ttl" '348
  Fields.Add "fldDedDsc35ttl" '349
  Fields.Add "fldDedVal26ttl" '350
  Fields.Add "fldDedVal27ttl" '351
  Fields.Add "fldDedVal28ttl" '352
  Fields.Add "fldDedVal29ttl" '353
  Fields.Add "fldDedVal30ttl" '354
  Fields.Add "fldDedVal31ttl" '355
  Fields.Add "fldDedVal32ttl" '356
  Fields.Add "fldDedVal33ttl" '357
  Fields.Add "fldDedVal34ttl" '358
  Fields.Add "fldDedVal35ttl" '359
  Fields.Add "fldDedDsc36ttl" '360
  Fields.Add "fldDedDsc37ttl" '361
  Fields.Add "fldDedDsc38ttl" '362
  Fields.Add "fldDedDsc39ttl" '363
  Fields.Add "fldDedDsc40ttl" '364
  Fields.Add "fldDedDsc41ttl" '365
  Fields.Add "fldDedDsc42ttl" '366
  Fields.Add "fldDedDsc43ttl" '367
  Fields.Add "fldDedDsc44ttl" '368
  Fields.Add "fldDedDsc45ttl" '369
  Fields.Add "fldDedVal36ttl" '370
  Fields.Add "fldDedVal37ttl" '371
  Fields.Add "fldDedVal38ttl" '372
  Fields.Add "fldDedVal39ttl" '373
  Fields.Add "fldDedVal40ttl" '374
  Fields.Add "fldDedVal41ttl" '375
  Fields.Add "fldDedVal42ttl" '376
  Fields.Add "fldDedVal43ttl" '377
  Fields.Add "fldDedVal44ttl" '378
  Fields.Add "fldDedVal45ttl" '379
  Fields.Add "fldDedDsc46ttl" '380
  Fields.Add "fldDedDsc47ttl" '381
  Fields.Add "fldDedDsc48ttl" '382
  Fields.Add "fldDedDsc49ttl" '383
  Fields.Add "fldDedDsc50ttl" '384
  Fields.Add "fldDedVal46ttl" '385
  Fields.Add "fldDedVal47ttl" '386
  Fields.Add "fldDedVal48ttl" '387
  Fields.Add "fldDedVal49ttl" '388
  Fields.Add "fldDedVal50ttl" '389
  Fields.Add "fldFedGrsttl" '390
  Fields.Add "fldStaGrsttl" '391
  Fields.Add "fldSocGrsttl" '392
  Fields.Add "fldMedGrsttl" '393
  Fields.Add "fldRetGrsttl" '394
  End Sub
Private Sub ActiveReport_FetchData(eof As Boolean)
  Dim sLine As String
  Dim arr() As String
  
  If VBA.eof(HFile) Then
    eof = True
    Exit Sub
  Else
    eof = False
  End If
  Line Input #HFile, sLine
  arr = Split(sLine, "~")
  ' Here we set the values of the fields that we defines as unbound
  ' or user defined.
  Fields("ghGrpHdr1").Value = arr(0)
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
  Fields("fldTransDateDet").Value = arr(59)
  Fields("fldCheckNo").Value = arr(60)
  Fields("fldTaxFr").Value = arr(61)
  Fields("fldRegHrsDet").Value = arr(62)
  Fields("fldVacDet").Value = arr(63)
  Fields("fldSickDet").Value = arr(64)
  Fields("fldHolDet").Value = arr(65)
  Fields("fldCompDet").Value = arr(66)
  Fields("fldPersDet").Value = arr(67)
  Fields("fldTotHrsDet").Value = arr(68)
  Fields("fldOTPaidDet").Value = arr(69)
  Fields("fldEICDet").Value = arr(70)
  Fields("fldRegEarnDet").Value = arr(71)
  Fields("fldOTEarnDet").Value = arr(72)
  Fields("fldEarn1Det").Value = arr(73)
  Fields("fldEarn2Det").Value = arr(74)
  Fields("fldEarn3Det").Value = arr(75)
  Fields("fldGrossPayDet").Value = arr(76)
  Fields("fldSocSecDet").Value = arr(77)
  Fields("fldMedDet").Value = arr(78)
  Fields("fldFWTDet").Value = arr(79)
  Fields("fldSWTDet").Value = arr(80)
  Fields("fldRetDet").Value = arr(81)
  Fields("fldNetPayDet").Value = arr(82)
  Fields("fldDedVal1Det").Value = arr(83)
  Fields("fldDedVal2Det").Value = arr(84)
  Fields("fldDedVal3Det").Value = arr(85)
  Fields("fldDedVal4Det").Value = arr(86)
  Fields("fldDedVal5Det").Value = arr(87)
  Fields("fldDedVal6Det").Value = arr(88)
  Fields("fldDedVal7Det").Value = arr(89)
  Fields("fldDedVal8Det").Value = arr(90)
  Fields("fldDedVal9Det").Value = arr(91)
  Fields("fldDedVal10Det").Value = arr(92)
  Fields("fldDedVal11Det").Value = arr(93)
  Fields("fldDedVal12Det").Value = arr(94)
  Fields("fldDedVal13Det").Value = arr(95)
  Fields("fldDedVal14Det").Value = arr(96)
  Fields("fldDedVal15Det").Value = arr(97)
  Fields("fldDedVal16Det").Value = arr(98)
  Fields("fldDedVal17Det").Value = arr(99)
  Fields("fldDedVal18Det").Value = arr(100)
  Fields("fldDedVal19Det").Value = arr(101)
  Fields("fldDedVal20Det").Value = arr(102)
  Fields("fldDedVal21Det").Value = arr(103)
  Fields("fldDedVal22Det").Value = arr(104)
  Fields("fldDedVal23Det").Value = arr(105)
  Fields("fldDedVal24Det").Value = arr(106)
  Fields("fldDedVal25Det").Value = arr(107)
  Fields("fldDedVal26Det").Value = arr(108)
  Fields("fldDedVal27Det").Value = arr(109)
  Fields("fldDedVal28Det").Value = arr(110)
  Fields("fldDedVal29Det").Value = arr(111)
  Fields("fldDedVal30Det").Value = arr(112)
  Fields("fldDedVal31Det").Value = arr(113)
  Fields("fldDedVal32Det").Value = arr(114)
  Fields("fldDedVal33Det").Value = arr(115)
  Fields("fldDedVal34Det").Value = arr(116)
  Fields("fldDedVal35Det").Value = arr(117)
  Fields("fldDedVal36Det").Value = arr(118)
  Fields("fldDedVal37Det").Value = arr(119)
  Fields("fldDedVal38Det").Value = arr(120)
  Fields("fldDedVal39Det").Value = arr(121)
  Fields("fldDedVal40Det").Value = arr(122)
  Fields("fldDedVal41Det").Value = arr(123)
  Fields("fldDedVal42Det").Value = arr(124)
  Fields("fldDedVal43Det").Value = arr(125)
  Fields("fldDedVal44Det").Value = arr(126)
  Fields("fldDedVal45Det").Value = arr(127)
  Fields("fldDedVal46Det").Value = arr(128)
  Fields("fldDedVal47Det").Value = arr(129)
  Fields("fldDedVal48Det").Value = arr(130)
  Fields("fldDedVal49Det").Value = arr(131)
  Fields("fldDedVal50Det").Value = arr(132)
  
  Fields("fldTaxFRftr").Value = arr(133)
  Fields("fldRegHrsftr").Value = arr(134)
  Fields("fldVacftr").Value = arr(135)
  Fields("fldSickftr").Value = arr(136)
  Fields("fldHolftr").Value = arr(137)
  
  Fields("fldCompftr").Value = arr(138)
  Fields("fldPersftr").Value = arr(139)
  Fields("fldTotHrsftr").Value = arr(140)
  Fields("fldEarnDsc3ftr").Value = arr(141)
  Fields("fldEarnDsc2ftr").Value = arr(142)
  
  Fields("fldEarnDsc1ftr").Value = arr(143)
  Fields("fldOTPaidftr").Value = arr(144)
  Fields("fldEICftr").Value = arr(145)
  Fields("fldRegEarnftr").Value = arr(146)
  Fields("fldOTEarnftr").Value = arr(147)
  
  Fields("fldEarn3ftr").Value = arr(148)
  Fields("fldEarn2ftr").Value = arr(149)
  Fields("fldEarn1ftr").Value = arr(150)
  Fields("fldGrossPayftr").Value = arr(151)
  Fields("fldSocSecftr").Value = arr(152)
  Fields("fldMedftr").Value = arr(153)
  Fields("fldDedDsc1ftr").Value = arr(154)
  Fields("fldDedDsc2ftr").Value = arr(155)
  Fields("fldDedDsc3ftr").Value = arr(156)
  Fields("fldDedDsc4ftr").Value = arr(157)
  Fields("fldDedDsc5ftr").Value = arr(158)
  Fields("fldFWTftr").Value = arr(159)
  Fields("fldSWTftr").Value = arr(160)
  Fields("fldRetftr").Value = arr(161)
  Fields("fldNetPayftr").Value = arr(162)
  Fields("fldDedVal1ftr").Value = arr(163)
  Fields("fldDedVal2ftr").Value = arr(164)
  Fields("fldDedVal3ftr").Value = arr(165)
  Fields("fldDedVal4ftr").Value = arr(166)
  Fields("fldDedVal5ftr").Value = arr(167)
  
  Fields("fldDedDsc6ftr").Value = arr(168)
  Fields("fldDedDsc7ftr").Value = arr(169)
  Fields("fldDedDsc8ftr").Value = arr(170)
  Fields("fldDedDsc9ftr").Value = arr(171)
  Fields("fldDedDsc10ftr").Value = arr(172)
  Fields("fldDedDsc11ftr").Value = arr(173)
  Fields("fldDedDsc12ftr").Value = arr(174)
  Fields("fldDedDsc13ftr").Value = arr(175)
  Fields("fldDedDsc14ftr").Value = arr(176)
  Fields("fldDedDsc15ftr").Value = arr(177)
  Fields("fldDedVal6ftr").Value = arr(178)
  Fields("fldDedVal7ftr").Value = arr(179)
  Fields("fldDedVal8ftr").Value = arr(180)
  Fields("fldDedVal9ftr").Value = arr(181)
  Fields("fldDedVal10ftr").Value = arr(182)
  Fields("fldDedVal11ftr").Value = arr(183)
  Fields("fldDedVal12ftr").Value = arr(184)
  Fields("fldDedVal13ftr").Value = arr(185)
  Fields("fldDedVal14ftr").Value = arr(186)
  Fields("fldDedVal15ftr").Value = arr(187)
  
  Fields("fldDedDsc16ftr").Value = arr(188)
  Fields("fldDedDsc17ftr").Value = arr(189)
  Fields("fldDedDsc18ftr").Value = arr(190)
  Fields("fldDedDsc19ftr").Value = arr(191)
  Fields("fldDedDsc20ftr").Value = arr(192)
  Fields("fldDedDsc21ftr").Value = arr(193)
  Fields("fldDedDsc22ftr").Value = arr(194)
  Fields("fldDedDsc23ftr").Value = arr(195)
  Fields("fldDedDsc24ftr").Value = arr(196)
  Fields("fldDedDsc25ftr").Value = arr(197)
  Fields("fldDedVal16ftr").Value = arr(198)
  Fields("fldDedVal17ftr").Value = arr(199)
  Fields("fldDedVal18ftr").Value = arr(200)
  Fields("fldDedVal19ftr").Value = arr(201)
  Fields("fldDedVal20ftr").Value = arr(202)
  Fields("fldDedVal21ftr").Value = arr(203)
  Fields("fldDedVal22ftr").Value = arr(204)
  Fields("fldDedVal23ftr").Value = arr(205)
  Fields("fldDedVal24ftr").Value = arr(206)
  Fields("fldDedVal25ftr").Value = arr(207)
  
  Fields("fldDedDsc26ftr").Value = arr(208)
  Fields("fldDedDsc27ftr").Value = arr(209)
  Fields("fldDedDsc28ftr").Value = arr(210)
  Fields("fldDedDsc29ftr").Value = arr(211)
  Fields("fldDedDsc30ftr").Value = arr(212)
  Fields("fldDedDsc31ftr").Value = arr(213)
  Fields("fldDedDsc32ftr").Value = arr(214)
  Fields("fldDedDsc33ftr").Value = arr(215)
  Fields("fldDedDsc34ftr").Value = arr(216)
  Fields("fldDedDsc35ftr").Value = arr(217)
  Fields("fldDedVal26ftr").Value = arr(218)
  Fields("fldDedVal27ftr").Value = arr(219)
  Fields("fldDedVal28ftr").Value = arr(220)
  Fields("fldDedVal29ftr").Value = arr(221)
  Fields("fldDedVal30ftr").Value = arr(222)
  Fields("fldDedVal31ftr").Value = arr(223)
  Fields("fldDedVal32ftr").Value = arr(224)
  Fields("fldDedVal33ftr").Value = arr(225)
  Fields("fldDedVal34ftr").Value = arr(226)
  Fields("fldDedVal35ftr").Value = arr(227)
  
  Fields("fldDedDsc36ftr").Value = arr(228)
  Fields("fldDedDsc37ftr").Value = arr(229)
  Fields("fldDedDsc38ftr").Value = arr(230)
  Fields("fldDedDsc39ftr").Value = arr(231)
  Fields("fldDedDsc40ftr").Value = arr(232)
  Fields("fldDedDsc41ftr").Value = arr(233)
  Fields("fldDedDsc42ftr").Value = arr(234)
  Fields("fldDedDsc43ftr").Value = arr(235)
  Fields("fldDedDsc44ftr").Value = arr(236)
  Fields("fldDedDsc45ftr").Value = arr(237)
  Fields("fldDedVal36ftr").Value = arr(238)
  Fields("fldDedVal37ftr").Value = arr(239)
  Fields("fldDedVal38ftr").Value = arr(240)
  Fields("fldDedVal39ftr").Value = arr(241)
  Fields("fldDedVal40ftr").Value = arr(242)
  Fields("fldDedVal41ftr").Value = arr(243)
  Fields("fldDedVal42ftr").Value = arr(244)
  Fields("fldDedVal43ftr").Value = arr(245)
  Fields("fldDedVal44ftr").Value = arr(246)
  Fields("fldDedVal45ftr").Value = arr(247)
  
  Fields("fldDedDsc46ftr").Value = arr(248)
  Fields("fldDedDsc47ftr").Value = arr(249)
  Fields("fldDedDsc48ftr").Value = arr(250)
  Fields("fldDedDsc49ftr").Value = arr(251)
  Fields("fldDedDsc50ftr").Value = arr(252)
  
  Fields("fldDedVal46ftr").Value = arr(253)
  Fields("fldDedVal47ftr").Value = arr(254)
  Fields("fldDedVal48ftr").Value = arr(255)
  Fields("fldDedVal49ftr").Value = arr(256)
  Fields("fldDedVal50ftr").Value = arr(257)
  
  Fields("fldFedGrsftr").Value = arr(258)
  Fields("fldStaGrsftr").Value = arr(259)
  Fields("fldSocGrsftr").Value = arr(260)
  Fields("fldMedGrsftr").Value = arr(261)
  Fields("fldRetGrsftr").Value = arr(262)
  Fields("fldEmpFNameftr").Value = arr(263)
  Fields("fldEmpLNameftr").Value = arr(264)
  Fields("fldTaxFRttl").Value = arr(265)
  Fields("fldRegHrsttl").Value = arr(266)
  Fields("fldVacttl").Value = arr(267)
  Fields("fldSickttl").Value = arr(268)
  Fields("fldHolttl").Value = arr(269)
  Fields("fldCompttl").Value = arr(270)
  Fields("fldPersttl").Value = arr(271)
  Fields("fldTotHrsttl").Value = arr(272)
  Fields("fldEarnDsc3ttl").Value = arr(273)
  Fields("fldEarnDsc2ttl").Value = arr(274)
  Fields("fldEarnDsc1ttl").Value = arr(275)
  Fields("fldOTPaidttl").Value = arr(276)
  Fields("fldEICttl").Value = arr(277)
  Fields("fldRegEarnttl").Value = arr(278)
  Fields("fldOTEarnttl").Value = arr(279)
  Fields("fldEarn3ttl").Value = arr(280)
  Fields("fldEarn2ttl").Value = arr(281)
  Fields("fldEarn1ttl").Value = arr(282)
  Fields("fldGrossPayttl").Value = arr(283)
  Fields("fldSocSecttl").Value = arr(284)
  Fields("fldMedttl").Value = arr(285)
  Fields("fldDedDsc1ttl").Value = arr(286)
  Fields("fldDedDsc2ttl").Value = arr(287)
  Fields("fldDedDsc3ttl").Value = arr(288)
  Fields("fldDedDsc4ttl").Value = arr(289)
  Fields("fldDedDsc5ttl").Value = arr(290)
  Fields("fldFWTttl").Value = arr(291)
  Fields("fldSWTttl").Value = arr(292)
  Fields("fldRetttl").Value = arr(293)
  Fields("fldNetPayttl").Value = arr(294)
  Fields("fldDedVal1ttl").Value = arr(295)
  Fields("fldDedVal2ttl").Value = arr(296)
  Fields("fldDedVal3ttl").Value = arr(297)
  Fields("fldDedVal4ttl").Value = arr(298)
  Fields("fldDedVal5ttl").Value = arr(299)
  Fields("fldDedDsc6ttl").Value = arr(300)
  Fields("fldDedDsc7ttl").Value = arr(301)
  Fields("fldDedDsc8ttl").Value = arr(302)
  Fields("fldDedDsc9ttl").Value = arr(303)
  Fields("fldDedDsc10ttl").Value = arr(304)
  Fields("fldDedDsc11ttl").Value = arr(305)
  Fields("fldDedDsc12ttl").Value = arr(306)
  Fields("fldDedDsc13ttl").Value = arr(307)
  Fields("fldDedDsc14ttl").Value = arr(308)
  Fields("fldDedDsc15ttl").Value = arr(309)
  Fields("fldDedVal6ttl").Value = arr(310)
  Fields("fldDedVal7ttl").Value = arr(311)
  Fields("fldDedVal8ttl").Value = arr(312)
  Fields("fldDedVal9ttl").Value = arr(313)
  Fields("fldDedVal10ttl").Value = arr(314)
  Fields("fldDedVal11ttl").Value = arr(315)
  Fields("fldDedVal12ttl").Value = arr(316)
  Fields("fldDedVal13ttl").Value = arr(317)
  Fields("fldDedVal14ttl").Value = arr(318)
  Fields("fldDedVal15ttl").Value = arr(319)
  Fields("fldDedDsc16ttl").Value = arr(320)
  Fields("fldDedDsc17ttl").Value = arr(321)
  Fields("fldDedDsc18ttl").Value = arr(322)
  Fields("fldDedDsc19ttl").Value = arr(323)
  Fields("fldDedDsc20ttl").Value = arr(324)
  Fields("fldDedDsc21ttl").Value = arr(325)
  Fields("fldDedDsc22ttl").Value = arr(326)
  Fields("fldDedDsc23ttl").Value = arr(327)
  Fields("fldDedDsc24ttl").Value = arr(328)
  Fields("fldDedDsc25ttl").Value = arr(329)
  Fields("fldDedVal16ttl").Value = arr(330)
  Fields("fldDedVal17ttl").Value = arr(331)
  Fields("fldDedVal18ttl").Value = arr(332)
  Fields("fldDedVal19ttl").Value = arr(333)
  Fields("fldDedVal20ttl").Value = arr(334)
  Fields("fldDedVal21ttl").Value = arr(335)
  Fields("fldDedVal22ttl").Value = arr(336)
  Fields("fldDedVal23ttl").Value = arr(337)
  Fields("fldDedVal24ttl").Value = arr(338)
  Fields("fldDedVal25ttl").Value = arr(339)
  Fields("fldDedDsc26ttl").Value = arr(340)
  Fields("fldDedDsc27ttl").Value = arr(341)
  Fields("fldDedDsc28ttl").Value = arr(342)
  Fields("fldDedDsc29ttl").Value = arr(343)
  Fields("fldDedDsc30ttl").Value = arr(344)
  Fields("fldDedDsc31ttl").Value = arr(345)
  Fields("fldDedDsc32ttl").Value = arr(346)
  Fields("fldDedDsc33ttl").Value = arr(347)
  Fields("fldDedDsc34ttl").Value = arr(348)
  Fields("fldDedDsc35ttl").Value = arr(349)
  Fields("fldDedVal26ttl").Value = arr(350)
  Fields("fldDedVal27ttl").Value = arr(351)
  Fields("fldDedVal28ttl").Value = arr(352)
  Fields("fldDedVal29ttl").Value = arr(353)
  Fields("fldDedVal30ttl").Value = arr(354)
  Fields("fldDedVal31ttl").Value = arr(355)
  Fields("fldDedVal32ttl").Value = arr(356)
  Fields("fldDedVal33ttl").Value = arr(357)
  Fields("fldDedVal34ttl").Value = arr(358)
  Fields("fldDedVal35ttl").Value = arr(359)
  Fields("fldDedDsc36ttl").Value = arr(360)
  Fields("fldDedDsc37ttl").Value = arr(361)
  Fields("fldDedDsc38ttl").Value = arr(362)
  Fields("fldDedDsc39ttl").Value = arr(363)
  Fields("fldDedDsc40ttl").Value = arr(364)
  Fields("fldDedDsc41ttl").Value = arr(365)
  Fields("fldDedDsc42ttl").Value = arr(366)
  Fields("fldDedDsc43ttl").Value = arr(367)
  Fields("fldDedDsc44ttl").Value = arr(368)
  Fields("fldDedDsc45ttl").Value = arr(369)
  Fields("fldDedVal36ttl").Value = arr(370)
  Fields("fldDedVal37ttl").Value = arr(371)
  Fields("fldDedVal38ttl").Value = arr(372)
  Fields("fldDedVal39ttl").Value = arr(373)
  Fields("fldDedVal40ttl").Value = arr(374)
  Fields("fldDedVal41ttl").Value = arr(375)
  Fields("fldDedVal42ttl").Value = arr(376)
  Fields("fldDedVal43ttl").Value = arr(377)
  Fields("fldDedVal44ttl").Value = arr(378)
  Fields("fldDedVal45ttl").Value = arr(379)
  Fields("fldDedDsc46ttl").Value = arr(380)
  Fields("fldDedDsc47ttl").Value = arr(381)
  Fields("fldDedDsc48ttl").Value = arr(382)
  Fields("fldDedDsc49ttl").Value = arr(383)
  Fields("fldDedDsc50ttl").Value = arr(384)
  Fields("fldDedVal46ttl").Value = arr(385)
  Fields("fldDedVal47ttl").Value = arr(386)
  Fields("fldDedVal48ttl").Value = arr(387)
  Fields("fldDedVal49ttl").Value = arr(388)
  Fields("fldDedVal50ttl").Value = arr(389)
  Fields("fldFedGrsttl").Value = arr(390)
  Fields("fldStaGrsttl").Value = arr(391)
  Fields("fldSocGrsttl").Value = arr(392)
  Fields("fldMedGrsttl").Value = arr(393)
  Fields("fldRetGrsttl").Value = arr(394)
  If Len(Fields("fldEarnDsc1").Value) = 0 Then
    Fields("fldEarn1Det").Value = ""
  End If
  If Len(Fields("fldEarnDsc2").Value) = 0 Then
    Fields("fldEarn2Det").Value = ""
  End If
  If Len(Fields("fldEarnDsc3").Value) = 0 Then
    Fields("fldEarn3Det").Value = ""
  End If
  
End Sub
Private Sub ActiveReport_ReportEnd()
  Unload frmLoadingRpt
  If HFile <> 0 Then
    Close #HFile
  End If
End Sub

Private Sub ActiveReport_ReportStart()
  Dim DedCnt As Integer
  Dim DedRec As DedCodeRecType
  Dim DHandle As Integer
  Dim X As Integer
  Me.Zoom = -1
  
  OpenDedCodeFile DHandle
  DedCnt = LOF(DHandle) / Len(DedRec)
  Close DHandle
  Select Case DedCnt
    Case 1 To 5
      PageHeader.Height = 2300
      Detail.Height = 560
      GroupFooter1.Height = 2000
      Line4.Y1 = 2200
      Line4.Y2 = 2200
      Line8.Y1 = 2200
      Line8.Y2 = 2200
    Case 6 To 15
      PageHeader.Height = 2350
      Detail.Height = 650
      GroupFooter1.Height = 2500
      Line4.Y1 = 2450
      Line4.Y2 = 2450
      Line8.Y1 = 2450
      Line8.Y2 = 2450
    Case 16 To 25
      PageHeader.Height = 2600
      Detail.Height = 875
      GroupFooter1.Height = 2850
      Line4.Y1 = 2750
      Line4.Y2 = 2750
      Line8.Y1 = 2875
      Line8.Y2 = 2875
    Case 26 To 35
      PageHeader.Height = 2850
      Detail.Height = 1000
      GroupFooter1.Height = 3175
      Line4.Y1 = 3000
      Line4.Y2 = 3000
      Line8.Y1 = 3325
      Line8.Y2 = 3325
    Case 36 To 45
      PageHeader.Height = 3100
      Detail.Height = 1190
      GroupFooter1.Height = 3600
      Line4.Y1 = 3300
      Line4.Y2 = 3300
      Line8.Y1 = 3775
      Line8.Y2 = 3775
    Case 46 To 50
      PageHeader.Height = 3375
      Detail.Height = 1650
      GroupFooter1.Height = 4100
      Line4.Y1 = 3550
      Line4.Y2 = 3550
      Line8.Y1 = 4225
      Line8.Y2 = 4225
    Case Else
      PageHeader.Height = 3650
      Detail.Height = 1470
      GroupFooter1.Height = 3850
      Line4.Y1 = 3650
      Line4.Y2 = 3650
      Line8.Y1 = 3650
      Line8.Y2 = 3650
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
  If Len(Fields("fldEarnDsc3ftr").Value) = 0 Then
    Fields("fldEarn3ftr").Value = ""
  End If
  If Len(Fields("fldEarnDsc2ftr").Value) = 0 Then
    Fields("fldEarn2ftr").Value = ""
  End If
  If Len(Fields("fldEarnDsc1ftr").Value) = 0 Then
    Fields("fldEarn1ftr").Value = ""
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

Private Sub PageHeader_Format()
  If EndReport = True Then
    PageHeader.Height = 1350
  End If
End Sub

Private Sub ReportFooter_Format()
  EndReport = True
  Label1.Visible = False
  Label2.Visible = False
  Label3.Visible = False
  Label4.Visible = False
  Label5.Visible = False
  Label6.Visible = False
  Label7.Visible = False
  Label8.Visible = False
  Label9.Visible = False
  Label10.Visible = False
  Label11.Visible = False
  Label12.Visible = False
  Label13.Visible = False
  Label14.Visible = False
  Label15.Visible = False
  Label16.Visible = False
  Label17.Visible = False
  Label18.Visible = False
  Label19.Visible = False
  Label20.Visible = False
  lblEIC.Visible = False
  
  Fields("fldEarnDsc1").Value = ""
  Fields("fldEarnDsc2").Value = ""
  Fields("fldEarnDsc3").Value = ""
  Fields("fldDedDsc1").Value = ""
  Fields("fldDedDsc2").Value = ""
  Fields("fldDedDsc3").Value = ""
  Fields("fldDedDsc4").Value = ""
  Fields("fldDedDsc5").Value = ""
  Fields("fldDedDsc6").Value = ""
  Fields("fldDedDsc7").Value = ""
  Fields("fldDedDsc8").Value = ""
  Fields("fldDedDsc9").Value = ""
  Fields("fldDedDsc10").Value = ""
  Fields("fldDedDsc11").Value = ""
  Fields("fldDedDsc12").Value = ""
  Fields("fldDedDsc13").Value = ""
  Fields("fldDedDsc14").Value = ""
  Fields("fldDedDsc15").Value = ""
  Fields("fldDedDsc16").Value = ""
  Fields("fldDedDsc17").Value = ""
  Fields("fldDedDsc18").Value = ""
  Fields("fldDedDsc19").Value = ""
  Fields("fldDedDsc20").Value = ""
  Fields("fldDedDsc21").Value = ""
  Fields("fldDedDsc22").Value = ""
  Fields("fldDedDsc23").Value = ""
  Fields("fldDedDsc24").Value = ""
  Fields("fldDedDsc25").Value = ""
  Fields("fldDedDsc26").Value = ""
  Fields("fldDedDsc27").Value = ""
  Fields("fldDedDsc28").Value = ""
  Fields("fldDedDsc29").Value = ""
  Fields("fldDedDsc30").Value = ""
  Fields("fldDedDsc31").Value = ""
  Fields("fldDedDsc32").Value = ""
  Fields("fldDedDsc33").Value = ""
  Fields("fldDedDsc34").Value = ""
  Fields("fldDedDsc35").Value = ""
  Fields("fldDedDsc36").Value = ""
  Fields("fldDedDsc37").Value = ""
  Fields("fldDedDsc38").Value = ""
  Fields("fldDedDsc39").Value = ""
  Fields("fldDedDsc40").Value = ""
  Fields("fldDedDsc41").Value = ""
  Fields("fldDedDsc42").Value = ""
  Fields("fldDedDsc43").Value = ""
  Fields("fldDedDsc44").Value = ""
  Fields("fldDedDsc45").Value = ""
  Fields("fldDedDsc46").Value = ""
  Fields("fldDedDsc47").Value = ""
  Fields("fldDedDsc48").Value = ""
  Fields("fldDedDsc49").Value = ""
  Fields("fldDedDsc50").Value = ""
  Call PageHeader_Format
End Sub

Private Sub ReportHeader_Format()
  ReportHeader.Height = 0
End Sub
