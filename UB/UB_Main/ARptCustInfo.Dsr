VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARptCustInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Information"
   ClientHeight    =   7788
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   9384
   Icon            =   "ARptCustInfo.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   16552
   _ExtentY        =   13737
   SectionData     =   "ARptCustInfo.dsx":08CA
End
Attribute VB_Name = "ARptCustInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ReportFile As String
Private hFile As Integer
Dim rpt As ActiveReport
Dim cnt As Integer
Dim headers(1 To 217) As String

Public Sub GetName(RName As String) ', SName As String, Detail As Integer, RevSource As Integer)
  ReportFile$ = RName$
End Sub

Private Sub ActiveReport_DataInitialize()
'On Local Error GoTo stopthiscrazything

    headers(1) = "AcctNo"
    headers(2) = "Book"
    headers(3) = "SeqNo"
    headers(4) = "Status"
    headers(5) = "OpenDate"
    headers(6) = "Search"
    headers(7) = "CustName"
    headers(8) = "Addr1"
    headers(9) = "Addr2"
    headers(10) = "SvcAddr"
    headers(11) = "City"
    headers(12) = "State"
    headers(13) = "Zip"
    headers(14) = "DPCode"
    headers(15) = "HPhone"
    headers(16) = "WPhone"
    headers(17) = "SocSec"
    headers(18) = "DrvLic"
    headers(19) = "CustType"
    headers(20) = "Addr911"
    headers(21) = "BillTo"
    headers(22) = "BillCopies"
    headers(23) = "PostRte"
    headers(24) = "BillCyc"
    headers(25) = "Zone"
    headers(26) = "Seq"
    headers(27) = "CashOnly"
    headers(28) = "LateFee"
    headers(29) = "Cutoff"
    headers(30) = "TaxEx"
    headers(31) = "SrCit"
    headers(32) = "UseDraft"
    headers(33) = "AccType"
    headers(34) = "BankName"
    headers(35) = "BankLoc"
    headers(36) = "Transit"
    headers(37) = "BankAcct"
    headers(38) = "AcctBal"
    headers(39) = "Current"
    headers(40) = "PastDue"
    headers(41) = "Deposit"
    headers(42) = "BillCmt"
    headers(43) = "PayCmt"
    headers(44) = "PumpCode"
    headers(45) = "UserCode1"
    headers(46) = "UserCode2"
    headers(47) = "ProRatePct"
    headers(48) = "HHMsg1"
    headers(49) = "HHMsg2"
    headers(50) = "HHMsg3"
    headers(51) = "RevName1"
    headers(52) = "RRate1"
    headers(53) = "RMtr1"
    headers(54) = "RevName2"
    headers(55) = "RRate2"
    headers(56) = "RMtr2"
    headers(57) = "RevName3"
    headers(58) = "RRate3"
    headers(59) = "RMtr3"
    headers(60) = "RevName4"
    headers(61) = "RRate4"
    headers(62) = "RMtr4"
    headers(63) = "RevName5"
    headers(64) = "RRate5"
    headers(65) = "RMtr5"
    headers(66) = "RevName6"
    headers(67) = "RRate6"
    headers(68) = "RMtr6"
    headers(69) = "RevName7"
    headers(70) = "RRate7"
    headers(71) = "RMtr7"
    headers(72) = "RevName8"
    headers(73) = "RRate8"
    headers(74) = "RMtr8"
    headers(75) = "RevName9"
    headers(76) = "RRate9"
    headers(77) = "RMtr9"
    headers(78) = "RevName10"
    headers(79) = "RRate10"
    headers(80) = "RMtr10"
    headers(81) = "RevName11"
    headers(82) = "RRate11"
    headers(83) = "RMtr11"
    headers(84) = "RevName12"
    headers(85) = "RRate12"
    headers(86) = "RMtr12"
    headers(87) = "RevName13"
    headers(88) = "RRate13"
    headers(89) = "RMtr13"
    headers(90) = "RevName14"
    headers(91) = "RRate14"
    headers(92) = "RMtr14"
    headers(93) = "RevName15"
    headers(94) = "RRate15"
    headers(95) = "RMtr15"
    headers(96) = "FRDesc1"
    headers(97) = "FRAmt1"
    headers(98) = "FRFreq1"
    headers(99) = "FRRev1"
    headers(100) = "FRMin1"
    headers(101) = "FRDesc2"
    headers(102) = "FRAmt2"
    headers(103) = "FRFreq2"
    headers(104) = "FRRev2"
    headers(105) = "FRMin2"
    headers(106) = "FRDesc3"
    headers(107) = "FRAmt3"
    headers(108) = "FRFreq3"
    headers(109) = "FRRev3"
    headers(110) = "FRMin3"
    headers(111) = "FRDesc4"
    headers(112) = "FRAmt4"
    headers(113) = "FRFreq4"
    headers(114) = "FRRev4"
    headers(115) = "FRMin4"
    headers(116) = "MAmtO1"
    headers(117) = "MAmtP1"
    headers(118) = "MPay1"
    headers(119) = "MRev1"
    headers(120) = "MAmtO2"
    headers(121) = "MAmtP2"
    headers(122) = "MPay2"
    headers(123) = "MRev2"
    headers(124) = "MFee1"
    headers(125) = "MFee2"
    headers(126) = "MtrNum1"
    headers(127) = "MtrMult1"
    headers(128) = "MtrType1"
    headers(129) = "MtrUnit1"
    headers(130) = "NumUser1"
    headers(131) = "InsDate1"
    headers(132) = "CurrRead1"
    headers(133) = "PrevRead1"
    headers(134) = "CurrDate1"
    headers(135) = "PrevDate1"
    headers(136) = "MtrIdNo1"
    headers(137) = "MtrLat1"
    headers(138) = "MtrLon1"
    headers(139) = "MtrNum2"
    headers(140) = "MtrMult2"
    headers(141) = "MtrType2"
    headers(142) = "MtrUnit2"
    headers(143) = "NumUser2"
    headers(144) = "InsDate2"
    headers(145) = "CurrRead2"
    headers(146) = "PrevRead2"
    headers(147) = "CurrDate2"
    headers(148) = "PrevDate2"
    headers(149) = "MtrIdNo2"
    headers(150) = "MtrLat2"
    headers(151) = "MtrLon2"
    headers(152) = "MtrNum3"
    headers(153) = "MtrMult3"
    headers(154) = "MtrType3"
    headers(155) = "MtrUnit3"
    headers(156) = "NumUser3"
    headers(157) = "InsDate3"
    headers(158) = "CurrRead3"
    headers(159) = "PrevRead3"
    headers(160) = "CurrDate3"
    headers(161) = "PrevDate3"
    headers(162) = "MtrIdNo3"
    headers(163) = "MtrLat3"
    headers(164) = "MtrLon3"
    headers(165) = "MtrNum4"
    headers(166) = "MtrMult4"
    headers(167) = "MtrType4"
    headers(168) = "MtrUnit4"
    headers(169) = "NumUser4"
    headers(170) = "InsDate4"
    headers(171) = "CurrRead4"
    headers(172) = "PrevRead4"
    headers(173) = "CurrDate4"
    headers(174) = "PrevDate4"
    headers(175) = "MtrIdNo4"
    headers(176) = "MtrLat4"
    headers(177) = "MtrLon4"
    headers(178) = "MtrNum5"
    headers(179) = "MtrMult5"
    headers(180) = "MtrType5"
    headers(181) = "MtrUnit5"
    headers(182) = "NumUser5"
    headers(183) = "InsDate5"
    headers(184) = "CurrRead5"
    headers(185) = "PrevRead5"
    headers(186) = "CurrDate5"
    headers(187) = "PrevDate5"
    headers(188) = "MtrIdNo5"
    headers(189) = "MtrLat5"
    headers(190) = "MtrLon5"
    headers(191) = "MtrNum6"
    headers(192) = "MtrMult6"
    headers(193) = "MtrType6"
    headers(194) = "MtrUnit6"
    headers(195) = "NumUser6"
    headers(196) = "InsDate6"
    headers(197) = "CurrRead6"
    headers(198) = "PrevRead6"
    headers(199) = "CurrDate6"
    headers(200) = "PrevDate6"
    headers(201) = "MtrIdNo6"
    headers(202) = "MtrLat6"
    headers(203) = "MtrLon6"
    headers(204) = "MtrNum7"
    headers(205) = "MtrMult7"
    headers(206) = "MtrType7"
    headers(207) = "MtrUnit7"
    headers(208) = "NumUser7"
    headers(209) = "InsDate7"
    headers(210) = "CurrRead7"
    headers(211) = "PrevRead7"
    headers(212) = "CurrDate7"
    headers(213) = "PrevDate7"
    headers(214) = "MtrIdNo7"
    headers(215) = "MtrLat7"
    headers(216) = "MtrLon7"
    headers(217) = "GrpCode"
    

    hFile = FreeFile
    Open ReportFile$ For Input As #hFile
'
'    ' This sets up the fields used in data binding
'    'Fields.Add "ProductID"
    For cnt = 1 To 217
      Fields.Add headers(cnt)
    Next
'stopthiscrazything:
'Stop
End Sub
Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
    Unload frmLoadingRpt
    MsgBox "Error Number: " & Str(Number) & " " & Description, vbOKOnly, "Printer Error"
    Unload Me
  End If
  CancelDisplay = True
End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)

Dim sLine As String
Dim arr() As String
On Error GoTo ERRORSTUFF
'    ' We reached the end of the file we exit leaving the
'    ' eof parameter as True (default except on first call) that will
'    ' tell AR that we are done feeding data
'    ' otherwise we have to set the eof parameter to False so that
'    ' AR continues fetching data, until we're done
'    ' if the report had a data control, the value of the parameter
'    ' will be ignored, AR will always follow the data control's recordset
'    ' EOF property
    If VBA.eof(hFile) Then
        eof = True
        Exit Sub
    Else
        eof = False
    End If
    frmLoadingRpt.ShowHowMuch
    Line Input #hFile, sLine
    arr = Split(sLine, "~")
'    ' Here we set the values of the fields that we defines as unbound
'    ' or user defined.
    For cnt = 1 To 217
      Fields(headers(cnt)) = arr(cnt - 1)
    Next
'If something wrong in file give message instead of crashing
Exit Sub
ERRORSTUFF:
      Unload frmLoadingRpt
'  Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "ARptVendHist", "Fetch Data", Erl)
'    Case emrExitProc:
'      Resume Proc_Exit
'    Case emrResume:
'      Resume
'    Case emrResumeNext:
'      Resume Next
'    Case Else
'      '--- Technically, this should never happen.
'      Resume Proc_Exit
'  End Select
   MsgBox Err.Number, Err.Description, Err.Source, vbOKOnly, "Error"
   GoSub Proc_Exit
Proc_Exit:
  '--- Cleanup code goes here...
    Close
   Unload Me
End Sub

Public Sub startrpt()
  Me.Run True
End Sub

Private Sub ActiveReport_Initialize()
  Me.Toolbar.Tools.Add "&Close"
  Me.Toolbar.Tools.Add "Save/&Excel"
  Me.Toolbar.Tools.Add "/&Text"

End Sub

'Private Sub ActiveReport_ReportStart()
'  If Det > 0 Then
'    If Det < 5 Then
'      Me.GroupFooter1.Height = 0
'    ElseIf Det > 4 And Det < 9 Then
'      Me.GroupFooter1.Height = 270
'    ElseIf Det > 8 And Det < 13 Then
'      Me.GroupFooter1.Height = 540
'    Else
'      Me.GroupFooter1.Height = 810
'    End If
'  Else
'    Me.GroupFooter1.Visible = False
'  End If
'  If Rev > 0 Then
'    Me.Label14.Visible = False
'    Me.Label15.Visible = False
'    txtcur.Visible = False
'    txtHead.Visible = True
'    Me.Label18.Visible = True
'    Me.Debit.Left = 7110
'    Me.txtTotCur.Left = 7110
'    Me.txtTotAcctBal.Visible = False
'    Me.txtTotPast.Visible = False
'  End If
'  Me.Label19 = Me.txtRptParm1
'  Me.Label21 = Me.txtRptParm2
'End Sub
'
'Private Sub PageHeader_Format()
'If Me.pageNumber = 1 Then
'  Label5.Visible = True
'  Shape1.Visible = True
'  txtRptParm1.Visible = True
'  txtRptParm2.Visible = True
'  Me.PageHeader.Height = 1620
'Else
'  Label5.Visible = False
'  Shape1.Visible = False
'  txtRptParm1.Visible = False
'  txtRptParm2.Visible = False
'  Me.PageHeader.Height = 1156
'  Label6.Top = 630
'  Label13.Top = 630
'  labloc.Top = 630
'  txtHead.Top = 630
'  txtcur.Top = 630
'  Label14.Top = 630
'  Label15.Top = 630
'  Label18.Top = 630
'End If
'End Sub

'Private Sub GroupFooter1_Format()
'  If Det > 0 Then
'    If Det < 5 Then
'      Me.GroupFooter1.Height = 270
'    ElseIf Det > 4 And Det < 9 Then
'      Me.GroupFooter1.Height = 540
'    ElseIf Det > 8 And Det < 13 Then
'      Me.GroupFooter1.Height = 810
'    Else
'      Me.GroupFooter1.Height = 1080
'    End If
'  Else
'    Me.GroupFooter1.Visible = False
'  End If
'End Sub

'Private Sub ReportFooter_Format()
'  If Det > 0 Then
'    Set Me.SubReport1.object = New ARSubTot
'  End If
'End Sub

Private Sub ActiveReport_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
  If KeyCode = vbKeyEscape Then
    Unload Me
    KeyCode = 0
  End If
  If Shift = 4 Then
    If KeyCode = vbKeyC Then
      Unload Me
    End If
    If KeyCode = vbKeyE Then
      Screen.MousePointer = vbHourglass
      ExportReport 1
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - UBInf.xls, created in the Citipak Directory.", vbOKOnly
    End If
    If KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - UBInf.txt, created in the Citipak Directory.", vbOKOnly
    End If
  End If
End Sub

'Private Sub Detail_AfterPrint()
'  Me.GroupHeader1.Visible = False
'End Sub
'
'
'Private Sub PageHeader_AfterPrint()
'  GHP = False
'End Sub

'Private Sub GroupHeader1_BeforePrint()
'  if
'End Sub

'Private Sub PageHeader_Format()
'  If Me.Fields("LineTyp").Value = "D" Then
'    Me.PageHeader.Height = 1752
'    Me.GroupHeader1.Visible = False
'  End If
'  If Fields("LineTyp").Value = "TE" Then
'    Me.PageHeader.Height = 1068
'  End If
'  If Fields("LineTyp").Value = "T" Then
'    Me.PageHeader.Height = 1380
'  End If
'  If Me.Fields("LineTyp").Value = "H" Then
'    If Not Me.Fields("Date").Value = "Account" Then
'       Me.PageHeader.Height = 1068
'    Else
'      Me.PageHeader.Height = 1380
'    End If
'  End If
'End Sub
'
'Private Sub Detail_Format()
'  If Me.Fields("Code").Value <> "" Then
'  If Me.Fields("Code").Value = 4 Then
'    totOpnPO = totOpnPO + Me.Fields("Credit").Value
'  End If
'  If Me.Fields("Code").Value = -4 Then
'    totClsPO = totClsPO + Me.Fields("Credit").Value
'  End If
'  If Me.Fields("Code").Value = 1 Then
'    totInv = totInv + Me.Fields("Credit").Value
'      If Left$(Me.Fields("Status").Value, 4) = "  Pd" Then
'        totPInv = totPInv + Me.Fields("Credit").Value
'      End If
'  End If
'  If Me.Fields("Code").Value = -1 Then
'    'totInv = totInv + Me.Fields("Credit").Value
'    totVInv = totVInv + Me.Fields("Credit").Value
'  End If
'  If Me.Fields("Code").Value = 3 Then
'    totCks = totCks + Me.Fields("Debit").Value
'  End If
'  If Me.Fields("Code").Value = -3 Then
'    totVCks = totVCks + Me.Fields("Debit").Value
'  End If
'  End If
'End Sub

'Private Sub GroupFooter1_AfterPrint()
'  totInv = 0
'  totOpnPO = 0
'  totPInv = 0
'  totVInv = 0
'  totClsPO = 0
'  totCks = 0
'  totVCks = 0
'End Sub

'Private Sub GroupFooter1_BeforePrint()
'Dim tmpPO As Double, Bal As Double
'Bal = 0
'tmpPO = 0
'  'Me.txtTotInv = totInv
'  Me.txtTotOpnInv.DataValue = Round(totInv - totPInv)
'  'Me.txtTotPO = totOpnPO
'  If totOpnPO > 0 Then
'    tmpPO = Round(totOpnPO)
'    If tmpPO > 0 Then
'      Me.txtTotOpnPO.DataValue = tmpPO
'    Else
'      Me.txtTotOpnPO.DataValue = 0
'    End If
'  Else
'    Me.txtTotOpnPO.DataValue = 0
'  End If
'  Me.txtTotCks.DataValue = totCks
'  Me.txtTotInv.DataValue = totInv
'  Bal = Round(totInv - totCks)
'  Me.txtTotBAlance.DataValue = Bal
'
'End Sub

'Private Sub GroupHeader1_AfterPrint()
'  Me.GroupHeader1.Visible = False
'  Me.GroupHeader1.GrpKeepTogether = ddGrpNone
'  GHP = False
'End Sub
'
'Private Sub GroupHeader1_Format()
'  If Det = 0 Then
'    Me.GroupHeader1.Visible = False
'  End If
'End Sub
'  If Fields("Linetyp").Value = "H" Then
'    Me.GroupHeader1.Visible = False
'    Me.GroupHeader1.GrpKeepTogether = ddGrpNone
'    'Me.Line1.Visible = False
'  End If
'  If Fields("LineTyp").Value = "D" Then
'    Me.GroupHeader1.Visible = True
'    Me.GroupHeader1.GrpKeepTogether = ddGrpFirstDetail
'    GHP = True
'  End If
''  If Fields("linetyp").Value = "T" Then
''    Me.GroupHeader1.Visible = False
''  End If
'  If Fields("linetyp").Value = "TE" Then
'     Me.GroupHeader1.Visible = False
'     Me.GroupFooter1.Visible = True
'     Me.GroupHeader1.GrpKeepTogether = ddGrpNone
'     'Me.Line1.Visible = True
'  End If
'End Sub

Private Sub ActiveReport_QueryClose(Cancel As Integer, CloseMode As Integer)
 ' KillFile ReportFile$
 ' KillFile SubFile$
End Sub
Private Sub ActiveReport_ReportEnd()
    If hFile <> 0 Then
        Close #hFile
    End If
  Unload frmLoadingRpt
  Me.Show 1
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
    MsgBox "File - UBInf.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "/&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - UBInf.txt, created in the Citipak Directory.", vbOKOnly
  End If
End Sub
Private Sub ExportReport(x As Integer)
  Dim oEXL As ActiveReportsExcelExport.ARExportExcel
  Dim oTXT As ActiveReportsTextExport.ARExportText
  Dim outfile As String
  If Right$(UBPath, 1) = ":" Then
    outfile = UBPath
  Else
    outfile = UBPath & "\"
  End If
  Select Case x
    Case 1   '"Excel"
        Set oEXL = New ActiveReportsExcelExport.ARExportExcel
        oEXL.FileName = outfile & "UBInf.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "UBInf.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub
