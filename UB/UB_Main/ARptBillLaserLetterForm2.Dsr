VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARptBillLaserLetterForm2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laser Bill "
   ClientHeight    =   6744
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   9192
   Icon            =   "ARptBillLaserLetterForm2.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   16214
   _ExtentY        =   11896
   SectionData     =   "ARptBillLaserLetterForm2.dsx":08CA
End
Attribute VB_Name = "ARptBillLaserLetterForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ReportFile As String
Private hFile As Integer
Dim rpt As ActiveReport
Dim tempTot As Integer
Dim cnt As Integer, dcnt As Integer
Dim headers(1 To 90) As String

Public Sub GetName(RName As String)
  ReportFile$ = RName$
End Sub
Private Sub ActiveReport_DataInitialize()
    headers(1) = "BLNum"
    headers(2) = "PDate"
    headers(3) = "RDate"
    headers(4) = "Days"
    headers(5) = "MtrNum1"
    headers(6) = "PrevR1"
    headers(7) = "CurrR1"
    headers(8) = "Use1"
    headers(9) = "MtrNum2"
    headers(10) = "PrevR2"
    headers(11) = "CurrR2"
    headers(12) = "Use2"
    headers(13) = "MtrNum3"
    headers(14) = "PrevR3"
    headers(15) = "CurrR3"
    headers(16) = "Use3"
    headers(17) = "MtrNum4"
    headers(18) = "PrevR4"
    headers(19) = "CurrR4"
    headers(20) = "Use4"
    headers(21) = "MtrNum5"
    headers(22) = "PrevR5"
    headers(23) = "CurrR5"
    headers(24) = "Use5"
    headers(25) = "MtrNum6"
    headers(26) = "PrevR6"
    headers(27) = "CurrR6"
    headers(28) = "Use6"
    headers(29) = "MtrNum7"
    headers(30) = "PrevR7"
    headers(31) = "CurrR7"
    headers(32) = "Use7"
    headers(33) = "RevN1"
    headers(34) = "RevAmt1"
    headers(35) = "RevN2"
    headers(36) = "RevAmt2"
    headers(37) = "RevN3"
    headers(38) = "RevAmt3"
    headers(39) = "RevN4"
    headers(40) = "RevAmt4"
    headers(41) = "RevN5"
    headers(42) = "RevAmt5"
    headers(43) = "RevN6"
    headers(44) = "RevAmt6"
    headers(45) = "RevN7"
    headers(46) = "RevAmt7"
    headers(47) = "RevN8"
    headers(48) = "RevAmt8"
    headers(49) = "RevN9"
    headers(50) = "RevAmt9"
    headers(51) = "RevN10"
    headers(52) = "RevAmt10"
    headers(53) = "RevN11"
    headers(54) = "RevAmt11"
    headers(55) = "RevN12"
    headers(56) = "RevAmt12"
    headers(57) = "RevN13"
    headers(58) = "RevAmt13"
    headers(59) = "RevN14"
    headers(60) = "RevAmt14"
    headers(61) = "RevN15"
    headers(62) = "RevAmt15"
    headers(63) = "TaxDesc"
    headers(64) = "TaxTot"
    headers(65) = "PrevDesc"
    headers(66) = "PrevTot"
    headers(67) = "CurrDesc"
    headers(68) = "CurTot"
    headers(69) = "DepDesc"
    headers(70) = "DepTot"
    headers(71) = "TotDesc"
    headers(72) = "TotAmt"
    headers(73) = "BillDate"
    headers(74) = "PastDate"
    headers(75) = "Msg1"
    headers(76) = "Msg2"
    headers(77) = "Msg3"
    headers(78) = "Msg4"
    headers(79) = "BillMsg"
    headers(80) = "Draft"
    headers(81) = "Acct"
    headers(82) = "SvcAddr"
    headers(83) = "CustName"
    headers(84) = "Addr1"
    headers(85) = "Addr2"
    headers(86) = "City"
    headers(87) = "PenAmt"
    headers(88) = "Zip"
    headers(89) = "Location"
    headers(90) = "BarZip"
'    headers(86) = "CDue"
'    headers(87) = "CGPrev"
'    headers(88) = "CGCurr"
'    headers(89) = "CGUse"
'    headers(90) = "CGCode"
'    headers(91) = "CWPrev"
'    headers(92) = "CWCurr"
'    headers(93) = "CWUse"
'    headers(94) = "CWCode"
'    headers(95) = "CSvcAddr"
'    headers(96) = "CCustName"
'    headers(97) = "CAddr1"
'    headers(98) = "CAddr2"
'    headers(99) = "CCityStZip"
'    headers(100) = "CAcctNo"
'    headers(101) = "CRev1"
'    headers(102) = "CAmt1"
'    headers(103) = "CRev2"
'    headers(104) = "CAmt2"
'    headers(105) = "CRev3"
'    headers(106) = "CAmt3"
'    headers(107) = "CRev4"
'    headers(108) = "CAmt4"
'    headers(109) = "CRev5"
'    headers(110) = "CAmt5"
'    headers(111) = "CRev6"
'    headers(112) = "CAmt6"
'    headers(113) = "CPrev"
'    headers(114) = "CPrevAmt"
'    headers(115) = "CDraft"
'    headers(116) = "CTotal"
'    headers(117) = "CMessage"


    hFile = FreeFile
    Open ReportFile$ For Input As #hFile
'
'    ' This sets up the fields used in data binding
'    'Fields.Add "ProductID"
    For cnt = 1 To 90
      Fields.Add headers(cnt)
    Next
 
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
'On Error GoTo ERRORSTUFF
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
    For cnt = 1 To 90
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
   MsgBox "Err.Number, Err.Description, Err.Source", vbOKOnly, "Error"
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
  dcnt = 0
End Sub

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
      MsgBox "File - UBLazBil.xls, created in the Citipak Directory.", vbOKOnly
    End If
    If KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - UBLazBil.txt, created in the Citipak Directory.", vbOKOnly
    End If
  End If
End Sub

Private Sub ActiveReport_QueryClose(Cancel As Integer, CloseMode As Integer)
 ' KillFile ReportFile$
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
    MsgBox "File - UBLazBil.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "/&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - UBLazBil.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "UBLazBil.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "UBLazBil.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

'  dcnt = dcnt + 1
'  If dcnt = 3 Then
'    GroupFooter1.Visible = False
'    dcnt = 0
'  ElseIf dcnt = 1 Then
'    GroupFooter1.Visible = True
'  End If
'End Sub

