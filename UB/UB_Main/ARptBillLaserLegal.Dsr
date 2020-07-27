VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARptBillLaserLegal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laser Bill "
   ClientHeight    =   7905
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   11025
   Icon            =   "ARptBillLaserLegal.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   19447
   _ExtentY        =   13944
   SectionData     =   "ARptBillLaserLegal.dsx":08CA
End
Attribute VB_Name = "ARptBillLaserLegal"
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
Dim headers(1 To 64) As String

Public Sub GetName(RName As String)
  ReportFile$ = RName$
End Sub
Private Sub ActiveReport_DataInitialize()
    headers(1) = "BLNum"
    headers(2) = "PDate"
    headers(3) = "RDate"
    headers(4) = "Days"
    headers(5) = "RevN1"
    headers(6) = "PrevR1"
    headers(7) = "CurrR1"
    headers(8) = "Use1"
    headers(9) = "RevAmt1"
    headers(10) = "RevN2"
    headers(11) = "PrevR2"
    headers(12) = "CurrR2"
    headers(13) = "Use2"
    headers(14) = "RevAmt2"
    headers(15) = "RevN3"
    headers(16) = "PrevR3"
    headers(17) = "CurrR3"
    headers(18) = "Use3"
    headers(19) = "RevAmt3"
    headers(20) = "RevN4"
    headers(21) = "PrevR4"
    headers(22) = "CurrR4"
    headers(23) = "Use4"
    headers(24) = "RevAmt4"
    headers(25) = "RevN5"
    headers(26) = "PrevR5"
    headers(27) = "CurrR5"
    headers(28) = "Use5"
    headers(29) = "RevAmt5"
    headers(30) = "RevN6"
    headers(31) = "PrevR6"
    headers(32) = "CurrR6"
    headers(33) = "Use6"
    headers(34) = "RevAmt6"
    headers(35) = "RevN7"
    headers(36) = "PrevR7"
    headers(37) = "CurrR7"
    headers(38) = "Use7"
    headers(39) = "RevAmt7"
    headers(40) = "RevN8"
    headers(41) = "PrevR8"
    headers(42) = "CurrR8"
    headers(43) = "Use8"
    headers(44) = "RevAmt8"
    headers(45) = "Current"
    headers(46) = "Deposit"
    headers(47) = "DepAmt"
    headers(48) = "BillDate"
    headers(49) = "PastDate"
    headers(50) = "Msg1"
    headers(51) = "Msg2"
    headers(52) = "Msg3"
    headers(53) = "Msg4"
    headers(54) = "Total"
    headers(55) = "TotAmt"
    headers(56) = "CustNo"
    headers(57) = "SvcAddr"
    headers(58) = "CustName"
    headers(59) = "Addr1"
    headers(60) = "Addr2"
    headers(61) = "CityStZip"
    headers(62) = "PenAmt"
    headers(63) = "zip"
    headers(64) = "Location"
'    headers(64) = "BRev2"
'    headers(65) = "BAmt2"
'    headers(66) = "BRev3"
'    headers(67) = "BAmt3"
'    headers(68) = "BRev4"
'    headers(69) = "BAmt4"
'    headers(70) = "BRev5"
'    headers(71) = "BAmt5"
'    headers(72) = "BRev6"
'    headers(73) = "BAmt6"
'    headers(74) = "BPrev"
'    headers(75) = "BPrevAmt"
'    headers(76) = "BDraft"
'    headers(77) = "BTotal"
'    headers(78) = "BMessage"
'    headers(79) = "CBLNum"
'    headers(80) = "CMo1"
'    headers(81) = "CDay1"
'    headers(82) = "CYr1"
'    headers(83) = "CMo2"
'    headers(84) = "CDay2"
'    headers(85) = "CYr2"
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
    For cnt = 1 To 64
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
    For cnt = 1 To 64
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

