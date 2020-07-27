VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARptBillLaser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laser Utility Bill"
   ClientHeight    =   12510
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   13350
   Icon            =   "ARptBillLaser.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   23548
   _ExtentY        =   22066
   SectionData     =   "ARptBillLaser.dsx":08CA
End
Attribute VB_Name = "ARptBillLaser"
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
Dim headers(1 To 117) As String

Public Sub GetName(RName As String)
  ReportFile$ = RName$
End Sub
Private Sub ActiveReport_DataInitialize()
    headers(1) = "ABLNum"
    headers(2) = "AMo1"
    headers(3) = "ADay1"
    headers(4) = "AYr1"
    headers(5) = "AMo2"
    headers(6) = "ADay2"
    headers(7) = "AYr2"
    headers(8) = "ADue"
    headers(9) = "AGPrev"
    headers(10) = "AGCurr"
    headers(11) = "AGUse"
    headers(12) = "AGCode"
    headers(13) = "AWPrev"
    headers(14) = "AWCurr"
    headers(15) = "AWUse"
    headers(16) = "AWCode"
    headers(17) = "ASvcAddr"
    headers(18) = "ACustName"
    headers(19) = "AAddr1"
    headers(20) = "AAddr2"
    headers(21) = "ACityStZip"
    headers(22) = "AAcctNo"
    headers(23) = "ARev1"
    headers(24) = "AAmt1"
    headers(25) = "ARev2"
    headers(26) = "AAmt2"
    headers(27) = "ARev3"
    headers(28) = "AAmt3"
    headers(29) = "ARev4"
    headers(30) = "AAmt4"
    headers(31) = "ARev5"
    headers(32) = "AAmt5"
    headers(33) = "ARev6"
    headers(34) = "AAmt6"
    headers(35) = "APrev"
    headers(36) = "APrevAmt"
    headers(37) = "ADraft"
    headers(38) = "ATotal"
    headers(39) = "AMessage"
    headers(40) = "BBLNum"
    headers(41) = "BMo1"
    headers(42) = "BDay1"
    headers(43) = "BYr1"
    headers(44) = "BMo2"
    headers(45) = "BDay2"
    headers(46) = "BYr2"
    headers(47) = "BDue"
    headers(48) = "BGPrev"
    headers(49) = "BGCurr"
    headers(50) = "BGUse"
    headers(51) = "BGCode"
    headers(52) = "BWPrev"
    headers(53) = "BWCurr"
    headers(54) = "BWUse"
    headers(55) = "BWCode"
    headers(56) = "BSvcAddr"
    headers(57) = "BCustName"
    headers(58) = "BAddr1"
    headers(59) = "BAddr2"
    headers(60) = "BCityStZip"
    headers(61) = "BAcctNo"
    headers(62) = "BRev1"
    headers(63) = "BAmt1"
    headers(64) = "BRev2"
    headers(65) = "BAmt2"
    headers(66) = "BRev3"
    headers(67) = "BAmt3"
    headers(68) = "BRev4"
    headers(69) = "BAmt4"
    headers(70) = "BRev5"
    headers(71) = "BAmt5"
    headers(72) = "BRev6"
    headers(73) = "BAmt6"
    headers(74) = "BPrev"
    headers(75) = "BPrevAmt"
    headers(76) = "BDraft"
    headers(77) = "BTotal"
    headers(78) = "BMessage"
    headers(79) = "CBLNum"
    headers(80) = "CMo1"
    headers(81) = "CDay1"
    headers(82) = "CYr1"
    headers(83) = "CMo2"
    headers(84) = "CDay2"
    headers(85) = "CYr2"
    headers(86) = "CDue"
    headers(87) = "CGPrev"
    headers(88) = "CGCurr"
    headers(89) = "CGUse"
    headers(90) = "CGCode"
    headers(91) = "CWPrev"
    headers(92) = "CWCurr"
    headers(93) = "CWUse"
    headers(94) = "CWCode"
    headers(95) = "CSvcAddr"
    headers(96) = "CCustName"
    headers(97) = "CAddr1"
    headers(98) = "CAddr2"
    headers(99) = "CCityStZip"
    headers(100) = "CAcctNo"
    headers(101) = "CRev1"
    headers(102) = "CAmt1"
    headers(103) = "CRev2"
    headers(104) = "CAmt2"
    headers(105) = "CRev3"
    headers(106) = "CAmt3"
    headers(107) = "CRev4"
    headers(108) = "CAmt4"
    headers(109) = "CRev5"
    headers(110) = "CAmt5"
    headers(111) = "CRev6"
    headers(112) = "CAmt6"
    headers(113) = "CPrev"
    headers(114) = "CPrevAmt"
    headers(115) = "CDraft"
    headers(116) = "CTotal"
    headers(117) = "CMessage"


    hFile = FreeFile
    Open ReportFile$ For Input As #hFile
'
'    ' This sets up the fields used in data binding
'    'Fields.Add "ProductID"
    For cnt = 1 To 117
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
    For cnt = 1 To 117
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
