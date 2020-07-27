VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARptLateNotice1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Late Notice "
   ClientHeight    =   6828
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   9072
   Icon            =   "ARptLateNotice1.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   16002
   _ExtentY        =   12044
   SectionData     =   "ARptLateNotice1.dsx":08CA
End
Attribute VB_Name = "ARptLateNotice1"
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
    headers(1) = "AAcctNum"
    headers(2) = "ACustName"
    headers(3) = "AMo1"
    headers(4) = "ADay1"
    headers(5) = "AYr1"
    headers(6) = "AMo2"
    headers(7) = "ADay2"
    headers(8) = "AYr2"
    headers(9) = "ADueDate"
    headers(10) = "ASvcAddr"
    headers(11) = "AAddr1"
    headers(12) = "AAddr2"
    headers(13) = "ACityStZip"
    headers(14) = "ARev1"
    headers(15) = "AAmt1"
    headers(16) = "ARev2"
    headers(17) = "AAmt2"
    headers(18) = "ARev3"
    headers(19) = "AAmt3"
    headers(20) = "ARev4"
    headers(21) = "AAmt4"
    headers(22) = "ARev5"
    headers(23) = "AAmt5"
    headers(24) = "ARev6"
    headers(25) = "AAmt6"
    headers(26) = "AOtherRev"
    headers(27) = "AOtherAmt"
    headers(28) = "ATotDue"
    headers(29) = "AMsgLine1"
    headers(30) = "AMsgLine2"
    headers(31) = "BAcctNum"
    headers(32) = "BCustName"
    headers(33) = "BMo1"
    headers(34) = "BDay1"
    headers(35) = "BYr1"
    headers(36) = "BMo2"
    headers(37) = "BDay2"
    headers(38) = "BYr2"
    headers(39) = "BDueDate"
    headers(40) = "BSvcAddr"
    headers(41) = "BAddr1"
    headers(42) = "BAddr2"
    headers(43) = "BCityStZip"
    headers(44) = "BRev1"
    headers(45) = "BAmt1"
    headers(46) = "BRev2"
    headers(47) = "BAmt2"
    headers(48) = "BRev3"
    headers(49) = "BAmt3"
    headers(50) = "BRev4"
    headers(51) = "BAmt4"
    headers(52) = "BRev5"
    headers(53) = "BAmt5"
    headers(54) = "BRev6"
    headers(55) = "BAmt6"
    headers(56) = "BOtherRev"
    headers(57) = "BOtherAmt"
    headers(58) = "BTotDue"
    headers(59) = "BMsgLine1"
    headers(60) = "BMsgLine2"
    headers(61) = "CAcctNum"
    headers(62) = "CCustName"
    headers(63) = "CMo1"
    headers(64) = "CDay1"
    headers(65) = "CYr1"
    headers(66) = "CMo2"
    headers(67) = "CDay2"
    headers(68) = "CYr2"
    headers(69) = "CDueDate"
    headers(70) = "CSvcAddr"
    headers(71) = "CAddr1"
    headers(72) = "CAddr2"
    headers(73) = "CCityStZip"
    headers(74) = "CRev1"
    headers(75) = "CAmt1"
    headers(76) = "CRev2"
    headers(77) = "CAmt2"
    headers(78) = "CRev3"
    headers(79) = "CAmt3"
    headers(80) = "CRev4"
    headers(81) = "CAmt4"
    headers(82) = "CRev5"
    headers(83) = "CAmt5"
    headers(84) = "CRev6"
    headers(85) = "CAmt6"
    headers(86) = "COtherRev"
    headers(87) = "COtherAmt"
    headers(88) = "CTotDue"
    headers(89) = "CMsgLine1"
    headers(90) = "CMsgLine2"

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
      MsgBox "File - UBLazLN.xls, created in the Citipak Directory.", vbOKOnly
    End If
    If KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - UBLazLN.txt, created in the Citipak Directory.", vbOKOnly
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
    MsgBox "File - UBLazLN.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "/&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - UBLazLN.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "UBLazLN.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "UBLazLN.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

'Private Sub Detail_Format()
'  dcnt = dcnt + 1
'  If dcnt = 3 Then
'    GroupFooter1.Visible = False
'    dcnt = 0
'  ElseIf dcnt = 1 Then
'    GroupFooter1.Visible = True
'  End If
'End Sub
