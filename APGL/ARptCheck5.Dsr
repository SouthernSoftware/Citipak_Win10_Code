VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARptCheck5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "A/P Check 5"
   ClientHeight    =   7200
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   8640
   Icon            =   "ARptCheck5.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   15240
   _ExtentY        =   12700
   SectionData     =   "ARptCheck5.dsx":08CA
End
Attribute VB_Name = "ARptCheck5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ReportFile As String
Private hFile As Integer
Dim rpt As ActiveReport
Dim headers(1 To 83) As String
Dim cnt As Integer

Public Sub GetName(RName As String)
  ReportFile$ = RName$
End Sub

Private Sub ActiveReport_DataInitialize()
   hFile = FreeFile
   Open ReportFile$ For Input As #hFile

    headers(1) = "TS1Date"
    headers(2) = "TS1Inv"
    headers(3) = "TS1PO"
    headers(4) = "TS1Amt"
    headers(5) = "TS2Date"
    headers(6) = "TS2Inv"
    headers(7) = "TS2PO"
    headers(8) = "TS2Amt"
    headers(9) = "TS3Date"
    headers(10) = "TS3Inv"
    headers(11) = "TS3PO"
    headers(12) = "TS3Amt"
    headers(13) = "TS4Date"
    headers(14) = "TS4Inv"
    headers(15) = "TS4PO"
    headers(16) = "TS4Amt"
    headers(17) = "TS5Date"
    headers(18) = "TS5Inv"
    headers(19) = "TS5PO"
    headers(20) = "TS5Amt"
    headers(21) = "TS6Date"
    headers(22) = "TS6Inv"
    headers(23) = "TS6PO"
    headers(24) = "TS6Amt"
    headers(25) = "TS7Date"
    headers(26) = "TS7Inv"
    headers(27) = "TS7PO"
    headers(28) = "TS7Amt"
    headers(29) = "TS8Date"
    headers(30) = "TS8Inv"
    headers(31) = "TS8PO"
    headers(32) = "TS8Amt"
    headers(33) = "TS9Date"
    headers(34) = "TS9Inv"
    headers(35) = "TS9PO"
    headers(36) = "TS9Amt"
    headers(37) = "TS10Date"
    headers(38) = "TS10Inv"
    headers(39) = "TS10PO"
    headers(40) = "TS10Amt"
    headers(41) = "TS11Date"
    headers(42) = "TS11Inv"
    headers(43) = "TS11PO"
    headers(44) = "TS11Amt"
    headers(45) = "TS12Date"
    headers(46) = "TS12Inv"
    headers(47) = "TS12PO"
    headers(48) = "TS12Amt"
    headers(49) = "TS13Date"
    headers(50) = "TS13Inv"
    headers(51) = "TS13PO"
    headers(52) = "TS13Amt"
    headers(53) = "TS14Date"
    headers(54) = "TS14Inv"
    headers(55) = "TS14PO"
    headers(56) = "TS14Amt"
    headers(57) = "TS15Date"
    headers(58) = "TS15Inv"
    headers(59) = "TS15PO"
    headers(60) = "TS15Amt"
    headers(61) = "TS16Date"
    headers(62) = "TS16Inv"
    headers(63) = "TS16PO"
    headers(64) = "TS16Amt"
    headers(65) = "TS17Date"
    headers(66) = "TS17Inv"
    headers(67) = "TS17PO"
    headers(68) = "TS17Amt"
    headers(69) = "TS18Date"
    headers(70) = "TS18Inv"
    headers(71) = "TS18PO"
    headers(72) = "TS18Amt"
    headers(73) = "CkNum"
    headers(74) = "CkDate"
    headers(75) = "CkAmt"
    headers(76) = "CkAmt1"
    headers(77) = "CkAmt2"
    headers(78) = "Pay1"
    headers(79) = "Pay2"
    headers(80) = "Pay3"
    headers(81) = "Pay4"
    headers(82) = "Vend"
    headers(83) = "Memo"
'
'    ' This sets up the fields used in data binding
'    'Fields.Add "ProductID"
    For cnt = 1 To 83
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

'
Private Sub ActiveReport_FetchData(reof As Boolean)

Dim sLine As String
Dim arr() As String
'
'    ' We reached the end of the file we exit leaving the
'    ' eof parameter as True (default except on first call) that will
'    ' tell AR that we are done feeding data
'    ' otherwise we have to set the eof parameter to False so that
'    ' AR continues fetching data, until we're done
'    ' if the report had a data control, the value of the parameter
'    ' will be ignored, AR will always follow the data control's recordset
'    ' EOF property
On Error GoTo ERRORSTUFF
    If VBA.eof(hFile) Then
        reof = True
        Exit Sub
    Else
        reof = False
    End If

    Line Input #hFile, sLine
    arr = Split(sLine, "~")

'    ' Here we set the values of the fields that we defines as unbound
'    ' or user defined.
    For cnt = 1 To 83
       Fields(headers(cnt)) = arr(cnt - 1)
    
    Next
  Exit Sub
'    ("Fund").Value = arr(0)
'    Fields("Dept").Value = arr(1)
'    Fields("DeptName").Value = arr(2)
'    Fields("AcctDesc").Value = arr(3)
'    Fields("Budget").Value = Val(arr(4))
'    Fields("MTD/Enc").Value = arr(5)
'    Fields("YTD").Value = Val(arr(6))
'    Fields("Variance").Value = Val(arr(7))
'    Fields("Pct").Value = arr(8)
ERRORSTUFF:
Stop
      Unload frmLoadingRpt
  Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "ARptVendHist", "Fetch Data", Erl)
    Case emrExitProc:
      Resume Proc_Exit
    Case emrResume:
      Resume
    Case emrResumeNext:
      Resume Next
    Case Else
      '--- Technically, this should never happen.
      Resume Proc_Exit
  End Select

Proc_Exit:
  '--- Cleanup code goes here...
    Close
    Unload Me
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
    End If
    If KeyCode = vbKeyE Then
      Screen.MousePointer = vbHourglass
      ExportReport 1
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - chk5.xls, created in the Citipak Directory.", vbOKOnly
    End If
    If KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - chk5.txt, created in the Citipak Directory.", vbOKOnly
    End If
  End If
End Sub

Private Sub ActiveReport_QueryClose(Cancel As Integer, CloseMode As Integer)
'  KillFile ReportFile$
End Sub

Private Sub ActiveReport_ReportEnd()
Dim STUF As Integer
  If hFile <> 0 Then
    Close #hFile
  End If
  Unload frmLoadingRpt
  DoEvents
  STUF = Me.Pages.Count
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
    MsgBox "File - chk5.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "/&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - chk5.txt, created in the Citipak Directory.", vbOKOnly
  End If
End Sub

Public Sub startrpt()
  Me.Run
End Sub
Private Sub ActiveReport_Initialize()
  Me.Toolbar.Tools.Add "&Close"
  Me.Toolbar.Tools.Add "Save/&Excel"
  Me.Toolbar.Tools.Add "/&Text"
  
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
        oEXL.FileName = outfile & "chk5.xls"
        oEXL.Export Me.Pages
        
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "chk5.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
''
''Me.Pages.Save "check.rdf"
End Sub

