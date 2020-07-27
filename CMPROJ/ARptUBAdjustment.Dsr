VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARptUBAdjustment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Utility Adjustment Report"
   ClientHeight    =   6696
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   9324
   Icon            =   "ARptUBAdjustment.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   16447
   _ExtentY        =   11811
   SectionData     =   "ARptUBAdjustment.dsx":08CA
End
Attribute VB_Name = "ARptUBAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ReportFile As String
Public SubFile As String
Private hFile As Integer
Dim rpt As ActiveReport
Dim cnt As Integer
Dim headers(1 To 55) As String

Public Sub GetName(RName As String)
  ReportFile$ = RName$
End Sub

Private Sub ActiveReport_DataInitialize()
    headers(1) = "CustName"
    headers(2) = "AcctNo"
    headers(3) = "TranDate"
    headers(4) = "TranType"
    headers(5) = "Addr1"
    headers(6) = "Note"
    headers(7) = "TransAmt"
    headers(8) = "Rev1N"
    headers(9) = "Rev1A"
    headers(10) = "Rev1A2"
    headers(11) = "Rev2N"
    headers(12) = "Rev2A"
    headers(13) = "Rev2A2"
    headers(14) = "Rev3N"
    headers(15) = "Rev3A"
    headers(16) = "Rev3A2"
    headers(17) = "Rev4N"
    headers(18) = "Rev4A"
    headers(19) = "Rev4A2"
    headers(20) = "Rev5N"
    headers(21) = "Rev5A"
    headers(22) = "Rev5A2"
    headers(23) = "Rev6N"
    headers(24) = "Rev6A"
    headers(25) = "Rev6A2"
    headers(26) = "Rev7N"
    headers(27) = "Rev7A"
    headers(28) = "Rev7A2"
    headers(29) = "Rev8N"
    headers(30) = "Rev8A"
    headers(31) = "Rev8A2"
    headers(32) = "Rev9N"
    headers(33) = "Rev9A"
    headers(34) = "Rev9A2"
    headers(35) = "Rev10N"
    headers(36) = "Rev10A"
    headers(37) = "Rev10A2"
    headers(38) = "Rev11N"
    headers(39) = "Rev11A"
    headers(40) = "Rev11A2"
    headers(41) = "Rev12N"
    headers(42) = "Rev12A"
    headers(43) = "Rev12A2"
    headers(44) = "Rev13N"
    headers(45) = "Rev13A"
    headers(46) = "Rev13A2"
    headers(47) = "Rev14N"
    headers(48) = "Rev14A"
    headers(49) = "Rev14A2"
    headers(50) = "Rev15N"
    headers(51) = "Rev15A"
    headers(52) = "Rev15A2"
    headers(53) = "AcctBal"
    headers(54) = "CurrBal"
    headers(55) = "PrevBal"
    hFile = FreeFile
    Open ReportFile$ For Input As #hFile
'
'    ' This sets up the fields used in data binding
'    'Fields.Add "ProductID"
    For cnt = 1 To 55
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
    For cnt = 1 To 55
      Fields(headers(cnt)) = arr(cnt - 1)
    Next
'If something wrong in file give message instead of crashing
Exit Sub
ERRORSTUFF:
      Unload frmLoadingRpt
  Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "ARptUBAdjustment", "Fetch Data", Erl)
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

Public Sub startrpt()
  Me.Run True
End Sub

Private Sub ActiveReport_Initialize()
  Me.Toolbar.Tools.Add "&Close"
  Me.Toolbar.Tools.Add "Save/&Excel"
  Me.Toolbar.Tools.Add "/&Text"

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
      MsgBox "File - AdjTran.xls, created in the Citipak Directory.", vbOKOnly
    End If
    If KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - AdjTran.txt, created in the Citipak Directory.", vbOKOnly
    End If
  End If
End Sub


Private Sub ActiveReport_QueryClose(Cancel As Integer, CloseMode As Integer)
'  KillFile ReportFile$
'  KillFile SubFile$
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
    MsgBox "File - AdjTran.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "/&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - AdjTran.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "AdjTran.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "AdjTran.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub



