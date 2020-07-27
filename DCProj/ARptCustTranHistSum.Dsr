VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARptCustTranHistSum 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Transaction Summary"
   ClientHeight    =   4380
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   8952
   Icon            =   "ARptCustTranHistSum.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   15790
   _ExtentY        =   7726
   SectionData     =   "ARptCustTranHistSum.dsx":08CA
End
Attribute VB_Name = "ARptCustTranHistSum"
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
Dim headers(1 To 52) As String
Dim Det As Boolean
Dim Rev As Integer, Mtr As Integer

Public Sub GetName(RName As String)
  ReportFile$ = RName$
'  SubFile$ = SName$
End Sub

Private Sub ActiveReport_DataInitialize()
    headers(1) = "CustName"
    headers(2) = "AcctNo"
    headers(3) = "TranDate"
    headers(4) = "TranType"
    headers(5) = "ReadDate"
    headers(6) = "PrevDate"
    headers(7) = "TRAmt"
    headers(8) = "RunBal"
    headers(9) = "Mtr1"
    headers(10) = "Curr1"
    headers(11) = "Prev1"
    headers(12) = "Use1"
    headers(13) = "Mtr2"
    headers(14) = "Curr2"
    headers(15) = "Prev2"
    headers(16) = "Use2"
    headers(17) = "Mtr3"
    headers(18) = "Curr3"
    headers(19) = "Prev3"
    headers(20) = "Use3"
    headers(21) = "Mtr4"
    headers(22) = "Curr4"
    headers(23) = "Prev4"
    headers(24) = "Use4"
    headers(25) = "Mtr5"
    headers(26) = "Curr5"
    headers(27) = "Prev5"
    headers(28) = "Use5"
    headers(29) = "Mtr6"
    headers(30) = "Curr6"
    headers(31) = "Prev6"
    headers(32) = "Use6"
    headers(33) = "Mtr7"
    headers(34) = "Curr7"
    headers(35) = "Prev7"
    headers(36) = "Use7"
    headers(37) = "Rev1"
    headers(38) = "Rev2"
    headers(39) = "Rev3"
    headers(40) = "Rev4"
    headers(41) = "Rev5"
    headers(42) = "Rev6"
    headers(43) = "Rev7"
    headers(44) = "Rev8"
    headers(45) = "Rev9"
    headers(46) = "Rev10"
    headers(47) = "Rev11"
    headers(48) = "Rev12"
    headers(49) = "Rev13"
    headers(50) = "Rev14"
    headers(51) = "Rev15"
    headers(52) = "TNum"
    hFile = FreeFile
    Open ReportFile$ For Input As #hFile
'
'    ' This sets up the fields used in data binding
'    'Fields.Add "ProductID"
    For cnt = 1 To 52
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
    For cnt = 1 To 52
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
      MsgBox "File - CustTran.xls, created in the Citipak Directory.", vbOKOnly
    End If
    If KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - CustTran.txt, created in the Citipak Directory.", vbOKOnly
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
    MsgBox "File - CustTran.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "/&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - CustTran.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "CustTran.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "CustTran.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub



