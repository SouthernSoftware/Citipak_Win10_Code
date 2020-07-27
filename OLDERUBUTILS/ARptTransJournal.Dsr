VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARptTransJournal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction Journal Report"
   ClientHeight    =   4380
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   8952
   Icon            =   "ARptTransJournal.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   15790
   _ExtentY        =   7726
   SectionData     =   "ARptTransJournal.dsx":08CA
End
Attribute VB_Name = "ARptTransJournal"
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
Dim headers(1 To 38) As String
Dim Det As Boolean
Dim Rev As Integer

Public Sub GetName(RName As String, SName As String, Detail As Boolean, RevSource As Integer)
  ReportFile$ = RName$
  SubFile$ = SName$
  Det = Detail
  Rev = RevSource
End Sub

Private Sub ActiveReport_DataInitialize()
    headers(1) = "TranNum"
    headers(2) = "TransDate"
    headers(3) = "Acct"
    headers(4) = "CustName"
    headers(5) = "Desc"
    headers(6) = "Operator"
    headers(7) = "TransDesc"
    headers(8) = "Amt"
    headers(9) = "Rev1N"
    headers(10) = "Rev1A"
    headers(11) = "Rev2N"
    headers(12) = "Rev2A"
    headers(13) = "Rev3N"
    headers(14) = "Rev3A"
    headers(15) = "Rev4N"
    headers(16) = "Rev4A"
    headers(17) = "Rev5N"
    headers(18) = "Rev5A"
    headers(19) = "Rev6N"
    headers(20) = "Rev6A"
    headers(21) = "Rev7N"
    headers(22) = "Rev7A"
    headers(23) = "Rev8N"
    headers(24) = "Rev8A"
    headers(25) = "Rev9N"
    headers(26) = "Rev9A"
    headers(27) = "Rev10N"
    headers(28) = "Rev10A"
    headers(29) = "Rev11N"
    headers(30) = "Rev11A"
    headers(31) = "Rev12N"
    headers(32) = "Rev12A"
    headers(33) = "Rev13N"
    headers(34) = "Rev13A"
    headers(35) = "Rev14N"
    headers(36) = "Rev14A"
    headers(37) = "Rev15N"
    headers(38) = "Rev15A"

    hFile = FreeFile
    Open ReportFile$ For Input As #hFile
'
'    ' This sets up the fields used in data binding
'    'Fields.Add "ProductID"
    For cnt = 1 To 38
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
    For cnt = 1 To 38
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

Private Sub ActiveReport_ReportStart()
  If Det = True Then
   'Label22.Visible = True
   Line22.Visible = True
    If Rev < 6 Then
      Me.Detail.Height = 270
    ElseIf Rev > 5 And Rev < 11 Then
      Me.Detail.Height = 540
    Else 'If Rev > 8 And Rev < 13 Then
      Me.Detail.Height = 810
'    Else
'      Me.Detail.Height = 810
    End If
  Else
    Me.Detail.Height = 0
    Me.Line1.Visible = False
  End If
  Me.Label19 = Me.txtRptParm1
  Me.Label21 = Me.txtRptParm2
End Sub

'Private Sub PageHeader_BeforePrint()
'  If txtPageNumber = txtPagecount Then
'    If Fields("TranNum").Value = 0 Then
'      Label1.Visible = False
'      Label2.Visible = False
'      Label3.Visible = False
'      Label4.Visible = False
'      Label5.Visible = False
'      Label6.Visible = False
'      Label7.Visible = False
'    End If
'  End If
'End Sub

Private Sub PageHeader_Format()
If Me.pageNumber = 1 Then
  Label12.Visible = True
  Shape1.Visible = True
  txtRptParm1.Visible = True
  txtRptParm2.Visible = True
  Me.PageHeader.Height = 1188
Else
  Label12.Visible = False
  Shape1.Visible = False
  txtRptParm1.Visible = False
  txtRptParm2.Visible = False
  Me.PageHeader.Height = 624
'  Label1.Top = 720
'  Label2.Top = 720
'  Label3.Top = 720
'  Label4.Top = 720
'  Label5.Top = 720
'  Label6.Top = 720
'  Label7.Top = 720
'  Line1.Y1 = 990
'  Line1.Y2 = 990
End If

End Sub


Private Sub ReportFooter_Format()
 ' If Rev > 0 Then
    Set Me.SubReport1.object = New ARSubTot
'  End If
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
      MsgBox "File - TranJrnl.xls, created in the Citipak Directory.", vbOKOnly
    End If
    If KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - TranJrnl.txt, created in the Citipak Directory.", vbOKOnly
    End If
  End If
End Sub


Private Sub ActiveReport_QueryClose(Cancel As Integer, CloseMode As Integer)
  KillFile ReportFile$
  KillFile SubFile$
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
    MsgBox "File - TranJrnl.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "/&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - TranJrnl.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "TranJrnl.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "TranJrnl.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

