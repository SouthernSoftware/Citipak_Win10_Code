VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARptCashReceiptsJournal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cash Receipts Journal"
   ClientHeight    =   4380
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   8952
   Icon            =   "ARptCashReceiptsJournal.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   15790
   _ExtentY        =   7726
   SectionData     =   "ARptCashReceiptsJournal.dsx":08CA
End
Attribute VB_Name = "ARptCashReceiptsJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ReportFile As String
Public SubFile As String
Public SubFile2 As String
Private hFile As Integer
Dim rpt As ActiveReport
Dim cnt As Integer, Chktot As Integer
Dim headers(1 To 9) As String
Public Sub GetName(RName As String, SName As String, SsName As String)
  ReportFile$ = RName$
  SubFile$ = SName$
  SubFile2$ = SsName$
  Chktot = 0
End Sub

Private Sub ActiveReport_DataInitialize()
   
    hFile = FreeFile
    Open ReportFile$ For Input As #hFile
    headers(1) = "Date"
    headers(2) = "AcctNo"
    headers(3) = "Customer"
    headers(4) = "Type"
    headers(5) = "CashAmt"
    headers(6) = "ChkAmt"
    headers(7) = "Charge"
    headers(8) = "AmtApplied"
    headers(9) = "Change"
'    ' This sets up the fields used in data binding
'    'Fields.Add "ProductID"
    For cnt = 1 To 9
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
On Error GoTo Errorstuff
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
    For cnt = 1 To 9
      Fields(headers(cnt)) = arr(cnt - 1)
    Next
'If something wrong in file give message instead of crashing
Exit Sub
Errorstuff:
      Unload frmLoadingRpt
  Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "ARptCashReceiptsJournal", "Fetch Data", Erl)
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


Private Sub Detail_Format()
  If Me.Fields("ChkAmt").Value > 0 Then
    Chktot = Chktot + 1
  End If
End Sub

Private Sub GroupFooter1_Format()
  Me.totchks.DataValue = Chktot
End Sub

Private Sub ReportFooter_Format()
' Set Me.SubReport1.object = New ARSubRevSum
'  Set Me.SubReport2.object = New ARSubBookSum
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
      KeyCode = 0
    End If
    If KeyCode = vbKeyE Then
      Screen.MousePointer = vbHourglass
      ExportReport 1
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - PayJourn.xls, created in the Citipak Directory.", vbOKOnly
    End If
    If KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - PayJourn.txt, created in the Citipak Directory.", vbOKOnly
    End If
  End If

End Sub
Private Sub ActiveReport_QueryClose(Cancel As Integer, CloseMode As Integer)
  KillFile ReportFile$
  KillFile SubFile$
  KillFile SubFile2$
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
    MsgBox "File - PayJourn.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "/&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - PayJourn.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "PayJourn.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "PayJourn.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub



