VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arTaxLateLtr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ActiveReport1"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8955
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   15796
   _ExtentY        =   7726
   SectionData     =   "arTaxLateLtr.dsx":0000
End
Attribute VB_Name = "arTaxLateLtr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ReportFile As String
Public SubFile As String
Private hFile As Integer
Dim rpt As ActiveReport
Dim Cnt As Integer
Dim headers(1 To 1) As String

Public Sub GetName(RName As String) ', Optional SName As String, Optional Detail As Integer, Optional RevSource As Integer)
  ReportFile$ = RName$
End Sub

Private Sub ActiveReport_DataInitialize()
  headers(1) = "One"

  hFile = FreeFile
  Open StartPath & "\TAXRPTS\TaxLatLtr.RPT" For Input As #hFile
'    ' This sets up the fields used in data binding
  For Cnt = 1 To 1
    Fields.Add headers(Cnt)
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
  'on error goto ERRORSTUFF
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
  Line Input #hFile, sLine
  If InStr(sLine$, Chr$(12)) Then
    sLine = Space$(79) + Chr$(126)
  End If
'    ' Here we set the values of the fields that we defines as unbound
'    ' or user defined.
  For Cnt = 1 To 1
    Fields(headers(Cnt)) = sLine '(cnt - 1)
  Next
'If something wrong in file give message instead of crashing
Exit Sub

ERRORSTUFF:
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
  Me.ToolBar.Tools.Add "&Close"
  Me.ToolBar.Tools.Add "Save/&Excel"
  Me.ToolBar.Tools.Add "/&Text"
End Sub

Private Sub ActiveReport_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
  'If LetsGo = True Then
    If KeyCode = vbKeyEscape Then
      'ActivateControls fromX
      Unload Me
    KeyCode = 0
  End If
  If Shift = 4 Then
    If KeyCode = vbKeyC Then
      'ActivateControls fromX
      Unload Me
    End If
    If KeyCode = vbKeyE Then
      Screen.MousePointer = vbHourglass
      ExportReport 1
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - TxLateLet.xls, created in the Citipak Directory.", vbOKOnly
    End If
    If KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - TxLateLet.txt, created in the Citipak Directory.", vbOKOnly
    End If
  End If
End Sub

Private Sub ActiveReport_ReportEnd()
    If hFile <> 0 Then
        Close #hFile
    End If
  Unload frmLoadingRpt
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
    MsgBox "File - TxLateLet.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "/&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - TxLateLet.txt, created in the Citipak Directory.", vbOKOnly
  End If
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
        oEXL.FileName = outfile & "TxLateLet.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "TxLateLet.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub


Private Sub Detail_Format()
  If Me.Fields("One").Value = Space$(79) + Chr$(126) Then
    Me.Detail.NewPage = ddNPAfter
  Else
    Me.Detail.NewPage = ddNPNone
  End If
End Sub


