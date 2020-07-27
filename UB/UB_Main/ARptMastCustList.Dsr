VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARptMastCustList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Master Customer Listing"
   ClientHeight    =   7104
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   12216
   Icon            =   "ARptMastCustList.dsx":0000
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   21548
   _ExtentY        =   12531
   SectionData     =   "ARptMastCustList.dsx":08CA
End
Attribute VB_Name = "ARptMastCustList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ReportFile As String
Private hFile As Integer
Dim rpt As ActiveReport
Dim headers(1 To 14) As String
Dim cnt As Integer
Public rptnum As Integer

Public Sub GetName(RName As String)
  ReportFile$ = RName$
End Sub

Private Sub ActiveReport_DataInitialize()
    headers(1) = "Acct"
    headers(2) = "Name"
    headers(3) = "Location"
    headers(4) = "Status"
    headers(5) = "Addr1"
    headers(6) = "HPhone"
    headers(7) = "Type"
    headers(8) = "Addr2"
    headers(9) = "WPhone"
    headers(10) = "Draft"
    headers(11) = "City"
    headers(12) = "State"
    headers(13) = "Zip"
    headers(14) = "ServAddr"
    
    hFile = FreeFile
    Open ReportFile$ For Input As #hFile
'
'    ' This sets up the fields used in data binding
'    'Fields.Add "ProductID"
    For cnt = 1 To 14
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
Private Sub ActiveReport_FetchData(eof As Boolean)

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
    For cnt = 1 To 14
      Fields(headers(cnt)) = arr(cnt - 1)
    Next
'    ("Fund").Value = arr(0)
'    Fields("Dept").Value = arr(1)
'    Fields("DeptName").Value = arr(2)
'    Fields("AcctDesc").Value = arr(3)
'    Fields("Budget").Value = Val(arr(4))
'    Fields("MTD/Enc").Value = arr(5)
'    Fields("YTD").Value = Val(arr(6))
'    Fields("Variance").Value = Val(arr(7))
'    Fields("Pct").Value = arr(8)

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
      MsgBox "File - MastCust.xls, created in the Citipak Directory.", vbOKOnly
    End If
    If KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - MastCust.txt, created in the Citipak Directory.", vbOKOnly
    End If
  End If
End Sub

Private Sub ActiveReport_QueryClose(Cancel As Integer, CloseMode As Integer)
  KillFile ReportFile$
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
    MsgBox "File - MastCust.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "/&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - MastCust.txt, created in the Citipak Directory.", vbOKOnly
  End If
End Sub

Public Sub startrpt()
  Me.Run True
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
  If Right$(UBPath, 1) = ":" Then
    outfile = UBPath
  Else
    outfile = UBPath & "\"
  End If

  Select Case x
    Case 1   '"Excel"
        Set oEXL = New ActiveReportsExcelExport.ARExportExcel
        oEXL.FileName = outfile & "MastCust.xls"
        oEXL.Export Me.Pages
        
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "MastCust.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

