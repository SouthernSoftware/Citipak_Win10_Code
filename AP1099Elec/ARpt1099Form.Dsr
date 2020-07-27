VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARpt1099Form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "1099 Form"
   ClientHeight    =   9045
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   13005
   Icon            =   "ARpt1099Form.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   22939
   _ExtentY        =   15954
   SectionData     =   "ARpt1099Form.dsx":08CA
End
Attribute VB_Name = "ARpt1099Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ReportFile As String
Private hFile As Integer
Dim rpt As ActiveReport
Dim headers(1 To 62) As String
Dim cnt As Integer

Public Sub GetName(RName As String)
  ReportFile$ = RName$
End Sub

Private Sub ActiveReport_DataInitialize()
   hFile = FreeFile
   Open ReportFile$ For Input As #hFile
    headers(1) = "PayName"
    headers(2) = "PayAddr1"
    headers(3) = "PayAddr2"
    headers(4) = "PayCSZ"
    headers(5) = "PayFedID"
    headers(6) = "RecID"
    headers(7) = "RecName"
    headers(8) = "DBA"
    headers(9) = "Addr1"
    headers(10) = "Addr2"
    headers(11) = "CSZ"
    headers(12) = "Acct"
    headers(13) = "Notice"
    headers(14) = "Box1"
    headers(15) = "Box2"
    headers(16) = "Box3"
    headers(17) = "Box4"
    headers(18) = "Box5"
    headers(19) = "Box6"
    headers(20) = "Box7"
    headers(21) = "Box8"
    headers(22) = "Box9"
    headers(23) = "Box10"
    headers(24) = "Box13"
    headers(25) = "Box14"
    headers(26) = "Box15"
    headers(27) = "Box16"
    headers(28) = "Box17"
    headers(29) = "Box18"
    headers(30) = "Vopt1"
    headers(31) = "Copt1"
    headers(32) = "PayName2"
    headers(33) = "PayAddr12"
    headers(34) = "PayAddr22"
    headers(35) = "PayCSZ2"
    headers(36) = "PayFedID2"
    headers(37) = "RecID2"
    headers(38) = "RecName2"
    headers(39) = "DBA2"
    headers(40) = "Addr12"
    headers(41) = "Addr22"
    headers(42) = "CSZ2"
    headers(43) = "Acct2"
    headers(44) = "Notice2"
    headers(45) = "Box12"
    headers(46) = "Box22"
    headers(47) = "Box32"
    headers(48) = "Box42"
    headers(49) = "Box52"
    headers(50) = "Box62"
    headers(51) = "Box72"
    headers(52) = "Box82"
    headers(53) = "Box92"
    headers(54) = "Box102"
    headers(55) = "Box132"
    headers(56) = "Box142"
    headers(57) = "Box152"
    headers(58) = "Box162"
    headers(59) = "Box172"
    headers(60) = "Box182"
    headers(61) = "Vopt2"
    headers(62) = "Copt2"
'    ' This sets up the fields used in data binding
'    'Fields.Add "ProductID"
    For cnt = 1 To 62
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
    For cnt = 1 To 62
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
      MsgBox "File - Ten99.xls, created in the Citipak Directory.", vbOKOnly
    End If
    If KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - Ten99.txt, created in the Citipak Directory.", vbOKOnly
    End If
  End If
End Sub

Private Sub ActiveReport_QueryClose(Cancel As Integer, CloseMode As Integer)
'  KillFile ReportFile$
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
    MsgBox "File - Ten99.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "/&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - Ten99.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "Ten99.xls"
        oEXL.Export Me.Pages
        
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "Ten99.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
''
''Me.Pages.Save "check.rdf"
End Sub
