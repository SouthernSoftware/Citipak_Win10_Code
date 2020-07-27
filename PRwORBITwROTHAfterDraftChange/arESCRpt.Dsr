VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arESCRpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ActiveReport1"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11580
   Icon            =   "arESCRpt.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20426
   _ExtentY        =   15637
   SectionData     =   "arESCRpt.dsx":08CA
End
Attribute VB_Name = "arESCRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Private hFile As Integer
Dim DetailCnt As Integer

Private Sub ActiveReport_Initialize()
  Me.ToolBar.Tools.Add "&Close"
  Me.ToolBar.Tools.Add "Save/&Excel"
  Me.ToolBar.Tools.Add "&Text"
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
    ElseIf KeyCode = vbKeyE Then
      Screen.MousePointer = vbHourglass
      ExportReport 1
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - ESCRpt.xls, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - ESCRpt.txt, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    End If
  End If
End Sub

Private Sub ActiveReport_Terminate()
  Close
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
    MsgBox "File - ESCRpt.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - ESCRpt.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "ESCRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "ESCRpt.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub ActiveReport_DataInitialize()
  hFile = FreeFile
  Select Case GlblQtr
    Case 1
      Open StartPath & "\PRRPTS\ESCQTR1.RPT" For Input As #hFile
    Case 2
      Open StartPath & "\PRRPTS\ESCQTR2.RPT" For Input As #hFile
    Case 3
      Open StartPath & "\PRRPTS\ESCQTR3.RPT" For Input As #hFile
    Case 4
      Open StartPath & "\PRRPTS\ESCQTR4.RPT" For Input As #hFile
    Case Else
      MsgBox "ERROR: No path to ESCQTR.RPT file found"
      Exit Sub
  End Select
  Fields.Add "fldEmployer" '0
  Fields.Add "fldStaid" '1
  Fields.Add "fldQtr" '2
  Fields.Add "fldYear" '3
  Fields.Add "fldSSN" '4
  Fields.Add "fldEmpName" '5
  Fields.Add "fldGPay" '6
  Fields.Add "fldTotGPay" '7
  Fields.Add "fldTotXWages" '8
  Fields.Add "fldTotTaxWage" '9
End Sub

Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
    Unload frmLoadingRpt
    MsgBox "Error Number: " & Str(Number) & " " & Description, vbOKOnly, "Printer Error"
    Unload Me
  End If
  CancelDisplay = True 'removes the error message
End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)
  Dim sLine As String
  Dim arr() As String
  
  If VBA.eof(hFile) Then
    eof = True
    Exit Sub
  Else
    eof = False
  End If
  Line Input #hFile, sLine
  arr = Split(sLine, "~")
  ' Here we set the values of the fields that we defines as unbound
  ' or user defined.
  DetailCnt = DetailCnt + 1
  Fields("fldEmployer").Value = arr(0)
  Fields("fldStaid").Value = arr(1)
  Fields("fldQtr").Value = arr(2)
  Fields("fldYear").Value = arr(3)
  Fields("fldSSN").Value = arr(4)
  Fields("fldEmpName").Value = arr(5)
  Fields("fldGPay").Value = arr(6)
  Fields("fldTotGPay").Value = arr(7)
  Fields("fldTotXWages").Value = arr(8)
  Fields("fldTotTaxWage").Value = arr(9)
End Sub

Private Sub ActiveReport_ReportEnd()
  Unload frmLoadingRpt
  If hFile <> 0 Then
    Close #hFile
  End If
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me

End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub ActiveReport_ReportStart()
  Unload frmLoadingRpt
  ReportHeader.Height = 0
  Me.Zoom = -1
End Sub

Private Sub Detail_Format()
  GroupHeader1.GroupValue = Fields("fldEmployer").Value
  If DetailCnt = 25 Then
    Detail.NewPage = ddNPAfter
    DetailCnt = 0
  Else
    Detail.NewPage = ddNPNone
  End If
  
  If DetailCnt = 1 Then
    fldQtr.Visible = True
    fldYear.Visible = True
  Else
    fldQtr.Visible = False
    fldYear.Visible = False
  End If
  
End Sub

Private Sub PageFooter_Format()
  If DetailCnt = 0 Then
    PageFooter.Visible = False
  Else
    PageFooter.Visible = True
  End If
End Sub

Private Sub PageHeader_Format()
  If Fields("fldEmployer").Value = "" Then
    txtPageNumber.Visible = False
    ReportHeader.Height = 0
    PageHeader.Height = 0
    PageFooter.Height = 0
    Detail.Height = 0
  End If

End Sub
